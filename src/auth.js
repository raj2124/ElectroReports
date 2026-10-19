// Sign-in, sessions and the audit trail (Session 6E).
//
// "Sign in with Zoho" is OpenID Connect against accounts.zoho.in with a Server-based client of its
// own (not the Projects Self Client). The e-mail comes from the id_token (AaaServer.profile.READ is
// the fallback); only @elegrow.com accounts that are ACTIVE users of our Zoho Projects portal get in.
// The user's own Zoho token is used once to read the profile and then dropped.
//
// Sessions are rows in the database; the cookie holds a random id (httpOnly, SameSite=Lax, Secure
// over HTTPS) and only an HMAC of it is stored. They end after `idleHours` without use, and an
// active portal user is re-checked at most every `recheckMinutes`.
// AUTH_MODE=dev is a user picker for development and tests: never the default, loopback only.

const crypto = require('crypto');
const axios = require('axios');
const { hasRole, mayCall } = require('./permissions');

const COOKIE = 'er_session';
const STATE_COOKIE = 'er_oauth_state';
const OPEN_API = [/^\/api\/health$/, /^\/api\/auth\/(me|dev-users)$/];

function parseCookies(header) {
  const cookies = {};
  String(header || '').split(';').forEach((part) => {
    const index = part.indexOf('=');
    if (index > 0) {
      cookies[part.slice(0, index).trim()] = decodeURIComponent(part.slice(index + 1).trim());
    }
  });
  return cookies;
}

function serializeCookie(name, value, { maxAgeSeconds, secure } = {}) {
  return [`${name}=${encodeURIComponent(value)}`, 'Path=/', 'HttpOnly', 'SameSite=Lax', ...(secure ? ['Secure'] : []), ...(maxAgeSeconds !== undefined ? [`Max-Age=${maxAgeSeconds}`] : [])].join('; ');
}

const lower = (value) => String(value || '').trim().toLowerCase();

// The claims of an id_token. It comes straight from Zoho's token endpoint over TLS, which is the
// case OpenID Connect allows without checking the signature; audience and expiry are checked.
function readIdToken(idToken, clientId) {
  const parts = String(idToken || '').split('.');
  if (parts.length < 2) {
    return null;
  }
  let claims;
  try {
    claims = JSON.parse(Buffer.from(parts[1].replace(/-/g, '+').replace(/_/g, '/'), 'base64').toString('utf8'));
  } catch (_error) {
    return null;
  }
  const audience = Array.isArray(claims.aud) ? claims.aud : [claims.aud];
  if (!audience.includes(clientId) || (claims.exp && claims.exp * 1000 < Date.now())) {
    return null;
  }
  return claims;
}

function parseDevUsers(entries) {
  return entries.map((entry) => {
    const [email, roles = 'engineer', name = '', designation = ''] = entry.split(':');
    return { email: lower(email), roles: roles.split('+').map((role) => lower(role)).filter(Boolean), name: name.trim() || lower(email).split('@')[0], designation: designation.trim() };
  });
}

function createAuth({ store, config, zoho, http = axios, log = console }) {
  const auth = config.auth;
  const secure = auth.zoho.redirectUri.startsWith('https:');
  const idleMs = auth.idleHours * 3600 * 1000;
  const recheckMs = auth.recheckMinutes * 60 * 1000;
  const devUsers = parseDevUsers(auth.devUsers);
  const recheckFailures = new Map();

  const hashOf = (token) => crypto.createHmac('sha256', auth.sessionSecret || 'dev-only').update(token).digest('hex');
  const sign = (value) => crypto.createHmac('sha256', auth.sessionSecret || 'dev-only').update(value).digest('hex');

  function audit(user, action, entityType, entityId, detail) {
    try {
      store.addAudit({ user, action, entityType, entityId, detail });
    } catch (error) {
      log.error(`audit write failed: ${error.message}`);
    }
  }

  // Seed administrators exist from the start, so the admin screen is reachable on a fresh database.
  auth.seedAdmins.forEach((email) => {
    if (!store.getUserByEmail(email)) {
      store.addUser({ email, name: email.split('@')[0], roles: ['admin', 'engineer'] });
    }
  });

  const modeReady = () => (auth.mode === 'dev' ? true : auth.mode === 'zoho' && Boolean(auth.zoho.clientId && auth.zoho.clientSecret && auth.sessionSecret));

  function startSession(res, user) {
    const token = crypto.randomBytes(32).toString('hex');
    store.createSession(hashOf(token), user.id);
    store.updateUser(user.id, { lastLoginAt: new Date().toISOString() });
    res.append('Set-Cookie', serializeCookie(COOKIE, token, { secure }));
  }

  // The checks every sign-in passes. Returns { user } or { error: code }.
  async function admit(profile) {
    const email = lower(profile.email);
    if (!email || !email.endsWith(`@${auth.allowedDomain}`)) {
      return { error: 'domain', email };
    }
    let portal = null;
    if (auth.mode === 'zoho') {
      let member;
      try {
        member = await zoho.lookupPortalUser(email);
      } catch (_error) {
        return { error: 'zoho_unreachable', email };
      }
      if (!member.found) return { error: 'not_in_portal', email };
      if (!member.active) return { error: 'inactive', email };
      profile = { ...profile, name: profile.name || member.name };
      portal = { found: member.found, active: member.active };
    }
    let user = store.getUserByEmail(email);
    if (!user) {
      user = store.addUser({ email, name: profile.name || email.split('@')[0], roles: profile.roles || ['engineer'], designation: profile.designation || '', zuid: profile.zuid || '' });
    } else {
      // a seeded user is created with the e-mail name; the first real sign-in gives the real name
      const patch = {};
      if (profile.zuid && !user.zuid) patch.zuid = profile.zuid;
      if (profile.name && user.name === email.split('@')[0]) patch.name = profile.name;
      if (Object.keys(patch).length) user = store.updateUser(user.id, patch);
    }
    if (!user.active) {
      return { error: 'deactivated', email };
    }
    return { user, portal };
  }

  // ---- the session on each request -------------------------------------------------------------
  async function attachUser(req, _res, next) {
    req.user = null;
    const token = parseCookies(req.headers.cookie)[COOKIE];
    if (!token) {
      return next();
    }
    const idHash = hashOf(token);
    const session = store.getSession(idHash);
    if (!session) {
      return next();
    }
    const now = Date.now();
    if (now - Date.parse(session.lastSeenAt) > idleMs) {
      store.deleteSession(idHash);
      return next();
    }
    const user = store.getUser(session.userId);
    if (!user || !user.active) {
      store.deleteSession(idHash);
      return next();
    }
    let checked = false;
    if (auth.mode === 'zoho' && now - Date.parse(session.checkedAt) > recheckMs && !(recheckFailures.get(user.id) > now)) {
      try {
        const member = await zoho.lookupPortalUser(user.email, { maxAgeMs: Math.min(recheckMs, 5 * 60 * 1000) });
        if (!member.found || !member.active) {
          store.deleteUserSessions(user.id);
          audit(user, 'auth.revoked', 'user', user.id, { reason: 'no longer an active Zoho Projects portal user' });
          return next();
        }
        checked = true;
      } catch (_error) {
        recheckFailures.set(user.id, now + 5 * 60 * 1000); // Zoho down: keep the session, try again later
      }
    }
    if (checked || now - Date.parse(session.lastSeenAt) > Math.min(30000, idleMs / 4)) {
      store.touchSession(idHash, { checked });
    }
    req.user = user;
    req.sessionHash = idHash;
    return next();
  }

  // Sign-in required for the API and the generated PDFs; a few open routes.
  function requireSignIn(req, res, next) {
    if (OPEN_API.some((pattern) => pattern.test(req.baseUrl + req.path))) {
      return next();
    }
    if (!req.user) {
      return res.status(401).json({ message: 'Sign in to continue.', code: 'SIGN_IN_REQUIRED' });
    }
    return next();
  }

  // A cross-site page must not be able to act with the cookie: a state-changing request that carries
  // an Origin must come from this host.
  function sameOrigin(req, res, next) {
    const origin = req.headers.origin;
    if (!['GET', 'HEAD', 'OPTIONS'].includes(req.method) && origin) {
      try {
        if (new URL(origin).host !== req.headers.host) {
          return res.status(403).json({ message: 'Cross-site request refused.', code: 'CROSS_SITE' });
        }
      } catch (_error) {
        return res.status(403).json({ message: 'Cross-site request refused.', code: 'CROSS_SITE' });
      }
    }
    return next();
  }

  // Route rules (src/permissions.js) for every /api request, then an audit line for each change.
  function authorize(req, res, next) {
    const pathname = req.baseUrl + req.path;
    if (OPEN_API.some((pattern) => pattern.test(pathname)) || !req.user) {
      return next();
    }
    if (!mayCall(req.user, req.method, pathname)) {
      audit(req.user, 'auth.forbidden', 'route', `${req.method} ${pathname}`, null);
      return res.status(403).json({ message: 'Your role does not allow this.', code: 'FORBIDDEN' });
    }
    if (!['GET', 'HEAD'].includes(req.method)) {
      res.on('finish', () => {
        if (res.statusCode < 400) {
          const { action, entityType, entityId } = describe(req.method, pathname);
          audit(req.user, action, entityType, entityId, { method: req.method, status: res.statusCode });
        }
      });
    }
    return next();
  }

  function describe(method, pathname) {
    const parts = pathname.split('/').filter(Boolean); // api, kind, id, ...
    const [, kind, id, a, b] = parts;
    const named = (action, entityType = kind, entityId = id) => ({ action, entityType: entityType === 'reports' ? 'report' : entityType === 'projects' ? 'project' : entityType, entityId: entityId || '' });
    if (kind === 'reports') {
      if (!id) return named('report.create');
      if (a === 'review') return b === 'ticks' ? named('review.ticks') : named(`review.${b}`);
      if (a === 'reissue') return named('review.reissue');
      if (a === 'pdf') return named('report.export');
      if (a === 'ai') return named('report.ai_narrative');
      if (a === 'client-logo') return named('logo.report');
      return named(method === 'DELETE' ? 'report.delete' : 'report.update');
    }
    if (kind === 'projects') {
      if (!id) return named('project.create');
      if (a === 'days') return named('project.day_save');
      if (a === 'final') return named('project.assemble');
      if (a === 'zoho-refresh') return named('project.zoho_refresh');
      return named('project.update');
    }
    if (kind === 'notices') return named(`notice.${a}`, 'notice');
    if (kind === 'client-logo') return named(`logo.${method === 'DELETE' ? 'delete' : 'set'}`, 'logo', '');
    if (kind === 'sheets') return named('sheets.print', 'sheets', '');
    if (kind === 'admin') return named('admin.user_update', 'user', a || '');
    if (kind === 'profile') return named('profile.update', 'user', '');
    if (kind === 'zoho') return named('zoho.poll', 'zoho', '');
    return { action: `${method} ${pathname}`, entityType: kind || '', entityId: id || '' };
  }

  // ---- routes ----------------------------------------------------------------------------------
  const ERRORS = {
    domain: `Only @${auth.allowedDomain} accounts can sign in.`,
    not_in_portal: 'This account is not a user of our Zoho Projects portal.',
    inactive: 'This account is not active in our Zoho Projects portal.',
    deactivated: 'This account has been deactivated in ElectroReports. Ask an admin.',
    zoho_unreachable: 'Zoho could not be reached to check your account. Try again in a minute.',
    config: 'Sign-in is not configured yet.',
    failed: 'Sign-in failed. Try again.',
    state: 'The sign-in link expired. Try again.'
  };

  function register(app) {
    app.get('/api/auth/me', (req, res) => {
      res.json({ mode: auth.mode || 'none', ready: modeReady(), allowedDomain: auth.allowedDomain, user: req.user ? publicUser(req.user) : null, errors: ERRORS });
    });

    app.get('/auth/login', (req, res) => {
      if (auth.mode !== 'zoho' || !modeReady()) {
        return res.redirect('/?auth_error=config');
      }
      const state = crypto.randomBytes(16).toString('hex');
      res.append('Set-Cookie', serializeCookie(STATE_COOKIE, `${state}.${sign(state)}`, { maxAgeSeconds: 600, secure }));
      const query = new URLSearchParams({ scope: auth.zoho.scope, client_id: auth.zoho.clientId, response_type: 'code', access_type: 'online', redirect_uri: auth.zoho.redirectUri, state });
      return res.redirect(`${auth.zoho.accountsUrl}/oauth/v2/auth?${query}`);
    });

    app.get('/auth/callback', async (req, res) => {
      const fail = (code, email, detail) => {
        audit(null, 'auth.rejected', 'user', email || '', { reason: code, ...(detail || {}) });
        res.append('Set-Cookie', serializeCookie(STATE_COOKIE, '', { maxAgeSeconds: 0, secure }));
        return res.redirect(`/?auth_error=${code}`);
      };
      if (auth.mode !== 'zoho' || !modeReady()) return fail('config');
      const [state, signature] = String(parseCookies(req.headers.cookie)[STATE_COOKIE] || '').split('.');
      if (!state || signature !== sign(state) || req.query.state !== state) return fail('state');
      if (req.query.error || !req.query.code) return fail('failed', '', { zoho: String(req.query.error || 'no code').slice(0, 60) });
      try {
        const token = await http.post(`${auth.zoho.accountsUrl}/oauth/v2/token`, new URLSearchParams({ grant_type: 'authorization_code', client_id: auth.zoho.clientId, client_secret: auth.zoho.clientSecret, redirect_uri: auth.zoho.redirectUri, code: String(req.query.code) }).toString(), { headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, timeout: 15000 });
        const claims = readIdToken(token.data && token.data.id_token, auth.zoho.clientId);
        let emailFrom = 'id_token';
        let profile = claims && claims.email && claims.email_verified !== false ? { email: claims.email, name: claims.name || [claims.given_name, claims.family_name].filter(Boolean).join(' '), zuid: String(claims.sub || '') } : null;
        if (!profile && token.data && token.data.access_token) {
          const info = await http.get(`${auth.zoho.accountsUrl}/oauth/user/info`, { headers: { Authorization: `Zoho-oauthtoken ${token.data.access_token}` }, timeout: 15000 });
          if (info.data && info.data.Email) {
            emailFrom = 'profile_endpoint';
            profile = { email: info.data.Email, name: info.data.Display_Name || [info.data.First_Name, info.data.Last_Name].filter(Boolean).join(' '), zuid: String(info.data.ZUID || '') };
          }
        }
        if (!profile) return fail('failed', '', { reason: 'no e-mail in the id_token or the profile (check AUTH_ZOHO_SCOPE)' });
        const admitted = await admit(profile);
        if (admitted.error) return fail(admitted.error, admitted.email);
        res.append('Set-Cookie', serializeCookie(STATE_COOKIE, '', { maxAgeSeconds: 0, secure }));
        startSession(res, admitted.user);
        audit(admitted.user, 'auth.signin', 'user', admitted.user.id, { via: 'zoho', emailFrom, portalMember: admitted.portal, roles: admitted.user.roles });
        return res.redirect('/');
      } catch (error) {
        log.error(`Sign-in failed: ${String(error.message || 'error').slice(0, 120)}`);
        return fail('failed');
      }
    });

    app.post('/auth/logout', (req, res) => {
      if (req.sessionHash) {
        store.deleteSession(req.sessionHash);
        audit(req.user, 'auth.signout', 'user', req.user.id, null);
      }
      res.append('Set-Cookie', serializeCookie(COOKIE, '', { maxAgeSeconds: 0, secure }));
      res.json({ signedOut: true });
    });

    // Development and tests only.
    const devAllowed = (req) => auth.mode === 'dev' && ['127.0.0.1', '::1', '::ffff:127.0.0.1'].includes(req.socket.remoteAddress) && process.env.NODE_ENV !== 'production';
    app.get('/api/auth/dev-users', (req, res) => {
      if (!devAllowed(req)) return res.status(404).json({ message: 'Not found.' });
      return res.json(devUsers.map((user) => ({ email: user.email, name: user.name, roles: user.roles })));
    });
    app.post('/auth/dev-login', async (req, res) => {
      if (!devAllowed(req)) return res.status(404).json({ message: 'Not found.' });
      const spec = devUsers.find((user) => user.email === lower(req.body && req.body.email));
      if (!spec) return res.status(403).json({ message: 'Unknown development user.', code: 'UNKNOWN_USER' });
      const admitted = await admit({ email: spec.email, name: spec.name, designation: spec.designation, roles: spec.roles });
      if (admitted.error) return res.status(403).json({ message: ERRORS[admitted.error], code: admitted.error });
      startSession(res, admitted.user);
      audit(admitted.user, 'auth.signin', 'user', admitted.user.id, { via: 'dev' });
      return res.json({ user: publicUser(admitted.user) });
    });

  }

  // Mounted AFTER the sign-in and role guard (app.use('/api', requireSignIn, authorize)).
  function registerProtected(app) {
    // My profile: name and designation (printed as Prepared / Checked / Approved by).
    app.put('/api/profile', (req, res) => {
      const name = String((req.body && req.body.name) || '').trim();
      const designation = String((req.body && req.body.designation) || '').trim();
      if (!name) return res.status(400).json({ message: 'Enter your name.' });
      return res.json({ user: publicUser(store.updateUser(req.user.id, { name, designation })) });
    });

    // Admin: users and roles, and the audit log (read only).
    app.get('/api/admin/users', (_req, res) => res.json(store.listUsers().map(publicUser)));
    app.put('/api/admin/users/:id', (req, res) => {
      const target = store.getUser(req.params.id);
      if (!target) return res.status(404).json({ message: 'User not found.' });
      const body = req.body || {};
      const patch = {};
      if (Array.isArray(body.roles)) {
        const roles = [...new Set(body.roles.map(lower))].filter((role) => ['engineer', 'reviewer', 'approver', 'admin'].includes(role));
        patch.roles = roles;
      }
      if (typeof body.active === 'boolean') patch.active = body.active;
      const nextRoles = patch.roles || target.roles;
      const nextActive = patch.active === undefined ? target.active : patch.active;
      const otherAdmins = store.listUsers().filter((user) => user.id !== target.id && user.active && user.roles.includes('admin'));
      if (!otherAdmins.length && !(nextActive && nextRoles.includes('admin'))) {
        return res.status(409).json({ message: 'There must be at least one active admin.', code: 'LAST_ADMIN' });
      }
      const updated = store.updateUser(target.id, patch);
      if (patch.active === false || (patch.roles && !patch.roles.length)) store.deleteUserSessions(target.id);
      audit(req.user, 'admin.user_change', 'user', target.id, { roles: patch.roles, active: patch.active, before: { roles: target.roles, active: target.active } });
      return res.json(publicUser(updated));
    });
    app.get('/api/admin/audit', (req, res) => res.json(store.listAudit({ userId: String(req.query.user || ''), action: String(req.query.action || ''), limit: Number.parseInt(req.query.limit || '100', 10) || 100, offset: Number.parseInt(req.query.offset || '0', 10) || 0 })));
  }

  const publicUser = (user) => ({ id: user.id, email: user.email, name: user.name, designation: user.designation, roles: user.roles, active: user.active, lastLoginAt: user.lastLoginAt });

  return { attachUser, requireSignIn, sameOrigin, authorize, register, registerProtected, audit, hashOf, publicUser, modeReady, hasRole };
}

module.exports = { createAuth, readIdToken, parseCookies, serializeCookie, parseDevUsers, COOKIE, STATE_COOKIE };
