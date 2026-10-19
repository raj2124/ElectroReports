// Client logos for the report cover: stored once per client (keyed by the Zoho client/account,
// with the client name as fallback) and reused by every report for that client. A report may
// override it. PNG or JPEG, at most 1 MB. They live in the database only, never in the repository.

const MAX_BYTES = 1024 * 1024;

class LogoError extends Error {
  constructor(message) {
    super(message);
    this.name = 'LogoError';
  }
}

// Validates a data URL and returns { mime, buffer }. The bytes, not the declared type, decide.
function parseLogoDataUrl(dataUrl) {
  const match = /^data:(image\/(?:png|jpeg|jpg));base64,([A-Za-z0-9+/=\s]+)$/i.exec(String(dataUrl || ''));
  if (!match) {
    throw new LogoError('The logo must be a PNG or JPEG image.');
  }
  const buffer = Buffer.from(match[2].replace(/\s/g, ''), 'base64');
  if (buffer.length > MAX_BYTES) {
    throw new LogoError(`The logo is ${(buffer.length / 1024 / 1024).toFixed(1)} MB; the limit is 1 MB.`);
  }
  const isPng = buffer.length > 8 && buffer.subarray(0, 8).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]));
  const isJpeg = buffer.length > 3 && buffer[0] === 0xff && buffer[1] === 0xd8 && buffer[2] === 0xff;
  if (!isPng && !isJpeg) {
    throw new LogoError('The file is not a valid PNG or JPEG image.');
  }
  return { mime: isPng ? 'image/png' : 'image/jpeg', buffer };
}

const normalizeName = (name) => String(name || '').trim().toLowerCase().replace(/\s+/g, ' ');

// Keys to try, most specific first.
function clientKeys(project = {}) {
  const keys = [];
  if (String(project.clientId || '').trim()) {
    keys.push(`zoho:${String(project.clientId).trim()}`);
  }
  if (normalizeName(project.clientName)) {
    keys.push(`name:${normalizeName(project.clientName)}`);
  }
  return keys;
}

// The logo a report's cover uses: its own override, else the client's stored logo.
function resolveClientLogo(store, report) {
  if (report.clientLogoOverride) {
    try {
      return parseLogoDataUrl(report.clientLogoOverride).buffer;
    } catch (_error) {
      return null;
    }
  }
  for (const key of clientKeys(report.project)) {
    const row = store.getClientLogo(key);
    if (row) {
      return row.data;
    }
  }
  return null;
}

module.exports = { LogoError, parseLogoDataUrl, clientKeys, resolveClientLogo, MAX_BYTES };
