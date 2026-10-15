// Sheet IDs: ER-XXXX-XXXXX (8 random characters + 1 check character), from an alphabet without
// look-alikes (no 0/O, 1/I). The QR code on every printed page carries `ER1|<id>`.
const crypto = require('crypto');

const ALPHABET = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
const QR_PREFIX = 'ER1|';

function checkChar(body) {
  const sum = [...body].reduce((total, ch, index) => total + ALPHABET.indexOf(ch) * (index + 1), 0);
  return ALPHABET[sum % ALPHABET.length];
}

function newSheetId() {
  const bytes = crypto.randomBytes(8);
  const body = [...bytes].map((byte) => ALPHABET[byte % ALPHABET.length]).join('');
  return `ER-${body.slice(0, 4)}-${body.slice(4)}${checkChar(body)}`;
}

// Normalises what a QR reader (or a person) gives back: any case, stray spaces. Returns the
// canonical ID or null if it is not a well-formed ID with a valid check character.
function parseSheetId(text) {
  const raw = String(text || '').trim().toUpperCase();
  const stripped = raw.startsWith(QR_PREFIX.toUpperCase()) ? raw.slice(QR_PREFIX.length) : raw;
  const match = /^ER-([A-Z0-9]{4})-([A-Z0-9]{4})([A-Z0-9])$/.exec(stripped);
  if (!match) {
    return null;
  }
  const body = match[1] + match[2];
  if ([...body].some((ch) => !ALPHABET.includes(ch)) || checkChar(body) !== match[3]) {
    return null;
  }
  return `ER-${match[1]}-${match[2]}${match[3]}`;
}

module.exports = { newSheetId, parseSheetId, QR_PREFIX };
