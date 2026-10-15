// Fonts for the printed sheets: Arial from the system (as the approved v3 layout), with
// Liberation Sans (metric-compatible with Arial, SIL OFL, bundled in assets/fonts) as the fallback.
const fs = require('fs');
const path = require('path');

const ARIAL_DIR = '/System/Library/Fonts/Supplemental';
const BUNDLED_DIR = path.join(__dirname, '..', '..', 'assets', 'fonts');

function resolveFonts() {
  const arial = { regular: path.join(ARIAL_DIR, 'Arial.ttf'), bold: path.join(ARIAL_DIR, 'Arial Bold.ttf') };
  if (fs.existsSync(arial.regular) && fs.existsSync(arial.bold)) {
    return { family: 'Arial', ...arial };
  }
  return {
    family: 'Liberation Sans',
    regular: path.join(BUNDLED_DIR, 'LiberationSans-Regular.ttf'),
    bold: path.join(BUNDLED_DIR, 'LiberationSans-Bold.ttf')
  };
}

module.exports = { resolveFonts, ARIAL_DIR, BUNDLED_DIR };
