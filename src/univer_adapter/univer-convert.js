const fs = require('fs');
const path = require('path');

const [,, inputPath, outputPath] = process.argv;

if (!inputPath || !outputPath) {
  console.error("Usage: node univer-convert.js <input.docx> <output.pdf>");
  process.exit(1);
}

// Univer headless facade conversion mockable stub & exchange hook
async function convert() {
  if (!fs.existsSync(inputPath)) {
    console.error(`Input file not found: ${inputPath}`);
    process.exit(2);
  }

  try {
    // If running in environment with @univerjs exchange client, invoke snapshot export
    // Otherwise throw to trigger graceful fallback to LibreOffice
    console.log(`[Univer] Converting ${inputPath} to ${outputPath}...`);
    // Placeholder trigger for headless exchange API:
    // When Univer Exchange is present, write PDF buffer
    // For now, if exchange module not configured, exit with 10 to fallback
    process.exit(10);
  } catch (err) {
    console.error("[Univer Error]", err);
    process.exit(1);
  }
}

convert();
