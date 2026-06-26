#!/usr/bin/env node
const { recognize } = require("tesseract.js");

async function main() {
  const imagePath = process.argv[2];
  if (!imagePath) {
    throw new Error("Missing image path");
  }

  const result = await recognize(imagePath, "chi_sim+eng");
  process.stdout.write(JSON.stringify({ text: result.data?.text || "" }));
}

main().catch((error) => {
  process.stderr.write(error.stack || error.message || String(error));
  process.exit(1);
});
