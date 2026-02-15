const path = require('path');
const fs = require('fs');
const html2pptx = require('/Users/bingran_you/Documents/GitHub_MacBook/skills/.agents/skills/pptx/scripts/html2pptx.js');
const pptxgen = require('/Users/bingran_you/Documents/GitHub_MacBook/skills/.agents/skills/pptx/scripts/node_modules/pptxgenjs');

async function main() {
  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_16x9';
  pptx.title = 'DeepSeek mHC 全量整合版';
  pptx.author = 'Codex';

  const slidesDir = path.join(__dirname, 'slides_full_integrated');
  const files = fs.readdirSync(slidesDir)
    .filter(f => f.endsWith('.html'))
    .sort();

  for (const file of files) {
    const filePath = path.join(slidesDir, file);
    await html2pptx(filePath, pptx);
  }

  const outputPath = path.join(__dirname, 'DeepSeek_mHC_Full_Integrated.pptx');
  await pptx.writeFile({ fileName: outputPath });
  console.log(`Created: ${outputPath}`);
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
