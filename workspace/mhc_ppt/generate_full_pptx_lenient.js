const html2pptx = require('../../.claude/skills/pptx/scripts/html2pptx.js');
const pptxgen = require('../../.claude/skills/pptx/scripts/node_modules/pptxgenjs');
const path = require('path');
const fs = require('fs');

async function main() {
  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_16x9';
  pptx.title = 'DeepSeek mHC: 流形约束超连接';
  pptx.author = 'Generated with Claude Code';

  const slidesDir = path.join(__dirname, 'slides_full');
  const files = fs.readdirSync(slidesDir)
    .filter(f => f.endsWith('.html'))
    .sort();

  console.log(`Processing ${files.length} slides...`);

  let successCount = 0;
  let errorCount = 0;

  for (const file of files) {
    const filePath = path.join(slidesDir, file);
    console.log(`  Adding: ${file}`);
    try {
      await html2pptx(filePath, pptx);
      successCount++;
    } catch (err) {
      errorCount++;
      console.error(`  ⚠️ Skipped ${file}: validation error`);
      // Skip problematic slides but continue
    }
  }

  const outputPath = path.join(__dirname, 'DeepSeek_mHC_Full.pptx');
  await pptx.writeFile({ fileName: outputPath });
  console.log(`\n✅ Created: ${outputPath}`);
  console.log(`   Success: ${successCount} slides`);
  console.log(`   Skipped: ${errorCount} slides`);
}

main().catch(err => {
  console.error('Fatal error:', err);
  process.exit(1);
});
