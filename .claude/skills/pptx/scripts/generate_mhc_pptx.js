const pptxgen = require('pptxgenjs');
const html2pptx = require('./html2pptx');
const path = require('path');

async function createPresentation() {
    const pptx = new pptxgen();
    pptx.layout = 'LAYOUT_16x9';
    pptx.author = 'DeepSeek';
    pptx.title = 'DeepSeek mHC: 流形约束超连接';
    pptx.subject = 'Manifold-Constrained Hyper-Connections 技术介绍';

    const slidesDir = '/Users/bingran_you/Documents/GitHub_MacBook/skills/workspace/mhc_ppt/slides';

    const slides = [
        'slide01_cover.html',
        'slide02_overview.html',
        'slide03_background.html',
        'slide04_hc_problems.html',
        'slide05_mhc_core.html',
        'slide06_technical.html',
        'slide07_architecture.html',
        'slide08_optimization.html',
        'slide09_results.html',
        'slide10_replication.html',
        'slide11_explosion.html',
        'slide12_summary.html'
    ];

    for (const slideFile of slides) {
        const slidePath = path.join(slidesDir, slideFile);
        console.log(`Processing: ${slideFile}`);
        await html2pptx(slidePath, pptx);
    }

    const outputPath = '/Users/bingran_you/Documents/GitHub_MacBook/skills/workspace/mhc_ppt/DeepSeek_mHC_Presentation.pptx';
    await pptx.writeFile({ fileName: outputPath });
    console.log(`Presentation saved to: ${outputPath}`);
}

createPresentation().catch(console.error);
