const puppeteer = require('puppeteer-core');
const pptxgen = require('pptxgenjs');
const fs = require('fs');
const path = require('path');

const chromePath = 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe';

async function exportAll() {
    console.log('Launching headless Chrome...');
    const browser = await puppeteer.launch({
        executablePath: chromePath,
        headless: true,
        args: [
            '--no-sandbox',
            '--disable-setuid-sandbox',
            '--disable-web-security',
            '--allow-file-access-from-files',
            '--window-size=1920,1080',
            '--force-device-scale-factor=2'
        ]
    });

    const page = await browser.newPage();
    await page.setViewport({ width: 1920, height: 1080, deviceScaleFactor: 2 });

    const htmlPath = 'd:/Slide TEMSA/temsa_presentation.html';
    const fileUrl = `file:///${htmlPath.replace(/\\/g, '/')}`;
    console.log('Navigating to:', fileUrl);

    await page.goto(fileUrl, { waitUntil: 'networkidle0', timeout: 60000 });

    // Wait for fonts to load
    await page.evaluateHandle('document.fonts.ready');
    await new Promise(r => setTimeout(r, 2000));

    // Inject styles to hide UI controls & make slides full-bleed 16:9
    await page.addStyleTag({
        content: `
            header, nav, #gridOverview, #commentDrawer, #notificationToast, .fixed.bottom-4, .slide-edit-controls-host, #floatingEditToolbar, #appToast {
                display: none !important;
            }
            body {
                background: #f8fafc !important;
                overflow: hidden !important;
            }
            .slide-item {
                position: fixed !important;
                inset: 0 !important;
                width: 100vw !important;
                height: 100vh !important;
                padding: 2.5rem 3.5rem !important;
            }
        `
    });

    const totalSlides = 32;
    console.log(`Starting capture of ${totalSlides} slides...`);

    const pptx = new pptxgen();
    pptx.layout = 'LAYOUT_16x9';
    pptx.title = 'TEMSA Level 1 Accreditation Presentation - Cho-airong Hospital';
    pptx.author = 'Cho-airong Hospital EMS Team';
    pptx.company = 'Cho-airong Hospital';

    const tempDir = 'd:/Slide TEMSA/temp_slide_exports';
    if (!fs.existsSync(tempDir)) fs.mkdirSync(tempDir, { recursive: true });

    const slideImages = [];

    for (let i = 1; i <= totalSlides; i++) {
        console.log(`Rendering Slide ${i}/${totalSlides}...`);
        
        await page.evaluate((slideNum) => {
            if (typeof window.goToSlide === 'function') {
                window.goToSlide(slideNum);
            } else {
                document.querySelectorAll('.slide-item').forEach(el => el.classList.add('hidden'));
                const target = document.querySelector(`.slide-item[data-slide="${slideNum}"]`);
                if (target) target.classList.remove('hidden');
            }
        }, i);

        await new Promise(r => setTimeout(r, 600));

        const imgPath = path.join(tempDir, `slide_${i.toString().padStart(2, '0')}.png`);
        await page.screenshot({ path: imgPath, type: 'png' });
        slideImages.push(imgPath);

        // Add slide to PPTX
        const slide = pptx.addSlide();
        slide.addImage({ path: imgPath, x: 0, y: 0, w: '100%', h: '100%' });
    }

    const pptxPath = 'd:/Slide TEMSA/temsa_presentation.pptx';
    console.log('Writing PowerPoint file to:', pptxPath);
    await pptx.writeFile({ fileName: pptxPath });
    console.log('PPTX created successfully!');

    console.log('Generating PDF from high-res slides...');
    const pdfPage = await browser.newPage();
    
    let pdfHtml = `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
        @page { size: 1920px 1080px; margin: 0; }
        body { margin: 0; padding: 0; background: #fff; }
        .page { width: 1920px; height: 1080px; page-break-after: always; display: flex; align-items: center; justify-content: center; overflow: hidden; }
        img { width: 1920px; height: 1080px; display: block; object-fit: contain; }
    </style></head><body>`;

    slideImages.forEach(img => {
        const base64 = fs.readFileSync(img).toString('base64');
        pdfHtml += `<div class="page"><img src="data:image/png;base64,${base64}" /></div>`;
    });

    pdfHtml += `</body></html>`;

    const pdfPath = 'd:/Slide TEMSA/temsa_presentation.pdf';
    await pdfPage.setContent(pdfHtml, { waitUntil: 'load' });
    await pdfPage.pdf({
        path: pdfPath,
        width: '1920px',
        height: '1080px',
        printBackground: true,
        margin: { top: 0, bottom: 0, left: 0, right: 0 }
    });

    console.log('PDF created successfully at:', pdfPath);

    // Cleanup temp images
    try {
        slideImages.forEach(img => fs.unlinkSync(img));
        fs.rmdirSync(tempDir);
    } catch (e) {}

    await browser.close();
    console.log('=== All Exports (PDF & PPTX) Generated Successfully! ===');
}

exportAll().catch(err => {
    console.error('Export Error:', err);
    process.exit(1);
});
