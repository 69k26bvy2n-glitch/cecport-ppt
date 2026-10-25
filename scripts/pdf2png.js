// pdf2png.js — macOS 自带 PDFKit 把 PDF 逐页渲染为 PNG（零依赖，JXA）
// 用法: osascript -l JavaScript pdf2png.js 输入.pdf 输出目录 [宽像素]
ObjC.import('AppKit');
ObjC.import('PDFKit');

function run(argv) {
    const pdfPath = argv[0];
    const outDir = argv[1];
    const targetW = parseInt(argv[2] || '1600', 10);

    const doc = $.PDFDocument.alloc.initWithURL($.NSURL.fileURLWithPath(pdfPath));
    if (doc.isNil()) throw new Error('打不开 PDF: ' + pdfPath);
    const box = $.kPDFDisplayBoxMediaBox;
    const n = doc.pageCount;
    let done = 0;
    for (let i = 0; i < n; i++) {
        const page = doc.pageAtIndex(i);
        const b = page.boundsForBox(box);
        const scale = targetW / b.size.width;
        const w = Math.round(b.size.width * scale);
        const h = Math.round(b.size.height * scale);

        const rep = $.NSBitmapImageRep.alloc.initWithBitmapDataPlanesPixelsWidePixelsHighBitsPerSampleSamplesPerPixelHasAlphaIsPlanarColorSpaceNameBytesPerRowBitsPerPixel(
            null, w, h, 8, 4, true, false, $.NSDeviceRGBColorSpace, 0, 32);
        const nsctx = $.NSGraphicsContext.graphicsContextWithBitmapImageRep(rep);
        $.NSGraphicsContext.setCurrentContext(nsctx);
        const ctx = nsctx.CGContext;
        // 不透明白底
        $.CGContextSaveGState(ctx);
        $.CGContextSetRGBFillColor(ctx, 1, 1, 1, 1);
        $.CGContextFillRect(ctx, $.CGRectMake(0, 0, w, h));
        $.CGContextScaleCTM(ctx, scale, scale);
        page.drawWithBoxToContext(box, ctx);
        $.CGContextRestoreGState(ctx);

        const png = rep.representationUsingTypeProperties($.NSPNGFileType, $.NSDictionary.dictionary);
        const name = 'slides.' + String(i + 1).padStart(3, '0') + '.png';
        png.writeToFileAtomically(outDir + '/' + name, true);
        done++;
    }
    return done;
}
