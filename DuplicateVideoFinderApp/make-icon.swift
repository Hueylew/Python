// Renders the Duplicate Video Finder icon (two stacked film frames under a
// magnifying glass, on a teal→green squircle) into an .iconset directory.
// Usage: swift make-icon.swift <output.iconset dir>
import AppKit
import CoreGraphics
import Foundation

func roundedRect(_ r: CGRect, _ radius: CGFloat) -> CGPath {
    CGPath(roundedRect: r, cornerWidth: radius, cornerHeight: radius, transform: nil)
}

func drawIcon(into ctx: CGContext, side S: CGFloat) {
    ctx.setAllowsAntialiasing(true)
    ctx.interpolationQuality = .high
    ctx.clear(CGRect(x: 0, y: 0, width: S, height: S))

    let inset = S * 0.086
    let rect = CGRect(x: inset, y: inset, width: S - 2 * inset, height: S - 2 * inset)
    let corner = rect.width * 0.2237
    let squircle = roundedRect(rect, corner)

    // ── Drop shadow under the tile ──────────────────────────────────────────
    ctx.saveGState()
    ctx.setShadow(offset: CGSize(width: 0, height: -S * 0.012),
                  blur: S * 0.05,
                  color: CGColor(red: 0, green: 0, blue: 0, alpha: 0.28))
    ctx.addPath(squircle)
    ctx.setFillColor(CGColor(red: 0, green: 0, blue: 0, alpha: 1))
    ctx.fillPath()
    ctx.restoreGState()

    // ── Gradient body (teal → green) ────────────────────────────────────────
    ctx.saveGState()
    ctx.addPath(squircle)
    ctx.clip()
    let cs = CGColorSpaceCreateDeviceRGB()
    let grad = CGGradient(colorsSpace: cs, colors: [
        CGColor(red: 0.16, green: 0.78, blue: 0.72, alpha: 1),   // top – teal
        CGColor(red: 0.10, green: 0.63, blue: 0.55, alpha: 1),   // mid
        CGColor(red: 0.06, green: 0.45, blue: 0.42, alpha: 1),   // btm – deep green
    ] as CFArray, locations: [0.0, 0.58, 1.0])!
    ctx.drawLinearGradient(grad,
                           start: CGPoint(x: rect.minX, y: rect.maxY),
                           end: CGPoint(x: rect.maxX, y: rect.minY),
                           options: [])
    let sheen = CGGradient(colorsSpace: cs, colors: [
        CGColor(red: 1, green: 1, blue: 1, alpha: 0.22),
        CGColor(red: 1, green: 1, blue: 1, alpha: 0.0),
    ] as CFArray, locations: [0, 1])!
    ctx.drawLinearGradient(sheen,
                           start: CGPoint(x: rect.midX, y: rect.maxY),
                           end: CGPoint(x: rect.midX, y: rect.maxY - rect.height * 0.5),
                           options: [])
    ctx.restoreGState()

    // ── Two offset "copies" (the duplicate pair) ────────────────────────────
    let cardW = S * 0.34, cardH = S * 0.25
    let cardR = S * 0.042
    let backRect = CGRect(x: S * 0.255, y: S * 0.455, width: cardW, height: cardH)
    let frontRect = CGRect(x: S * 0.325, y: S * 0.355, width: cardW, height: cardH)

    // back copy — translucent, sits behind
    ctx.saveGState()
    ctx.setShadow(offset: CGSize(width: 0, height: -S * 0.008), blur: S * 0.025,
                  color: CGColor(red: 0, green: 0.10, blue: 0.10, alpha: 0.35))
    ctx.addPath(roundedRect(backRect, cardR))
    ctx.setFillColor(CGColor(red: 1, green: 1, blue: 1, alpha: 0.55))
    ctx.fillPath()
    ctx.restoreGState()

    // front copy — solid white with a play triangle
    ctx.saveGState()
    ctx.setShadow(offset: CGSize(width: 0, height: -S * 0.008), blur: S * 0.03,
                  color: CGColor(red: 0, green: 0.10, blue: 0.10, alpha: 0.40))
    ctx.addPath(roundedRect(frontRect, cardR))
    ctx.setFillColor(CGColor(red: 1, green: 1, blue: 1, alpha: 1))
    ctx.fillPath()
    ctx.restoreGState()

    let teal = CGColor(red: 0.07, green: 0.52, blue: 0.47, alpha: 1)
    let cx = frontRect.midX, cy = frontRect.midY
    let tri = CGMutablePath()
    let th = S * 0.055
    tri.move(to: CGPoint(x: cx - th * 0.55, y: cy - th))
    tri.addLine(to: CGPoint(x: cx - th * 0.55, y: cy + th))
    tri.addLine(to: CGPoint(x: cx + th * 0.85, y: cy))
    tri.closeSubpath()
    ctx.addPath(tri)
    ctx.setFillColor(teal)
    ctx.fillPath()

    // ── Magnifying glass over the lower-left ────────────────────────────────
    let gx = S * 0.375, gy = S * 0.345
    let gr = S * 0.135
    let ring = S * 0.038

    ctx.saveGState()
    ctx.setShadow(offset: CGSize(width: 0, height: -S * 0.010), blur: S * 0.03,
                  color: CGColor(red: 0, green: 0.10, blue: 0.10, alpha: 0.45))
    // handle first, so the ring overlaps it cleanly
    let handle = CGMutablePath()
    let hStart = CGPoint(x: gx - gr * 0.70, y: gy - gr * 0.70)
    let hEnd = CGPoint(x: gx - gr * 1.55, y: gy - gr * 1.55)
    handle.move(to: hStart)
    handle.addLine(to: hEnd)
    ctx.addPath(handle.copy(strokingWithWidth: ring * 1.15, lineCap: .round,
                            lineJoin: .round, miterLimit: 10))
    ctx.setFillColor(CGColor(red: 1, green: 1, blue: 1, alpha: 1))
    ctx.fillPath()

    let lens = CGMutablePath()
    lens.addEllipse(in: CGRect(x: gx - gr, y: gy - gr, width: gr * 2, height: gr * 2))
    ctx.addPath(lens.copy(strokingWithWidth: ring, lineCap: .round,
                          lineJoin: .round, miterLimit: 10))
    ctx.setFillColor(CGColor(red: 1, green: 1, blue: 1, alpha: 1))
    ctx.fillPath()
    ctx.restoreGState()

    // glass tint inside the lens
    ctx.saveGState()
    ctx.addEllipse(in: CGRect(x: gx - gr + ring / 2, y: gy - gr + ring / 2,
                              width: (gr - ring / 2) * 2, height: (gr - ring / 2) * 2))
    ctx.setFillColor(CGColor(red: 0.60, green: 0.95, blue: 0.92, alpha: 0.42))
    ctx.fillPath()
    ctx.restoreGState()
}

func renderPNG(size: Int, to url: URL) {
    let cs = CGColorSpaceCreateDeviceRGB()
    guard let ctx = CGContext(data: nil, width: size, height: size,
                              bitsPerComponent: 8, bytesPerRow: 0, space: cs,
                              bitmapInfo: CGImageAlphaInfo.premultipliedLast.rawValue) else { return }
    drawIcon(into: ctx, side: CGFloat(size))
    guard let image = ctx.makeImage(),
          let dest = CGImageDestinationCreateWithURL(url as CFURL, "public.png" as CFString, 1, nil)
    else { return }
    CGImageDestinationAddImage(dest, image, nil)
    CGImageDestinationFinalize(dest)
}

let args = CommandLine.arguments
guard args.count >= 2 else {
    FileHandle.standardError.write(Data("usage: make-icon.swift <dir.iconset>\n".utf8))
    exit(1)
}
let outDir = URL(fileURLWithPath: args[1])
try? FileManager.default.createDirectory(at: outDir, withIntermediateDirectories: true)

let specs: [(String, Int)] = [
    ("icon_16x16.png", 16), ("icon_16x16@2x.png", 32),
    ("icon_32x32.png", 32), ("icon_32x32@2x.png", 64),
    ("icon_128x128.png", 128), ("icon_128x128@2x.png", 256),
    ("icon_256x256.png", 256), ("icon_256x256@2x.png", 512),
    ("icon_512x512.png", 512), ("icon_512x512@2x.png", 1024),
]
for (name, px) in specs {
    renderPNG(size: px, to: outDir.appendingPathComponent(name))
    print("  \(name)  (\(px)px)")
}
