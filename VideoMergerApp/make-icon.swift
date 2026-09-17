// Renders the Video Merger icon (two film clips converging into one, on an
// indigo→violet squircle) into an .iconset directory.
// Usage: swift make-icon.swift <output.iconset dir>
import AppKit
import CoreGraphics
import Foundation

func roundedRect(_ r: CGRect, _ radius: CGFloat) -> CGPath {
    CGPath(roundedRect: r, cornerWidth: radius, cornerHeight: radius, transform: nil)
}

/// A film clip: a rounded body with sprocket holes down its left edge.
func drawClip(_ ctx: CGContext, _ rect: CGRect, radius: CGFloat,
              body: CGColor, holes: CGColor) {
    ctx.saveGState()
    ctx.setShadow(offset: CGSize(width: 0, height: -rect.height * 0.05),
                  blur: rect.height * 0.14,
                  color: CGColor(red: 0, green: 0, blue: 0, alpha: 0.32))
    ctx.addPath(roundedRect(rect, radius))
    ctx.setFillColor(body)
    ctx.fillPath()
    ctx.restoreGState()

    let holeW = rect.width * 0.1, holeH = rect.height * 0.15
    let x = rect.minX + rect.width * 0.075
    ctx.setFillColor(holes)
    for i in 0..<3 {
        let y = rect.minY + rect.height * (0.18 + 0.29 * CGFloat(i))
        ctx.addPath(roundedRect(CGRect(x: x, y: y, width: holeW, height: holeH), holeW * 0.32))
    }
    ctx.fillPath()
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

    // ── Gradient body (indigo → violet) ─────────────────────────────────────
    ctx.saveGState()
    ctx.addPath(squircle)
    ctx.clip()
    let cs = CGColorSpaceCreateDeviceRGB()
    let grad = CGGradient(colorsSpace: cs, colors: [
        CGColor(red: 0.42, green: 0.40, blue: 0.94, alpha: 1),   // top – indigo
        CGColor(red: 0.34, green: 0.26, blue: 0.80, alpha: 1),   // mid
        CGColor(red: 0.24, green: 0.14, blue: 0.56, alpha: 1),   // btm – deep violet
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

    // ── Two source clips on the left, one merged clip on the right ──────────
    let white = CGColor(red: 1, green: 1, blue: 1, alpha: 0.96)
    let faded = CGColor(red: 1, green: 1, blue: 1, alpha: 0.74)
    let holeColour = CGColor(red: 0.28, green: 0.20, blue: 0.66, alpha: 1)

    let smallW = S * 0.215, smallH = S * 0.205
    drawClip(ctx, CGRect(x: S * 0.175, y: S * 0.545, width: smallW, height: smallH),
             radius: S * 0.035, body: faded, holes: holeColour)
    drawClip(ctx, CGRect(x: S * 0.175, y: S * 0.255, width: smallW, height: smallH),
             radius: S * 0.035, body: faded, holes: holeColour)

    let bigH = S * 0.32
    drawClip(ctx, CGRect(x: S * 0.585, y: S * 0.34, width: S * 0.24, height: bigH),
             radius: S * 0.042, body: white, holes: holeColour)

    // ── The arrow that says "into one" ──────────────────────────────────────
    ctx.saveGState()
    ctx.setStrokeColor(white)
    ctx.setLineWidth(S * 0.032)
    ctx.setLineCap(.round)
    ctx.setLineJoin(.round)
    let y = S * 0.5, x0 = S * 0.425, x1 = S * 0.545
    ctx.move(to: CGPoint(x: x0, y: y))
    ctx.addLine(to: CGPoint(x: x1, y: y))
    ctx.strokePath()
    ctx.move(to: CGPoint(x: x1 - S * 0.048, y: y + S * 0.048))
    ctx.addLine(to: CGPoint(x: x1, y: y))
    ctx.addLine(to: CGPoint(x: x1 - S * 0.048, y: y - S * 0.048))
    ctx.strokePath()
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
