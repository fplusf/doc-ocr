// Apple Vision OCR helper (macOS only).
// Usage: vision_ocr <image> [comma-separated languages, e.g. en-US,ru-RU] [fast]
// Prints JSON: {"width":W,"height":H,"lines":[{"text":..,"conf":..,"x":..,"y":..,"w":..,"h":..}]}
// Coordinates are normalised 0..1 with origin at the TOP-left.
import Foundation
import Vision
import AppKit

let args = CommandLine.arguments
guard args.count >= 2 else {
    FileHandle.standardError.write("usage: vision_ocr <image> [langs] [fast]\n".data(using: .utf8)!)
    exit(2)
}
if args[1] == "--list" {
    let r = VNRecognizeTextRequest()
    r.recognitionLevel = .accurate
    let langs = (try? r.supportedRecognitionLanguages()) ?? []
    print(langs.joined(separator: ","))
    exit(0)
}
let url = URL(fileURLWithPath: args[1])
guard let img = NSImage(contentsOf: url),
      let cg = img.cgImage(forProposedRect: nil, context: nil, hints: nil) else {
    FileHandle.standardError.write("cannot open image \(args[1])\n".data(using: .utf8)!)
    exit(3)
}

let req = VNRecognizeTextRequest()
req.recognitionLevel = (args.count >= 4 && args[3] == "fast") ? .fast : .accurate
req.usesLanguageCorrection = true
if args.count >= 3 && !args[2].isEmpty && args[2] != "auto" {
    req.recognitionLanguages = args[2].split(separator: ",").map { String($0) }
} else if #available(macOS 13.0, *) {
    req.automaticallyDetectsLanguage = true
}

let handler = VNImageRequestHandler(cgImage: cg, options: [:])
do { try handler.perform([req]) } catch {
    FileHandle.standardError.write("vision failed: \(error)\n".data(using: .utf8)!)
    exit(4)
}

var lines: [[String: Any]] = []
for obs in req.results ?? [] {
    guard let top = obs.topCandidates(1).first else { continue }
    let b = obs.boundingBox
    lines.append([
        "text": top.string,
        "conf": Double(top.confidence),
        "x": Double(b.origin.x),
        "y": Double(1.0 - b.origin.y - b.size.height),
        "w": Double(b.size.width),
        "h": Double(b.size.height),
    ])
}
let out: [String: Any] = ["width": cg.width, "height": cg.height, "lines": lines]
let data = try! JSONSerialization.data(withJSONObject: out, options: [])
FileHandle.standardOutput.write(data)
