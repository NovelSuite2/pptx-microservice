const express = require("express");
const PptxGenJS = require("pptxgenjs");
const fs = require("fs");
const path = require("path");
const os = require("os");
const crypto = require("crypto");
const https = require("https");
const http = require("http");

const app = express();
app.use(express.json({ limit: "50mb" }));

app.get("/health", (req, res) => res.json({ status: "ok" }));

function downloadImage(url) {
  return new Promise((resolve, reject) => {
    const client = url.startsWith("https") ? https : http;
    client.get(url, (res) => {
      const chunks = [];
      res.on("data", (chunk) => chunks.push(chunk));
      res.on("end", () => resolve(Buffer.concat(chunks).toString("base64")));
      res.on("error", reject);
    }).on("error", reject);
  });
}

app.post("/generate-slide", async (req, res) => {
  try {
    const { topBarColor, heading, bodyText, imageUrl, imageBase64 } = req.body;

    if (!heading || !bodyText) {
      return res.status(400).json({ error: "Missing required fields: heading, bodyText" });
    }

    let imgBase64 = imageBase64;
    if (!imgBase64 && imageUrl) {
      imgBase64 = await downloadImage(imageUrl);
    }

    const pres = new PptxGenJS();
    pres.layout = "LAYOUT_16x9";

    const slide = pres.addSlide();
    slide.background = { color: "F0F0F0" };

    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 10, h: 0.8,
      fill: { color: topBarColor || "4A72B0" },
      line: { color: topBarColor || "4A72B0", width: 0 },
    });

    slide.addShape(pres.shapes.ISOSCELES_TRIANGLE, {
      x: 0.44, y: 0.1, w: 0.42, h: 0.22,
      fill: { color: "FFFFFF" },
      line: { color: "FFFFFF", width: 0 },
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: 0.3, w: 0.3, h: 0.26,
      fill: { color: "FFFFFF" },
      line: { color: "FFFFFF", width: 0 },
    });

    if (imgBase64) {
      slide.addImage({
        data: "image/png;base64," + imgBase64,
        x: 0.5, y: 1.35, w: 3.3, h: 3.3,
      });
    }

    slide.addText(heading, {
      x: 4.3, y: 1.35, w: 5.2, h: 0.85,
      fontSize: 28,
      fontFace: "Calibri",
      bold: true,
      color: "111111",
      align: "left",
      valign: "top",
      wrap: true,
      margin: 0,
    });

    slide.addText([{ text: bodyText }], {
      x: 4.3, y: 2.2, w: 5.2, h: 2.0,
      fontSize: 18,
      fontFace: "Calibri",
      color: "111111",
      align: "left",
      valign: "top",
      wrap: true,
      lineSpacingMultiple: 1.2,
      margin: 0,
    });

    const btnX = 8.3, btnY = 4.675, btnW = 1.2, btnH = 0.45;
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: btnX, y: btnY, w: btnW, h: btnH,
      fill: { color: "E0E0E0" },
      line: { color: "BBBBBB", width: 1 },
      rectRadius: 0.05,
    });
    slide.addText("Next >", {
      x: btnX, y: btnY, w: btnW, h: btnH,
      fontSize: 14,
      fontFace: "Calibri",
      color: "111111",
      align: "center",
      valign: "middle",
      margin: 0,
    });

    const tmpFile = path.join(os.tmpdir(), `slide-${crypto.randomUUID()}.pptx`);
    await pres.writeFile({ fileName: tmpFile });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    res.setHeader("Content-Disposition", 'attachment; filename="Slide_1.pptx"');

    const stream = fs.createReadStream(tmpFile);
    stream.pipe(res);
    stream.on("end", () => fs.unlink(tmpFile, () => {}));
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`PPTX microservice running on port ${PORT}`));
