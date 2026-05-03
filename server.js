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

    client
      .get(url, (res) => {
        const chunks = [];
        res.on("data", (chunk) => chunks.push(chunk));
        res.on("end", () =>
          resolve(Buffer.concat(chunks).toString("base64"))
        );
        res.on("error", reject);
      })
      .on("error", reject);
  });
}

function num(value, fallback) {
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}

function bool(value, fallback) {
  if (value === true || value === "true") return true;
  if (value === false || value === "false") return false;
  return fallback;
}

app.post("/generate-slide", async (req, res) => {
  try {
    const {
      heading,
      bodyText,
      imageUrl,
      imageBase64,

      backgroundColor,
      topBarColor,
      topBarHeight,

      contentTop,

      leftColumnX,
      leftColumnWidth,
      rightColumnX,
      rightColumnWidth,

      headingFontSize,
      headingFontFace,
      headingColor,
      headingBold,
      headingAlign,

      bodyFontSize,
      bodyFontFace,
      bodyColor,
      bodyAlign,

      spacingBelowHeading,
      paragraphSpacingPt,

      imageX,
      imageY,
      imageWidth,
      imageHeight
    } = req.body;

    if (!heading || !bodyText) {
      return res.status(400).json({
        error: "Missing required fields: heading, bodyText"
      });
    }

    const safeBackgroundColor = backgroundColor || "F0F0F0";
    const safeTopBarColor = topBarColor || "4A72B0";
    const safeTopBarHeight = num(topBarHeight, 0.8);

    const safeContentTop = num(contentTop, 1.35);

    const safeLeftColumnX = num(leftColumnX, 0.5);
    const safeLeftColumnWidth = num(leftColumnWidth, 3.3);
    const safeRightColumnX = num(rightColumnX, 4.3);
    const safeRightColumnWidth = num(rightColumnWidth, 5.2);

    const safeHeadingFontSize = num(headingFontSize, 28);
    const safeHeadingFontFace = headingFontFace || "Calibri";
    const safeHeadingColor = headingColor || "111111";
    const safeHeadingBold = bool(headingBold, true);
    const safeHeadingAlign = headingAlign || "left";

    const safeBodyFontSize = num(bodyFontSize, 18);
    const safeBodyFontFace = bodyFontFace || "Calibri";
    const safeBodyColor = bodyColor || "111111";
    const safeBodyAlign = bodyAlign || "left";

    const safeSpacingBelowHeading = num(spacingBelowHeading, 0.3);
    const safeParagraphSpacingPt = num(paragraphSpacingPt, 3);

    const safeImageX = num(imageX, safeLeftColumnX);
    const safeImageY = num(imageY, safeContentTop);
    const safeImageWidth = num(imageWidth, safeLeftColumnWidth);
    const safeImageHeight = num(imageHeight, 3.3);

    let imgBase64 = imageBase64;

    if (!imgBase64 && imageUrl) {
      imgBase64 = await downloadImage(imageUrl);
    }

    const pres = new PptxGenJS();
    pres.layout = "LAYOUT_16x9";

    const slide = pres.addSlide();

    slide.background = {
      color: safeBackgroundColor
    };

    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0,
      y: 0,
      w: 10,
      h: safeTopBarHeight,
      fill: { color: safeTopBarColor },
      line: { color: safeTopBarColor, width: 0 }
    });

    slide.addShape(pres.shapes.ISOSCELES_TRIANGLE, {
      x: 0.44,
      y: 0.1,
      w: 0.42,
      h: 0.22,
      fill: { color: "FFFFFF" },
      line: { color: "FFFFFF", width: 0 }
    });

    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5,
      y: 0.3,
      w: 0.3,
      h: 0.26,
      fill: { color: "FFFFFF" },
      line: { color: "FFFFFF", width: 0 }
    });

    if (imgBase64) {
      slide.addImage({
        data: "image/png;base64," + imgBase64,
        x: safeImageX,
        y: safeImageY,
        w: safeImageWidth,
        h: safeImageHeight
      });
    }

    slide.addText(heading, {
      x: safeRightColumnX,
      y: safeContentTop,
      w: safeRightColumnWidth,
      h: 0.85,
      fontSize: safeHeadingFontSize,
      fontFace: safeHeadingFontFace,
      bold: safeHeadingBold,
      color: safeHeadingColor,
      align: safeHeadingAlign,
      valign: "top",
      wrap: true,
      margin: 0
    });

    const bodyY = safeContentTop + 0.85 + safeSpacingBelowHeading;

    slide.addText([{ text: bodyText }], {
      x: safeRightColumnX,
      y: bodyY,
      w: safeRightColumnWidth,
      h: 2.0,
      fontSize: safeBodyFontSize,
      fontFace: safeBodyFontFace,
      color: safeBodyColor,
      align: safeBodyAlign,
      valign: "top",
      wrap: true,
      lineSpacingMultiple: 1.2,
      breakLine: false,
      paraSpaceAfterPt: safeParagraphSpacingPt,
      margin: 0
    });

    const btnX = 8.3;
    const btnY = 4.675;
    const btnW = 1.2;
    const btnH = 0.45;

    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: btnX,
      y: btnY,
      w: btnW,
      h: btnH,
      fill: { color: "E0E0E0" },
      line: { color: "BBBBBB", width: 1 },
      rectRadius: 0.05
    });

    slide.addText("Next >", {
      x: btnX,
      y: btnY,
      w: btnW,
      h: btnH,
      fontSize: 14,
      fontFace: "Calibri",
      color: "111111",
      align: "center",
      valign: "middle",
      margin: 0
    });

    const tmpFile = path.join(
      os.tmpdir(),
      `slide-${crypto.randomUUID()}.pptx`
    );

    await pres.writeFile({ fileName: tmpFile });

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    );

    res.setHeader(
      "Content-Disposition",
      'attachment; filename="Slide_1.pptx"'
    );

    const stream = fs.createReadStream(tmpFile);
    stream.pipe(res);
    stream.on("end", () => fs.unlink(tmpFile, () => {}));
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, () =>
  console.log(`PPTX microservice running on port ${PORT}`)
);
