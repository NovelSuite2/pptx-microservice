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
      .get(url, (response) => {
        const chunks = [];

        response.on("data", (chunk) => chunks.push(chunk));

        response.on("end", () => {
          resolve(Buffer.concat(chunks).toString("base64"));
        });

        response.on("error", reject);
      })
      .on("error", reject);
  });
}

function num(value, fallback) {
  const cleaned = String(value ?? "").replace(/=/g, "").trim();
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : fallback;
}

function str(value, fallback) {
  if (value === undefined || value === null || value === "") return fallback;
  return String(value).replace(/=/g, "").trim();
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

      logoUrl,
      logoX,
      logoY,
      logoWidth,
      logoHeight,

      contentTop,

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

    const safeBackgroundColor = str(backgroundColor, "F0F0F0");
    const safeTopBarColor = str(topBarColor, "4A72B0");
    const safeTopBarHeight = num(topBarHeight, 0.8);

    const safeLogoUrl = str(logoUrl, "");
    const safeLogoX = num(logoX, 0.4);
    const safeLogoY = num(logoY, 0.15);
    const safeLogoWidth = num(logoWidth, 0.5);
    const safeLogoHeight = num(logoHeight, 0.5);

    const safeContentTop = num(contentTop, 1.35);

    const safeTextX = num(rightColumnX, 4.3);
    const safeTextWidth = num(rightColumnWidth, 5.2);

    const safeHeadingFontSize = num(headingFontSize, 28);
    const safeHeadingFontFace = str(headingFontFace, "Calibri");
    const safeHeadingColor = str(headingColor, "111111");
    const safeHeadingBold = bool(headingBold, true);
    const safeHeadingAlign = str(headingAlign, "left");

    const safeBodyFontSize = num(bodyFontSize, 18);
    const safeBodyFontFace = str(bodyFontFace, "Calibri");
    const safeBodyColor = str(bodyColor, "111111");
    const safeBodyAlign = str(bodyAlign, "left");

    const safeSpacingBelowHeading = num(spacingBelowHeading, 0.3);
    const safeParagraphSpacingPt = num(paragraphSpacingPt, 2);

    const safeImageX = num(imageX, 0.5);
    const safeImageY = num(imageY, 1.35);
    const safeImageWidth = num(imageWidth, 3.3);
    const safeImageHeight = num(imageHeight, 3.3);

    let imgBase64 = imageBase64;

    if (!imgBase64 && imageUrl) {
      imgBase64 = await downloadImage(imageUrl);
    }

    let logoBase64 = "";

    if (safeLogoUrl) {
      logoBase64 = await downloadImage(safeLogoUrl);
    }

    const pres = new PptxGenJS();
    pres.layout = "LAYOUT_16x9";

    const slide = pres.addSlide();

    slide.background = {
      color: safeBackgroundColor
    };

    // Top bar
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0,
      y: 0,
      w: 10,
      h: safeTopBarHeight,
      fill: { color: safeTopBarColor },
      line: { color: safeTopBarColor, width: 0 }
    });

    // Logo from global YAML
    if (logoBase64) {
      slide.addImage({
        data: "image/png;base64," + logoBase64,
        x: safeLogoX,
        y: safeLogoY,
        w: safeLogoWidth,
        h: safeLogoHeight
      });
    }

    // Main slide image from slide YAML URL + global image position rules
    if (imgBase64) {
      slide.addImage({
        data: "image/png;base64," + imgBase64,
        x: safeImageX,
        y: safeImageY,
        w: safeImageWidth,
        h: safeImageHeight
      });
    }

    // Heading on right
    slide.addText(heading, {
      x: safeTextX,
      y: safeContentTop,
      w: safeTextWidth,
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

    // Body text on right
    slide.addText([{ text: bodyText }], {
      x: safeTextX,
      y: bodyY,
      w: safeTextWidth,
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

    // Next button
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

    const tmpFile = path.join(os.tmpdir(), `slide-${crypto.randomUUID()}.pptx`);

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

app.listen(PORT, () => {
  console.log(`PPTX microservice running on port ${PORT}`);
});
