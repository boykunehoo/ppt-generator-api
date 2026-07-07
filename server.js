const express = require("express");
const cors = require("cors");
const pptxgen = require("pptxgenjs");

const app = express();

app.use(cors());
app.use(express.json({ limit: "50mb" }));

app.get("/", (req, res) => {
  res.send("Church PPT Generator API is running 🚀");
});

function addWelcomeSlide(pptx) {
  const slide = pptx.addSlide();
  slide.background = { color: "000000" };

  slide.addText("WELCOME TO", {
    x: 0.5,
    y: 1.75,
    w: 12.3,
    h: 0.5,
    fontSize: 26,
    bold: true,
    color: "D9D9D9",
    align: "center"
  });

  slide.addText("CHURCH OF GOD", {
    x: 0.5,
    y: 2.35,
    w: 12.3,
    h: 0.9,
    fontSize: 48,
    bold: true,
    color: "4F64D8",
    align: "center"
  });

  slide.addText("IN CHRIST JESUS AND BY THE HOLY SPIRIT", {
    x: 0.7,
    y: 3.35,
    w: 12,
    h: 0.45,
    fontSize: 20,
    bold: true,
    color: "D9D9D9",
    align: "center"
  });
}

function addTitleSlide(pptx, title) {
  const slide = pptx.addSlide();
  slide.background = { color: "000000" };

  slide.addText(title || "", {
    x: 0.7,
    y: 2.5,
    w: 12,
    h: 1.6,
    fontSize: 44,
    bold: true,
    color: "FFFFFF",
    align: "center",
    valign: "mid",
    fit: "shrink"
  });
}

function addLyricsSlide(pptx, section, content) {
  const slide = pptx.addSlide();
  slide.background = { color: "000000" };

  slide.addText(section ? `[${section}]` : "", {
    x: 0.5,
    y: 0.55,
    w: 12.3,
    h: 0.45,
    fontSize: 18,
    bold: true,
    color: "FFFFFF",
    align: "center"
  });

  slide.addText(content || "", {
    x: 0.8,
    y: 1.35,
    w: 11.8,
    h: 5.4,
    fontSize: 30,
    bold: true,
    color: "FFFFFF",
    align: "center",
    valign: "mid",
    fit: "shrink",
    breakLine: false
  });
}

app.post("/generate-ppt", async (req, res) => {
  try {
    const songs = req.body.songs;

    if (!songs || !Array.isArray(songs)) {
      return res.status(400).json({ error: "songs array is required" });
    }

    const pptx = new pptxgen();
    pptx.layout = "LAYOUT_WIDE";
    pptx.author = "Church Song Presenter";
    pptx.subject = "Generated Worship Presentation";
    pptx.title = "Church Worship Songs";

    addWelcomeSlide(pptx);

    for (const song of songs) {
      const slides = song.slides || [];

      for (const slideData of slides) {
        if (slideData.section === "TITLE") {
          addTitleSlide(pptx, slideData.content || song.song_title || "");
        } else {
          addLyricsSlide(pptx, slideData.section, slideData.content);
        }
      }
    }

    const buffer = await pptx.write("nodebuffer");

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    );

    res.setHeader(
      "Content-Disposition",
      "attachment; filename=Church Worship Songs.pptx"
    );

    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log(`🚀 PPT API running on port ${PORT}`);
});