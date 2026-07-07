const express = require("express");
const cors = require("cors");
const pptxgen = require("pptxgenjs");

const app = express();

app.use(cors());
app.use(express.json({ limit: "50mb" }));

// Health Check
app.get("/", (req, res) => {
    res.send("Church PPT Generator API is running 🚀");
});

// Generate PowerPoint
app.post("/generate-ppt", async (req, res) => {
    try {
        const songs = req.body.songs;

        if (!songs || !Array.isArray(songs)) {
            return res.status(400).json({
                error: "songs array is required"
            });
        }

        const pptx = new pptxgen();

        pptx.layout = "LAYOUT_WIDE";
        pptx.author = "Winter Chaos";
        pptx.company = "Church Song Presenter";
        pptx.subject = "Generated Worship Presentation";
        pptx.title = "Church Worship Songs";

        for (const song of songs) {

            for (const slideData of song.slides) {

                const slide = pptx.addSlide();

                slide.background = {
                    color: "000000"
                };

                // Section
                slide.addText(slideData.section, {
                    x: 0.3,
                    y: 0.25,
                    w: 12.8,
                    h: 0.4,
                    align: "center",
                    bold: true,
                    color: "AAAAAA",
                    fontSize: 16
                });

                // Lyrics
                slide.addText(slideData.content, {
                    x: 0.6,
                    y: 1.1,
                    w: 12.1,
                    h: 5.7,
                    align: "center",
                    valign: "mid",
                    color: "FFFFFF",
                    fontSize: 30,
                    bold: true,
                    fit: "shrink"
                });

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

        res.status(500).json({
            error: err.message
        });
    }
});

app.get("/test-ppt", async (req, res) => {
    const pptx = new pptxgen();

    const slide = pptx.addSlide();

    slide.background = { color: "000000" };

    slide.addText("TEST PPT", {
        x: 1,
        y: 2,
        w: 11,
        h: 1,
        fontSize: 40,
        bold: true,
        color: "FFFFFF",
        align: "center"
    });

    const buffer = await pptx.write("nodebuffer");

    res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    );

    res.setHeader(
        "Content-Disposition",
        "attachment; filename=test.pptx"
    );

    res.send(buffer);
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
    console.log(`🚀 PPT API running on port ${PORT}`);
});