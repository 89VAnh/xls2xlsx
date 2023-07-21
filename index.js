const express = require("express");
const multer = require("multer");
const path = require("path");
const { ConvertXlsToXlsx } = require("./convert.js");

const app = express();
const port = 3000;
app.use(express.json());
app.use(
  express.urlencoded({
    extended: true,
  })
);

const storage = multer.diskStorage({
  destination: function (req, file, callback) {
    if (path.extname(file.originalname) === ".xls") {
      callback(null, __dirname + "/upload");
    }
  },
  filename: function (req, file, callback) {
    callback(null, file.originalname);
  },
});

const upload = multer({ storage: storage });

app.post("/", upload.array("files"), async (req, res) => {
  if (!req.files || req.files.length === 0) {
    return res.status(400).send("No file uploaded.");
  } else {
    try {
      for (let file of req.files) {
        await ConvertXlsToXlsx(file.path);
      }
      return res.send(`Convert file successfully!`);
    } catch (err) {
      console.error(err);
      return res.status(500).send("An error occurred during conversion.");
    }
  }
});

app.listen(port, () => {
  console.log(`Example app listening at http://localhost:${port}`);
});
