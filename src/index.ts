import express, { Application, Request, Response } from "express";
import multer from "multer";
import path from "path";
import { ConvertXlsToXlsx } from "./convert";

const app: Application = express();
const port = 3000;
app.use(express.json());
app.use(
  express.urlencoded({
    extended: true,
  })
);

const storage = multer.diskStorage({
  destination: function (
    req: Request,
    file: Express.Multer.File,
    callback: (error: Error | null, filename: string) => void
  ) {
    if (path.extname(file.originalname) === ".xls") {
      callback(null, __dirname + "/upload");
    }
  },
  filename: function (
    req: Request,
    file: Express.Multer.File,
    callback: (error: Error | null, filename: string) => void
  ) {
    callback(null, file.originalname);
  },
});

const upload = multer({ storage: storage });

app.post("/", upload.array("files"), async (req: Request, res: Response) => {
  if (!req.files || req.files.length === 0) {
    return res.status(400).send("No file uploaded.");
  } else {
    try {
      if (Array.isArray(req.files)) {
        for (let file of req.files) {
          try {
            await ConvertXlsToXlsx(file.path);
          } catch (err) {
            console.error(`Error converting ${file.originalname}:`, err);
            throw new Error("An error occurred during file conversion.");
          }
        }
        return res.send("Files converted successfully!");
      }
    } catch (err) {
      console.error(err);
      return res.status(500).send("An error occurred during conversion.");
    }
  }
});

app.listen(port, () => {
  console.log(`Example app listening at http://localhost:${port}`);
});
