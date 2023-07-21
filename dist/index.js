"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const express_1 = __importDefault(require("express"));
const multer_1 = __importDefault(require("multer"));
const path_1 = __importDefault(require("path"));
const convert_1 = require("./convert");
const app = (0, express_1.default)();
const port = 3000;
app.use(express_1.default.json());
app.use(express_1.default.urlencoded({
    extended: true,
}));
const storage = multer_1.default.diskStorage({
    destination: function (req, file, callback) {
        if (path_1.default.extname(file.originalname) === ".xls") {
            callback(null, __dirname + "/upload");
        }
    },
    filename: function (req, file, callback) {
        callback(null, file.originalname);
    },
});
const upload = (0, multer_1.default)({ storage: storage });
app.post("/", upload.array("files"), (req, res) => __awaiter(void 0, void 0, void 0, function* () {
    if (!req.files || req.files.length === 0) {
        return res.status(400).send("No file uploaded.");
    }
    else {
        try {
            if (Array.isArray(req.files)) {
                for (let file of req.files) {
                    try {
                        yield (0, convert_1.ConvertXlsToXlsx)(file.path);
                    }
                    catch (err) {
                        console.error(`Error converting ${file.originalname}:`, err);
                        throw new Error("An error occurred during file conversion.");
                    }
                }
                return res.send("Files converted successfully!");
            }
        }
        catch (err) {
            console.error(err);
            return res.status(500).send("An error occurred during conversion.");
        }
    }
}));
app.listen(port, () => {
    console.log(`Example app listening at http://localhost:${port}`);
});
