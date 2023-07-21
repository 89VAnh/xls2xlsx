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
exports.ConvertXlsToXlsx = void 0;
const child_process_1 = require("child_process");
const path_1 = __importDefault(require("path"));
function runPowerShellCommand(command) {
    return new Promise((resolve, reject) => {
        (0, child_process_1.exec)(`powershell.exe -Command "${command}"`, (error, stdout, stderr) => {
            if (error) {
                reject(new Error(`PowerShell command execution failed: ${error.message}`));
                return;
            }
            if (stderr) {
                reject(new Error(`PowerShell command execution failed: ${stderr}`));
                return;
            }
            resolve(stdout.trim());
        });
    });
}
function ConvertXlsToXlsx(inputPath) {
    return __awaiter(this, void 0, void 0, function* () {
        const shellScript = path_1.default.join(__dirname, "/xls2xlsx.ps1");
        const outputPath = path_1.default.join(__dirname, "/xlsx/", path_1.default.basename(inputPath) + "x");
        yield runPowerShellCommand(`${shellScript} -InputPath ${inputPath} -OutputPath ${outputPath}`)
            .then((output) => {
            console.log("Converted successfully");
            console.log(output);
        })
            .catch((error) => {
            console.error("Failed to execute PowerShell command:", error);
        });
    });
}
exports.ConvertXlsToXlsx = ConvertXlsToXlsx;
