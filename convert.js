const { exec } = require("child_process");
const path = require("path");

function runPowerShellCommand(command) {
  return new Promise((resolve, reject) => {
    exec(`powershell.exe -Command "${command}"`, (error, stdout, stderr) => {
      if (error) {
        reject(
          new Error(`PowerShell command execution failed: ${error.message}`)
        );
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

async function ConvertXlsToXlsx(inputPath) {
  const shellScript = path.join(__dirname, "/xls2xlsx.ps1");
  outputPath = path.join(__dirname, "/xlsx/", path.basename(inputPath) + "x");
  await runPowerShellCommand(
    `${shellScript} -InputPath ${inputPath} -OutputPath ${outputPath}`
  )
    .then((output) => {
      console.log("Converted successfully");
      console.log(output);
    })
    .catch((error) => {
      console.error("Failed to execute PowerShell command:", error);
    });
}

module.exports = {
  ConvertXlsToXlsx,
};
