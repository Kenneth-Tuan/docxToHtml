const fs = require("fs");
const path = require("path");
const mammoth = require("mammoth");

const inputDir = "./input";
const outputDir = "./output";

// 遞迴讀取目錄下的檔案
function readFiles(dir) {
  const files = fs.readdirSync(dir);

  files.forEach((file) => {
    const filePath = path.join(dir, file);
    const stats = fs.statSync(filePath);

    if (stats.isDirectory()) {
      // 如果是目錄，遞迴處理子目錄
      readFiles(filePath);
    } else if (path.extname(file) === ".docx") {
      // 如果是 DOCX 檔案，進行轉換
      convertDocxToHtml(filePath);
    }
  });
}

// 將 DOCX 檔案轉換為 HTML
function convertDocxToHtml(filePath) {
  const fileContent = fs.readFileSync(filePath);

  mammoth
    .convertToHtml({ path: filePath })
    .then((result) => {
      const html = result.value;

      // 建立相同的目錄結構
      const relativePath = path.relative(inputDir, filePath);
      const outputFilePath = path.join(outputDir, relativePath);
      const outputHtmlPath = outputFilePath.replace(".docx", ".html");

      // 將 HTML 寫入檔案
      fs.mkdirSync(path.dirname(outputHtmlPath), { recursive: true });
      fs.writeFileSync(outputHtmlPath, html);

      console.log(`Converted ${filePath} to ${outputHtmlPath}`);
    })
    .catch((error) => {
      console.error(`Error converting ${filePath}: ${error}`);
    });
}

// 開始處理檔案夾
readFiles(inputDir);
