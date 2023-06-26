const fs = require("fs");
const path = require("path");
const mammoth = require("mammoth");

const articles = { list: [] };

const inputDir = "./input";
const outputDir = "./output";

const options = {
  styleMap: [
    "p[style-name='Section Title'] => h1:fresh",
    "p[style-name='Subsection Title'] => h2:fresh",
  ],
};

// 遞迴讀取目錄下的檔案
async function readFiles(dir) {
  const files = fs.readdirSync(dir);

  for (const file of files) {
    const filePath = path.join(dir, file);
    const stats = fs.statSync(filePath);

    if (stats.isDirectory()) {
      // 如果是目錄，遞迴處理子目錄
      this.category = path.join(file);
      await readFiles(filePath); // 使用 await 等待遞迴操作完成
    } else if (path.extname(file) === ".docx") {
      // 如果是 DOCX 檔案，進行轉換
      await convertDocxToHtml(filePath, this.category); // 使用 await 等待轉換操作完成
    }
  }

  fs.writeFile(
    "./output/articles.json",
    JSON.stringify(articles),
    "utf8",
    function () {
      // 在寫入完成後執行的程式碼
      console.log("Write file completed.");
    }
  );
}

// 將 DOCX 檔案轉換為 HTML
async function convertDocxToHtml(filePath, fileCategory) {
  try {
    const result = await mammoth.convertToHtml({ path: filePath }, options);
    const html = result.value;

    // 建立相同的目錄結構
    const relativePath = path.relative(inputDir, filePath);
    const outputFilePath = path.join(outputDir, relativePath);
    const outputHtmlPath = outputFilePath.replace(".docx", ".html");

    const fileJSON = {
      title: filePath,
      category: fileCategory,
      content: html,
    };

    articles.list.push(fileJSON);

    // 將 HTML 寫入檔案
    fs.mkdirSync(path.dirname(outputHtmlPath), { recursive: true });
    fs.writeFileSync(outputHtmlPath, html);

    console.log(`Converted ${filePath} to ${outputHtmlPath}`);
  } catch (error) {
    console.error(`Error converting ${filePath}: ${error}`);
  }
}

// 開始處理檔案夾
readFiles(inputDir);
