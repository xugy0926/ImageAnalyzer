/**
 * 通过 kimi 或者 ideaTALK 来实现上传图片，获得分析结果（json），保存到本地json文件，最后汇总结果到excel。
 * 
 * 注意：使用ideaTALK 时，需手动选择 gemini-2.0-pro 模型，识别图片效果更好。
 * 
 * 使用方法：
 * > node parseImages.js kimi /path/to/your_folder
 * > node parseImages.js ideaTALK /path/to/your_folder
 */

const fs = require("fs").promises;
const path = require("path");
const { chromium } = require("playwright");
const xlsx = require("xlsx");
const { spawn } = require("child_process");

const CONFIG = {
  sites: {
    kimi: {
      URL: "https://kimi.moonshot.cn/chat/empty",
      SELECTORS: {
        attachmentButton: "label.attachment-button",
        fileInput: "input.hidden-input",
        chatInputEditor: "div.chat-input-editor",
        sendButton: "div.send-button-container",
        analyzedMark: "button.stop-message-btn",
        copyResultButton: 'span:has-text("复制")',
      },
      MAX_UPLOADS_PER_SESSION: 15,
      ANALYSIS_PROMPT:
        "按照json的格式分析图片，格式严格遵循[{store_name, product_name, price}, {store_name, product_name, price}]。不要任何废话。最后的内容给我文本字符串格式。",
    },
    ideaTALK: {
      URL: "https://aistudio.alibaba-inc.com/#/ideaTALK", // 注意：请手动选择 gemini-2.0-pro 模型以获得更好的图片识别效果
      SELECTORS: {
        attachmentButton: "div.placing-ashes-upload-btn-icon",
        fileInput: 'input[type="file"][name="file"][id="image"]',
        chatInputEditor: "textarea#question",
        sendButton: "button.submit-btn", // 占位符，实际未使用
        analyzedMark: "div.once-more-img",
        copyResultButton: "div.dislike-copy",
      },
      MAX_UPLOADS_PER_SESSION: 100,
      ANALYSIS_PROMPT:
        "按照json的格式分析图片，格式严格遵循[{store_name, product_name, price}, {store_name, product_name, price}]。不要任何废话。最后的内容给我文本字符串格式。",
    },
  },
};

const logger = {
  info: (msg) => console.log(`[INFO] ${msg}`),
  error: (msg, err) => console.error(`[ERROR] ${msg}`, err || ""),
};

/**
 * 启动 Chrome 并返回进程对象
 * @returns {Promise<ChildProcess>}
 */
async function startChrome() {
  const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
  const args = [
    "--remote-debugging-port=9222",
    "--no-first-run", // 避免首次运行提示
    "--no-default-browser-check", // 跳过默认浏览器检查
    "--user-data-dir=/tmp/chrome-debug-profile", // 使用临时用户数据目录，避免与现有实例冲突
    "--new-window"
  ];

  logger.info("启动 Chrome 服务...");
  const chromeProcess = spawn(chromePath, args, {
    stdio: "inherit",
  });

  return new Promise((resolve, reject) => {
    chromeProcess.on("error", (err) => {
      reject(new Error(`无法启动 Chrome: ${err.message}`));
    });

    chromeProcess.on("spawn", async () => {
      logger.info("Chrome 已启动，等待端口就绪...");
      // 重试机制，确保端口可用
      for (let i = 0; i < 5; i++) {
        try {
          await chromium.connectOverCDP("http://localhost:9222");
          logger.info("Chrome 调试端口已就绪");
          resolve(chromeProcess);
          return;
        } catch (err) {
          logger.info(`尝试连接端口 (第 ${i + 1}/5)...`);
          await new Promise((r) => setTimeout(r, 2000)); // 每 2 秒重试
        }
      }
      reject(new Error("Chrome 调试端口在 10 秒内未就绪"));
    });
  });
}

/**
 * 检查文件夹是否存在
 * @param {string} folderPath
 * @returns {Promise<boolean>}
 */
async function folderExists(folderPath) {
  try {
    await fs.access(folderPath, fs.constants.F_OK);
    return true;
  } catch (err) {
    return false;
  }
}

/**
 * 获取图片文件列表
 * @param {string} folderPath - 文件夹路径
 * @returns {Promise<string[]>}
 */
async function getImageFiles(folderPath) {
  const files = await fs.readdir(folderPath);
  const imageExts = [".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"];
  const imageFiles = files.filter((file) =>
    imageExts.includes(path.extname(file).toLowerCase())
  );

  if (!imageFiles.length) {
    throw new Error("文件夹中没有找到图片文件！");
  }
  logger.info(`找到的图片文件: ${imageFiles.length} 个`);
  return imageFiles;
}

/**
 * 初始化浏览器和页面
 * @returns {Promise<{browser: Browser, page: Page}>}
 */
async function initializeBrowser() {
  logger.info("连接到浏览器...");
  const browser = await chromium.connectOverCDP("http://localhost:9222");
  const context = browser.contexts()[0];
  if (!context) throw new Error("无法获取浏览器上下文！");
  const page = await context.newPage();
  return { browser, page };
}

/**
 * 网站特定的操作逻辑
 */
const siteActions = {
  kimi: async (page, filePath, fileNameWithoutExt, siteConfig) => {
    const { SELECTORS, ANALYSIS_PROMPT } = siteConfig;

    logger.info(`正在上传图片: ${filePath}`);
    await page.evaluate((selector) => {
      const button = document.querySelector(selector);
      if (button) button.addEventListener("click", (e) => e.preventDefault());
    }, SELECTORS.attachmentButton);
    await page.locator(SELECTORS.attachmentButton).click();
    await page.setInputFiles(SELECTORS.fileInput, filePath);
    logger.info(`图片上传完成: ${filePath}`);

    logger.info("输入提示文本...");
    await page.locator(SELECTORS.chatInputEditor).click();
    await page.locator(SELECTORS.chatInputEditor).fill(ANALYSIS_PROMPT);

    await waitForSendButton(page, siteConfig);
    await page.keyboard.press("Enter");
    logger.info("消息已发送！");

    await page.waitForSelector(SELECTORS.analyzedMark, {
      state: "detached",
      timeout: 60000,
    });
    logger.info("分析完成！");

    return await copyAndSaveResult(page, filePath, fileNameWithoutExt, siteConfig);
  },
  ideaTALK: async (page, filePath, fileNameWithoutExt, siteConfig) => {
    const { SELECTORS, ANALYSIS_PROMPT } = siteConfig;

    logger.info(`正在上传图片: ${filePath}`);
    await page.locator(SELECTORS.attachmentButton).first().click();
    await page.setInputFiles(SELECTORS.fileInput, filePath);
    logger.info(`图片上传完成: ${filePath}`);

    logger.info("输入提示文本...");
    await page.locator(SELECTORS.chatInputEditor).fill(ANALYSIS_PROMPT);

    await page.waitForTimeout(3000);

    await page.keyboard.press("Enter");
    logger.info("消息已发送！");

    await page.waitForSelector(SELECTORS.analyzedMark, {
      state: "attached",
      timeout: 60000,
    });
    logger.info("分析完成！");

    return await copyAndSaveResult(page, filePath, fileNameWithoutExt, siteConfig);
  },
};

/**
 * 分析单张图片（根据网站选择操作）
 * @param {Page} page
 * @param {string} filePath
 * @param {string} fileNameWithoutExt
 * @param {string} site
 */
async function analyzeImage(page, filePath, fileNameWithoutExt, site) {
  const siteConfig = CONFIG.sites[site];
  if (!siteConfig) throw new Error(`未知的网站: ${site}`);
  const action = siteActions[site];
  if (!action) throw new Error(`未定义 ${site} 的操作逻辑`);
  return await action(page, filePath, fileNameWithoutExt, siteConfig);
}

/**
 * 等待发送按钮可用（站点特定）
 * @param {Page} page
 * @param {Object} siteConfig
 */
async function waitForSendButton(page, siteConfig) {
  const { SELECTORS } = siteConfig;
  logger.info("等待发送按钮可用...");
  while (true) {
    const classAttr = await page.locator(SELECTORS.sendButton).getAttribute("class");
    if (!classAttr.includes("disabled")) {
      logger.info("发送按钮已可用！");
      break;
    }
    await page.waitForTimeout(2000);
  }
  await page.waitForTimeout(3000);
}

/**
 * 复制结果并保存（站点特定）
 * @param {Page} page
 * @param {string} filePath
 * @param {string} fileNameWithoutExt
 * @param {Object} siteConfig
 * @returns {Promise<string>}
 */
async function copyAndSaveResult(page, filePath, fileNameWithoutExt, siteConfig) {
  const { SELECTORS } = siteConfig;
  const clipboardy = (await import("clipboardy")).default;

  await page.locator(SELECTORS.copyResultButton).last().click();
  logger.info("复制按钮点击完成！");

  let content = await clipboardy.read();
  content = content.replace("```json", "").replace("```", "").trim();
  const outputPath = path.join(path.dirname(filePath), `${fileNameWithoutExt}.json`);
  await fs.writeFile(outputPath, content, "utf-8");
  logger.info(`结果已保存到: ${outputPath}`);
  return content;
}

/**
 * 将 JSON 文件合并为 Excel
 * @param {string} folderPath
 */
async function mergeJsonToExcel(folderPath) {
  logger.info("合并 JSON 文件到 Excel...");
  const jsonFiles = (await fs.readdir(folderPath)).filter((file) =>
    path.extname(file).toLowerCase() === ".json"
  );

  const allData = [];
  for (const file of jsonFiles) {
    const filePath = path.join(folderPath, file);
    try {
      const data = JSON.parse(await fs.readFile(filePath, "utf-8"));
      if (Array.isArray(data)) {
        data.forEach((item) =>
          allData.push({
            store_name: item.store_name || item.name || "",
            product_name: item.product_name || "",
            price: item.price || "",
          })
        );
      }
    } catch (err) {
      logger.error(`解析 ${file} 失败，跳过`, err);
    }
  }

  const workbook = xlsx.utils.book_new();
  const worksheet = xlsx.utils.json_to_sheet(allData);
  xlsx.utils.book_append_sheet(workbook, worksheet, "Data");
  const excelPath = path.join(folderPath, "merged_data.xlsx");
  xlsx.writeFile(workbook, excelPath);
  logger.info(`Excel 文件生成: ${excelPath}`);
}

/**
 * 重试机制包装 analyzeImage
 * @param {Page} page
 * @param {string} filePath
 * @param {string} fileNameWithoutExt
 * @param {string} site
 * @param {number} maxRetries
 */
async function retryAnalyzeImage(page, filePath, fileNameWithoutExt, site, maxRetries = 3) {
  const siteConfig = CONFIG.sites[site];
  let attempt = 1;

  while (attempt <= maxRetries) {
    try {
      const result = await analyzeImage(page, filePath, fileNameWithoutExt, site);
      logger.info(`图片 ${filePath} 分析成功（第 ${attempt} 次尝试）`);
      return result;
    } catch (err) {
      logger.error(`图片 ${filePath} 分析失败（第 ${attempt} 次尝试）`, err);
      if (attempt === maxRetries) {
        logger.error(`图片 ${filePath} 已达最大重试次数 (${maxRetries})，跳过`);
        throw err;
      }
      await page.waitForTimeout(2000);
      attempt++;
    }
  }
}

/**
 * 主函数
 * @param {string} folderPath
 * @param {string} site
 */
async function main(folderPath, site = "kimi") {
  let chromeProcess;
  try {
    const siteConfig = CONFIG.sites[site];
    if (!siteConfig) throw new Error(`不支持的网站: ${site}`);

    // 处理相对路径或绝对路径
    const absoluteFolderPath = path.resolve(folderPath);
    if (!(await folderExists(absoluteFolderPath))) {
      throw new Error(`文件夹 ${absoluteFolderPath} 不存在！`);
    }

    // 如果使用 ideaTALK，提醒选择模型
    if (site === "ideaTALK") {
      logger.info("使用 ideaTALK，请确保在页面中手动选择 gemini-2.0-pro 模型以获得最佳图片识别效果");
    }

    // 启动 Chrome
    chromeProcess = await startChrome();

    const imageFiles = await getImageFiles(absoluteFolderPath);
    const { browser, page } = await initializeBrowser();
    let currentIndex = 0;

    while (currentIndex < imageFiles.length) {
      await page.goto(siteConfig.URL);
      await page.waitForLoadState("networkidle", { timeout: 10000 });
      logger.info("页面加载完成！");

      let uploadsThisSession = 0;
      while (
        uploadsThisSession < siteConfig.MAX_UPLOADS_PER_SESSION &&
        currentIndex < imageFiles.length
      ) {
        const imageFile = imageFiles[currentIndex];
        const filePath = path.join(absoluteFolderPath, imageFile);
        const fileNameWithoutExt = path.basename(imageFile, path.extname(imageFile));

        try {
          await retryAnalyzeImage(page, filePath, fileNameWithoutExt, site);
          currentIndex++;
          uploadsThisSession++;
        } catch (err) {
          logger.error(`跳过图片 ${filePath}，继续处理下一张`);
          currentIndex++;
          uploadsThisSession++;
        }
      }

      logger.info("当前页面上传完成，准备重新加载...");
      await page.waitForTimeout(5000);
    }

    await mergeJsonToExcel(absoluteFolderPath);
    await browser.close();
    logger.info("任务完成，浏览器已关闭！");
  } catch (err) {
    logger.error("主流程出错", err);
  } finally {
    // 确保 Chrome 进程关闭
    if (chromeProcess && !chromeProcess.killed) {
      logger.info("关闭 Chrome 服务...");
      chromeProcess.kill("SIGTERM");
      await new Promise((resolve) => chromeProcess.on("exit", resolve));
      logger.info("Chrome 服务已关闭");
    }
  }
}

// 从命令行获取参数
const site = process.argv[2] || "kimi";
const folderPath = process.argv[3];
if (!folderPath) {
  console.error("错误：请提供文件夹路径！示例：node parseImages.js kimi /path/to/folder");
  process.exit(1);
}
main(folderPath, site);
