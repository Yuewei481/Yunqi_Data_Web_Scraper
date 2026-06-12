import { chromium } from "playwright";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import { spawn } from "node:child_process";

const ROOT = path.resolve(import.meta.dirname);
const OUTPUT_DIR = process.env.OUTPUT_DIR
  ? path.resolve(ROOT, process.env.OUTPUT_DIR)
  : path.join(ROOT, "outputs", "yunqi-pop-up-greeting-card");
const IMAGE_DIR = path.join(OUTPUT_DIR, "images");
const DATA_PATH = path.join(OUTPUT_DIR, "products.json");
const EXISTING_PRODUCT_IDS_PATH = process.env.EXISTING_PRODUCT_IDS_PATH || "";
const TEMPLATE_PATH = process.env.EXCEL_TEMPLATE || path.resolve(ROOT, "..", "templates", "选品表格-模板.xlsx");
const PYTHON = process.env.CODEX_PYTHON || process.env.PYTHON || "python3";
const KEYWORD = process.env.YUNQI_KEYWORD || "pop up greeting card";
const KEYWORDS = splitEnvList(KEYWORD);
const CATEGORY_PARENT = normalizeText(process.env.YUNQI_CATEGORY_PARENT || "");
const CATEGORY_CHILDREN = splitEnvList(process.env.YUNQI_CATEGORY_CHILDREN || "");
const USERNAME = process.env.YUNQI_USERNAME || "";
const PASSWORD = process.env.YUNQI_PASSWORD || "";
const LOGIN_URL = "https://www.yunqishuju.com/login";
const HEADLESS = process.env.HEADLESS === "1";
const CHROME_PATH = process.env.CHROME_PATH || "";
const CHROME_ARGS = [
  "--disable-session-crashed-bubble",
  "--disable-infobars",
  "--no-first-run",
  "--no-default-browser-check",
];
const LOGIN_WAIT_MS = Number(process.env.LOGIN_WAIT_MS || "180000");
const USER_DATA_DIR = process.env.USER_DATA_DIR || path.join(OUTPUT_DIR, "chrome-profile");
const START_URL = "https://www.yunqishuju.com/temu/home/";
const MAX_PAGES = Number(process.env.MAX_PAGES || "50");
const DAILY_MIN = Number(process.env.DAILY_MIN || "30");
const MONTHLY_MIN = Number(process.env.MONTHLY_MIN || "1000");
const MIN_IMAGES_PER_PRODUCT = Number(process.env.MIN_IMAGES_PER_PRODUCT || "2");
const IMAGE_RETRY_MAX = Number(process.env.IMAGE_RETRY_MAX || "6");
const IMAGE_RETRY_WAIT_MS = Number(process.env.IMAGE_RETRY_WAIT_MS || "10000");
const HOVER_IMAGE_WAIT_MS = Number(process.env.HOVER_IMAGE_WAIT_MS || "3000");
const DUPLICATE_IMAGE_RETRY_MAX = Number(process.env.DUPLICATE_IMAGE_RETRY_MAX || "2");
const USE_PERSISTENT_CONTEXT = process.env.PERSISTENT_CONTEXT !== "0";
const MAIN_ROW_SELECTOR = ".el-table__body-wrapper tbody tr.el-table__row";
const DEBUG_RESPONSES = process.env.DEBUG_RESPONSES === "1";
const responseDebugRecords = [];

function splitEnvList(value) {
  return String(value || "")
    .split(",")
    .map((item) => normalizeText(item))
    .filter(Boolean);
}

const SELECTORS = {
  username: [
    'input[name="username"]',
    'input[name="phone"]',
    'input[type="tel"]',
    'input[placeholder*="手机号"]',
    'input[placeholder*="账号"]',
    'input[placeholder*="用户名"]',
    'input[placeholder*="手机"]',
  ],
  password: [
    'input[name="password"]',
    'input[type="password"]',
    'input[placeholder*="密码"]',
  ],
  loginSubmit: [
    "button.sgin",
    'button:has-text("登录")',
    '.login button:has-text("登录")',
    '[role="button"]:has-text("登录")',
  ],
  searchInput: [
    'input[placeholder*="搜索"]',
    'input[placeholder*="请输入"]',
    'input[placeholder*="关键词"]',
    'input[placeholder*="中文/英文"]',
    'input[type="search"]',
    '.search input',
  ],
  searchSubmit: [
    'button:has-text("搜索")',
    '[role="button"]:has-text("搜索")',
    '.search button',
  ],
  tableRows: [
    MAIN_ROW_SELECTOR,
    "tr.el-table__row",
    ".ant-table-tbody tr",
    "table tbody tr",
  ],
  tableHeaders: [
    ".el-table__header-wrapper thead th",
    ".ant-table-thead th",
    "table thead th",
  ],
  nextPage: [
    ".el-pagination button.btn-next:not([disabled])",
    ".ant-pagination-next:not(.ant-pagination-disabled) button",
    'button:has-text("下一页")',
  ],
};

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const jitter = (min = 220, max = 850) => Math.floor(min + Math.random() * (max - min));

async function humanPause(min, max) {
  await sleep(jitter(min, max));
}

async function firstVisible(page, selectors, timeout = 1200) {
  for (const selector of selectors) {
    const loc = page.locator(selector).first();
    try {
      await loc.waitFor({ state: "visible", timeout });
      return loc;
    } catch {
      // Try the next selector.
    }
  }
  return null;
}

async function typeLikeHuman(locator, value) {
  try {
    await locator.click({ delay: jitter(30, 90), timeout: 8000 });
    await locator.fill("");
    for (const char of value) {
      await locator.type(char, { delay: jitter(45, 135) });
    }
  } catch {
    await locator.evaluate((el, text) => {
      const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")?.set;
      setter?.call(el, "");
      el.dispatchEvent(new Event("input", { bubbles: true }));
      setter?.call(el, text);
      el.dispatchEvent(new Event("input", { bubbles: true }));
      el.dispatchEvent(new Event("change", { bubbles: true }));
    }, value);
  }
}

async function gotoSoft(page, url) {
  let lastError = null;
  for (let attempt = 1; attempt <= 4; attempt += 1) {
    try {
      await page.goto(url, { waitUntil: "commit", timeout: 60000 });
      lastError = null;
      break;
    } catch (error) {
      lastError = error;
      const currentUrl = page.url();
      if (currentUrl.startsWith(url.replace(/\/$/, ""))) {
        lastError = null;
        break;
      }
      await sleep(2500 * attempt);
    }
  }
  if (lastError) throw lastError;
  await page.waitForLoadState("domcontentloaded", { timeout: 8000 }).catch(() => {});
  await humanPause(1200, 2200);
}

async function login(page) {
  if (!USERNAME || !PASSWORD) {
    throw new Error("请先设置 YUNQI_USERNAME 和 YUNQI_PASSWORD。可以复制 .env.example 为 .env 后填写账号密码。");
  }
  for (let attempt = 1; attempt <= 2; attempt += 1) {
    console.log(attempt === 1 ? "打开云启数据登录页..." : "未确认登录成功，重新打开登录页重试...");
    await gotoSoft(page, LOGIN_URL);
    await waitForLoginOrAuthenticated(page);
    if (await isLoggedInPage(page)) {
      console.log("当前账号已处于登录状态。");
      await closePopups(page);
      return;
    }
    console.log("登录页已显示，切换到账号密码登录...");
    await ensurePasswordLoginMode(page);

    const username = await firstVisible(page, SELECTORS.username, 2500);
    const password = await firstVisible(page, SELECTORS.password, 2500);
    if (!username || !password) {
      await saveDebug(page, "login-fields-not-found");
      throw new Error("没有找到登录页的账号或密码输入框。请先用可视模式运行，并根据页面更新 SELECTORS。");
    }

    await typeLikeHuman(username, USERNAME);
    await humanPause();
    await typeLikeHuman(password, PASSWORD);
    await humanPause();

    const submit = await firstVisible(page, SELECTORS.loginSubmit, 2500);
    if (!submit) throw new Error("没有找到登录按钮。");
    console.log("已输入账号密码，点击登录并等待响应...");
    await clickLikeHuman(page, submit);
    await page.waitForTimeout(1800);
    if (page.url().includes("/login") && await submit.isVisible().catch(() => false)) {
      await clickLikeHuman(page, submit);
    }
    if (await waitForLoginResult(page)) {
      await closePopups(page);
      return;
    }
    await saveDebug(page, `login-attempt-${attempt}-not-confirmed`);
  }

  throw new Error("登录后没有进入已登录状态，可能需要验证码、短信验证或账号状态确认。");
}

async function isLoginPageReady(page) {
  return await page.evaluate(() => {
    const text = document.body?.innerText || "";
    return Boolean(
      document.querySelector(".MainContainerRight") ||
      document.querySelector(".switchtab") ||
      document.querySelector(".tips") ||
      document.querySelector(".tabList .tabItem") ||
      document.querySelector("input[placeholder*='手机']") ||
      text.includes("微信扫码注册/登录") ||
      text.includes("验证码登录") ||
      text.includes("云启数据用户登录")
    );
  }).catch(() => false);
}

async function waitForLoginPageReady(page) {
  const deadline = Date.now() + Number(process.env.LOGIN_PAGE_READY_MS || "180000");
  let reloaded = false;
  while (Date.now() < deadline) {
    if (await isLoginPageReady(page)) {
      await humanPause(1200, 2200);
      return;
    }
    if (!reloaded && Date.now() + 45000 < deadline) {
      await page.waitForTimeout(8000);
      const stillBlank = await page.evaluate(() => (document.querySelector("#__layout")?.textContent || "").trim().length < 20).catch(() => false);
      if (stillBlank) {
        await page.reload({ waitUntil: "commit", timeout: 60000 }).catch(() => {});
        reloaded = true;
      }
    }
    await page.waitForTimeout(2500);
  }
  if (await isLoginPageReady(page)) {
    await humanPause(1200, 2200);
    return;
  }
  await saveDebug(page, "login-page-not-ready");
  throw new Error("登录页长时间没有渲染完成。");
}

async function waitForLoginOrAuthenticated(page) {
  const deadline = Date.now() + Number(process.env.LOGIN_PAGE_READY_MS || "180000");
  let reloaded = false;
  while (Date.now() < deadline) {
    if (await isLoggedInPage(page) || await isLoginPageReady(page)) {
      await humanPause(1200, 2200);
      return;
    }
    if (!reloaded && Date.now() + 45000 < deadline) {
      await page.waitForTimeout(8000);
      const stillBlank = await page.evaluate(() => (document.querySelector("#__layout")?.textContent || "").trim().length < 20).catch(() => false);
      if (stillBlank) {
        await page.reload({ waitUntil: "commit", timeout: 60000 }).catch(() => {});
        reloaded = true;
      }
    }
    await page.waitForTimeout(2500);
  }
  await page.waitForTimeout(8000);
  if (await isLoggedInPage(page) || await isLoginPageReady(page)) {
    await humanPause(1200, 2200);
    return;
  }
  await saveDebug(page, "login-or-authenticated-not-ready");
  throw new Error("登录页或已登录首页长时间没有渲染完成。");
}

async function clickLikeHuman(page, locator) {
  await locator.scrollIntoViewIfNeeded().catch(() => {});
  const box = await locator.boundingBox().catch(() => null);
  if (box) {
    const x = box.x + box.width / 2 + Math.random() * 8 - 4;
    const y = box.y + box.height / 2 + Math.random() * 6 - 3;
    await page.mouse.move(x, y, { steps: 12 });
    await humanPause(180, 360);
    await page.mouse.down();
    await humanPause(80, 180);
    await page.mouse.up();
  } else {
    await locator.click({ delay: jitter(80, 180), force: true });
  }
}

async function isLoginLikePage(page) {
  return await page.evaluate(() => {
    const text = document.body?.innerText || "";
    const loginPanel = document.querySelector(".MainContainerRight");
    const panelText = loginPanel?.textContent || "";
    return Boolean(
      (loginPanel && (panelText.includes("微信扫码注册/登录") || panelText.includes("云启数据用户登录"))) ||
      document.querySelector(".switchtab") ||
      text.includes("微信扫码注册/登录") ||
      text.includes("云启数据用户登录") ||
      (text.includes("验证码登录") && text.includes("用户协议"))
    );
  }).catch(() => false);
}

async function isLoggedInPage(page) {
  return await page.evaluate(() => {
    const text = document.body?.innerText || "";
    const loginPanel = document.querySelector(".MainContainerRight");
    const panelText = loginPanel?.textContent || "";
    const loginLike = Boolean(
      (loginPanel && (panelText.includes("微信扫码注册/登录") || panelText.includes("云启数据用户登录"))) ||
      text.includes("微信扫码注册/登录") ||
      text.includes("云启数据用户登录")
    );
    if (loginLike) return false;
    if (text.includes("请登录")) return false;
    return Boolean(
      text.includes("VIP天数") ||
      text.includes("Captain") ||
      text.includes("购买/续费") ||
      text.includes("工作台") ||
      (
        document.querySelector("input[placeholder*='中文'][placeholder*='英文'][placeholder*='关键词']") &&
        text.includes("商品库")
      )
    );
  }).catch(() => false);
}

async function waitForLoginResult(page) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < LOGIN_WAIT_MS) {
    if (page.isClosed()) return false;
    await page.waitForLoadState("networkidle", { timeout: 5000 }).catch(() => {});
    await closePopups(page);
    if (await isLoggedInPage(page)) return true;
    if (page.isClosed()) return false;
    await page.waitForTimeout(3500).catch(() => {});
  }
  return false;
}

async function ensurePasswordLoginMode(page) {
  await page.waitForTimeout(5000);
  for (let attempt = 0; attempt < 20; attempt += 1) {
    if (await page.locator('input[type="password"], input[placeholder*="密码"]').first().isVisible().catch(() => false)) {
      return;
    }

    const qrLoginVisible = await page.getByText("微信扫码注册/登录", { exact: true }).isVisible().catch(() => false);
    if (qrLoginVisible) {
      const switcherTips = page.locator(".tips").filter({ hasText: "验证码登录" }).first();
      const switcherIcon = page.locator(".switchtab").first();
      let clickedSwitcher = false;
      if (await switcherIcon.isVisible().catch(() => false)) {
        await clickLikeHuman(page, switcherIcon).catch(() => {});
        clickedSwitcher = true;
        await humanPause(800, 1400);
      }
      if (await switcherTips.isVisible().catch(() => false)) {
        await clickLikeHuman(page, switcherTips).catch(() => {});
        clickedSwitcher = true;
      }
      await humanPause(clickedSwitcher ? 2200 : 1200, clickedSwitcher ? 3600 : 2000);
      continue;
    }

    await page.evaluate(() => {
      const passwordTab = Array.from(document.querySelectorAll(".tabList .tabItem")).find(
        (el) => (el.textContent || "").trim() === "密码"
      );
      if (passwordTab) passwordTab.click();
    }).catch(() => {});
    await humanPause(1400, 2400);
  }
}

async function closePopups(page) {
  await closeExternalTemuTabs(page);
  const selectors = [
    ".el-dialog__headerbtn",
    ".el-dialog__close",
    ".ant-modal-close",
    ".modal-close",
    ".close",
    ".popup-close",
    ".notice-close",
    'text=×',
    'text=✕',
    '[aria-label="Close"]',
    'button:has-text("关闭")',
    'button:has-text("我知道了")',
    'button:has-text("知道了")',
    'button:has-text("取消")',
  ];
  for (let round = 0; round < 4; round += 1) {
    let clicked = false;
    if (await closeEmbeddedTemuProductPage(page)) {
      await humanPause(500, 1000);
      clicked = true;
    }
    for (const selector of selectors) {
      const loc = page.locator(selector).first();
      if (await loc.isVisible().catch(() => false)) {
        await loc.click({ delay: jitter(30, 100), force: true }).catch(() => {});
        await humanPause(500, 1000);
        clicked = true;
      }
    }
    const clickedTextClose = await page.evaluate(() => {
      const isVisible = (el) => {
        const style = window.getComputedStyle(el);
        const rect = el.getBoundingClientRect();
        return style.visibility !== "hidden" && style.display !== "none" && rect.width > 0 && rect.height > 0;
      };
      const closers = Array.from(document.querySelectorAll("button, span, i, div")).filter((el) => {
        const text = (el.textContent || "").trim();
        const className = String(el.className || "").toLowerCase();
        return isVisible(el) && (text === "×" || text === "✕" || className.includes("close"));
      });
      const target = closers
        .map((el) => ({ el, rect: el.getBoundingClientRect(), z: Number(window.getComputedStyle(el).zIndex) || 0 }))
        .filter(({ rect }) => rect.top < window.innerHeight * 0.75 && rect.left > window.innerWidth * 0.35)
        .sort((a, b) => b.z - a.z || a.rect.width * a.rect.height - b.rect.width * b.rect.height)[0]?.el;
      if (target) {
        target.click();
        return true;
      }
      return false;
    }).catch(() => false);
    if (clickedTextClose) {
      await humanPause(500, 1000);
      clicked = true;
    }
    await page.keyboard.press("Escape").catch(() => {});
    if (!clicked) break;
  }
}

async function closeExternalTemuTabs(page) {
  const context = page.context();
  for (const other of context.pages()) {
    if (other === page) continue;
    const url = other.url();
    const title = await other.title().catch(() => "");
    if (/temu\.com/i.test(url) || /^Temu\b/i.test(title)) {
      await other.close({ runBeforeUnload: false }).catch(() => {});
    }
  }
  await page.bringToFront().catch(() => {});
}

async function closeEmbeddedTemuProductPage(page) {
  return await page.evaluate(() => {
    const hasTemuSearch = Boolean(
      document.querySelector("input[placeholder*='中文'][placeholder*='英文'][placeholder*='关键词']")
    );
    if (!hasTemuSearch) return false;

    const visible = (el) => {
      const style = window.getComputedStyle(el);
      const rect = el.getBoundingClientRect();
      return style.display !== "none" && style.visibility !== "hidden" && rect.width > 0 && rect.height > 0;
    };
    const overlays = Array.from(document.querySelectorAll(".temp-page, [class*='temp-page']")).filter((el) => {
      const text = (el.textContent || "").replace(/\s+/g, "");
      return (
        visible(el) &&
        (text.includes("立即购买") ||
          text.includes("订单和账户") ||
          text.includes("客服") ||
          text.includes("已售") ||
          text.includes("CA$") ||
          text.includes("Temu"))
      );
    });
    if (!overlays.length) return false;

    for (const overlay of overlays) {
      const closer = Array.from(overlay.querySelectorAll("button, span, i, div")).find((el) => {
        if (!visible(el)) return false;
        const text = (el.textContent || "").trim();
        const cls = String(el.className || "").toLowerCase();
        const rect = el.getBoundingClientRect();
        return (
          text === "×" ||
          text === "✕" ||
          text === "关闭" ||
          (cls.includes("close") && rect.width <= 80 && rect.height <= 80)
        );
      });
      if (closer) {
        closer.click();
      } else {
        overlay.remove();
      }
    }
    return true;
  }).catch(() => false);
}

async function saveDebug(page, name) {
  await fs.mkdir(OUTPUT_DIR, { recursive: true });
  const safeName = `${new Date().toISOString().replace(/[:.]/g, "-")}-${name}`;
  if (page.isClosed()) {
    await fs.writeFile(path.join(OUTPUT_DIR, `${safeName}.txt`), "Page was already closed before debug capture.\n").catch(() => {});
    return;
  }
  await page.screenshot({ path: path.join(OUTPUT_DIR, `${safeName}.png`), fullPage: true }).catch(() => {});
  await fs.writeFile(path.join(OUTPUT_DIR, `${safeName}.html`), await page.content()).catch(() => {});
}

async function searchKeyword(page, keyword, applyFilters = false) {
  console.log("登录完成，进入 Temu 页面...");
  await enterTemuHome(page);
  await waitForTemuOrLoginShell(page);
  if (await isLoginLikePage(page)) {
    console.log("进入 Temu 后仍显示未登录，重新登录一次...");
    await login(page);
    await enterTemuHome(page);
    await waitForTemuOrLoginShell(page);
  }
  if (await isLoginLikePage(page)) {
    await saveDebug(page, "temu-redirected-to-login");
    throw new Error("进入 Temu 后仍被重定向到登录页。");
  }
  await waitForTemuLoaded(page);

  const input = await getTopKeywordInput(page);
  if (!input) {
    await saveDebug(page, "search-input-not-found");
    throw new Error("没有找到左上角关键词搜索框。");
  }
  if (applyFilters) await applyCategoryFilter(page);
  console.log(`在 Temu 左上角搜索框搜索：${keyword}`);
  await typeLikeHuman(input, keyword);
  await humanPause();

  await clickTopKeywordSearchButton(page, input);
  await page.waitForTimeout(1200);
  console.log("已点击左上角搜索按钮，等待搜索结果...");
  await waitForSearchResults(page);
}

async function openCategoryDropdown(page) {
  if (!CATEGORY_PARENT) return false;
  const clickDropdown = async () => page.evaluate(() => {
    const visible = (el) => {
      const style = window.getComputedStyle(el);
      const rect = el.getBoundingClientRect();
      return style.display !== "none" && style.visibility !== "hidden" && rect.width > 0 && rect.height > 0;
    };
    const labels = Array.from(document.querySelectorAll("body *")).filter((el) => {
      const text = (el.textContent || "").replace(/\s+/g, "").trim();
      const rect = el.getBoundingClientRect();
      return visible(el) && text === "分类" && rect.top > 180 && rect.top < 420;
    });
    const label = labels.sort((a, b) => a.getBoundingClientRect().left - b.getBoundingClientRect().left)[0];
    const input = Array.from(document.querySelectorAll("input[placeholder='分类筛选']"))
      .find((el) => visible(el));
    const cascader = input?.closest(".el-cascader") ||
      label?.closest(".el-form-item")?.querySelector(".el-cascader") ||
      label?.parentElement?.querySelector(".el-cascader");
    const target = cascader?.querySelector(".el-input__suffix, .el-input__suffix-inner, .el-input") || cascader;
    if (!target) return false;
    target.dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));
    target.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
    target.dispatchEvent(new MouseEvent("mouseup", { bubbles: true }));
    target.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    return true;
  });

  let opened = await clickDropdown();
  if (!opened) {
    await saveDebug(page, "category-dropdown-not-found");
    throw new Error("没有找到分类筛选下拉框。");
  }
  for (let attempt = 1; attempt <= 2; attempt += 1) {
    const loaded = await page.waitForFunction(
      () => {
        const visible = (el) => {
          const style = window.getComputedStyle(el);
          const rect = el.getBoundingClientRect();
          return style.display !== "none" && style.visibility !== "hidden" && rect.width > 0 && rect.height > 0;
        };
        const panel = Array.from(document.querySelectorAll(".el-popper.outCascader, .el-cascader__dropdown, .el-cascader-panel"))
          .find((el) => visible(el));
        if (!panel) return false;
        return Array.from(panel.querySelectorAll(".el-cascader-menu__list .el-cascader-node"))
          .some((node) => visible(node) && (node.textContent || "").replace(/\s+/g, "").trim());
      },
      null,
      { timeout: 15000 }
    ).then(() => true).catch(() => false);
    if (loaded) return true;
    await page.keyboard.press("Escape").catch(() => {});
    await page.waitForTimeout(800);
    opened = await clickDropdown();
    if (!opened) break;
  }
  await saveDebug(page, "category-options-not-loaded");
  throw new Error("分类列表没有加载出来。");
}

async function clickCategoryOption(page, label, options = {}) {
  const clicked = await page.evaluate(({ label, preferCheckbox, preferExpand }) => {
    const normalize = (value) => String(value || "")
      .replace(/\s+/g, " ")
      .replace(/[（(].*?[）)]/g, "")
      .trim()
      .toLowerCase();
    const needle = normalize(label);
    const visible = (el) => {
      const style = window.getComputedStyle(el);
      const rect = el.getBoundingClientRect();
      return style.display !== "none" && style.visibility !== "hidden" && rect.width > 0 && rect.height > 0;
    };
    const roots = Array.from(document.querySelectorAll(".el-popper.outCascader, .el-cascader__dropdown, .el-cascader-panel, .el-cascader-menu"))
      .filter(visible);
    const candidates = [];
    for (const root of roots) {
      for (const el of Array.from(root.querySelectorAll(".el-cascader-node, li, label, .el-checkbox, .el-select-dropdown__item, div, span"))) {
        const node = el.closest(".el-cascader-node") || el;
        if (!visible(node)) continue;
        const text = (node.textContent || "").replace(/\s+/g, " ").trim();
        if (!text || text.length > 140) continue;
        const normalized = normalize(text);
        if (normalized === needle || normalized.startsWith(`${needle} `) || normalized.includes(needle)) {
          candidates.push(node);
        }
      }
    }
    const uniqueCandidates = Array.from(new Set(candidates));
    const candidate = uniqueCandidates
      .map((el) => ({ el, rect: el.getBoundingClientRect(), score: Math.abs(normalize(el.textContent || "").length - needle.length) }))
      .sort((a, b) => a.score - b.score || a.rect.left - b.rect.left)[0]?.el;
    if (!candidate) return false;

    let target = candidate;
    if (preferCheckbox) {
      const checkbox = candidate.querySelector(".el-checkbox__input");
      const input = candidate.querySelector("input[type='checkbox']");
      if (checkbox?.classList.contains("is-checked") || input?.checked) return true;
      target = candidate.querySelector(".el-checkbox__inner, .el-checkbox__input, input[type='checkbox']") ||
        candidate.closest("label")?.querySelector(".el-checkbox__input, input[type='checkbox']") ||
        candidate;
    } else if (preferExpand) {
      target = candidate.querySelector(".el-cascader-node__postfix, .el-icon-arrow-right") ||
        Array.from(candidate.querySelectorAll("span, i, svg")).at(-1) ||
        candidate;
    }

    target.dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));
    target.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
    target.dispatchEvent(new MouseEvent("mouseup", { bubbles: true }));
    target.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    return true;
  }, { label, preferCheckbox: Boolean(options.preferCheckbox), preferExpand: Boolean(options.preferExpand) });
  if (!clicked) {
    await saveDebug(page, `category-option-not-found-${label}`);
    throw new Error(`没有找到分类选项：${label}`);
  }
  await page.waitForTimeout(700);
  if (options.preferCheckbox) {
    const checked = await page.waitForFunction(
      (label) => {
        const normalize = (value) => String(value || "")
          .replace(/\s+/g, " ")
          .replace(/[（(].*?[）)]/g, "")
          .trim()
          .toLowerCase();
        const needle = normalize(label);
        const visible = (el) => {
          const style = window.getComputedStyle(el);
          const rect = el.getBoundingClientRect();
          return style.display !== "none" && style.visibility !== "hidden" && rect.width > 0 && rect.height > 0;
        };
        const nodes = Array.from(document.querySelectorAll(".el-popper.outCascader .el-cascader-node, .el-cascader__dropdown .el-cascader-node"));
        for (const node of nodes) {
          if (!visible(node)) continue;
          const text = (node.textContent || "").replace(/\s+/g, " ").trim();
          const normalized = normalize(text);
          if (!(normalized === needle || normalized.startsWith(`${needle} `) || normalized.includes(needle))) continue;
          const checkbox = node.querySelector(".el-checkbox__input");
          const input = node.querySelector("input[type='checkbox']");
          return Boolean(checkbox?.classList.contains("is-checked") || input?.checked);
        }
        return false;
      },
      label,
      { timeout: 3000 }
    ).then(() => true).catch(() => false);
    if (!checked) {
      await saveDebug(page, `category-option-not-checked-${label}`);
      throw new Error(`分类选项没有成功勾选：${label}`);
    }
  }
}

async function applyCategoryFilter(page) {
  if (!CATEGORY_PARENT) return;
  console.log(
    CATEGORY_CHILDREN.length
      ? `应用分类筛选：${CATEGORY_PARENT} / ${CATEGORY_CHILDREN.join(", ")}`
      : `应用分类筛选：${CATEGORY_PARENT}`
  );
  await page.evaluate(() => {
    window.scrollTo(0, 0);
    document.querySelector(".body")?.scrollTo?.(0, 0);
    document.querySelector(".el-table__body-wrapper")?.scrollTo?.(0, 0);
  }).catch(() => {});
  await page.waitForTimeout(500);
  await openCategoryDropdown(page);
  if (CATEGORY_CHILDREN.length) {
    await clickCategoryOption(page, CATEGORY_PARENT, { preferExpand: true });
    for (const child of CATEGORY_CHILDREN) {
      await clickCategoryOption(page, child, { preferCheckbox: true });
    }
  } else {
    await clickCategoryOption(page, CATEGORY_PARENT, { preferCheckbox: true });
  }
  await page.keyboard.press("Escape").catch(() => {});
  await page.waitForTimeout(700);
}

async function ensureTemuHomeReady(page) {
  console.log("进入 Temu 页面前检查登录状态...");
  await gotoSoft(page, START_URL);
  await closePopups(page);
  await waitForTemuOrLoginShell(page);

  if (await isLoginLikePage(page)) {
    console.log("Temu 页面显示未登录，先登录云启数据...");
    await login(page);
    console.log("登录确认完成，重新进入 Temu 页面...");
    await gotoSoft(page, START_URL);
    await closePopups(page);
    await waitForTemuOrLoginShell(page);
  }

  if (await isLoginLikePage(page)) {
    await saveDebug(page, "temu-still-login-after-login");
    throw new Error("登录后进入 Temu 仍显示登录页，可能需要验证码、短信验证或账号状态确认。");
  }

  await waitForTemuLoaded(page);
}

async function waitForTemuOrLoginShell(page) {
  const deadline = Date.now() + Number(process.env.TEMU_READY_WAIT_MS || "90000");
  while (Date.now() < deadline) {
    await closePopups(page);
    const hasKeywordInput = Boolean(await getTopKeywordInput(page, 400).catch(() => null));
    if (hasKeywordInput || await isLoginLikePage(page) || await isLoggedInPage(page)) return;
    await page.waitForTimeout(2500);
  }
}

async function waitForManualTemuReady(page) {
  console.log("请在打开的 Chrome 中手动登录云启数据，并进入顶部 Temu 栏。脚本会等待页面就绪后继续。");
  const deadline = Date.now() + Number(process.env.MANUAL_WAIT_MS || "600000");
  while (Date.now() < deadline) {
    await closePopups(page);
    const url = page.url();
    const input = await firstVisible(page, SELECTORS.searchInput, 500).catch(() => null);
    if (/temu/i.test(url) && input) return;
    await page.waitForTimeout(2500);
  }
  await saveDebug(page, "manual-temu-not-ready");
  throw new Error("等待手动登录并进入 Temu 页面超时。");
}

async function waitForTemuLoaded(page) {
  const deadline = Date.now() + Number(process.env.TEMU_READY_WAIT_MS || "90000");
  while (Date.now() < deadline) {
    await closePopups(page);
    const input = await getTopKeywordInput(page, 500).catch(() => null);
    const loadingVisible = await page.getByText("Loading", { exact: true }).isVisible().catch(() => false);
    if (input && !loadingVisible) return;
    await page.waitForTimeout(2500);
  }
}

async function enterTemuHome(page) {
  if (/\/temu\/home\/?/i.test(page.url())) return;
  await gotoSoft(page, START_URL);
  await closePopups(page);
}

async function getTopKeywordInput(page, timeout = 5000) {
  const selectors = [
    'input[placeholder*="中文/英文关键词"]',
    'input[placeholder*="中文"][placeholder*="英文"][placeholder*="关键词"]',
  ];
  for (const selector of selectors) {
    const loc = page.locator(selector).first();
    try {
      await loc.waitFor({ state: "visible", timeout });
      return loc;
    } catch {
      // Try next exact keyword search selector.
    }
  }
  return null;
}

async function clickTopKeywordSearchButton(page, input) {
  const inputBox = await input.boundingBox();
  const buttonBox = await page.evaluate(() => {
    const visible = (el) => {
      const style = window.getComputedStyle(el);
      const rect = el.getBoundingClientRect();
      return style.display !== "none" && style.visibility !== "hidden" && rect.width > 0 && rect.height > 0;
    };
    const keywordInput = Array.from(document.querySelectorAll("input")).find((el) => {
      const placeholder = el.getAttribute("placeholder") || "";
      return visible(el) && placeholder.includes("中文") && placeholder.includes("英文") && placeholder.includes("关键词");
    });
    if (!keywordInput) return null;
    const inputRect = keywordInput.getBoundingClientRect();
    const buttons = Array.from(document.querySelectorAll("button, [role='button'], .el-button")).filter((el) => {
      const text = (el.textContent || "").replace(/\s+/g, "");
      const rect = el.getBoundingClientRect();
      const sameRow = rect.top < inputRect.bottom + 12 && rect.bottom > inputRect.top - 12;
      return visible(el) && text === "搜索" && sameRow && rect.left >= inputRect.right - 8;
    });
    const button = buttons
      .map((el) => ({ el, rect: el.getBoundingClientRect() }))
      .sort((a, b) => a.rect.left - b.rect.left)[0]?.el;
    if (!button) return null;
    const rect = button.getBoundingClientRect();
    return { x: rect.x, y: rect.y, width: rect.width, height: rect.height };
  });
  if (buttonBox) {
    await page.mouse.move(buttonBox.x + buttonBox.width / 2, buttonBox.y + buttonBox.height / 2, { steps: 8 });
    await humanPause(120, 240);
    await page.mouse.down();
    await humanPause(60, 140);
    await page.mouse.up();
    return;
  }
  if (inputBox) {
    await input.press("Enter");
    return;
  }
  throw new Error("没有找到左上角关键词搜索按钮。");
}

async function waitForSearchResults(page) {
  const deadline = Date.now() + Number(process.env.SEARCH_READY_WAIT_MS || "120000");
  while (Date.now() < deadline) {
    await closePopups(page);
    const loadingVisible = await page.getByText("Loading", { exact: true }).isVisible().catch(() => false);
    const rowCount = await page.locator(SELECTORS.tableRows.join(",")).count().catch(() => 0);
    const noData = await page.getByText("暂无数据", { exact: true }).isVisible().catch(() => false);
    if (!loadingVisible && (rowCount > 0 || noData)) {
      await humanPause(1500, 2600);
      return;
    }
    await page.waitForTimeout(2500);
  }
  await saveDebug(page, "search-results-timeout");
}

function normalizeText(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function parseNumber(value) {
  const text = normalizeText(value).toLowerCase().replace(/,/g, "");
  const match = text.match(/-?\d+(?:\.\d+)?/);
  if (!match) return 0;
  let n = Number(match[0]);
  if (/[k千]/i.test(text)) n *= 1000;
  if (/[w万]/i.test(text)) n *= 10000;
  return Math.round(n);
}

function parseUsd(value) {
  const text = normalizeText(value).replace(/,/g, "");
  const usd = text.match(/\$\s*(\d+(?:\.\d+)?)/);
  if (usd) return Number(usd[1]);
  const nums = text.match(/\d+(?:\.\d+)?/g);
  return nums?.length ? Number(nums[0]) : "";
}

function canonicalImageKey(url) {
  try {
    const parsed = new URL(url);
    return `${parsed.origin}${parsed.pathname}`;
  } catch {
    return String(url || "").split("?")[0];
  }
}

function extractProductId(value) {
  const text = String(value || "");
  return (
    text.match(/商品ID[:：]?\s*(\d{9,})/)?.[1] ||
    text.match(/\/Main\/(\d+)/)?.[1] ||
    text.match(/\/good\/(\d+)/)?.[1] ||
    text.match(/(?:goodsId|goods_id|productId|product_id|id)["'=:\s]+(\d{9,})/)?.[1] ||
    ""
  );
}

function isProductImageUrl(url) {
  const text = String(url || "");
  if (!/^https?:\/\//i.test(text)) return false;
  if (/\.(?:mp4|webm|mov)(?:\?|$)|goods-vod|video/i.test(text)) return false;
  if (/_nuxt|favicon|logo|qrcode|avatar|douyin|weixin|user_group|supplier-public-tag|\/admin\/|\/Malls\//i.test(text)) return false;
  return (
    /yunqi-temu-img|yunqi.*temu.*img|\/Main\/|kwcdn|temu.*img|img.*temu|oss-accelerate/i.test(text) ||
    /\.(?:png|jpe?g|webp|gif)(?:\?|$)/i.test(text)
  );
}

function looksLikeFullProductImage(item) {
  const width = item?.naturalWidth || item?.width || 0;
  const height = item?.naturalHeight || item?.height || 0;
  return width >= 220 || height >= 220;
}

function normalizeCssImageUrl(value) {
  const match = String(value || "").match(/url\(["']?(.+?)["']?\)/);
  return match ? match[1] : "";
}

function imageUrlCandidates(url, baseUrl) {
  if (!url) return [];
  const absolute = new URL(url, baseUrl).toString();
  const candidates = [];
  if (absolute.includes("?imageView2")) {
    candidates.push(absolute.split("?imageView2")[0]);
  }
  try {
    const parsed = new URL(absolute);
    if (parsed.searchParams.has("x-oss-process")) {
      parsed.searchParams.delete("x-oss-process");
      candidates.push(parsed.toString());
    }
  } catch {
    // Keep the original URL below.
  }
  candidates.push(absolute);
  return Array.from(new Set(candidates));
}

function isSmallDownloadedImage(image) {
  return (image.width && image.width < 250) || (image.height && image.height < 250);
}

function imageExtension(url, contentType = "") {
  const fromType = contentType.match(/image\/(png|jpe?g|webp|gif)/i)?.[1];
  if (fromType) return fromType.replace("jpeg", "jpg").toLowerCase();
  return url.split("?")[0].match(/\.(png|jpe?g|webp|gif)$/i)?.[1]?.toLowerCase() || "jpg";
}

function readImageSize(buffer) {
  if (buffer.length >= 24 && buffer.toString("ascii", 1, 4) === "PNG") {
    return { width: buffer.readUInt32BE(16), height: buffer.readUInt32BE(20) };
  }
  if (buffer.length >= 10 && buffer.toString("ascii", 0, 4) === "RIFF" && buffer.toString("ascii", 8, 12) === "WEBP") {
    const chunk = buffer.toString("ascii", 12, 16);
    if (chunk === "VP8X" && buffer.length >= 30) {
      return {
        width: 1 + buffer.readUIntLE(24, 3),
        height: 1 + buffer.readUIntLE(27, 3),
      };
    }
    if (chunk === "VP8 " && buffer.length >= 30) {
      return {
        width: buffer.readUInt16LE(26) & 0x3fff,
        height: buffer.readUInt16LE(28) & 0x3fff,
      };
    }
    if (chunk === "VP8L" && buffer.length >= 25) {
      const b0 = buffer[21];
      const b1 = buffer[22];
      const b2 = buffer[23];
      const b3 = buffer[24];
      return {
        width: 1 + (((b1 & 0x3f) << 8) | b0),
        height: 1 + (((b3 & 0x0f) << 10) | (b2 << 2) | ((b1 & 0xc0) >> 6)),
      };
    }
  }
  if (buffer.length >= 4 && buffer[0] === 0xff && buffer[1] === 0xd8) {
    let offset = 2;
    while (offset + 9 < buffer.length) {
      if (buffer[offset] !== 0xff) {
        offset += 1;
        continue;
      }
      const marker = buffer[offset + 1];
      const length = buffer.readUInt16BE(offset + 2);
      if (marker >= 0xc0 && marker <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(marker)) {
        return { height: buffer.readUInt16BE(offset + 5), width: buffer.readUInt16BE(offset + 7) };
      }
      offset += 2 + length;
    }
  }
  return { width: 0, height: 0 };
}

function collectImageUrlsDeep(value, depth = 0, result = []) {
  if (depth > 8 || value == null) return result;
  if (typeof value === "string") {
    if (isProductImageUrl(value)) result.push(value);
    return result;
  }
  if (Array.isArray(value)) {
    for (const item of value) collectImageUrlsDeep(item, depth + 1, result);
    return result;
  }
  if (typeof value === "object") {
    for (const item of Object.values(value)) collectImageUrlsDeep(item, depth + 1, result);
  }
  return result;
}

function installResponseDebug(page) {
  if (!DEBUG_RESPONSES) return;
  page.on("response", async (response) => {
    const headers = response.headers();
    const contentType = headers["content-type"] || "";
    if (!contentType.includes("json")) return;
    const url = response.url();
    if (!/yunqi|temu|product|goods|search|list/i.test(url)) return;
    try {
      const data = await response.json();
      const imageUrls = dedupeImageUrls(collectImageUrlsDeep(data));
      if (!imageUrls.length) return;
      responseDebugRecords.push({
        url,
        imageCount: imageUrls.length,
        imageUrls: imageUrls.slice(0, 30),
        sample: JSON.stringify(data).slice(0, 5000),
      });
    } catch {
      // Ignore non-JSON or already-consumed bodies.
    }
  });
}

async function clickDailySalesDescending(page) {
  console.log("点击销量数据下的日销降序排序...");
  const clicked = await page.evaluate(() => {
    const visible = (el) => {
      const style = window.getComputedStyle(el);
      const rect = el.getBoundingClientRect();
      return style.display !== "none" && style.visibility !== "hidden" && rect.width > 0 && rect.height > 0;
    };
    const headers = Array.from(document.querySelectorAll("th, [role='columnheader'], .el-table__cell"))
      .filter(visible)
      .map((el) => ({ el, rect: el.getBoundingClientRect(), text: (el.textContent || "").replace(/\s+/g, "") }));
    const salesHeader = headers.find((h) => h.text.includes("销量数据"));
    const candidates = headers.filter((h) => h.text === "日" || h.text === "日销" || h.text === "日销量");
    const header = salesHeader
      ? candidates.find((h) => h.rect.left >= salesHeader.rect.left - 2 && h.rect.right <= salesHeader.rect.right + 2)?.el
      : candidates[0]?.el;
    if (!header) return false;

    const desc =
      header.querySelector(".sort-caret.descending, .descending, [aria-label*='降序'], [title*='降序']") ||
      header;
    desc.dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));
    desc.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    return true;
  });
  if (!clicked) {
    await saveDebug(page, "daily-sort-not-found");
    throw new Error("没有找到日销列的降序排序控件。");
  }
  await page.waitForLoadState("networkidle", { timeout: 30000 }).catch(() => {});
  await waitForSearchResults(page);
  await humanPause(2200, 3600);
}

async function getHeaderMap(page) {
  for (const selector of SELECTORS.tableHeaders) {
    const headers = await page.locator(selector).evaluateAll((ths) =>
      ths.map((th) => (th.textContent || "").replace(/\s+/g, " ").trim())
    ).catch(() => []);
    if (headers.length) {
      const find = (...needles) => headers.findIndex((h) => needles.some((n) => h.includes(n)));
      return {
        headers,
        daily: find("日销", "日销量", "日"),
        monthly: find("月销", "月销量", "月"),
        price: find("售价", "价格", "售卖"),
        listedAt: find("上架", "创建"),
        category: find("分类", "类目", "前台", "后台"),
        title: find("标题", "商品", "名称"),
      };
    }
  }
  return null;
}

async function extractRowsOnPage(page, headerMap) {
  let rowSelector = null;
  for (const candidate of SELECTORS.tableRows) {
    if (await page.locator(candidate).count().catch(() => 0)) {
      rowSelector = candidate;
      break;
    }
  }
  if (!rowSelector) return [];

  const rows = await page.locator(rowSelector).evaluateAll((rows, map) => {
    const isVisible = (el) => {
      const style = window.getComputedStyle(el);
      const rect = el.getBoundingClientRect();
      return style.display !== "none" && style.visibility !== "hidden" && rect.width > 0 && rect.height > 0;
    };
    const cellText = (cells, index) => {
      if (index == null || index < 0 || index >= cells.length) return "";
      return (cells[index].textContent || "").replace(/\s+/g, " ").trim();
    };
    const pickText = (cells, fallbackIndex) => cellText(cells, fallbackIndex);
    const extractDate = (text, label) => {
      const match = text.match(new RegExp(`${label}[:：]?\\s*(\\d{4}-\\d{2}-\\d{2}(?:\\s+\\d{2}:\\d{2}:\\d{2})?)`));
      return match ? match[1] : "";
    };
    const cleanTitle = (text) => {
      const raw = (text || "").replace(/\s+/g, " ").replace(/^播放视频\s*/, "").trim();
      return raw
        .split(/\s+(?:全托管|广告|明星卖家|同款|最新\d+天|更新[:：]|上架[:：]|美区\s|欧区\s|全球\s|\d+个\b)/)[0]
        .trim();
    };
    return rows.map((row, rowIndex) => ({ row, rowIndex })).filter(({ row }) => isVisible(row)).map(({ row, rowIndex }) => {
      const cells = Array.from(row.querySelectorAll("td, [role='cell']"));
      const rowText = (row.textContent || "").replace(/\s+/g, " ").trim();
      const img = row.querySelector("img");
      const hrefs = Array.from(row.querySelectorAll("a[href]")).map((a) => a.href || a.getAttribute("href") || "");
      const attrIds = Array.from(row.querySelectorAll("*")).flatMap((el) =>
        ["data-id", "data-goods-id", "data-product-id", "goods-id", "product-id"]
          .map((name) => el.getAttribute(name))
          .filter(Boolean)
      );
      const imgUrls = Array.from(row.querySelectorAll("img")).map((el) => el.currentSrc || el.src || el.getAttribute("src") || "");
      const productText = cellText(cells, 0);
      const titleIndex = map.title >= 0 ? map.title : 0;
      const categoryParts = [];
      const categoryText = pickText(cells, 2);
      if (categoryText) categoryParts.push(categoryText);
      const front = rowText.match(/前台分类[:：]?\s*([^后台]+?)(?:后台分类|$)/);
      const back = rowText.match(/后台分类[:：]?\s*(.+)$/);
      if (front) categoryParts.push(`前台分类: ${front[1].trim()}`);
      if (back) categoryParts.push(`后台分类: ${back[1].trim()}`);
      const listedAt = extractDate(productText, "上架") || extractDate(rowText, "上架");
      return {
        rowIndex,
        title: cleanTitle(cellText(cells, titleIndex)) || cleanTitle(rowText.slice(0, 260)),
        listedAt,
        dailySalesText: pickText(cells, 5),
        monthlySalesText: pickText(cells, 7),
        priceText: pickText(cells, 4),
        category: Array.from(new Set(categoryParts.filter(Boolean))).join(" | "),
        imageUrl: img ? img.currentSrc || img.src || img.getAttribute("src") || "" : "",
        productIdText: [...hrefs, ...attrIds, ...imgUrls, rowText].join(" "),
        rowText,
      };
    }).filter((row) => row.title && row.dailySalesText);
  }, headerMap);
  return rows.map((row) => ({ ...row, rowSelector }));
}

async function collectProductImageUrlsOnPage(page) {
  return await page.evaluate(() => {
    const visible = (el) => {
      const style = window.getComputedStyle(el);
      const rect = el.getBoundingClientRect();
      return style.display !== "none" && style.visibility !== "hidden" && rect.width > 0 && rect.height > 0;
    };
    const urls = [];
    for (const img of document.querySelectorAll("img")) {
      if (!visible(img)) continue;
      const url = img.currentSrc || img.src || img.getAttribute("src") || "";
      if (url) urls.push(url);
    }
    for (const el of document.querySelectorAll("*")) {
      if (!visible(el)) continue;
      const bg = window.getComputedStyle(el).backgroundImage;
      const match = bg && bg.match(/url\(["']?(.+?)["']?\)/);
      if (match) urls.push(match[1]);
    }
    return urls;
  }).catch(() => []);
}

async function collectHoverImageUrls(page, row, rowSelector) {
  const fallbackUrl = row.imageUrl || "";
  const rowLocator = page.locator(rowSelector || MAIN_ROW_SELECTOR).nth(row.rowIndex);
  const imageLocator = rowLocator.locator("img").first();
  if (!(await imageLocator.isVisible().catch(() => false))) {
    return fallbackUrl ? [fallbackUrl] : [];
  }

  const collectAfterHover = async (before, targetRect) => {
    await page.waitForFunction(
      ({ existing, target }) => {
        const visible = (el) => {
          const style = window.getComputedStyle(el);
          const rect = el.getBoundingClientRect();
          return style.display !== "none" && style.visibility !== "hidden" && rect.width > 0 && rect.height > 0;
        };
        const key = (url) => {
          try {
            const parsed = new URL(url);
            return `${parsed.origin}${parsed.pathname}`;
          } catch {
            return String(url || "").split("?")[0];
          }
        };
        const productUrl = (url) =>
          /^https?:\/\//i.test(url) &&
          !/supplier-public-tag|logo|avatar|qrcode|\/admin\//i.test(url) &&
          !existing.includes(key(url));
        const candidateContainers = Array.from(document.querySelectorAll("body *"))
          .filter(visible)
          .map((el) => {
            const rect = el.getBoundingClientRect();
            const urls = Array.from(el.querySelectorAll("img"))
              .map((img) => img.currentSrc || img.src || img.getAttribute("src") || "")
              .filter(productUrl);
            return { rect, count: new Set(urls.map(key)).size };
          })
          .filter(({ rect, count }) => {
            if (count < 2) return false;
            const closeVertically = rect.top <= target.bottom + 360 && rect.bottom >= target.top - 80;
            const nearHorizontal = rect.right >= target.left - 260 && rect.left <= target.right + 1500;
            return closeVertically && nearHorizontal;
          });
        return candidateContainers.length > 0;
      },
      { existing: Array.from(before), target: targetRect },
      { timeout: Number(process.env.HOVER_IMAGE_WAIT_MS || "10000") }
    ).catch(() => {});
    await humanPause(900, 1600);
    return await page.evaluate(({ existing, target }) => {
    const visible = (el) => {
      const style = window.getComputedStyle(el);
      const rect = el.getBoundingClientRect();
      return style.display !== "none" && style.visibility !== "hidden" && rect.width > 0 && rect.height > 0;
    };
    const key = (url) => {
      try {
        const parsed = new URL(url);
        return `${parsed.origin}${parsed.pathname}`;
      } catch {
        return String(url || "").split("?")[0];
      }
    };
    const existingSet = new Set(existing);
    const isCandidateUrl = (url) =>
      /^https?:\/\//i.test(url) &&
      !/supplier-public-tag|logo|avatar|qrcode|\/admin\//i.test(url) &&
      !existingSet.has(key(url));
    const collectUrl = (url, el) => {
      if (!url || !isCandidateUrl(url)) return null;
      const rect = el.getBoundingClientRect();
      return {
        url,
        x: rect.x,
        y: rect.y,
        width: rect.width,
        height: rect.height,
        naturalWidth: el.naturalWidth || 0,
        naturalHeight: el.naturalHeight || 0,
      };
    };
    const collectItems = (root) => {
      const items = [];
      for (const img of root.querySelectorAll("img")) {
        if (!visible(img)) continue;
        const item = collectUrl(img.currentSrc || img.src || img.getAttribute("src") || "", img);
        if (item) items.push(item);
      }
      for (const el of root.querySelectorAll("*")) {
        if (!visible(el)) continue;
        const bg = window.getComputedStyle(el).backgroundImage;
        const match = bg && bg.match(/url\(["']?(.+?)["']?\)/);
        if (match && isCandidateUrl(match[1])) {
          const rect = el.getBoundingClientRect();
          items.push({ url: match[1], x: rect.x, y: rect.y, width: rect.width, height: rect.height, naturalWidth: 0, naturalHeight: 0 });
        }
      }
      return items;
    };
    const containers = Array.from(document.querySelectorAll("body *"))
      .filter(visible)
      .map((el) => {
        const rect = el.getBoundingClientRect();
        const items = collectItems(el);
        const uniqueCount = new Set(items.map((item) => key(item.url))).size;
        const closeVertically = rect.top <= target.bottom + 360 && rect.bottom >= target.top - 80;
        const nearHorizontal = rect.right >= target.left - 260 && rect.left <= target.right + 1500;
        const topDistance = Math.abs(rect.top - target.bottom);
        const area = rect.width * rect.height;
        return { el, rect, items, uniqueCount, closeVertically, nearHorizontal, topDistance, area };
      })
      .filter((entry) => entry.uniqueCount >= 2 && entry.closeVertically && entry.nearHorizontal)
      .sort((a, b) => {
        if (a.area !== b.area) return a.area - b.area;
        if (b.uniqueCount !== a.uniqueCount) return b.uniqueCount - a.uniqueCount;
        return a.topDistance - b.topDistance;
      });
    return containers[0]?.items || [];
    }, { existing: Array.from(before), target: targetRect }).catch(() => []);
  };

  let bestHoverUrls = [];
  let lastTargetRect = null;
  for (let attempt = 0; attempt < 3; attempt += 1) {
    await page.mouse.move(20, 20, { steps: 6 }).catch(() => {});
    await page.waitForTimeout(500);
    const before = new Set((await collectProductImageUrlsOnPage(page)).filter(isProductImageUrl).map(canonicalImageKey));
    await imageLocator.evaluate((el) => el.scrollIntoView({ block: "center", inline: "center" })).catch(() => {});
    await page.waitForTimeout(500);
    const box = await imageLocator.boundingBox().catch(() => null);
    if (box) {
      const points = [
        [box.x + box.width / 2, box.y + box.height / 2],
        [box.x + Math.min(box.width - 4, Math.max(4, box.width * 0.25)), box.y + box.height / 2],
        [box.x + Math.min(box.width - 4, Math.max(4, box.width * 0.75)), box.y + box.height / 2],
      ];
      const [x, y] = points[Math.min(attempt, points.length - 1)];
      await page.mouse.move(x, y, { steps: 16 });
    } else {
      await imageLocator.hover({ force: true }).catch(() => {});
    }
    await page.waitForTimeout(HOVER_IMAGE_WAIT_MS);
    const targetRect = box
      ? { left: box.x, right: box.x + box.width, top: box.y, bottom: box.y + box.height }
      : await imageLocator.evaluate((el) => {
          const rect = el.getBoundingClientRect();
          return { left: rect.left, right: rect.right, top: rect.top, bottom: rect.bottom };
        }).catch(() => ({ left: 0, right: 0, top: 0, bottom: 0 }));
    lastTargetRect = targetRect;
    const after = await collectAfterHover(before, targetRect);
    const hoverUrls = after
      .filter((item) => isProductImageUrl(item.url))
      .filter((item) => !before.has(canonicalImageKey(item.url)))
      .filter(looksLikeFullProductImage)
      .sort((a, b) => (b.naturalWidth * b.naturalHeight || b.width * b.height) - (a.naturalWidth * a.naturalHeight || a.width * a.height))
      .map((item) => item.url);
    if (hoverUrls.length > bestHoverUrls.length) bestHoverUrls = hoverUrls;
    if (bestHoverUrls.length >= 2) break;
  }

  if (bestHoverUrls.length < 2 && lastTargetRect) {
    const after = await collectAfterHover(new Set(), lastTargetRect);
    const hoverUrls = after
      .filter((item) => isProductImageUrl(item.url))
      .filter(looksLikeFullProductImage)
      .sort((a, b) => (b.naturalWidth * b.naturalHeight || b.width * b.height) - (a.naturalWidth * a.naturalHeight || a.width * a.height))
      .map((item) => item.url);
    if (hoverUrls.length > bestHoverUrls.length) bestHoverUrls = hoverUrls;
  }

  await page.mouse.move(20, 20, { steps: 6 }).catch(() => {});
  return dedupeImageUrls([...bestHoverUrls, fallbackUrl]);
}

async function openProductDetailAndExtract(page, row, rowSelector) {
  await closeExternalTemuTabs(page);
  const rowLocator = page.locator(rowSelector || MAIN_ROW_SELECTOR).nth(row.rowIndex);
  const popupPromise = page.context().waitForEvent("page", { timeout: 5000 }).catch(() => null);
  let clicked = await rowLocator.evaluate((row) => {
    const visible = (el) => {
      const style = window.getComputedStyle(el);
      const rect = el.getBoundingClientRect();
      return style.display !== "none" && style.visibility !== "hidden" && rect.width > 0 && rect.height > 0;
    };
    const targets = Array.from(row.querySelectorAll("a, button, span, div")).filter((el) => {
      const text = (el.textContent || "").replace(/\s+/g, "");
      return text === "商品详情" || text.endsWith("商品详情");
    });
    const target = targets.find(visible) || targets[0];
    if (!target) return false;
    target.dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));
    target.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    return true;
  }).catch(() => false);

  if (!clicked) {
    const detailLinks = page.getByText("商品详情", { exact: true });
    const count = await detailLinks.count().catch(() => 0);
    if (count > row.rowIndex) {
      await detailLinks.nth(row.rowIndex).click({ force: true, timeout: 8000 }).catch(() => {});
      clicked = true;
    }
  }

  if (!clicked) return { productId: "", imageUrls: [] };
  const popup = await popupPromise;
  if (popup && popup !== page) {
    await popup.close({ runBeforeUnload: false }).catch(() => {});
    await page.bringToFront().catch(() => {});
  }
  await closeExternalTemuTabs(page);

  await page.waitForFunction(
    () => {
      const text = document.body?.innerText || "";
      return /商品ID[:：]?\s*\d{9,}/.test(text);
    },
    null,
    { timeout: Number(process.env.DETAIL_MODAL_WAIT_MS || "35000") }
  ).catch(() => {});
  await humanPause(1200, 2200);

  const detail = await page.evaluate(() => {
    const visible = (el) => {
      const style = window.getComputedStyle(el);
      const rect = el.getBoundingClientRect();
      return style.display !== "none" && style.visibility !== "hidden" && rect.width > 0 && rect.height > 0;
    };
    const modals = Array.from(document.querySelectorAll(".el-dialog, .el-dialog__wrapper, [role='dialog'], .el-overlay"))
      .filter(visible)
      .filter((el) => (el.textContent || "").includes("商品ID"));
    const root = modals.sort((a, b) => b.getBoundingClientRect().width * b.getBoundingClientRect().height - a.getBoundingClientRect().width * a.getBoundingClientRect().height)[0];
    if (!root) return { productId: "", imageUrls: [] };
    const text = root.textContent || "";
    const productId = text.match(/商品ID[:：]?\s*(\d{9,})/)?.[1] || "";
    const imageUrls = [];
    for (const img of root.querySelectorAll("img")) {
      const url = img.currentSrc || img.src || img.getAttribute("src") || "";
      const rect = img.getBoundingClientRect();
      const largeEnough = (img.naturalWidth || rect.width || 0) >= 220 || (img.naturalHeight || rect.height || 0) >= 220;
      if (url && largeEnough) imageUrls.push(url);
    }
    for (const el of root.querySelectorAll("*")) {
      const bg = window.getComputedStyle(el).backgroundImage;
      const match = bg && bg.match(/url\(["']?(.+?)["']?\)/);
      const rect = el.getBoundingClientRect();
      if (match && (rect.width >= 220 || rect.height >= 220)) imageUrls.push(match[1]);
    }
    return { productId, imageUrls };
  }).catch(() => ({ productId: "", imageUrls: [] }));

  await closeProductDetailModal(page);
  await closeExternalTemuTabs(page);
  return {
    productId: detail.productId || "",
    imageUrls: dedupeImageUrls(detail.imageUrls || []),
  };
}

async function closeProductDetailModal(page) {
  const closed = await page.evaluate(() => {
    const visible = (el) => {
      const style = window.getComputedStyle(el);
      const rect = el.getBoundingClientRect();
      return style.display !== "none" && style.visibility !== "hidden" && rect.width > 0 && rect.height > 0;
    };
    const candidates = Array.from(document.querySelectorAll(".el-dialog, .el-dialog__wrapper, [role='dialog'], .el-overlay"))
      .filter(visible)
      .filter((el) => (el.textContent || "").includes("商品ID"));
    const root = candidates[0];
    if (!root) return false;
    const closer =
      root.querySelector(".el-dialog__headerbtn, .el-dialog__close, [aria-label='Close']") ||
      Array.from(root.querySelectorAll("button, span, i, div")).find((el) => {
        if (!visible(el)) return false;
        const text = (el.textContent || "").trim();
        const cls = String(el.className || "").toLowerCase();
        const rect = el.getBoundingClientRect();
        return text === "×" || text === "✕" || (cls.includes("close") && rect.width <= 80 && rect.height <= 80);
      });
    if (closer) {
      closer.dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));
      closer.dispatchEvent(new MouseEvent("click", { bubbles: true }));
      return true;
    }
    root.remove();
    return true;
  }).catch(() => false);
  if (!closed) {
    await page.keyboard.press("Escape").catch(() => {});
  }
  await page.waitForTimeout(800);
}

async function fetchProductDetailImageUrls(page, productId) {
  if (!productId) return [];
  const detailUrl = new URL(`/api/proxytemu/good/${productId}`, page.url()).toString();
  const response = await page.request.get(detailUrl, {
    timeout: 45000,
    headers: {
      referer: page.url(),
      accept: "application/json, text/plain, */*",
    },
  }).catch(() => null);
  if (!response || !response.ok()) return [];

  const data = await response.json().catch(() => null);
  const direct = data?.data?.image_urls;
  const directUrls = Array.isArray(direct) ? dedupeImageUrls(direct) : [];
  if (directUrls.length) return directUrls;
  const deep = dedupeImageUrls(collectImageUrlsDeep(data));
  if (deep.length) return deep;
  const thumb = data?.data?.thumb_url || data?.data?.local_thumb_url || "";
  return dedupeImageUrls(thumb ? [thumb] : []);
}

function dedupeImageUrls(urls) {
  const seen = new Set();
  const result = [];
  for (const url of urls.filter(Boolean)) {
    if (!isProductImageUrl(url)) continue;
    const key = canonicalImageKey(url);
    if (seen.has(key)) continue;
    seen.add(key);
    result.push(url);
  }
  return result;
}

async function downloadOneBestImage(page, url, fileStem) {
  let best = null;
  for (const candidate of imageUrlCandidates(url, page.url())) {
    const response = await page.request.get(candidate, { timeout: 30000 }).catch(() => null);
    if (!response || !response.ok()) continue;
    const body = await response.body();
    const contentType = response.headers()["content-type"] || "";
    const size = readImageSize(body);
    const score = (size.width || 0) * (size.height || 0);
    if (!best || score > best.score) {
      best = { body, url: candidate, contentType, score, size };
    }
  }
  if (!best) return "";
  const ext = imageExtension(best.url, best.contentType);
  const filePath = path.join(IMAGE_DIR, `${fileStem}.${ext}`);
  await fs.writeFile(filePath, best.body);
  return {
    path: filePath,
    hash: crypto.createHash("sha1").update(best.body).digest("hex"),
    width: best.size.width || 0,
    height: best.size.height || 0,
  };
}

async function downloadImages(page, urls, productIndex) {
  const paths = [];
  const smallImages = [];
  const seen = new Set();
  for (let imageIndex = 0; imageIndex < urls.length; imageIndex += 1) {
    const fileStem = `${String(productIndex).padStart(3, "0")}-${String(imageIndex + 1).padStart(2, "0")}`;
    const image = await downloadOneBestImage(page, urls[imageIndex], fileStem);
    if (!image) continue;
    const key = image.hash || canonicalImageKey(urls[imageIndex]);
    if (seen.has(key)) {
      await fs.rm(image.path, { force: true }).catch(() => {});
      continue;
    }
    seen.add(key);
    if (isSmallDownloadedImage(image)) {
      smallImages.push(image.path);
    } else {
      paths.push(image.path);
    }
  }
  if (paths.length) {
    await Promise.all(smallImages.map((imagePath) => fs.rm(imagePath, { force: true }).catch(() => {})));
    return paths;
  }
  return smallImages;
}

async function fileSha1(filePath) {
  const body = await fs.readFile(filePath);
  return crypto.createHash("sha1").update(body).digest("hex");
}

async function findDuplicateImageProducts(products) {
  const seen = new Map();
  const duplicates = [];
  for (let productIndex = 0; productIndex < products.length; productIndex += 1) {
    const product = products[productIndex];
    for (const imagePath of product.imagePaths || []) {
      const hash = await fileSha1(imagePath).catch(() => "");
      if (!hash) continue;
      const old = seen.get(hash);
      if (old && old.productIndex !== productIndex) {
        duplicates.push({
          hash,
          firstProduct: old.productIndex + 1,
          secondProduct: productIndex + 1,
          firstTitle: products[old.productIndex]?.title || "",
          secondTitle: product.title || "",
        });
      } else if (!old) {
        seen.set(hash, { productIndex, imagePath });
      }
    }
  }
  return duplicates;
}

async function removeProductImageFiles(product) {
  for (const imagePath of product.imagePaths || []) {
    await fs.rm(imagePath, { force: true }).catch(() => {});
  }
  product.imageUrls = [];
  product.imagePaths = [];
  product.imagePath = "";
}

async function removeDuplicateImagesFromCurrentProduct(products, currentIndex) {
  const currentProduct = products[currentIndex];
  if (!currentProduct?.imagePaths?.length) return 0;

  const previousHashes = new Set();
  for (let productIndex = 0; productIndex < currentIndex; productIndex += 1) {
    for (const imagePath of products[productIndex]?.imagePaths || []) {
      const hash = await fileSha1(imagePath).catch(() => "");
      if (hash) previousHashes.add(hash);
    }
  }

  const kept = [];
  let removed = 0;
  for (const imagePath of currentProduct.imagePaths || []) {
    const hash = await fileSha1(imagePath).catch(() => "");
    if (hash && previousHashes.has(hash)) {
      await fs.rm(imagePath, { force: true }).catch(() => {});
      removed += 1;
    } else {
      kept.push(imagePath);
      if (hash) previousHashes.add(hash);
    }
  }

  currentProduct.imagePaths = kept;
  currentProduct.imagePath = kept[0] || "";
  return removed;
}

async function removeDuplicateImagesAcrossProducts(products) {
  let totalRemoved = 0;
  for (let productIndex = 0; productIndex < products.length; productIndex += 1) {
    totalRemoved += await removeDuplicateImagesFromCurrentProduct(products, productIndex);
  }
  return totalRemoved;
}

async function loadExistingProductIds() {
  if (!EXISTING_PRODUCT_IDS_PATH) return new Set();
  const body = await fs.readFile(EXISTING_PRODUCT_IDS_PATH, "utf8").catch(() => "");
  if (!body) return new Set();
  const parsed = JSON.parse(body);
  const ids = Array.isArray(parsed) ? parsed : parsed.productIds;
  return new Set((ids || []).map((id) => String(id).trim()).filter(Boolean));
}

function hasEnoughProductImages(product) {
  return (product.imagePaths || []).length >= MIN_IMAGES_PER_PRODUCT;
}

async function collectAndDownloadProductImages(page, row, product, productIndex, detailInfo = { imageUrls: [] }) {
  await removeProductImageFiles(product);
  let latestDetailInfo = detailInfo;
  for (let attempt = 1; attempt <= IMAGE_RETRY_MAX; attempt += 1) {
    product.imageUrls = await collectImageUrlsForProduct(page, row, product, latestDetailInfo);
    product.imagePaths = await downloadImages(page, product.imageUrls, productIndex + 1);
    product.imagePath = product.imagePaths?.[0] || "";
    if (hasEnoughProductImages(product)) return true;

    await removeProductImageFiles(product);
    if (attempt < IMAGE_RETRY_MAX) {
      console.log(`商品「${product.title}」图片未完全加载，等待 ${IMAGE_RETRY_WAIT_MS / 1000} 秒后第 ${attempt + 1} 次重试悬浮采图...`);
      await sleep(IMAGE_RETRY_WAIT_MS);
      latestDetailInfo = await openProductDetailAndExtract(page, row, row.rowSelector);
      product.productId = latestDetailInfo.productId || product.productId;
    }
  }
  return false;
}

async function refreshProductImages(page, product, productIndex) {
  const row = product._row;
  if (!row) return;
  let detailInfo = await openProductDetailAndExtract(page, row, row.rowSelector);
  product.productId = detailInfo.productId || product.productId;
  await collectAndDownloadProductImages(page, row, product, productIndex, detailInfo);
}

async function resolveCurrentPageImageDuplicates(page, results, currentIndex, pageNo) {
  for (let attempt = 1; attempt <= DUPLICATE_IMAGE_RETRY_MAX; attempt += 1) {
    const duplicates = await findDuplicateImageProducts(results);
    const related = duplicates.filter((dup) => (
      dup.firstProduct - 1 === currentIndex || dup.secondProduct - 1 === currentIndex
    ));
    if (!related.length) return;

    const currentProduct = results[currentIndex];
    if (!currentProduct?._row || currentProduct._pageNo !== pageNo) return;

    const relatedIndexes = Array.from(new Set(related.flatMap((dup) => [dup.firstProduct, dup.secondProduct])));
    console.log(`检测到当前商品 ${currentIndex + 1} 与商品序号 ${relatedIndexes.filter((index) => index !== currentIndex + 1).join(", ")} 图片重复，停止悬浮 ${IMAGE_RETRY_WAIT_MS / 1000} 秒后只重新采集当前商品（第 ${attempt} 次）`);
    await page.mouse.move(8, 8).catch(() => {});
    await sleep(IMAGE_RETRY_WAIT_MS);
    await refreshProductImages(page, currentProduct, currentIndex);
  }

  const remaining = (await findDuplicateImageProducts(results)).filter((dup) => (
    dup.firstProduct - 1 === currentIndex || dup.secondProduct - 1 === currentIndex
  ));
  if (remaining.length) {
    const removed = await removeDuplicateImagesFromCurrentProduct(results, currentIndex);
    console.log(`当前商品 ${currentIndex + 1} 重采 ${DUPLICATE_IMAGE_RETRY_MAX} 次后仍有重复图片，已删除当前商品中的 ${removed} 张重复图片，保留剩余图片继续记录。`);
  }
}

async function goNextPage(page) {
  const before = await page.evaluate(() => {
    const active = document.querySelector(".el-pagination .active")?.textContent?.trim() || "";
    const firstRow = document.querySelector(".el-table__body-wrapper tbody tr.el-table__row")?.textContent || "";
    return { active, firstRow: firstRow.replace(/\s+/g, " ").trim() };
  }).catch(() => ({ active: "", firstRow: "" }));

  const clicked = await page.evaluate(() => {
    const visible = (el) => {
      const style = window.getComputedStyle(el);
      const rect = el.getBoundingClientRect();
      return style.display !== "none" && style.visibility !== "hidden" && rect.width > 0 && rect.height > 0;
    };
    const next = Array.from(document.querySelectorAll(".el-pagination .btn-next, button, [role='button']")).find((el) => {
      const text = (el.textContent || "").replace(/\s+/g, "");
      const className = String(el.className || "");
      return (
        visible(el) &&
        (className.includes("btn-next") || text === "下一页") &&
        !el.disabled &&
        !className.includes("disabled")
      );
    });
    if (!next) return false;
    next.dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));
    next.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    return true;
  }).catch(() => false);

  if (!clicked) return false;

  const changed = await page.waitForFunction(
    (old) => {
      const active = document.querySelector(".el-pagination .active")?.textContent?.trim() || "";
      const firstRow = (document.querySelector(".el-table__body-wrapper tbody tr.el-table__row")?.textContent || "")
        .replace(/\s+/g, " ")
        .trim();
      const loading = Array.from(document.querySelectorAll("body *")).some((el) => (el.textContent || "").trim() === "Loading");
      return !loading && (active !== old.active || firstRow !== old.firstRow);
    },
    before,
    { timeout: 30000 }
  ).catch(() => null);

  await humanPause(2200, 3800);
  return Boolean(changed);
}

function dedupeKey(item) {
  return `${item.title}|${item.listedAt}|${item.dailySales}|${item.monthlySales}`;
}

async function collectImageUrlsForProduct(page, row, item, detailInfo = { imageUrls: [] }) {
  const hoverUrls = await collectHoverImageUrls(page, row, row.rowSelector);
  const detailUrls = await fetchProductDetailImageUrls(page, item.productId);
  let imageUrls = dedupeImageUrls([...hoverUrls, ...(detailInfo.imageUrls || []), ...detailUrls, item.imageUrl]);

  if (imageUrls.length < MIN_IMAGES_PER_PRODUCT) {
    await humanPause(1600, 2600);
    const retryHoverUrls = await collectHoverImageUrls(page, row, row.rowSelector);
    const retryDetailUrls = await fetchProductDetailImageUrls(page, item.productId);
    imageUrls = dedupeImageUrls([...retryHoverUrls, ...retryDetailUrls, ...imageUrls]);
  }

  return imageUrls;
}

async function scrapeProducts(page, existingProductIds = new Set(), productIndexOffset = 0) {
  await clickDailySalesDescending(page);
  const headerMap = await getHeaderMap(page) || {};

  const results = [];
  const seen = new Set();
  const debugRows = [];
  for (let pageNo = 1; pageNo <= MAX_PAGES; pageNo += 1) {
    const rows = await extractRowsOnPage(page, headerMap);
    if (!rows.length) break;

    let pageHasDailyCandidate = false;
    let pageQualifiedCount = 0;
    for (const row of rows) {
      const dailySales = parseNumber(row.dailySalesText);
      const monthlySales = parseNumber(row.monthlySalesText);
      debugRows.push({
        pageNo,
        title: row.title,
        productId: extractProductId(`${row.productIdText || ""} ${row.imageUrl || ""}`),
        dailySalesText: row.dailySalesText,
        dailySales,
        monthlySalesText: row.monthlySalesText,
        monthlySales,
        priceText: row.priceText,
        listedAt: row.listedAt,
      });
      if (dailySales < DAILY_MIN) {
        continue;
      }
      pageHasDailyCandidate = true;
      if (monthlySales < MONTHLY_MIN) continue;
      pageQualifiedCount += 1;

      const item = {
        listedAt: row.listedAt,
        dailySales,
        monthlySales,
        priceUsd: parseUsd(row.priceText),
        title: row.title,
        productId: extractProductId(`${row.productIdText || ""} ${row.imageUrl || ""}`),
        imageUrl: row.imageUrl,
        _pageNo: pageNo,
        _row: row,
      };
      const key = dedupeKey(item);
      if (!seen.has(key)) {
        seen.add(key);
        const detailInfo = await openProductDetailAndExtract(page, row, row.rowSelector);
        item.productId = detailInfo.productId || item.productId;
        if (!item.productId) {
          const retryDetailInfo = await openProductDetailAndExtract(page, row, row.rowSelector);
          item.productId = retryDetailInfo.productId || item.productId;
          detailInfo.imageUrls = dedupeImageUrls([...(detailInfo.imageUrls || []), ...(retryDetailInfo.imageUrls || [])]);
        }
        if (!item.productId) {
          await saveDebug(page, `product-id-not-found-${results.length + 1}`);
          throw new Error(`商品「${item.title}」没有从商品详情里采集到商品ID。`);
        }
        if (existingProductIds.has(String(item.productId))) {
          console.log(`商品ID ${item.productId} 已存在于输入表格，跳过采图和记录。`);
          continue;
        }
        const gotImages = await collectAndDownloadProductImages(page, row, item, productIndexOffset + results.length, detailInfo);
        if (!gotImages) {
          await saveDebug(page, `not-enough-images-${results.length + 1}`);
          throw new Error(`商品「${item.title}」只采集到 ${item.imagePaths.length} 张参考图，低于要求的 ${MIN_IMAGES_PER_PRODUCT} 张。`);
        }
        results.push(item);
        await resolveCurrentPageImageDuplicates(page, results, results.length - 1, pageNo);
      }
    }

    console.log(`第 ${pageNo} 页：发现 ${pageQualifiedCount} 个符合条件商品。`);
    if (!pageHasDailyCandidate) break;
    if (!(await goNextPage(page))) break;
  }
  await fs.writeFile(path.join(OUTPUT_DIR, "debug-rows.json"), JSON.stringify(debugRows, null, 2));
  const duplicates = await findDuplicateImageProducts(results);
  if (duplicates.length) {
    await fs.writeFile(path.join(OUTPUT_DIR, "debug-cross-product-duplicates.json"), JSON.stringify(duplicates, null, 2));
    const removed = await removeDuplicateImagesAcrossProducts(results);
    const remainingDuplicates = await findDuplicateImageProducts(results);
    if (remainingDuplicates.length) {
      await fs.writeFile(path.join(OUTPUT_DIR, "debug-cross-product-duplicates-after-cleanup.json"), JSON.stringify(remainingDuplicates, null, 2));
    }
    console.log(`最终检测到 ${duplicates.length} 处跨商品重复图片，已删除 ${removed} 张后面商品里的重复图片，剩余重复 ${remainingDuplicates.length} 处。`);
  }
  console.log(`抓取完成，符合日销>=${DAILY_MIN} 且月销>=${MONTHLY_MIN} 的商品数：${results.length}`);
  return results.map(({ _pageNo, _row, ...item }) => item);
}

function runPython(args) {
  return new Promise((resolve, reject) => {
    const child = spawn(PYTHON, args, { stdio: "inherit", cwd: ROOT });
    child.on("exit", (code) => {
      if (code === 0) resolve();
      else reject(new Error(`Python 写表失败，退出码 ${code}`));
    });
  });
}

async function main() {
  await fs.rm(IMAGE_DIR, { recursive: true, force: true });
  await fs.mkdir(IMAGE_DIR, { recursive: true });
  let browser = null;
  let context;
  const executablePathOption = CHROME_PATH ? { executablePath: CHROME_PATH } : {};
  if (USE_PERSISTENT_CONTEXT) {
    context = await chromium.launchPersistentContext(USER_DATA_DIR, {
      ...executablePathOption,
      headless: HEADLESS,
      args: CHROME_ARGS,
      slowMo: HEADLESS ? 0 : 80,
      viewport: { width: 1440, height: 1000 },
      userAgent:
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    });
  } else {
    browser = await chromium.launch({
      ...executablePathOption,
      headless: HEADLESS,
      args: CHROME_ARGS,
      slowMo: HEADLESS ? 0 : 80,
    });
    context = await browser.newContext({
      viewport: { width: 1440, height: 1000 },
      userAgent:
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    });
  }
  const existingYunqiPage = context.pages().find((item) => /yunqishuju\.com/i.test(item.url()));
  const page = existingYunqiPage || await context.newPage();
  await closeExternalTemuTabs(page);
  installResponseDebug(page);
  let completed = false;

  try {
    await login(page);
    const existingProductIds = await loadExistingProductIds();
    if (existingProductIds.size) {
      console.log(`增量模式：已从输入表格读取 ${existingProductIds.size} 个商品ID，旧商品将跳过采图。`);
    }
    const products = [];
    for (let keywordIndex = 0; keywordIndex < KEYWORDS.length; keywordIndex += 1) {
      const keyword = KEYWORDS[keywordIndex];
      console.log(`开始处理搜索词：${keyword}`);
      await searchKeyword(page, keyword, Boolean(CATEGORY_PARENT));
      const keywordProducts = await scrapeProducts(page, existingProductIds, products.length);
      for (const product of keywordProducts) {
        const productId = String(product.productId || "").trim();
        if (productId && existingProductIds.has(productId)) continue;
        products.push(product);
        if (productId) existingProductIds.add(productId);
      }
      console.log(`搜索词「${keyword}」完成：新增 ${keywordProducts.length} 个商品，当前合计 ${products.length} 个商品。`);
    }
    if (DEBUG_RESPONSES) {
      await fs.writeFile(path.join(OUTPUT_DIR, "debug-responses.json"), JSON.stringify(responseDebugRecords, null, 2));
    }
    await fs.writeFile(DATA_PATH, JSON.stringify(products, null, 2));
    const outputPath = path.join(OUTPUT_DIR, `选品表格-pop-up-greeting-card-${new Date().toISOString().slice(0, 10)}.xlsx`);
    await runPython(["fill_yunqi_excel.py", DATA_PATH, TEMPLATE_PATH, outputPath]);
    console.log(`完成：找到 ${products.length} 个符合条件的商品`);
    console.log(`Excel：${outputPath}`);
    completed = true;
  } finally {
    const keepOnError = process.env.KEEP_BROWSER_ON_ERROR === "1";
    if (process.env.KEEP_BROWSER !== "1" && (completed || !keepOnError)) {
      await context.close();
      if (browser) await browser.close();
    }
  }
}

main().catch((error) => {
  console.error(error.stack || error.message);
  process.exitCode = 1;
});
