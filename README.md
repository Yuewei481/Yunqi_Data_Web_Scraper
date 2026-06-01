# Yunqi_Data_Web_Scraper

云启数据爬虫。项目包含三个脚本，用于从云启数据 Temu 页面采集满足条件的商品，并生成、对比或更新 Excel 选品表格。

## 功能

- 脚本一：生成当天完整选品表格。
- 脚本二：生成当天表格，并与昨天表格对比，输出新增/移除商品。
- 脚本三：输入一个已有表格，只把当天新增商品追加到这个表格底部。

默认筛选条件：

- 日销 `>= 30`
- 月销 `>= 1000`
- 关键词 `pop up greeting card`

这些都可以在 `.env` 中修改。

## 一、Mac 安装

### 1. 安装基础软件

需要先安装：

- Git
- Node.js 22 LTS 或更高版本
- Python 3
- Google Chrome，可选；默认使用 Playwright 自带的 Chromium

如果使用 Homebrew：

```bash
brew install git node python
```

如果你想改用自己电脑上的 Google Chrome，再安装 Chrome：

```bash
brew install --cask google-chrome
```

检查版本：

```bash
git --version
node -v
python3 --version
```

### 2. 下载项目

```bash
cd ~/Documents
mkdir -p project
cd project
git clone https://github.com/Yuewei481/Yunqi_Data_Web_Scraper.git
cd Yunqi_Data_Web_Scraper
```

### 3. 创建虚拟环境

Mac 使用：

```bash
python3 -m venv venv
source venv/bin/activate
```

### 4. 安装依赖

这里的 `npx playwright install chromium` 会安装 Playwright 自带的 Chromium。默认情况下，脚本会使用这个浏览器，不需要额外配置 Chrome 路径。

```bash
npm install
pip install -r requirements.txt
npx playwright install chromium
```

### 5. 配置 `.env`

```bash
cp .env.example .env
nano .env
```

Mac 示例：

```bash
YUNQI_USERNAME="你的云启账号"
YUNQI_PASSWORD="你的云启密码"
YUNQI_KEYWORD="pop up greeting card"
DAILY_MIN=30
MONTHLY_MIN=1000
EXCEL_TEMPLATE=templates/选品表格-模板.xlsx
HEADLESS=0
PYTHON=python3
```

默认不用填写 `CHROME_PATH`。如果 Playwright 自带 Chromium 无法正常运行，或者你想强制使用本机 Chrome，再添加：

```bash
CHROME_PATH="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
```

## 二、Windows 安装

### 1. 安装基础软件

Windows 需要先安装：

- Git for Windows
- Node.js 22 LTS 或更高版本
- Python 3
- Google Chrome，可选；默认使用 Playwright 自带的 Chromium

Chrome 不是默认必需项。只有当 Playwright 自带 Chromium 无法正常运行，或者你想强制使用本机 Chrome 时，才需要安装并配置 `CHROME_PATH`。

安装 Python 时建议勾选：

```text
Add python.exe to PATH
```

### 2. 打开 Git Bash

本项目的启动脚本是 `.sh`，Windows 推荐使用 **Git Bash** 运行。

### 3. 下载项目

例如放在 `D:\Yunqi_Data_Web_Scraper`：

```bash
cd /d
git clone https://github.com/Yuewei481/Yunqi_Data_Web_Scraper.git
cd Yunqi_Data_Web_Scraper
```

如果放在 Documents：

```bash
cd ~/Documents
git clone https://github.com/Yuewei481/Yunqi_Data_Web_Scraper.git
cd Yunqi_Data_Web_Scraper
```

### 4. 创建虚拟环境

Windows Git Bash 使用：

```bash
python -m venv venv
source venv/Scripts/activate
```

如果使用 PowerShell 创建虚拟环境，则使用：

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
```

脚本运行仍建议回到 Git Bash。

### 5. 安装依赖

这里的 `npx playwright install chromium` 会安装 Playwright 自带的 Chromium。默认情况下，脚本会使用这个浏览器，不需要额外配置 Chrome 路径。

```bash
npm install
pip install -r requirements.txt
npx playwright install chromium
```

### 6. 配置 `.env`

```bash
cp .env.example .env
notepad .env
```

Windows 示例：

```bash
YUNQI_USERNAME="你的云启账号"
YUNQI_PASSWORD="你的云启密码"
YUNQI_KEYWORD="pop up greeting card"
DAILY_MIN=30
MONTHLY_MIN=1000
EXCEL_TEMPLATE=templates/选品表格-模板.xlsx
HEADLESS=0
PYTHON=python
```

默认不用填写 `CHROME_PATH`。如果 Playwright 自带 Chromium 无法正常运行，或者你想强制使用本机 Chrome，再添加：

```bash
CHROME_PATH="C:/Program Files/Google/Chrome/Application/chrome.exe"
```

如果 Chrome 安装在另一个目录，可以改成：

```bash
CHROME_PATH="C:/Program Files (x86)/Google/Chrome/Application/chrome.exe"
```

## 三、运行脚本

下面的路径请换成你本机项目的绝对路径。

### Mac 运行脚本一

```bash
cd /Users/你的用户名/Documents/project/Yunqi_Data_Web_Scraper/script1_yunqi_scraper
./run_yunqi_scraper.sh
```

### Windows 运行脚本一

Git Bash 中可以使用：

```bash
cd /d/Yunqi_Data_Web_Scraper/script1_yunqi_scraper
./run_yunqi_scraper.sh
```

或者：

```bash
cd /c/Users/你的用户名/Documents/Yunqi_Data_Web_Scraper/script1_yunqi_scraper
./run_yunqi_scraper.sh
```

脚本一输出完整选品表格。

默认输出位置：

```text
script1_yunqi_scraper/outputs/yunqi-pop-up-greeting-card/选品表格-pop-up-greeting-card-YYYY-MM-DD.xlsx
```

### Mac 运行脚本二

```bash
cd /Users/你的用户名/Documents/project/Yunqi_Data_Web_Scraper/script2_daily_compare
./run_yunqi_daily_compare.sh "/Users/你的用户名/Desktop/昨日表格.xlsx"
```

### Windows 运行脚本二

```bash
cd /d/Yunqi_Data_Web_Scraper/script2_daily_compare
./run_yunqi_daily_compare.sh "C:/Users/你的用户名/Desktop/昨日表格.xlsx"
```

也可以指定差异表输出路径：

```bash
./run_yunqi_daily_compare.sh "C:/Users/你的用户名/Desktop/昨日表格.xlsx" "C:/Users/你的用户名/Desktop/差异表.xlsx"
```

默认输出位置：

```text
script2_daily_compare/outputs/yunqi-pop-up-greeting-card/选品表格-pop-up-greeting-card-差异-YYYY-MM-DD.xlsx
```

### Mac 运行脚本三

```bash
cd /Users/你的用户名/Documents/project/Yunqi_Data_Web_Scraper/script3_append_new_products
./run_yunqi_append_new_products.sh "/Users/你的用户名/Desktop/已有表格.xlsx"
```

### Windows 运行脚本三

```bash
cd /d/Yunqi_Data_Web_Scraper/script3_append_new_products
./run_yunqi_append_new_products.sh "C:/Users/你的用户名/Desktop/已有表格.xlsx"
```

脚本三不会生成新的最终表格，而是直接修改输入的原有表格：

- 已存在的商品 ID：跳过
- 今日新增商品 ID：采集完整信息并追加到表格底部
- 今日消失的商品 ID：不删除，继续保留

## 四、修改输出位置

默认情况下，脚本会输出到当前脚本文件夹里的：

```text
outputs/yunqi-pop-up-greeting-card/
```

### Mac 输出目录示例

在 `.env` 里添加：

```bash
OUTPUT_DIR="/Users/你的用户名/Desktop/yunqi-outputs"
```

### Windows 输出目录示例

在 `.env` 里添加：

```bash
OUTPUT_DIR="C:/Users/你的用户名/Desktop/yunqi-outputs"
```

也可以只对某一次运行临时指定：

Mac：

```bash
OUTPUT_DIR="/Users/你的用户名/Desktop/yunqi-outputs" ./run_yunqi_scraper.sh
```

Windows Git Bash：

```bash
OUTPUT_DIR="C:/Users/你的用户名/Desktop/yunqi-outputs" ./run_yunqi_scraper.sh
```

脚本一会把完整选品表格输出到这个目录。脚本二会把今日临时表格和差异表输出到这个目录。脚本三会把临时今日表格放到这个目录，但最终仍然直接修改你输入的 Excel 表格本身。

## 五、常见注意事项

- `.env` 中只要值里有空格，就必须使用英文双引号，例如 `YUNQI_KEYWORD="pop up greeting card"`。
- Windows 路径推荐使用 `/d/...` 或 `C:/...`，不要混用中文引号。
- 第一次运行会打开浏览器并登录云启数据，网站响应较慢时请等待。
- 如果登录失败，多数是网络或网站超时，可以重新运行。

## 六、Codex 自动化

本项目也可以链接到 Codex 自动化里，让 Codex 每天自动运行脚本来爬取数据。
