# 云启数据爬虫

云启数据爬虫用于从云启数据 Temu 页面采集满足条件的商品，并生成、对比或更新 Excel 选品表格。

项目里有三个脚本：

- 脚本一：生成一张新的今日选品表格。
- 脚本二：先生成今日选品表格，再和输入的昨日表格对比，输出新增/移除商品差异表。
- 脚本三：输入一张已有选品表格，只把今日新增商品追加到这张表格底部，不删除原有商品。

默认筛选条件：

- 日销 `>= 30`
- 月销 `>= 1000`
- 商品图片悬浮后等待 `3000ms` 再读取
- 支持多个关键词，使用英文逗号分隔
- 支持 Temu 大分类/小分类筛选；配置分类后，每个关键词搜索前都会重新应用分类

这些条件都在项目根目录的 `.env` 文件里修改。

## 一、项目结构

```text
Yunqi_Data_Web_Scraper/
├── .env.example
├── README.md
├── requirements.txt
├── package.json
├── templates/
│   └── 选品表格-模板.xlsx
├── script1_yunqi_scraper/
│   └── run_yunqi_scraper.sh
├── script2_daily_compare/
│   └── run_yunqi_daily_compare.sh
└── script3_append_new_products/
    └── run_yunqi_append_new_products.sh
```

注意：`.env` 必须放在项目根目录，也就是和 `README.md` 同一层。

## 二、Mac 安装教程

下面步骤只需要在第一次安装时做一次。

### 1. 安装基础软件

Mac 需要：

- Git
- Node.js 22 LTS 或更高版本
- Python 3
- Playwright 自带 Chromium，安装依赖时会下载
- Google Chrome 可选，默认不需要

如果你使用 Homebrew，可以运行：

```bash
brew install git node python
```

检查是否安装成功：

```bash
git --version
node -v
python3 --version
```

如果 `node -v` 显示的是 `v18` 或更低版本，建议先升级到 Node.js 22 LTS 或更高版本。

### 2. 下载项目

下面示例把项目放在：

```text
/Users/你的用户名/Documents/project/Yunqi_Data_Web_Scraper
```

运行：

```bash
cd ~/Documents
mkdir -p project
cd project
git clone https://github.com/Yuewei481/Yunqi_Data_Web_Scraper.git
cd Yunqi_Data_Web_Scraper
```

如果你已经下载过项目，只需要进入项目根目录：

```bash
cd /Users/你的用户名/Documents/project/Yunqi_Data_Web_Scraper
```

### 3. 创建并启动 Python 虚拟环境

必须在项目根目录运行：

```bash
python3 -m venv venv
source venv/bin/activate
```

启动成功后，终端前面通常会出现：

```text
(venv)
```

如果没有看到 `(venv)`，说明虚拟环境没有启动成功，后面运行脚本可能会出现 `No module named openpyxl` 之类的错误。

### 4. 安装依赖

确认终端前面已经有 `(venv)`，然后运行：

```bash
npm install
pip install -r requirements.txt
npx playwright install chromium
```

`npx playwright install chromium` 会安装 Playwright 自带的 Chromium。默认情况下脚本使用这个浏览器，不需要另外配置 Chrome 路径。

### 5. 创建并修改 `.env`

仍然在项目根目录运行：

```bash
cp .env.example .env
open -e .env
```

如果 `open -e .env` 没打开，也可以用：

```bash
nano .env
```

Mac 推荐配置：

```bash
YUNQI_USERNAME="你的云启账号"
YUNQI_PASSWORD="你的云启密码"

YUNQI_KEYWORD="pop up greeting card"
DAILY_MIN=30
MONTHLY_MIN=1000
HOVER_IMAGE_WAIT_MS=3000

YUNQI_CATEGORY_PARENT=""
YUNQI_CATEGORY_CHILDREN=""

EXCEL_TEMPLATE=templates/选品表格-模板.xlsx
OUTPUT_DIR="/Users/你的用户名/Desktop/excel_output"

HEADLESS=0
PYTHON=python3
```

修改完 `.env` 后记得保存。

## 三、Windows 安装教程

下面步骤只需要在第一次安装时做一次。

Windows 推荐全程使用 **Git Bash** 操作本项目。原因是项目启动脚本是 `.sh` 文件，CMD 不能直接运行 `.sh`。

### 1. 安装基础软件

Windows 需要：

- Git for Windows，安装后会有 Git Bash
- Node.js 22 LTS 或更高版本
- Python 3
- Playwright 自带 Chromium，安装依赖时会下载
- Google Chrome 可选，默认不需要

安装 Python 时，一定建议勾选：

```text
Add python.exe to PATH
```

### 2. 打开 Git Bash

在 Windows 搜索框里搜索：

```text
Git Bash
```

然后打开它。

后面安装环境、启动环境、运行脚本，都建议使用这个 Git Bash。

### 3. 下载项目

下面示例把项目放在桌面：

```text
C:\Users\你的用户名\Desktop\Yunqi_Data_Web_Scraper
```

在 Git Bash 里运行：

```bash
cd ~/Desktop
git clone https://github.com/Yuewei481/Yunqi_Data_Web_Scraper.git
cd Yunqi_Data_Web_Scraper
```

如果你已经下载过项目，只需要进入项目根目录。比如项目在桌面：

```bash
cd /c/Users/你的用户名/Desktop/Yunqi_Data_Web_Scraper
```

如果你不知道路径怎么写，可以输入 `cd `，注意 `cd` 后面有一个空格，然后把项目文件夹拖进 Git Bash，再按回车。

### 4. 创建并启动 Python 虚拟环境

必须在项目根目录运行：

```bash
python -m venv venv
. venv/Scripts/activate
```

启动成功后，Git Bash 前面通常会出现：

```text
(venv)
```

如果没有看到 `(venv)`，说明虚拟环境没有启动成功。后面运行脚本可能会出现：

```text
ModuleNotFoundError: No module named 'openpyxl'
```

或者：

```text
Error: Python 写表失败
```

Windows 重点说明：

- 如果你用 Git Bash 创建了 `venv`，以后运行脚本前也要在 Git Bash 里用 `. venv/Scripts/activate` 启动这个环境。
- 不要在 CMD 里运行 `./run_yunqi_scraper.sh`，CMD 不能直接运行 `.sh` 文件。
- PowerShell/CMD 可以创建虚拟环境，但本项目实际运行脚本仍推荐回到 Git Bash。

### 5. 安装依赖

确认 Git Bash 前面已经有 `(venv)`，然后运行：

```bash
npm install
pip install -r requirements.txt
npx playwright install chromium
```

`npx playwright install chromium` 会安装 Playwright 自带的 Chromium。默认情况下脚本使用这个浏览器，不需要另外配置 Chrome 路径。

如果你在同一台电脑、同一个 Windows 用户下已经安装过 Playwright Chromium，可以先不重复安装。直接运行脚本，如果报找不到浏览器，再运行：

```bash
npx playwright install chromium
```

### 6. 创建并修改 `.env`

仍然在项目根目录运行：

```bash
cp .env.example .env
notepad .env
```

Windows 推荐配置：

```bash
YUNQI_USERNAME="你的云启账号"
YUNQI_PASSWORD="你的云启密码"

YUNQI_KEYWORD="pop up greeting card"
DAILY_MIN=30
MONTHLY_MIN=1000
HOVER_IMAGE_WAIT_MS=3000

YUNQI_CATEGORY_PARENT=""
YUNQI_CATEGORY_CHILDREN=""

EXCEL_TEMPLATE=templates/选品表格-模板.xlsx
OUTPUT_DIR="C:/Users/你的用户名/Desktop/excel_output"

HEADLESS=0
PYTHON=python
```

修改完 `.env` 后按 `Ctrl + S` 保存。

## 四、`.env` 配置说明

`.env` 里不要在等号两边加空格。

正确：

```bash
YUNQI_KEYWORD="keyboard"
```

错误：

```bash
YUNQI_KEYWORD = "keyboard"
```

常用配置：

```bash
YUNQI_USERNAME="你的云启账号"
YUNQI_PASSWORD="你的云启密码"
```

云启数据账号和密码。

```bash
YUNQI_KEYWORD="pop up greeting card, keyboard"
```

搜索关键词。多个关键词用英文逗号分隔。脚本会依次搜索，并把符合条件的商品写进同一张表格。

```bash
DAILY_MIN=30
MONTHLY_MIN=1000
```

筛选日销和月销。表示只记录日销大于等于 30、月销大于等于 1000 的商品。

```bash
HOVER_IMAGE_WAIT_MS=3000
```

鼠标悬浮到商品图后等待多久再读取图片。`3000` 表示等待 3 秒。

```bash
YUNQI_CATEGORY_PARENT=""
YUNQI_CATEGORY_CHILDREN=""
```

分类筛选。两个都留空时，不限制分类。

只限制大分类：

```bash
YUNQI_CATEGORY_PARENT="Electronics"
YUNQI_CATEGORY_CHILDREN=""
```

限制大分类下的小分类：

```bash
YUNQI_CATEGORY_PARENT="Electronics"
YUNQI_CATEGORY_CHILDREN="Keyboards, Mice & Accessories"
```

多个小分类用英文逗号分隔：

```bash
YUNQI_CATEGORY_PARENT="Appliances"
YUNQI_CATEGORY_CHILDREN="Air Quality, Ice Maker"
```

```bash
EXCEL_TEMPLATE=templates/选品表格-模板.xlsx
```

Excel 模板路径。一般不要改。

```bash
OUTPUT_DIR="C:/Users/你的用户名/Desktop/excel_output"
```

输出目录。建议使用绝对路径。Mac 可以写 `/Users/...`，Windows 推荐写 `C:/Users/...`。

```bash
HEADLESS=0
```

是否显示浏览器。`0` 表示显示浏览器，`1` 表示无头模式。建议使用 `0`。

```bash
PYTHON=python
```

Python 命令。Windows 通常写 `python`，Mac 通常写 `python3`。

```bash
CHROME_PATH="C:/Program Files/Google/Chrome/Application/chrome.exe"
```

默认不要写这一行。只有 Playwright 自带 Chromium 不能运行，或者你想强制使用本机 Chrome 时才填写。

## 五、以后每次运行前都要做什么

每次重新打开终端，都需要先进入项目根目录并启动虚拟环境。

### Mac 每次运行前

```bash
cd /Users/你的用户名/Documents/project/Yunqi_Data_Web_Scraper
source venv/bin/activate
```

看到 `(venv)` 后，再进入脚本文件夹运行脚本。

### Windows 每次运行前

打开 Git Bash，然后运行：

```bash
cd /c/Users/你的用户名/Desktop/Yunqi_Data_Web_Scraper
. venv/Scripts/activate
```

看到 `(venv)` 后，再进入脚本文件夹运行脚本。

如果项目不在桌面，可以输入 `cd `，把项目文件夹拖入 Git Bash，再按回车。

## 六、脚本一：生成新的今日选品表格

用途：创建一张新的选品表格，内容是今天搜索后符合条件的商品。

### Mac 运行脚本一

```bash
cd /Users/你的用户名/Documents/project/Yunqi_Data_Web_Scraper
source venv/bin/activate
cd script1_yunqi_scraper
./run_yunqi_scraper.sh
```

### Windows 运行脚本一

在 Git Bash 里运行：

```bash
cd /c/Users/你的用户名/Desktop/Yunqi_Data_Web_Scraper
. venv/Scripts/activate
cd script1_yunqi_scraper
./run_yunqi_scraper.sh
```

脚本一输出完整选品表格。

如果 `.env` 设置了：

```bash
OUTPUT_DIR="C:/Users/你的用户名/Desktop/excel_output"
```

那么输出 Excel 会在这个文件夹里。

如果没有设置 `OUTPUT_DIR`，默认输出在：

```text
script1_yunqi_scraper/outputs/yunqi-pop-up-greeting-card/
```

## 七、脚本二：生成今日表格并和昨日表格对比

用途：输入一张昨日表格，脚本会先跑出今日表格，然后比较商品 ID，输出差异表。

差异表会标记：

- 今日新增：今天满足条件，昨日不在表格里
- 今日移除：昨日满足条件，今天不在表格里

### Mac 运行脚本二

```bash
cd /Users/你的用户名/Documents/project/Yunqi_Data_Web_Scraper
source venv/bin/activate
cd script2_daily_compare
./run_yunqi_daily_compare.sh "/Users/你的用户名/Desktop/昨日表格.xlsx"
```

指定差异表输出位置：

```bash
./run_yunqi_daily_compare.sh "/Users/你的用户名/Desktop/昨日表格.xlsx" "/Users/你的用户名/Desktop/差异表.xlsx"
```

### Windows 运行脚本二

在 Git Bash 里运行：

```bash
cd /c/Users/你的用户名/Desktop/Yunqi_Data_Web_Scraper
. venv/Scripts/activate
cd script2_daily_compare
./run_yunqi_daily_compare.sh "C:/Users/你的用户名/Desktop/昨日表格.xlsx"
```

指定差异表输出位置：

```bash
./run_yunqi_daily_compare.sh "C:/Users/你的用户名/Desktop/昨日表格.xlsx" "C:/Users/你的用户名/Desktop/差异表.xlsx"
```

默认输出在：

```text
script2_daily_compare/outputs/yunqi-pop-up-greeting-card/
```

如果 `.env` 设置了 `OUTPUT_DIR`，则输出到 `OUTPUT_DIR`。

## 八、脚本三：把今日新增商品追加到已有表格

用途：输入一张已有选品表格，脚本会搜索今天符合条件的商品。如果商品 ID 已经在输入表格里，就跳过；如果商品 ID 不在输入表格里，就追加到表格底部。

脚本三不会删除旧商品。

脚本三最终修改的是你输入的那张 Excel 表格本身。

### Mac 运行脚本三

```bash
cd /Users/你的用户名/Documents/project/Yunqi_Data_Web_Scraper
source venv/bin/activate
cd script3_append_new_products
./run_yunqi_append_new_products.sh "/Users/你的用户名/Desktop/已有表格.xlsx"
```

### Windows 运行脚本三

在 Git Bash 里运行：

```bash
cd /c/Users/你的用户名/Desktop/Yunqi_Data_Web_Scraper
. venv/Scripts/activate
cd script3_append_new_products
./run_yunqi_append_new_products.sh "C:/Users/你的用户名/Desktop/已有表格.xlsx"
```

运行完成后，直接打开你输入的那张 Excel 表格查看结果。新增商品会追加在表格最下面。

## 九、运行时注意事项

- 运行脚本时不要关闭脚本打开的浏览器。
- 浏览器可以放到后面或拖到屏幕边上，但不要最小化。
- 云启数据登录和 Temu 页面加载可能较慢，请等待。
- 如果卡在登录界面，可以手动登录。
- 如果网络错误或网站超时，可以按 `Ctrl + C` 停止脚本，然后重新运行。
- `.env` 修改后必须保存，再重新运行脚本才会生效。
- Windows 路径推荐使用 `C:/Users/...`，不要使用中文引号。
- 关键词、分类名、路径里的逗号和引号都建议使用英文输入法。

## 十、常见错误

### 1. `No module named 'openpyxl'`

原因：没有启动虚拟环境，或者没有安装 Python 依赖。

Mac：

```bash
cd /Users/你的用户名/Documents/project/Yunqi_Data_Web_Scraper
source venv/bin/activate
pip install -r requirements.txt
```

Windows Git Bash：

```bash
cd /c/Users/你的用户名/Desktop/Yunqi_Data_Web_Scraper
. venv/Scripts/activate
pip install -r requirements.txt
```

### 2. `Failed to launch chromium`

原因：Playwright 自带 Chromium 没安装，或者浏览器缓存损坏。

运行：

```bash
npx playwright install chromium
```
