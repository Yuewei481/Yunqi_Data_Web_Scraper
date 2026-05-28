# Yunqi_Data_Web_Scraper

本项目包含三个脚本. 三个脚本用于从云启数据 Temu 页面采集满足条件的商品，并生成或更新 Excel 选品表格。

## 功能

- 脚本一：生成当天完整选品表格。
- 脚本二：生成当天表格，并与昨天表格对比，输出新增/移除商品。
- 脚本三：输入一个已有表格，只把当天新增商品追加到这个表格底部。

默认筛选条件：

- 日销 `>= 30`
- 月销 `>= 1000`
- 关键词 `pop up greeting card`

这些都可以在 `.env` 中修改。


## 安装

请先安装以下软件：

* Node.js
* Python 3
* Google Chrome

然后克隆项目并进入项目目录：

```bash
git clone <你的仓库地址>
cd Yunqi-Data-Scraper
```

创建 Python 独立虚拟环境：

```bash
python3 -m venv venv
source venv/bin/activate
```

安装项目依赖：

```bash
npm install
pip install -r requirements.txt
npx playwright install chromium
```

复制环境配置文件：

```bash
cp .env.example .env
```

随后编辑 `.env`，填写自己的云启数据账号密码：

```env
YUNQI_USERNAME=your_username
YUNQI_PASSWORD=your_password
```

默认情况下，脚本会使用以下 Excel 模板：

```text
templates/选品表格-模板.xlsx
```

如果需要自定义 Excel 模板路径，可在 `.env` 中配置：

```env
EXCEL_TEMPLATE=/absolute/path/to/选品表格-模板.xlsx
```

完成以上配置后，即可运行对应脚本。


## 启动方式

### 脚本一：生成今日完整选品表格

```bash
cd script1_yunqi_scraper
./run_yunqi_scraper.sh
```

输出位置：

```text
script1_yunqi_scraper/outputs/yunqi-pop-up-greeting-card/选品表格-pop-up-greeting-card-YYYY-MM-DD.xlsx
```

### 脚本二：对比昨天和今天

```bash
cd script2_daily_compare
./run_yunqi_daily_compare.sh /path/to/yesterday.xlsx
```

也可以指定差异表输出路径：

```bash
./run_yunqi_daily_compare.sh /path/to/yesterday.xlsx /path/to/diff.xlsx
```

默认输出位置：

```text
script2_daily_compare/outputs/yunqi-pop-up-greeting-card/选品表格-pop-up-greeting-card-差异-YYYY-MM-DD.xlsx
```

脚本二会先生成当天表格，再与输入的昨天表格按 `商品ID` 对比。

### 脚本三：只追加新商品到已有表格

```bash
cd script3_append_new_products
./run_yunqi_append_new_products.sh /path/to/base.xlsx
```

脚本三不会生成新的最终表格，而是直接修改输入的原有表格 `/path/to/base.xlsx`：

- 已存在的商品 ID：跳过
- 今日新增商品 ID：采集完整信息并追加到表格底部
- 今日消失的商品 ID：不删除，继续保留

## 常用配置

`.env` 支持：

```bash
YUNQI_USERNAME=
YUNQI_PASSWORD=
YUNQI_KEYWORD=pop up greeting card
DAILY_MIN=30
MONTHLY_MIN=1000
EXCEL_TEMPLATE=templates/选品表格-模板.xlsx
HEADLESS=0
```

## 注意事项

- 不要提交 `.env`，里面有账号密码。
- 不要提交 `outputs/`，里面有采集结果、图片、调试 HTML、Chrome profile。
- 第一次运行会打开浏览器并登录云启数据，网站响应较慢时请等待。
- 如果登录失败，多数是网络或网站超时，可以重新运行。

## 额外功能

本项目也可以链接到Codex的自动化里, 让Codex每天自动运行该脚本来爬取数据. 本脚本适配于Codex
