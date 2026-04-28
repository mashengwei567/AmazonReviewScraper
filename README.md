<div align="center">

# 🛒 Amazon Review Scraper

**Amazon 用户评论爬虫 — 一键爬取用户全部评论并自动分析刷评嫌疑**

[![Python](https://img.shields.io/badge/Python-3.8+-blue?logo=python&logoColor=white)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)]()

</div>

---

## ✨ 功能特性

- 📥 从 Amazon 用户概况页自动爬取**全部评论**（评分、日期、VP 状态等）
- 📋 支持**批量输入**多个用户 URL，一次爬取多人
- 📊 输出专业排版的 **Excel 报告**（条件着色、斑马纹、自动筛选）
- 🔍 自动**刷评分析**（6 项规则打分判定，生成分析报告）
- 💾 同时输出 **JSON 原始数据**，便于二次处理
- 🖥️ 提供 **GUI 图形界面** + CLI 命令行双模式
- 🔐 支持**代理配置**和自定义 User-Agent

---

## 📸 界面预览

> 启动后进入图形界面，三步完成爬取：登录 → 输入 URL → 开始

| 登录 | 输入 URL | 爬取中 |
|:----:|:-------:|:------:|
| 点击按钮登录 Amazon | 粘贴用户概况页 URL | 实时查看进度和日志 |

---

## 🚀 快速开始

### 环境要求

- Python 3.8+
- Google Chrome 浏览器

### 安装

```bash
# 克隆仓库
git clone https://github.com/你的用户名/AmazonReviewScraper.git
cd AmazonReviewScraper

# 安装依赖
pip install -r requirements.txt

# 安装 Playwright 浏览器驱动
playwright install chromium
```

### 运行

```bash
# GUI 模式（默认）
python amazon_review_scraper.py

# CLI 命令行模式
python amazon_review_scraper.py --cli
```

---

## 📖 使用说明

### 第 1 步：登录 Amazon

> 登录后可查看 Verified Purchase 等详情，且不易触发反爬验证。

1. 点击 **「登录 Amazon」** 按钮
2. 在弹出的浏览器窗口中**手动登录**
3. 登录成功后回到程序点击 **「确定」**

> 💡 登录状态保存在 `browser_data/` 中，下次无需重新登录。删除该文件夹可清除登录状态。

### 第 2 步：输入用户概况页 URL

在 Amazon 上找到目标用户的概况页，URL 格式：

```
https://www.amazon.com/gp/profile/amzn1.account.XXXXXXXXXXXXXXXXXXXXXXXXXXXX
```

**获取方式：** 在商品评论区点击评论者的用户名 → 复制地址栏 URL

### 第 3 步：开始爬取

1. 粘贴 URL（支持多个，每行一个）
2. 选择是否启用无头模式
3. 点击 **「开始爬取」**

> ⚠️ 爬取过程中如遇到验证码，程序会自动暂停，在浏览器中手动完成验证后点击确定继续。

---

## 📂 输出文件

所有输出保存在 `output/` 文件夹中：

```
output/
  ├── 用户名_reviews.json    # JSON 原始数据
  └── batch_时间戳.xlsx       # Excel 综合报告
```

### Excel 报告内容

| Sheet | 内容 |
|-------|------|
| 用户名 | 该用户的所有评论数据 |
| **分析报告** | 所有用户的刷评分析汇总 |

### 评论字段说明

| 字段 | 说明 |
|------|------|
| Review ID | Amazon 评论唯一 ID |
| Product Name | 商品名称 |
| ASIN | Amazon 商品编号（10 位） |
| Title | 评论标题 |
| Description | 评论正文 |
| Rating | 评分（1-5 ⭐） |
| Review Date | 评论日期 |
| Verified Purchase | 是否已验证购买 |
| Helpful Votes | "有帮助" 投票数 |
| Hearts | 点赞数 |
| Review URL | 评论详情页链接 |

---

## 🔍 刷评分析规则

程序会根据以下 6 项规则对用户进行评分：

| 规则条件 | 分值 | 说明 |
|----------|:----:|------|
| 5 星评价占比 > 80% | +2 | 高比例好评可能为虚假评论 |
| 已验证购买占比 < 40% | +2 | 大量评论无购买记录 |
| 评论平均长度 < 80 字符 | +1 | 内容空洞，可能为模板评论 |
| 同日发布 ≥ 3 条评论 | +2 | 短时间大量发布 |
| 所有评论无人点赞（≥5 条时） | +1 | 评论缺乏互动 |
| 评论总数 > 50 条 | +1 | 活跃度异常高 |

### 结论判定

| 分数 | 结论 |
|:----:|------|
| 0-2 | ✅ 未发现明显异常 |
| 3-4 | ⚠️ 存在刷评嫌疑 |
| ≥5 | 🚨 高度疑似刷评 |

---

## ⚙️ 配置参数

配置保存在 `config.json` 中，可通过 GUI 设置界面或 CLI 菜单修改。

| 参数 | 默认值 | 说明 |
|------|:------:|------|
| `headless` | `false` | 无头模式（`true` = 后台运行，不弹窗） |
| `page_timeout` | `90000` | 概况页加载超时（毫秒） |
| `detail_timeout` | `30000` | 详情页加载超时（毫秒） |
| `delay_min` | `4` | 评论间最小等待时间（秒） |
| `delay_max` | `8` | 评论间最大等待时间（秒） |
| `browse_min` | `2` | 详情页模拟浏览最小时间（秒） |
| `browse_max` | `5` | 详情页模拟浏览最大时间（秒） |
| `max_scrolls` | `100` | 概况页最大滚动次数 |
| `max_no_change` | `5` | 连续无新评论时停止的次数 |
| `user_agent` | `""` | 自定义 User-Agent（留空使用默认） |
| `proxy_type` | `""` | 代理协议（http / https / socks4 / socks5） |
| `proxy_ip` | `""` | 代理 IP |
| `proxy_port` | `""` | 代理端口 |

> 💡 **反爬建议：** `delay_min` 和 `delay_max` 建议保持 4-8 秒，过快容易触发 Amazon 反爬机制。

---

## ❓ 常见问题

<details>
<summary><b>Q: 爬取时出现验证码怎么办？</b></summary>

确保使用有头模式（不勾选「无头模式」），程序会自动暂停并在浏览器中提示你手动完成验证。
</details>

<details>
<summary><b>Q: 登录状态过期了怎么办？</b></summary>

重新点击「登录 Amazon」，在浏览器中重新登录即可。
</details>

<details>
<summary><b>Q: 概况页评论加载不全？</b></summary>

在设置中增大 `max_scrolls` 和 `max_no_change` 的值。
</details>

<details>
<summary><b>Q: 爬取速度太慢？</b></summary>

可适当减小 `delay_min` 和 `delay_max`，但 **不建议低于 2 秒**，否则容易触发反爬。
</details>

<details>
<summary><b>Q: 打包后的 exe 打开没反应？</b></summary>

- 确保没有被杀毒软件拦截（PyInstaller 打包的 exe 可能被误报）
- 右键「以管理员身份运行」
- 确保整个文件夹完整，不要只拷贝 exe
</details>

<details>
<summary><b>Q: 支持哪些 Amazon 站点？</b></summary>

目前仅支持 Amazon.com（美国站）。URL 必须包含 `amazon.com/gp/profile/`。
</details>

---

## 📁 项目结构

```
AmazonReviewScraper/
├── amazon_review_scraper.py   # 主程序（GUI + CLI）
├── requirements.txt           # Python 依赖
├── build.bat                  # PyInstaller 打包脚本
├── config.json                # 运行时配置（自动生成）
├── 使用说明.md                # 中文使用文档
├── 使用说明.pdf               # PDF 版文档
├── browser_data/              # 浏览器登录数据（自动生成，已 gitignore）
└── output/                    # 爬取结果输出（自动生成，已 gitignore）
```

---

## 📜 开源协议

本项目基于 [MIT License](LICENSE) 开源。

---

<div align="center">

**如果这个项目对你有帮助，请给个 ⭐ Star 支持一下！**

</div>
