# wiki2docx

一键将 RoboMaster 论坛 Wiki 文档导出为格式完整的 Word（.docx）文档。

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-green.svg)](https://www.python.org/)

## 功能特色

- 🔐 **微信扫码登录** — 弹出浏览器 → 微信扫码 → 自动检测登录 → 开始抓取，不碰密码
- 📑 **完整抓取** — 自动递归展开侧边栏（最多 12 轮），74 个节点全覆盖
- 📊 **格式保留** — 表格、图片、列表、标题层级、粗体/斜体/颜色/字号，全部转换到 Word
- 📄 **专业排版** — 自动生成封面、目录、分页，中文排版优化
- 🚀 **两种模式** — CLI 一键完成，或纯浏览器 Console 脚本（无需装 Python）

## 快速开始

### 前提条件

- **Python 3.10+**
- **Google Chrome**（系统已安装的即可）
- **RoboMaster 论坛账号**（已绑定微信）

### 安装

```bash
# 1. 克隆仓库
git clone https://github.com/icbk/rm-wiki-docx.git
cd rm-wiki-docx

# 2. 安装 Python 依赖
pip install -r requirements.txt
```

> **注意**：无需额外安装 Chromium。本工具使用你系统已安装的 Chrome 浏览器。

### 使用

```bash
# 最简单用法：替换 WIKI_ID 即可
python -m wiki2docx --wiki-id 46368458
```

**运行流程**：

1. 自动弹出 Chrome 浏览器，打开 DJI 登录页面
2. 切换至"微信登录"，截图二维码保存到本地
3. **用微信扫描二维码**完成登录
4. 脚本自动检测登录成功 → 展开侧边栏 → 爬取所有章节 → 生成 Word

```
=======================================================
 微信登录二维码已保存至：
   ./_browser_profile/wechat_login_qr.png

 请打开该图片，用微信扫描二维码完成登录。
 脚本将自动检测登录状态，无需操作。
=======================================================

检测到已重定向至 bbs.robomaster.com，登录成功！
导航到 wiki 页面: https://bbs.robomaster.com/wiki/46368458
wiki 页面访问成功。
展开侧边栏所有节点...
  共 74 个节点 (61 个叶子节点)
抓取 74 个章节内容...
  [1/74] 1. 团队建设分析 [27 chars] ...
  ...
抓取完成: 73/74 个章节有内容
Word 文档已生成: output.docx
```

### 高级用法

```bash
# 自定义输出路径
python -m wiki2docx --wiki-id 46368458 --output 我的文档.docx

# 使用配置文件（自定义封面标题、作者等）
python -m wiki2docx --wiki-id 46368458 --config config.json

# 仅抓取 JSON（不生成 Word）
python -m wiki2docx --wiki-id 46368458 --json-only

# 从已有的 JSON 文件生成 Word（队友发你的 data.json）
python -m wiki2docx --json-file data.json --output 输出.docx
```

### 给不想装 Python 的队友

如果你不想装 Python，可以用纯浏览器脚本完成抓取：

1. 确保已在 Chrome 中登录 RoboMaster 论坛
2. 打开目标 Wiki 页面，按 `F12` → **Console**
3. 复制 `wiki2docx/bookmarklet.js` 全部代码 → 粘贴到 Console → Enter
4. 等待 `=== DONE ===` → 运行 `copy(JSON.stringify(allData))`
5. 粘贴到记事本，保存为 `data.json`
6. 把 `data.json` 发给有 Python 环境的队友生成 Word

## 配置文件

创建 `config.json`（可选）：

```json
{
  "wiki_url": "https://bbs.robomaster.com/wiki/46368458",
  "title": "RMUC2026 广州航海学院 ICBK破冰船 赛季总结文档",
  "subtitle": "广州航海学院  ICBK破冰船 战队",
  "output": "output.docx",
  "font": "微软雅黑"
}
```

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `wiki_url` | Wiki 页面完整 URL | `https://bbs.robomaster.com/wiki/{wiki_id}` |
| `title` | Word 封面标题 | 自动从页面获取 |
| `subtitle` | Word 封面副标题 | 空 |
| `font` | 正文/标题字体 | `微软雅黑` |

## CLI 参数

```
python -m wiki2docx --help

  --wiki-id ID       Wiki 页面 ID（例如 46368458）
  --output PATH      输出 Word 路径（默认: output.docx）
  --config PATH      配置文件（JSON 格式）
  --profile-dir DIR  工作目录（默认: ./_browser_profile）
  --json-only        仅抓取 JSON，不生成 Word
  --json-file PATH   跳过抓取，从已有 JSON 文件生成 Word
```

## 项目结构

```
rm-wiki-docx/
├── wiki2docx/
│   ├── __init__.py
│   ├── __main__.py         # python -m wiki2docx 入口
│   ├── cli.py              # argparse CLI
│   ├── scraper.py          # Playwright 浏览器抓取 + 微信扫码登录
│   ├── converter.py        # HTML → Word 转换引擎
│   └── bookmarklet.js      # 纯浏览器 Console 脚本（无需 Python）
├── config.example.json     # 配置模板
├── requirements.txt
├── LICENSE                 # MIT
└── README.md
```

## 工作原理

```
┌──────────────┐     ┌─────────────────┐     ┌──────────────┐
│  DJI 登录页    │ ──→ │ 微信扫码 QR 截图  │ ──→ │ 轮询等待登录  │
│ (Chrome)     │     │ (本地 PNG)      │     │ (自动检测)    │
└──────────────┘     └─────────────────┘     └──────┬───────┘
                                                    │
                                        ┌───────────▼───────────┐
                                        │  登录成功              │
                                        │  → 导航到 wiki 页面     │
                                        │  → 展开侧边栏（12 轮）  │
                                        │  → 逐节点点击抓取 HTML  │
                                        └───────────┬───────────┘
                                                    │
                                        ┌───────────▼───────────┐
                                        │  HTML → Word 转换引擎  │
                                        │  · 表格、列表、图片     │
                                        │  · 标题层级、文本样式   │
                                        │  · 封面、目录、分页     │
                                        └───────────┬───────────┘
                                                    │
                                        ┌───────────▼───────────┐
                                        │   📄 output.docx       │
                                        └───────────────────────┘
```

## 适配范围

- ✅ `bbs.robomaster.com/wiki/` — RM 论坛 Wiki 页面
- ✅ DJI 账号微信扫码登录
- ✅ Ant Design 树形侧边栏（懒加载）
- ✅ Slate 编辑器富文本（表格/列表/图片/样式）
- 🔧 其他类似结构的文档站点可提 Issue 适配

## FAQ

<details>
<summary>Q: 扫码后脚本没反应？</summary>

请确认扫码后在浏览器中看到了 `bbs.robomaster.com` 首页（不是 DJI 登录页）。脚本检测到 URL 变为 `https://bbs.robomaster.com/` 即判定登录成功。
</details>

<details>
<summary>Q: 抓取到 0 个章节？</summary>

确认 Chrome 已安装在默认路径（`C:\Program Files\Google\Chrome\Application\chrome.exe`）。如果安装在别处，可以提 Issue。
</details>

<details>
<summary>Q: 生成的 Word 排版不对？</summary>

`config.json` 中可调整 `font` 参数。中文推荐 `微软雅黑`、`宋体`、`思源黑体`。
</details>

<details>
<summary>Q: 支持 macOS / Linux 吗？</summary>

支持。Python + Playwright 全平台兼容，使用系统安装的 Chrome。唯一区别是二维码截图路径不同。
</details>

<details>
<summary>Q: 爬取中途报错退出？</summary>

可能是网络问题或页面结构变化。重试通常可以解决。如果持续报错，请提 Issue 附上错误信息。
</details>

## 贡献

欢迎 PR！如果你适配了其他 RM 系列的文档站点、或新增了导出格式，请提 Issue 或直接 PR。

## License

MIT © ICBK RoboMaster
