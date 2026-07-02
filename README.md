# 02324 离散数学 · 思维导图

全国自考 02324 离散数学考点思维导图（交互式 HTML）。

在线访问：[https://hanxiaodao.github.io/02324-mindmap/](https://hanxiaodao.github.io/02324-mindmap/)

## 结构

```
├── index.html               # 思维导图页面（浏览器打开 / Pages 部署）
├── update.bat               # 双击入口，调用 scripts/update.ps1
├── scripts/                 # 自动化构建脚本
│   ├── download_wps.py      # 从 WPS 云文档下载最新 .docx
│   ├── generate.py          # 从 .docx 生成 index.html（零 LLM 依赖）
│   ├── update.ps1           # 一键运行：下载 + 生成 + 推送
│   └── requirements.txt     # Python 依赖
└── .github/workflows/       # GitHub Actions 自动部署
```

## 更新思维导图

### 一键流程

1. **双击** `update.bat`
2. 脚本自动完成全部：
   - 从 WPS 云文档下载最新笔记
   - 生成 `index.html`
   - 推送到 GitHub → Actions 自动部署 Pages
3. 等待 1-2 分钟后访问 Pages URL

### 首次使用

第一次运行需要短暂关闭 Edge（获取 WPS 登录态），之后无需再关。

### 分步执行

```bash
# 只下载
python scripts\download_wps.py

# 只生成（指定 docx）
python scripts\generate.py --docx 文件名.docx

# 只生成（自动找目录下的 docx）
python scripts\generate.py
```

## 发布到 GitHub Pages

`update.bat` 已自动执行 git push，无需手动操作。

如果只想手动发布：
```bash
git add index.html
git commit -m "update"
git push
```

## 依赖

```bash
pip install python-docx playwright
playwright install chromium
```
