# DocAutoReviewer 📄✨

DocAutoReviewer 是一个用于批量评阅学生实验报告（.docx）的桌面工具。  
它会自动抽取“实验心得/反思”内容，调用 DeepSeek API 生成教师评语，并回填到指定位置后保存到 output/。

## 功能亮点 ✅
- 批量处理 .docx，自动保存到 `output/`
- 支持“开始标记 / 结束标记 / 评语标记”三段定位
- 支持风格与期望字数配置
- 进度条 + 当前文件 + 日志输出
- 断点续跑（`output/progress.json`）
- 暂停 / 继续 / 停止

## 快速开始 🚀

### 1) 创建虚拟环境并安装依赖
```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

### 2) 运行
```powershell
python ui_main.py
```

### 3) 填写设置
- Base URL / API Key / Model
- 开始标记（例如：实验心得与反思）
- 结束标记（例如：教师评语与评分）
- 评语标记（例如：教师总评：）

## 国内镜像安装依赖 🇨🇳
```powershell
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```

## 一键打包 EXE 🧩
（单文件、无控制台、图标内嵌）
```powershell
pyinstaller -w -F --icon=report.ico --add-data "report.ico;." ui_main.py
```

打包后 EXE 位于：
```
dist\ui_main.exe
```

## 目录结构 📂
```
DocAutoReviewer/
├─ ui_main.py
├─ worker.py
├─ docx_io.py
├─ deepseek_client.py
├─ settings.py
├─ requirements.txt
├─ report.ico
└─ output/            # 运行时自动生成
```

## 说明与注意事项 ⚠️
- `settings.json` 会保存 API Key，请勿上传到仓库
- `output/progress.json` 用于断点续跑，删除后会重新处理全部文件
- 如果输出文件被 Word 打开导致写入失败，请关闭文档后重试

## 配置示例截图 🖼️
截图API配置示例：

![配置示例](screenshot.png)

## 常见问题 / 排错 🛠️
**1) 提示 “未找到反思内容”**
- 说明开始/结束标记不匹配文档标题，请改成更短的关键字（例如“实验心得与反思”）。

**2) 提示 “无法写入，请关闭文档后重试”**
- 说明输出 docx 正被 Word 打开占用，请关闭文件后再次运行。

**3) UI 能打开但无结果**
- 检查是否配置了 API Key，且网络可访问 DeepSeek API。
- 确认 `output/progress.json` 未将文件标记为 completed。

**4) 打包后的 exe 没图标**
- 确认打包命令包含 `--icon=report.ico --add-data "report.ico;."`
