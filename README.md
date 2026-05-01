# PPTX 智能讲解视频生成器

自动将 PowerPoint 文件转换成带 AI 讲解的视频。

## 功能特点

- ✅ 自动解析 PPT 内容（文本、标题、备注）
- ✅ 使用 Claude/智谱AI 生成专业讲解脚本
- ✅ 多语言 TTS 语音合成（支持中英日等多种语言）
- ✅ 保留原始 PPT 样式（通过 LibreOffice）
- ✅ 自动合成高清视频（1920x1080）
- ✅ 音画同步防护机制

## 工作流程

```
PPT 文件 → AI 分析 → 脚本生成 → 语音合成 → 视频合成
```

1. **PPT 解析** - 提取幻灯片内容
2. **AI 分析** - Claude/智谱AI 理解内容结构
3. **脚本生成** - 为每页生成讲解文案
4. **TTS 合成** - 将文案转为语音
5. **视频合成** - 图片 + 音频 → 视频

## 安装依赖

### 1. Python 依赖

```bash
pip install -r requirements.txt
```

### 2. FFmpeg（视频合成）

**Windows - Chocolatey:**
```powershell
choco install ffmpeg
```

**Windows - Scoop:**
```powershell
scoop install ffmpeg
```

**手动安装:**
- 下载: https://ffmpeg.org/download.html
- 解压并添加到 PATH

### 3. LibreOffice（保留原始 PPT 样式）

- 下载: https://www.libreoffice.org/download/
- 安装到任意位置（代码会自动检测常见路径）

### 4. Poppler（PDF 转图片）

**Windows - Chocolatey:**
```powershell
choco install poppler
```

**手动安装:**
```powershell
# 下载
Invoke-WebRequest -Uri "https://github.com/oschwartz10612/poppler-windows/releases/download/v24.08.0-0/Release-24.08.0-0.zip" -OutFile "$env:TEMP\poppler.zip"

# 解压
Expand-Archive -Path "$env:TEMP\poppler.zip" -DestinationPath "D:\Soft\poppler" -Force

# 添加到 PATH
$env:PATH += ";D:\Soft\poppler\poppler-24.08.0\Library\bin"
```

然后安装 Python 库：
```bash
pip install pdf2image
```

## 配置

### 1. 复制配置文件

```bash
cp .env.example .env
```

### 2. 编辑 .env 文件

```env
# LLM 提供商选择
LLM_PROVIDER=claude  # 或 zhipu

# Claude API Key
ANTHROPIC_API_KEY=sk-ant-your-key-here

# 智谱 AI API Key（备用）
ZHIPUAI_API_KEY=your-key-here

# TTS 语音配置
# 中文女声（默认）
TTS_VOICE=zh-CN-XiaoxiaoNeural

# 英文女声
# TTS_VOICE=en-US-AriaNeural

# 中文男声
# TTS_VOICE=zh-CN-YunxiNeural

# 语速调整（-50% 到 +100%）
TTS_RATE=+0%

# 音调调整（-50Hz 到 +50Hz）
TTS_PITCH=+0Hz
```

## 使用方法

### 基本使用

```bash
# 1. 将 PPT 文件放入 input 目录
cp your_presentation.pptx pptx_to_video/input/

# 2. 进入项目目录
cd pptx_to_video

# 3. 运行程序（自动处理 input 目录下所有 .pptx 文件）
python -m pptx_to_video

# 4. 查看输出
# 视频: output/<文件名>/<文件名>.mp4
# 脚本: output/<文件名>/scripts.json
```

### 高级选项

```bash
# 列出 input 目录中的 PPT 文件
python -m pptx_to_video --list

# 处理指定文件
python -m pptx_to_video --input path/to/file.pptx

# 指定输出路径
python -m pptx_to_video --input demo.pptx --output my_video.mp4

# 切换 LLM 提供商
python -m pptx_to_video --provider zhipu

# 跳过 TTS（仅生成脚本）
python -m pptx_to_video --skip-tts

# 跳过视频合成
python -m pptx_to_video --skip-video
```

## TTS 语音选项

### 中文语音

| 语音代码 | 性别 | 特点 |
|---------|------|------|
| `zh-CN-XiaoxiaoNeural` | 女 | 温柔、自然 |
| `zh-CN-YunxiNeural` | 男 | 沉稳、专业 |
| `zh-CN-XiaoyiNeural` | 女 | 甜美、活泼 |
| `zh-CN-YunjianNeural` | 男 | 磁性、成熟 |

### 英文语音

| 语音代码 | 性别 | 特点 |
|---------|------|------|
| `en-US-AriaNeural` | 女 | 美式英语 |
| `en-US-GuyNeural` | 男 | 美式英语 |
| `en-GB-SoniaNeural` | 女 | 英式英语 |
| `en-GB-RyanNeural` | 男 | 英式英语 |

### 日文语音

| 语音代码 | 性别 | 特点 |
|---------|------|------|
| `ja-JP-NanamiNeural` | 女 | 标准日语 |
| `ja-JP-KeitaNeural` | 男 | 标准日语 |

更多语音: https://speech.microsoft.com/portal/voicegallery

## 项目结构

```
pptx_to_video/
├── pptx_to_video.py      # 主程序
├── ppt_parser.py         # PPT 解析
├── script_generator.py   # AI 脚本生成
├── tts_service.py        # 语音合成
├── video_creator.py      # 视频合成
├── config.py             # 配置管理
├── prompts/              # AI 提示词
│   ├── analyze_ppt.txt
│   └── generate_script.txt
├── requirements.txt      # Python 依赖
├── .env                  # 配置文件（需自行创建）
├── .env.example          # 配置模板
├── .gitignore            # Git 忽略文件
├── README.md             # 使用文档
├── input/                # 输入 PPT 目录
├── output/               # 输出视频目录
└── temp/                 # 临时文件目录
```

## 输出文件

运行后会在 `output/<文件名>/` 目录下生成：

```
output/
└── 你的PPT名称/
    ├── 你的PPT名称.mp4      # 最终视频
    └── scripts.json         # 生成的讲解脚本
```

临时文件在 `temp/<文件名>/` 目录：

```
temp/
└── 你的PPT名称/
    ├── content.txt          # PPT 文本内容
    ├── slide_001.mp3        # 各页音频
    ├── slide_002.mp3
    ├── ...
    ├── 你的PPT名称_slide_001.png  # 各页图片
    ├── 你的PPT名称_slide_002.png
    └── ...
```

## 常见问题

### 1. FFmpeg 未找到

**错误**: `警告: FFmpeg 未安装或未在 PATH 中`

**解决**:
- 安装 FFmpeg 并添加到系统 PATH
- 重启 PowerShell/终端

### 2. LibreOffice 未检测到

**现象**: 使用文本渲染模式，而不是原始 PPT 样式

**解决**:
- 确保 LibreOffice 已安装
- 检查代码中的路径是否包含你的安装路径
- 或手动添加到 PATH

### 3. API 配额用完

**错误**: `Error code: 429 - rate_limit_error`

**解决**:
- 等待配额重置
- 切换到其他 LLM 提供商
- 使用其他 API Key

### 4. 中文路径乱码

**现象**: 视频合成失败，路径显示乱码

**解决**: 已修复，程序会自动使用 UTF-8 编码处理中文路径

### 5. 音画不同步

**现象**: 视频中音频和图片不匹配

**解决**: 程序已内置防御性检查，会在数量不匹配时报错并拒绝合成

## 技术栈

- **AI 模型**: Claude Sonnet 4.6 / 智谱 GLM-4
- **TTS**: Microsoft Edge TTS
- **PPT 处理**: python-pptx
- **PPT 渲染**: LibreOffice
- **PDF 转图片**: pdf2image + Poppler
- **视频合成**: FFmpeg
- **图像处理**: Pillow

## 开发计划

- [ ] 支持更多 LLM 提供商（OpenAI、Gemini）
- [ ] 支持自定义提示词模板
- [ ] 支持视频转场效果
- [ ] 支持背景音乐
- [ ] Web UI 界面
- [ ] 批量处理优化

## 许可证

MIT License

## 贡献

欢迎提交 Issue 和 Pull Request！

## 联系方式

如有问题，请提交 Issue。
