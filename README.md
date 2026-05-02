# PPTX 智能讲解视频生成器

自动将 PowerPoint 文件转换成带 AI 讲解的视频。

## 功能特点

- ✅ 自动解析 PPT 内容（文本、标题、备注）
- ✅ 使用 Claude/智谱AI 生成专业讲解脚本
- ✅ 多语言 TTS 语音合成（支持中英日等多种语言）
- ✅ 保留原始 PPT 样式（通过 LibreOffice）
- ✅ 自动合成高清视频（1920x1080）
- ✅ 音画同步防护机制
- ✅ **TTS 并发处理，速度提升 6 倍**
- ✅ **自动缓存，节省 API 费用**
- ✅ **自动检查和安装依赖**

## 快速开始

### 1. 安装依赖

**方式 1：自动安装（推荐）**

直接运行程序，会自动检查并安装缺失的依赖：

```bash
python -m pptx_to_video
```

**方式 2：手动安装**

```bash
pip install -r requirements.txt
```

### 2. 安装外部工具

- **FFmpeg**（必需）：视频合成
  ```bash
  # Windows
  choco install ffmpeg
  # 或下载: https://ffmpeg.org/download.html
  ```

- **LibreOffice**（可选）：保留原始 PPT 样式
  - 下载: https://www.libreoffice.org/download/

- **Poppler**（可选）：配合 LibreOffice 使用
  ```bash
  choco install poppler
  pip install pdf2image
  ```

### 3. 配置 API Key

```bash
# 复制配置模板
cp .env.example .env

# 编辑 .env 文件
ANTHROPIC_API_KEY=sk-ant-your-key-here
```

### 4. 运行程序

```bash
# 将 PPT 放入 input 目录
cp your_presentation.pptx input/

# 运行程序
python -m pptx_to_video

# 查看输出
# 视频: output/<文件名>/<文件名>.mp4
# 脚本: output/<文件名>/scripts.json
```

## 配置说明

### 基础配置（.env 文件）

```env
# LLM 提供商（支持 4 种）
LLM_PROVIDER=claude  # claude, zhipu, deepseek, qianwen

# API Keys（根据使用的提供商配置对应的 Key）
ANTHROPIC_API_KEY=sk-ant-your-key-here
ZHIPUAI_API_KEY=your-key-here
DEEPSEEK_API_KEY=your-key-here
QIANWEN_API_KEY=your-key-here

# TTS 语音
TTS_VOICE=zh-CN-XiaoxiaoNeural  # 中文女声
TTS_RATE=+0%   # 语速（-50% 到 +100%）
TTS_PITCH=+0Hz # 音调（-50Hz 到 +50Hz）
```

### 支持的 LLM 提供商

| 提供商 | 模型 | API Key 获取 | 特点 |
|--------|------|-------------|------|
| **claude** | Claude Sonnet 4.6 | https://console.anthropic.com/ | 推理能力强，支持长文本 |
| **zhipu** | GLM-4 | https://open.bigmodel.cn/ | 国内访问快，中文友好 |
| **deepseek** | DeepSeek Chat | https://platform.deepseek.com/ | 性价比高，推理能力强 |
| **qianwen** | 通义千问 Plus | https://dashscope.console.aliyun.com/ | 阿里云服务，稳定可靠 |

### 高级配置（config.ini 文件）

创建 `config.ini` 文件进行更详细的配置：

```ini
[paths]
# 自定义 LibreOffice 路径（留空则自动检测）
libreoffice_path = D:\Soft\LibreOffice\program\soffice.exe

# 自定义 FFmpeg 路径（留空则自动检测）
ffmpeg_path = D:\Soft\ffmpeg\bin\ffmpeg.exe
ffprobe_path = D:\Soft\ffmpeg\bin\ffprobe.exe

# 自定义 Poppler 路径（留空则自动检测）
# pdf2image 依赖，用于 LibreOffice PDF 转图片
poppler_path = D:\Soft\poppler\Library\bin

# 自定义字体路径
font_title = C:/Windows/Fonts/msyh.ttc  # 微软雅黑
font_body = C:/Windows/Fonts/simhei.ttf  # 黑体

[llm]
provider = claude

[tts]
voice = zh-CN-XiaoxiaoNeural
rate = +20%%  # 注意：百分号要写成 %%
pitch = +0Hz

[video]
width = 1920
height = 1080

[performance]
# 性能优化选项
enable_concurrent_tts = true  # TTS 并发处理（速度提升 6 倍）
enable_cache = true           # 启用缓存（节省 API 费用）
max_retries = 3               # API 最大重试次数
base_retry_delay = 2          # 基础重试延迟（秒）
```

**配置优先级**：`config.ini > .env > 默认值`

### 查看当前配置

```bash
python -c "import config; config.print_config()"
```

## 使用方法

### 基本命令

```bash
# 处理 input 目录下所有 PPT
python -m pptx_to_video

# 列出 input 目录中的文件
python -m pptx_to_video --list

# 处理指定文件
python -m pptx_to_video --input path/to/file.pptx

# 指定输出路径
python -m pptx_to_video --input demo.pptx --output my_video.mp4
```

### 高级选项

```bash
# 切换 LLM 提供商
python -m pptx_to_video --provider zhipu
python -m pptx_to_video --provider deepseek
python -m pptx_to_video --provider qianwen

# 只生成脚本（跳过 TTS）
python -m pptx_to_video --skip-tts

# 只生成音频（跳过视频）
python -m pptx_to_video --skip-video
```

## TTS 语音选项

### 常用语音

| 语言 | 语音代码 | 性别 | 特点 |
|------|---------|------|------|
| 中文 | `zh-CN-XiaoxiaoNeural` | 女 | 温柔、自然（默认） |
| 中文 | `zh-CN-YunxiNeural` | 男 | 沉稳、专业 |
| 英文 | `en-US-AriaNeural` | 女 | 美式英语 |
| 英文 | `en-US-GuyNeural` | 男 | 美式英语 |
| 日文 | `ja-JP-NanamiNeural` | 女 | 标准日语 |

完整列表：https://speech.microsoft.com/portal/voicegallery

## 性能优化

### 自动缓存机制

程序会自动缓存中间结果，避免重复计算：

- **脚本缓存**：`output/<文件名>/scripts.json`
- **音频缓存**：`temp/<文件名>/slide_*.mp3`

如果发现缓存，程序会询问是否使用。要强制重新生成，删除对应的缓存文件即可。

### 并发处理

TTS 语音合成使用并发处理，性能提升约 **6 倍**：

- 20 页 PPT：60 秒 → 10 秒
- 可通过 `config.ini` 中的 `enable_concurrent_tts` 控制

### 智能重试

API 调用失败时使用指数退避策略：

- 第 1 次重试：等待 2 秒
- 第 2 次重试：等待 4 秒
- 第 3 次重试：等待 8 秒

## 常见问题

### 1. 依赖安装

**Q: 每次都要运行 `pip install` 吗？**

A: 不需要。首次运行程序会自动检查并安装依赖，之后直接运行即可。

**Q: 自动安装失败怎么办？**

A: 手动运行 `pip install -r requirements.txt`，或使用国内镜像：
```bash
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```

### 2. FFmpeg 未找到

**错误**：`警告: FFmpeg 未安装或未在 PATH 中`

**解决**：
- 安装 FFmpeg 并添加到系统 PATH
- 重启终端

### 3. LibreOffice 未检测到

**现象**：使用文本渲染模式，而不是原始 PPT 样式

**解决**：
- 安装 LibreOffice
- 在 `config.ini` 中配置路径：
  ```ini
  [paths]
  libreoffice_path = D:\Soft\LibreOffice\program\soffice.exe
  ```

### 4. API 配额用完

**错误**：`Error code: 429 - rate_limit_error`

**解决**：
- 等待配额重置
- 切换到其他 LLM 提供商：`--provider zhipu`
- 使用其他 API Key

### 5. 缓存问题

**Q: 如何清除缓存？**

A: 删除缓存文件：
```bash
rm -rf output/*/scripts.json
rm -rf temp/*/slide_*.mp3
```

或在 `config.ini` 中禁用缓存：
```ini
[performance]
enable_cache = false
```

### 6. 配置未生效

**Q: 修改了配置但没有生效？**

A: 检查配置优先级和语法：
```bash
# 查看当前配置
python -c "import config; config.print_config()"

# 检查 config.ini 语法
cat config.ini
```

注意：`config.ini` 中的百分号要写成 `%%`

## 项目结构

```
pptx_to_video/
├── pptx_to_video.py      # 主程序
├── ppt_parser.py         # PPT 解析
├── script_generator.py   # AI 脚本生成
├── tts_service.py        # 语音合成（并发处理）
├── video_creator.py      # 视频合成
├── config.py             # 配置管理
├── utils.py              # 工具函数
├── check_dependencies.py # 依赖检查
├── prompts/              # AI 提示词
├── requirements.txt      # Python 依赖
├── .env                  # 配置文件（需自行创建）
├── config.ini            # 高级配置（可选）
├── input/                # 输入 PPT 目录
├── output/               # 输出视频目录
└── temp/                 # 临时文件目录
```

## 输出文件

```
output/
└── 你的PPT名称/
    ├── 你的PPT名称.mp4      # 最终视频
    └── scripts.json         # 生成的讲解脚本

temp/
└── 你的PPT名称/
    ├── content.txt          # PPT 文本内容
    ├── slide_001.mp3        # 各页音频
    └── 你的PPT名称_slide_001.png  # 各页图片
```

## 技术栈

- **AI 模型**: Claude Sonnet 4.6 / 智谱 GLM-4 / DeepSeek Chat / 通义千问 Plus
- **TTS**: Microsoft Edge TTS（并发处理）
- **PPT 处理**: python-pptx
- **PPT 渲染**: LibreOffice
- **PDF 转图片**: pdf2image + Poppler
- **视频合成**: FFmpeg
- **图像处理**: Pillow

## 开发文档

如果你想了解代码架构或进行二次开发，请查看 [CLAUDE.md](CLAUDE.md)。

## 许可证

MIT License

## 贡献

欢迎提交 Issue 和 Pull Request！

## 联系方式

如有问题，请提交 Issue。
