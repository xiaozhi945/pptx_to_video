# PPTX 智能讲解视频生成器

自动将 PowerPoint 文件转换成带 AI 讲解的视频。

## 功能特点

- ✅ 自动解析 PPT 内容（文本、标题、备注）
- ✅ 使用 Claude/智谱AI 生成专业讲解脚本
- ✅ 多语言 TTS 语音合成（支持中英日等多种语言）
- ✅ **PPT 动画支持（Windows + PowerPoint）**
- ✅ **自适应渲染后端（PowerPoint/LibreOffice/Pillow）**
- ✅ 保留原始 PPT 样式
- ✅ 自动合成高清视频（1920x1080）
- ✅ 音画同步防护机制
- ✅ **TTS 并发处理，速度提升 6 倍**
- ✅ **自动缓存，节省 API 费用**
- ✅ **自动检查和安装依赖**
- ✅ **跨平台支持（Windows/Linux/macOS）**

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

- **LibreOffice**（推荐）：高质量 PPT 渲染
  - 下载: https://www.libreoffice.org/download/

- **Poppler**（推荐）：配合 LibreOffice 使用
  ```bash
  choco install poppler
  pip install pdf2image
  ```

- **Microsoft PowerPoint**（可选，仅 Windows）：动画支持
  - 安装 PowerPoint 后运行：
    ```bash
    pip install pywin32
    ```
  - 启用后可导出 PPT 动画效果

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

[rendering]
# 渲染后端: auto, powerpoint, libreoffice, pillow
backend = auto  # 自动选择最佳后端
enable_animation = true  # 启用动画支持（仅 PowerPoint 后端有效）

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

## PPT 动画支持 🎬

### 渲染后端说明

系统支持 **三种渲染后端**，自动选择最佳方案：

| 后端 | 平台 | 动画支持 | 质量 | 依赖 |
|------|------|---------|------|------|
| **PowerPoint** | Windows | ✅ 完整 | ⭐⭐⭐⭐⭐ | Microsoft PowerPoint + pywin32 |
| **LibreOffice** | 跨平台 | ❌ 静态 | ⭐⭐⭐⭐ | LibreOffice + Poppler |
| **Pillow** | 跨平台 | ❌ 静态 | ⭐⭐ | 无（内置） |

### Windows 动画支持

**前提条件**：
- Windows 操作系统
- 已安装 Microsoft PowerPoint

**安装步骤**：
```bash
# 安装 pywin32
pip install pywin32

# 运行程序（自动使用 PowerPoint 渲染器）
python -m pptx_to_video
```

**效果**：
```
[4/5] 生成幻灯片缩略图...
  使用 PowerPoint 渲染器
  ✓ 支持动画渲染
  ✓ 已生成 25 个渲染帧
  ✓ 检测到动画效果，共 25 帧
    第 3 页: 3 个动画步骤
    第 5 页: 2 个动画步骤
```

### Linux/macOS 自动降级

在非 Windows 环境下，系统自动降级到静态渲染：

```bash
# 直接运行（自动使用 LibreOffice 或 Pillow）
python -m pptx_to_video
```

**效果**：
```
[4/5] 生成幻灯片缩略图...
  使用 LibreOffice 渲染器
  ⚠️  仅支持静态渲染（不支持动画）
  ✓ 已生成 20 个渲染帧
```

### 配置渲染后端

在 `config.ini` 中配置：

```ini
[rendering]
# auto: 自动选择最佳后端
# powerpoint: 强制使用 PowerPoint（仅 Windows）
# libreoffice: 强制使用 LibreOffice
# pillow: 强制使用 Pillow
backend = auto

# 是否启用动画支持（仅 PowerPoint 后端有效）
enable_animation = true
```

**详细文档**：查看 [ANIMATION_SUPPORT.md](ANIMATION_SUPPORT.md)

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

### 7. PowerPoint 动画不工作

**Q: 为什么没有检测到动画？**

A: 检查以下几点：
1. 确认在 Windows 系统上
2. 确认已安装 Microsoft PowerPoint
3. 安装 pywin32：`pip install pywin32`
4. 运行测试：`python test_renderer.py`

如果 PowerPoint 渲染器显示"不可用"，系统会自动降级到静态渲染。

### 8. 动画帧与音频不匹配

**Q: 为什么动画帧数量与音频不匹配？**

A: 当前脚本生成器按幻灯片生成脚本，不支持动画步骤级别。

**临时解决方案**：
- 禁用动画：在 `config.ini` 中设置 `enable_animation = false`
- 或手动为每个动画步骤编写脚本

**未来改进**：将更新 LLM prompt 以支持动画步骤级别的脚本生成。

## 项目结构

```
pptx_to_video/
├── pptx_to_video.py      # 主程序
├── ppt_parser.py         # PPT 解析
├── ppt_renderer.py       # 渲染器（PowerPoint/LibreOffice/Pillow）
├── script_generator.py   # AI 脚本生成
├── tts_service.py        # 语音合成（并发处理）
├── video_creator.py      # 视频合成
├── config.py             # 配置管理
├── utils.py              # 工具函数
├── ffmpeg_utils.py       # FFmpeg 工具
├── check_dependencies.py # 依赖检查
├── test_renderer.py      # 渲染器测试
├── prompts/              # AI 提示词
├── requirements.txt      # Python 依赖
├── .env                  # 配置文件（需自行创建）
├── config.ini            # 高级配置（可选）
├── config.ini.example    # 配置示例
├── input/                # 输入 PPT 目录
├── output/               # 输出视频目录
├── temp/                 # 临时文件目录
├── README.md             # 项目说明
├── CLAUDE.md             # 开发文档
└── ANIMATION_SUPPORT.md  # 动画支持文档
```
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
    ├── 你的PPT名称_slide_001.png  # 静态渲染图片
    └── 你的PPT名称_slide_003_step_002.png  # 动画帧（如果启用动画）
```

## 技术栈

- **AI 模型**: Claude Sonnet 4.6 / 智谱 GLM-4 / DeepSeek Chat / 通义千问 Plus
- **TTS**: Microsoft Edge TTS（并发处理）
- **PPT 处理**: python-pptx, PowerPoint COM API (Windows), LibreOffice
- **PDF 转图片**: pdf2image + Poppler
- **视频合成**: FFmpeg
- **图像处理**: Pillow

## 相关文档

- [CLAUDE.md](CLAUDE.md) - 开发文档和架构说明
- [ANIMATION_SUPPORT.md](ANIMATION_SUPPORT.md) - 动画支持详细文档
- [IMPLEMENTATION_SUMMARY.md](IMPLEMENTATION_SUMMARY.md) - 实现总结

## 许可证

MIT License

## 贡献

欢迎提交 Issue 和 Pull Request！

## 联系方式

如有问题，请提交 Issue。

如有问题，请提交 Issue。
