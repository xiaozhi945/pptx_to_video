# PPTX 智能讲解视频生成器

自动将 PowerPoint 文件转换成带 AI 讲解的视频。

## 功能特点

- ✅ 自动解析 PPT 内容（文本、标题、备注、表格、组合对象）
- ✅ 使用 Claude/智谱AI/DeepSeek/通义千问 生成专业讲解脚本
- ✅ 多语言 TTS 语音合成（支持中英日等 140+ 种语言）
- ✅ **智能语言匹配（根据 TTS 语音自动生成对应语言脚本）**
- ✅ **完美音视频同步（FFmpeg 帧级精度合成）**
- ✅ **PPT 动画完整保留（Windows + PowerPoint）**
- ✅ 保留原始 PPT 样式和动画效果
- ✅ 自动合成高清视频（1920x1080）
- ✅ **TTS 并发处理，速度提升 6 倍**
- ✅ **自动缓存，节省 API 费用**
- ✅ **自动检查和安装依赖**

## 工作原理

```
PPT 文件 → AI 生成讲解脚本 → TTS 语音合成 → PowerPoint 导出视频 → FFmpeg 精确合成音频 → 完成
```

**核心技术**：
- PowerPoint 导出保留所有动画效果
- FFmpeg 以帧级精度（30fps = 33.3ms）合成音频
- 音频在幻灯片出现时立即播放，与动画完美同步

## 快速开始

### 1. 安装依赖

**方式 1：自动安装（推荐）**

```bash
python pptx_to_video.py
```

程序会自动检查并安装缺失的依赖。

**方式 2：手动安装**

```bash
pip install -r requirements.txt
```

### 2. 安装外部工具

**必需工具**：

1. **FFmpeg**（音视频合成）
   ```bash
   # Windows (使用 Chocolatey)
   choco install ffmpeg
   
   # 或手动下载: https://ffmpeg.org/download.html
   ```

2. **Microsoft PowerPoint**（仅 Windows，用于动画支持）
   - 下载: https://www.microsoft.com/microsoft-365/powerpoint
   - 安装后运行：`pip install pywin32`
   - pywin32 下载: https://pypi.org/project/pywin32/

### 3. 配置 API Key

编辑 `config.ini` 或创建 `.env` 文件：

```ini
# config.ini
[llm]
anthropic_api_key = sk-ant-your-key-here
```

或

```bash
# .env
ANTHROPIC_API_KEY=sk-ant-your-key-here
```

### 4. 运行程序

```bash
# 将 PPT 放入 input 目录
cp your_presentation.pptx input/

# 运行程序
python pptx_to_video.py

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

# Claude API 中转站（可选）
# 如果使用第三方 Claude API 中转服务，配置此项
ANTHROPIC_BASE_URL=https://api.anthropic.com  # 默认官方地址
# ANTHROPIC_BASE_URL=https://your-proxy-url.com  # 中转站地址示例

# 自定义模型（可选）
# 如果不配置，将使用各提供商的默认模型
# LLM_MODEL=claude-opus-4-7  # Claude 自定义模型
# LLM_MODEL=glm-4-plus       # 智谱自定义模型

# TTS 语音配置
TTS_VOICE=zh-CN-XiaoxiaoNeural  # 中文女声（系统会自动识别语言）
TTS_RATE=+0%   # 语速（-50% 到 +100%）
TTS_PITCH=+0Hz # 音调（-50Hz 到 +50Hz）
```

### 支持的 LLM 提供商

| 提供商 | 默认模型 | API Key 获取 | 特点 |
|--------|---------|-------------|------|
| **claude** | claude-sonnet-4-6 | https://console.anthropic.com/ | 推理能力强，支持长文本 |
| **zhipu** | glm-4 | https://open.bigmodel.cn/ | 国内访问快，中文友好 |
| **deepseek** | deepseek-chat | https://platform.deepseek.com/ | 性价比高，推理能力强 |
| **qianwen** | qwen-plus | https://dashscope.console.aliyun.com/ | 阿里云服务，稳定可靠 |

#### 自定义模型

如果不配置模型，将使用上表中的默认模型。可以通过以下方式自定义：

**方式 1：通过 .env 文件**
```env
LLM_PROVIDER=claude
LLM_MODEL=claude-opus-4-7  # 使用更强大的 Opus 模型
```

**方式 2：通过 config.ini 文件**
```ini
[llm]
provider = claude
model = claude-opus-4-7
```

#### Claude API 中转站配置

如果使用第三方 Claude API 中转服务（例如国内中转站），可以配置自定义 base_url：

```env
# .env
ANTHROPIC_BASE_URL=https://your-proxy-url.com
```

或

```ini
# config.ini
[llm]
anthropic_base_url = https://your-proxy-url.com
```

### 高级配置（config.ini 文件）

```ini
[paths]
# FFmpeg 路径（用于音视频合成）
ffmpeg_path = D:\Soft\ffmpeg\bin\ffmpeg.exe
ffprobe_path = D:\Soft\ffmpeg\bin\ffprobe.exe

[llm]
provider = claude
anthropic_api_key = sk-ant-your-key-here
# anthropic_base_url = https://your-proxy-url.com  # 可选，用于 Claude API 中转站

[tts]
voice = zh-CN-XiaoxiaoNeural
rate = +0%%  # 注意：百分号要写成 %%
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
python pptx_to_video.py

# 列出 input 目录中的文件
python pptx_to_video.py --list

# 处理指定文件
python pptx_to_video.py --input path/to/file.pptx

# 指定输出路径
python pptx_to_video.py --input demo.pptx --output my_video.mp4
```

### 高级选项

```bash
# 切换 LLM 提供商
python pptx_to_video.py --provider zhipu
python pptx_to_video.py --provider deepseek
python pptx_to_video.py --provider qianwen

# 只生成脚本（跳过 TTS）
python pptx_to_video.py --skip-tts

# 只生成音频（跳过视频）
python pptx_to_video.py --skip-video
```

## TTS 语音选项

### 智能语言匹配

**核心特性**：系统会根据配置的 TTS 语音**自动识别目标语言**，并指示 AI 生成对应语言的脚本。

#### 工作原理

1. **语言识别**：从 TTS 语音配置中提取语言代码
   - `zh-CN-XiaoxiaoNeural` → 识别为中文
   - `en-US-AriaNeural` → 识别为英文
   - `ja-JP-NanamiNeural` → 识别为日文

2. **脚本生成**：AI 自动使用识别出的语言生成讲解脚本

3. **语音合成**：TTS 使用配置的语音合成音频

**优势**：
- ✅ 语音和脚本语言始终匹配，避免 TTS 合成失败
- ✅ 无需手动指定脚本语言
- ✅ 支持 20+ 种语言自动识别

#### 示例

```ini
# config.ini
[tts]
voice = zh-CN-XiaoxiaoNeural  # 配置中文语音
```

系统自动：
1. 识别语言为中文
2. 指示 AI 生成中文脚本
3. 使用中文语音合成音频

**重要**：如果手动修改 TTS 语音配置，系统会自动调整脚本语言，无需其他操作。

### 常用语音

| 语言 | 语音代码 | 性别 | 特点 |
|------|---------|------|------|
| 中文 | `zh-CN-XiaoxiaoNeural` | 女 | 温柔、自然（默认） |
| 中文 | `zh-CN-YunxiNeural` | 男 | 沉稳、专业 |
| 中文 | `zh-CN-XiaoyiNeural` | 女 | 活泼、年轻 |
| 中文 | `zh-CN-YunjianNeural` | 男 | 成熟、磁性 |
| 英文 | `en-US-AriaNeural` | 女 | 美式英语 |
| 英文 | `en-US-GuyNeural` | 男 | 美式英语 |
| 日文 | `ja-JP-NanamiNeural` | 女 | 标准日语 |
| 日文 | `ja-JP-KeitaNeural` | 男 | 标准日语 |
| 韩语 | `ko-KR-SunHiNeural` | 女 | 标准韩语 |
| 法语 | `fr-FR-DeniseNeural` | 女 | 法国法语 |
| 德语 | `de-DE-KatjaNeural` | 女 | 标准德语 |
| 西班牙语 | `es-ES-ElviraNeural` | 女 | 西班牙语 |

完整列表（支持 140+ 种语言）：https://speech.microsoft.com/portal/voicegallery

## 音视频同步原理

### FFmpeg 后期合成方案

系统使用 **FFmpeg 后期合成** 实现帧级精度的音视频同步：

**工作流程**：
1. PowerPoint 导出视频（无音频，保留所有动画）
2. FFmpeg 按顺序拼接所有音频文件
3. FFmpeg 将音频精确合成到视频中

**优势**：
- ✅ **帧级精度**：30fps = 33.3ms，60fps = 16.7ms
- ✅ **音频立即播放**：幻灯片出现时立即播放，不等待动画
- ✅ **完全可控**：精确控制音频时间点
- ✅ **业界标准**：专业视频编辑的标准流程

**技术细节**：
```bash
# 1. 拼接音频
ffmpeg -f concat -safe 0 -i concat_list.txt -c copy merged_audio.mp3

# 2. 合成到视频
ffmpeg -i video_no_audio.mp4 -i merged_audio.mp3 \
  -c:v copy \      # 不重新编码视频
  -c:a aac \       # 音频编码为 AAC
  -b:a 192k \      # 音频比特率
  -shortest \      # 匹配最短流
  output.mp4
```

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

### 1. FFmpeg 未找到

**错误**：`FFmpeg 未找到或不可用`

**解决**：
1. 安装 FFmpeg：`choco install ffmpeg`（Windows）
2. 或在 `config.ini` 中配置路径：
   ```ini
   [paths]
   ffmpeg_path = D:\Soft\ffmpeg\bin\ffmpeg.exe
   ffprobe_path = D:\Soft\ffmpeg\bin\ffprobe.exe
   ```
3. 重启终端

### 2. PowerPoint 未检测到

**现象**：程序提示需要 PowerPoint

**解决**：
1. **确认在原生 Windows 系统上运行**
   - ⚠️ **不支持 WSL (Windows Subsystem for Linux)**
   - 必须在 Windows 命令提示符 (CMD) 或 PowerShell 中运行
   - WSL 环境中 `platform.system()` 返回 'Linux'，导致 PowerPoint COM API 不可用
2. 安装 Microsoft PowerPoint
3. 安装 pywin32：`pip install pywin32`

### 3. API 配额用完

**错误**：`Error code: 429 - rate_limit_error`

**解决**：
- 等待配额重置
- 切换到其他 LLM 提供商：`--provider zhipu`
- 使用其他 API Key

### 4. 视频生成失败但文件存在

**现象**：PowerPoint 显示"失败"，但视频文件已生成

**原因**：PowerPoint API 的已知问题，状态码不可靠

**解决**：程序会自动检查文件是否存在，如果文件 >100KB 则继续处理

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
```

注意：`config.ini` 中的百分号要写成 `%%`

### 7. TTS 语音合成失败

**错误**：`No audio was received. The voice does not match the language of the input text.`

**原因**：TTS 语音配置与脚本语言不匹配（例如使用英文语音合成中文文本）

**解决**：
- 系统会自动根据 TTS 语音配置生成对应语言的脚本
- 如果仍然失败，检查 `config.ini` 中的 `voice` 配置是否正确
- 常见配置：
  - 中文：`zh-CN-XiaoxiaoNeural`
  - 英文：`en-US-AriaNeural`
  - 日文：`ja-JP-NanamiNeural`

## 已知问题

### 1. 不支持 WSL 环境

**限制**：程序必须在 Windows 原生环境中运行，不支持 WSL (Windows Subsystem for Linux)

**原因**：
- PowerPoint COM API 只能在 Windows 原生环境中使用
- WSL 环境中 `platform.system()` 返回 'Linux'，无法调用 PowerPoint

**解决**：在 Windows 命令提示符 (CMD) 或 PowerShell 中运行程序

### 2. 动画脚本生成功能暂不可用

**问题**：`generate_with_animation()` 方法引用了不存在的 `self.animation_script_prompt` 属性

**影响**：如果尝试使用动画感知的脚本生成功能会报错

**状态**：已知 bug，待修复

**临时方案**：使用默认的 `generate_script()` 方法（当前默认行为）

### 3. PowerPoint 导出状态码不可靠

**问题**：PowerPoint API 可能返回失败状态，但视频文件实际已成功生成

**解决**：程序会自动检查文件大小，如果 >100KB 则认为导出成功并继续处理

### 4. 表格和组合对象解析

**支持情况**：
- ✅ 已支持表格 (TABLE) 文本提取
- ✅ 已支持组合对象 (GROUP) 文本提取
- ⚠️ 复杂嵌套结构可能解析不完整

**建议**：如果发现某些文本未被提取，可以在 PPT 备注中手动添加关键信息

## 项目结构

```
pptx_to_video/
├── pptx_to_video.py      # 主程序入口
├── __main__.py           # python -m 入口
├── __init__.py           # 包初始化
├── ppt_parser.py         # PPT 解析
├── ppt_renderer.py       # PowerPoint 视频渲染
├── script_generator.py   # AI 脚本生成
├── tts_service.py        # 语音合成（并发处理）
├── ffmpeg_utils.py       # FFmpeg 音频合成
├── config.py             # 配置管理
├── check_dependencies.py # 依赖检查
├── utils.py              # 工具函数
├── prompts/              # AI 提示词
├── requirements.txt      # Python 依赖
├── config.ini            # 配置文件
├── .env.example          # 环境变量示例
├── input/                # 输入 PPT 目录
├── output/               # 输出视频目录
├── temp/                 # 临时文件目录
├── README.md             # 项目说明
├── CLAUDE.md             # 开发文档
└── MODELS.md             # 模型说明
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
    └── slide_001.mp3        # 各页音频
```

## 技术栈

- **AI 模型**: Claude Sonnet 4.6 / 智谱 GLM-4 / DeepSeek Chat / 通义千问 Plus
- **TTS**: Microsoft Edge TTS（并发处理）
- **PPT 处理**: python-pptx（解析）, PowerPoint COM API（Windows，视频导出）
- **音视频**: FFmpeg（音频合成）, FFprobe（时长检测）

## 相关文档

- [CLAUDE.md](CLAUDE.md) - 开发文档和架构说明

## 许可证

MIT License

## 贡献

欢迎提交 Issue 和 Pull Request！
