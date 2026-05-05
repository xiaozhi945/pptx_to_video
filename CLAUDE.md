# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

PPTX 智能讲解视频生成器 - Automatically converts PowerPoint presentations into narrated videos using AI-generated scripts and text-to-speech.

**Pipeline**: PPT parsing → LLM analysis → script generation → TTS synthesis → FFmpeg video composition

**Tech Stack**:
- Python 3.10+
- LLM: Claude Sonnet 4.6 / 智谱 GLM-4 / DeepSeek Chat / 通义千问 Plus
- TTS: Microsoft Edge TTS (concurrent processing)
- PPT processing: python-pptx, PowerPoint COM API (Windows), LibreOffice (cross-platform)
- PDF to image: pdf2image + Poppler
- Video: FFmpeg
- **Animation support**: PowerPoint COM API (Windows only)

## Architecture

### Core Modules

1. **pptx_to_video.py** - Main orchestrator
   - Coordinates the full pipeline
   - Validates resource counts (slides, audio, scripts) before video synthesis
   - Handles CLI arguments and error reporting
   - Checks and auto-installs dependencies on startup

2. **ppt_parser.py** - PPT content extraction
   - Extracts text, titles, notes from PPTX using python-pptx
   - Uses adaptive rendering backend (PowerPoint/LibreOffice/Pillow)
   - Exports structured text for LLM consumption

2a. **ppt_renderer.py** - Multi-backend PPT rendering
   - **PowerPointRenderer**: Windows + PowerPoint, supports animation export
   - **LibreOfficeRenderer**: Cross-platform, static rendering only
   - **PillowRenderer**: Fallback text rendering
   - **RendererFactory**: Auto-detects best available backend

3. **script_generator.py** - LLM script generation
   - Supports Claude (Anthropic), 智谱AI, DeepSeek, and 通义千问 providers
   - Two-phase: analyze PPT structure, then generate per-slide narration scripts
   - Includes exponential backoff retry logic (2s → 4s → 8s)
   - Parses JSON from LLM responses (handles markdown code blocks)
   - Uses OpenAI-compatible API for DeepSeek and 通义千问

4. **tts_service.py** - Text-to-speech synthesis
   - Uses edge-tts for multi-language voice synthesis
   - **Concurrent processing**: generates all MP3s in parallel using asyncio.gather()
   - Generates one MP3 per slide (slide_001.mp3, slide_002.mp3, etc.)
   - Configurable voice, rate, pitch via .env or config.ini

5. **video_creator.py** - FFmpeg video composition
   - Creates segments (image + audio) then concatenates
   - Enforces 1920x1080 resolution with aspect-preserving scaling
   - Validates input counts before processing
   - Uses utils.decode_subprocess_error() for error handling

6. **config.py** - Configuration management
   - Loads config.ini (highest priority), then .env, then defaults
   - Defines paths: input/, output/, temp/
   - Model selection based on LLM_PROVIDER
   - Performance settings: concurrent TTS, caching, retry strategy

7. **utils.py** - Utility functions
   - decode_subprocess_error(): handles UTF-8/GBK encoding for subprocess errors

8. **ffmpeg_utils.py** - FFmpeg path detection
   - Centralized FFmpeg/FFprobe path detection
   - Checks config.ini, .env, system PATH, and fallback search paths
   - Used by video_creator.py for video synthesis

9. **check_dependencies.py** - Dependency checker
   - Auto-checks required packages on startup
   - Prompts user to auto-install missing dependencies
   - Uses importlib.util.find_spec() to avoid importing problematic modules

### Rendering Backend Architecture

The system uses an **adaptive multi-backend architecture** for PPT rendering:

**Backend Selection Priority**:
1. **PowerPoint COM API** (Windows + PowerPoint installed)
   - ✅ Full animation support
   - ✅ Perfect fidelity to original PPT
   - ✅ Exports each animation step as separate frame
   - ❌ Windows-only, requires Microsoft PowerPoint

2. **LibreOffice** (Cross-platform)
   - ✅ Cross-platform (Windows/Linux/macOS)
   - ✅ High-quality static rendering
   - ❌ No animation support (exports final state only)
   - Requires: LibreOffice + Poppler + pdf2image

3. **Pillow** (Fallback)
   - ✅ Always available
   - ✅ No external dependencies
   - ❌ Text-only rendering (no images/shapes)
   - ❌ No animation support

**Configuration**:
```ini
[rendering]
backend = auto  # auto, powerpoint, libreoffice, pillow
enable_animation = true  # Only effective with PowerPoint backend
```

**Auto-detection logic** (when `backend = auto`):
- Windows + PowerPoint installed → PowerPointRenderer
- LibreOffice available → LibreOfficeRenderer  
- Otherwise → PillowRenderer (fallback)

**Animation Frame Format**:
Each rendered frame is a dict with:
- `path`: Image file path
- `slide_index`: Slide number (0-indexed)
- `animation_step`: Animation step (0 = initial state, 1+ = animation steps)

Static renderers always return `animation_step = 0`.

### Directory Structure

```
input/          # Place .pptx files here
output/         # Final videos and scripts.json
  <filename>/
    <filename>.mp4
    scripts.json
temp/           # Intermediate files (images, audio, PDFs)
  <filename>/
    content.txt
    slide_*.mp3
    *_slide_*.png
prompts/        # LLM prompt templates
  analyze_ppt.txt
  generate_script.txt
```

## Development Commands

### Setup

```bash
# Install dependencies (auto-prompted on first run)
pip install -r requirements.txt

# Configure environment
cp .env.example .env
# Edit .env with API keys

# Optional: create config.ini for advanced settings
```

### Running

```bash
# Process all .pptx files in input/
python -m pptx_to_video

# Process specific file
python -m pptx_to_video --input path/to/file.pptx

# List available files
python -m pptx_to_video --list

# Skip TTS (script generation only)
python -m pptx_to_video --skip-tts

# Skip video synthesis
python -m pptx_to_video --skip-video

# Switch LLM provider
python -m pptx_to_video --provider zhipu
```

### Testing Individual Components

```python
# Test PPT parsing
from ppt_parser import PPTParser
parser = PPTParser("input", "temp")
data = parser.parse(Path("input/example.pptx"))

# Test script generation
from script_generator import ScriptGenerator
generator = ScriptGenerator(api_key, "prompts", "claude")
result = generator.generate(ppt_text, ppt_data)

# Test TTS (concurrent)
from tts_service import TTSService
tts = TTSService("temp/example")
audio_files = tts.synthesize(scripts)

# Test video creation
from video_creator import VideoCreator
creator = VideoCreator("output/example", "temp/example")
creator.create_video(thumbnails, audio_files, "output.mp4")

# Check dependencies
python check_dependencies.py

# View current config
python -c "import config; config.print_config()"
```

## Key Implementation Details

### Auto Dependency Installation

check_dependencies.py runs on startup and:
1. Uses importlib.util.find_spec() to check for missing packages (avoids importing)
2. Prompts user to auto-install if dependencies are missing
3. Calls pip install via subprocess if user agrees
4. python-magic is excluded (can cause segfaults)

### Caching Mechanism

The system caches intermediate results to avoid redundant computation:
- **Script cache**: Checks for `output/<name>/scripts.json` before calling LLM
- **Audio cache**: Checks for `temp/<name>/slide_*.mp3` before running TTS
- Controlled by `ENABLE_CACHE` in config (default: true)
- Automatically uses cache when available, saving API costs and time

### TTS Concurrent Processing

tts_service.py uses `asyncio.gather()` to generate all audio files concurrently:
- Old: serial processing with asyncio.run() per file (~60s for 20 slides)
- New: batch processing with asyncio.gather() (~10s for 20 slides)
- ~6x speedup for multi-slide presentations
- Controlled by `ENABLE_CONCURRENT_TTS` in config (default: true)

### API Retry Strategy

script_generator.py uses exponential backoff for retries:
- Attempt 1: wait 2s (base_retry_delay * 2^0)
- Attempt 2: wait 4s (base_retry_delay * 2^1)
- Attempt 3: wait 8s (base_retry_delay * 2^2)
- Configurable via MAX_RETRIES and BASE_RETRY_DELAY in config

### Script Index Handling

The system auto-detects whether LLM returns 0-indexed or 1-indexed slide_index and normalizes to 0-indexed (pptx_to_video.py:143-147). Always validate script count matches slide count.

### Audio-Video Synchronization

Critical validation happens at multiple stages:
- After script generation: check script count vs slide count
- After TTS: check audio count vs script count  
- Before video synthesis: check thumbnails vs audio files (pptx_to_video.py:196-203)

The system refuses to synthesize video if counts don't match to prevent audio-video desync.

### Configuration System

Three-tier configuration with priority: `config.ini > .env > defaults`

**config.py** provides:
- `get_config(section, key, env_key, default)` - reads from config.ini or env
- `get_config_bool()` - for boolean values
- `get_config_int()` - for integer values

**config.ini** uses INI format:
- Percent signs must be escaped: `rate = +20%%` (not `+20%`)
- Supports sections: [paths], [llm], [tts], [video], [performance]

### LibreOffice Detection

ppt_parser.py checks for LibreOffice in:
1. LIBREOFFICE_PATH from config.ini or .env
2. System PATH (shutil.which("soffice"))
3. LIBREOFFICE_SEARCH_PATHS from config.py

If unavailable, falls back to text-based thumbnail rendering with PIL.

### Poppler Detection

ppt_parser._find_poppler() checks for Poppler (pdf2image dependency) in:
1. POPPLER_PATH from config.ini or .env
2. System PATH (shutil.which("pdftoppm"))
3. POPPLER_SEARCH_PATHS from config.py

Used by LibreOffice workflow to convert PDF to images. If unavailable, pdf2image will attempt to use system default.

### Font Loading

ppt_parser._load_font() loads fonts with priority:
1. FONT_TITLE/FONT_BODY from config.ini or .env
2. FONT_TITLE_FALLBACK/FONT_BODY_FALLBACK from config.py
3. PIL default font

Supports Windows/Linux/macOS font paths.

### FFmpeg Path Handling

ffmpeg_utils.py provides centralized FFmpeg path detection:
- get_ffmpeg_path() and get_ffprobe_path() check:
  1. FFMPEG_PATH/FFPROBE_PATH from config.ini or .env
  2. System PATH
  3. FFMPEG_SEARCH_PATHS/FFPROBE_SEARCH_PATHS from config.py

video_creator.py:
- Converts Windows backslashes to forward slashes for FFmpeg compatibility
- Uses UTF-8 encoding for concat lists to handle Chinese filenames
- Error messages decoded via utils.decode_subprocess_error() (tries UTF-8, GBK, then UTF-8 with errors ignored)

### LLM Provider Abstraction

script_generator.py provides a unified interface for multiple LLM providers:
- **Claude**: Uses Anthropic SDK
- **智谱AI**: Uses ZhipuAI SDK
- **DeepSeek**: Uses OpenAI-compatible API (openai SDK)
- **通义千问**: Uses OpenAI-compatible API (openai SDK)

The `_call_llm()` method handles API differences. Model names configured in config.py based on LLM_PROVIDER.

Supported providers: `claude`, `zhipu`, `deepseek`, `qianwen`

## Configuration

### Environment Variables (.env)

```bash
LLM_PROVIDER=claude          # or zhipu
ANTHROPIC_API_KEY=sk-ant-... # Claude API key
ZHIPUAI_API_KEY=...          # 智谱AI API key (backup)
TTS_VOICE=zh-CN-XiaoxiaoNeural  # Voice ID
TTS_RATE=+0%                 # Speech rate (-50% to +100%)
TTS_PITCH=+0Hz               # Pitch adjustment (-50Hz to +50Hz)
```

### Configuration File (config.ini) - Recommended

The project supports `config.ini` for detailed configuration with higher priority than `.env`:

```ini
[paths]
libreoffice_path = D:\Soft\LibreOffice\program\soffice.exe
font_title = C:/Windows/Fonts/msyh.ttc
font_body = C:/Windows/Fonts/simhei.ttf

[llm]
provider = claude

[rendering]
backend = auto  # auto, powerpoint, libreoffice, pillow
enable_animation = true  # Enable animation support (PowerPoint only)

[tts]
voice = zh-CN-XiaoxiaoNeural
rate = +20%%  # Note: % must be escaped as %%
pitch = +0Hz

[video]
width = 1920
height = 1080

[performance]
enable_concurrent_tts = true  # ~6x speedup
enable_cache = true           # saves API costs
max_retries = 3
base_retry_delay = 2
```

**Configuration Priority**: `config.ini > .env > defaults`

### Common TTS Voices

- Chinese: zh-CN-XiaoxiaoNeural (female), zh-CN-YunxiNeural (male)
- English: en-US-AriaNeural (female), en-US-GuyNeural (male)
- Japanese: ja-JP-NanamiNeural (female), ja-JP-KeitaNeural (male)

Full list: https://speech.microsoft.com/portal/voicegallery

## External Dependencies

Must be installed and in PATH:
- **FFmpeg** - Video encoding (ffmpeg, ffprobe) - REQUIRED
- **LibreOffice** - PPT rendering (optional, for static high-quality rendering)
- **Poppler** - PDF to image conversion (required if using LibreOffice)
- **Microsoft PowerPoint** - Animation support (optional, Windows only)

Windows installation via Chocolatey:
```powershell
choco install ffmpeg poppler
```

For animation support on Windows, install Microsoft PowerPoint and:
```bash
pip install pywin32
```

## Common Issues

### Script Count Mismatch
If LLM returns wrong number of scripts, check prompts/generate_script.txt. The system will warn but continue processing.

### Chinese Path Encoding
All file operations use UTF-8 encoding. FFmpeg concat lists use forward slashes for cross-platform compatibility.

### API Rate Limits
script_generator.py uses exponential backoff retry logic. If still hitting rate limits, switch providers or wait for quota reset.

### Missing Audio Files
If TTS fails silently, check edge-tts installation and network connectivity. The system validates audio file existence before video synthesis.

### Using Cache
The system automatically detects and uses cached scripts/audio. To force regeneration:
- Delete `output/<name>/scripts.json` for script cache
- Delete `temp/<name>/slide_*.mp3` for audio cache
- Or set `enable_cache = false` in config.ini

### Dependency Issues
Run `python check_dependencies.py` to check for missing packages. The main program auto-checks on startup.

### Config Not Working
- Check priority: config.ini > .env > defaults
- Verify INI syntax: `%` must be `%%`
- Run `python -c "import config; config.print_config()"` to see active config

## Recent Optimizations (2026-05-05)

1. **PPT Animation Support** - PowerPoint COM API for Windows animation rendering
2. **Adaptive Rendering Backend** - Auto-selects best renderer (PowerPoint/LibreOffice/Pillow)
3. **Cross-platform Compatibility** - Graceful degradation on Linux (static rendering)
4. **TTS Concurrent Processing** - ~6x speedup using asyncio.gather()
5. **Intermediate Result Caching** - saves API costs by reusing scripts/audio
6. **Exponential Backoff Retry** - smarter API retry strategy
7. **Configurable Paths** - LibreOffice and fonts via config.ini
8. **Unified Error Handling** - utils.decode_subprocess_error()
9. **Auto Dependency Installation** - check_dependencies.py on startup

See git history for detailed changes.
