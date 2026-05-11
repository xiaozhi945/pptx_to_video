# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

PPTX 智能讲解视频生成器 - Converts PowerPoint presentations into narrated videos using AI-generated scripts and text-to-speech.

**Pipeline**: PPT parsing → LLM script generation → TTS synthesis → PowerPoint video export (no audio) → FFmpeg audio composition

**Tech Stack**: Python 3.10+, python-pptx, PowerPoint COM API (Windows only), edge-tts, FFmpeg

## Commands

```bash
# Basic usage (processes all PPTX in input/)
python pptx_to_video.py

# Process a specific file
python pptx_to_video.py --input path/to/file.pptx

# Skip steps during development
python pptx_to_video.py --skip-tts --skip-video    # scripts only
python pptx_to_video.py --skip-video               # scripts + TTS

# Switch LLM provider
python pptx_to_video.py --provider claude|zhipu|deepseek|qianwen

# Print current configuration
python -c "import config; config.print_config()"

# List available PPTX files in input/
python pptx_to_video.py --list

# Clear cache to force regeneration
rm -rf output/*/scripts.json temp/*/slide_*.mp3
```

## Architecture

### Module dependency and import pattern

Every module uses a dual-import pattern to support both `python -m` and direct execution:

```python
try:
    from .config import ...
except ImportError:
    sys.path.insert(0, str(Path(__file__).parent))
    from config import ...
```

**Dependency graph**:
- `config.py` — no internal deps, loaded first by all others
- `ppt_parser.py` — only depends on python-pptx
- `script_generator.py` — depends on `config.py`, prompts/
- `tts_service.py` — depends on `config.py`
- `subtitle_generator.py` — depends on `config.py`, ffprobe for audio duration
- `ffmpeg_utils.py` — depends on `config.py`
- `ppt_renderer.py` — depends on `ffmpeg_utils.py`, win32com (Windows)
- `pptx_to_video.py` — orchestrator, imports everything except `utils.py`

### PPT parsing with special shape handling

`ppt_parser.py` extracts text from multiple shape types:

1. **Text frames** (`has_text_frame=True`) — standard text boxes, titles, placeholders
2. **Tables** (`shape_type=19`) — iterates through all cells to extract text
3. **Groups** (`shape_type=6`) — recursively extracts text from grouped shapes

The `_extract_shape_text()` method handles all three cases. **Important**: Office lock files (`~$*.pptx`) are automatically filtered out by `find_pptx_files()`.

### Two-phase LLM script generation

`script_generator.py` uses two prompts in sequence:

1. **`analyze_ppt.txt`** — LLM analyzes the full PPT text dump and returns structured JSON (topics, flow, key points per slide)
2. **`generate_script.txt`** — Uses the analysis + per-slide content to generate one narration script per slide

Both phases share the same retry logic (exponential backoff: 2s → 4s → 8s, configurable via `[performance]` in config.ini).

**Empty scripts**: If a slide has no extractable text (pure images, blank slides), the LLM may return an empty script. TTS will skip these slides, causing audio count mismatch. The parser now extracts table and group text to minimize this.

### Multi-provider LLM support

Four LLM providers are supported with unified interface in `script_generator.py`:

- **Claude**: Uses Anthropic SDK with streaming, supports custom `base_url` for proxy/relay services
- **ZhipuAI**: Uses ZhipuAI SDK
- **DeepSeek**: Uses OpenAI-compatible SDK with custom base URL
- **Qianwen**: Uses OpenAI-compatible SDK with Alibaba Cloud endpoint

All providers use the same retry logic and prompt templates.

**Model customization**: Models can be customized via `config.ini` or `.env` using the `model` or `LLM_MODEL` key. If not specified, defaults are used:
- Claude: `claude-sonnet-4-6`
- ZhipuAI: `glm-4`
- DeepSeek: `deepseek-chat`
- Qianwen: `qwen-plus`

**Adding new providers**: 
- OpenAI-compatible providers (Moonshot, Baichuan, MiniMax, etc.) can reuse the DeepSeek/Qianwen code path by adding a new `elif` branch with custom `base_url`
- Non-compatible providers (Gemini, Cohere) require dedicated SDK integration and separate handling in `_call_llm()`

### Language auto-detection from TTS voice

The system automatically detects script language from the configured TTS voice to ensure voice-text language matching:

- `config.py` contains `get_language_from_voice()` that extracts language code from voice name (e.g., `zh-CN-XiaoxiaoNeural` → Chinese)
- `SCRIPT_LANGUAGE` is automatically set based on `TTS_VOICE` configuration
- Supports 20+ languages including Chinese, English, Japanese, Korean, French, German, Spanish, etc.
- The language is passed to LLM prompts via `{language}` placeholder in `generate_script.txt`
- **Important**: Voice and script language must match for TTS to work. English voices cannot synthesize Chinese text and vice versa.

### TTS concurrent processing with word-level timestamps

`tts_service.py` uses `asyncio.gather()` to synthesize multiple audio files concurrently, providing ~6x speedup. 

**Critical**: Must set `boundary='WordBoundary'` in `edge_tts.Communicate()` to enable word-level timestamps for subtitle generation. Without this parameter, only `SentenceBoundary` events are emitted, resulting in empty `.timing.json` files.

Each audio generation produces two files:
- `slide_NNN.mp3` — audio file
- `slide_NNN.timing.json` — word-level timestamps (offset, duration, text)

Empty scripts are skipped (no audio generated), which can cause slide-audio count mismatch if not handled properly.

### FFmpeg composition (not embedding)

The system does NOT embed audio into PowerPoint. Instead:
1. PowerPoint exports video with slide timings matched to audio durations (no audio track)
2. FFmpeg concatenates all MP3s in order, then merges with `-c:v copy` (no video re-encode)
3. This gives frame-level precision (33.3ms at 30fps) and avoids PowerPoint's known animation-audio sync bugs

### Configuration priority

`config.ini` > `.env` > defaults. The `get_config()` helper in `config.py` enforces this chain.

**Important**: In `config.ini`, percent signs must be escaped as `%%` (e.g., `rate = +10%%`).

### Caching

- Scripts: `output/{pptx_name}/scripts.json`
- Audio: `temp/{pptx_name}/slide_*.mp3`
- Word timestamps: `temp/{pptx_name}/slide_*.timing.json`
- Cache is automatically used when `enable_cache = true` and file counts match

### Subtitle generation (three modes)

`subtitle_generator.py` supports three subtitle modes configured via `config.ini`:

1. **SRT mode** (`mode = srt`): Generates standalone `.srt` file
   - Fastest option (no video re-encoding)
   - Players load subtitle file separately
   - Can be manually edited

2. **Soft subtitle** (`mode = soft`): Embeds subtitle track into MP4
   - Subtitle can be toggled on/off in player
   - Uses `language` config for metadata (chi/eng/jpn/etc.)
   - No video re-encoding required

3. **Hard subtitle** (`mode = hard`): Burns subtitle into video frames
   - Best compatibility (all players)
   - Cannot be disabled
   - Requires video re-encoding (slow)
   - Uses white text + black outline (configurable `outline_width`) for visibility on any background

**Subtitle types**:
- `type = word`: Word-by-word display using `.timing.json` timestamps, grouped by `max_chars_per_line`
- `type = sentence`: Full sentence display using audio duration

**Auto-fallback**: If `.timing.json` is missing or empty, word-level generation automatically falls back to sentence-level.

**Important**: `config.ini` values cannot have inline comments. Use separate comment lines:
```ini
# Correct
outline_width = 2

# Wrong - will fail to parse
outline_width = 2  # comment here
```

## Platform Requirements

**Windows-only for video generation**: The PowerPoint COM API (`ppt_renderer.py`) requires:
- Native Windows environment (not WSL)
- Microsoft PowerPoint installed
- pywin32 package

`platform.system()` must return `'Windows'` for the renderer to be available. Running in WSL will cause `PowerPoint 不可用` error even if using Windows Python.

## Known Issues

1. **PowerPoint CreateVideoStatus unreliable**: Status 3 ("failed") may still produce valid video — code now checks file existence and size (>100KB) and marks as success if file is valid
2. **Windows-only**: Requires Windows + PowerPoint + pywin32
3. **COM API**: Sometimes need to save, close, and reopen PPT for slide timing changes to take effect
4. **test_renderer.py is stale**: Imports `RendererFactory`, `LibreOfficeRenderer`, `PillowRenderer` which were removed. The file will fail on import — don't use it as a reference
5. **Empty scripts cause audio mismatch**: If LLM generates empty script for a slide, TTS skips it, causing `len(audio_files) != len(slides)` error. Solution: ensure parser extracts all text (tables, groups) or handle empty scripts specially
6. **WSL incompatibility**: Must run in native Windows command prompt/PowerShell, not WSL terminal, for PowerPoint COM API to work
7. **generate_with_animation() broken**: `ScriptGenerator.generate_with_animation()` references `self.animation_script_prompt` which is never loaded in `__init__()`. This method will raise `AttributeError` if called. Only `analyze_ppt.txt` and `generate_script.txt` are loaded
8. **Slide indexing**: `scripts.json` uses `slide_index` (0-based) but audio files use `slide_NNN.mp3` (1-based). Subtitle generator handles this conversion: `slide_num = slide_index + 1`
9. **edge-tts WordBoundary**: Must explicitly set `boundary='WordBoundary'` parameter, otherwise only `SentenceBoundary` events are emitted and `.timing.json` files will be empty arrays.
