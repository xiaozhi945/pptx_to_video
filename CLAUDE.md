# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

PPTX 智能讲解视频生成器 - Automatically converts PowerPoint presentations into narrated videos using AI-generated scripts and text-to-speech.

**Pipeline**: PPT parsing → LLM script generation → TTS synthesis → PowerPoint video export (no audio) → FFmpeg audio composition

**Tech Stack**:
- Python 3.10+
- LLM: Claude Sonnet 4.6 / 智谱 GLM-4 / DeepSeek Chat / 通义千问 Plus
- TTS: Microsoft Edge TTS (concurrent processing)
- PPT processing: python-pptx (parsing), PowerPoint COM API (Windows, video export)
- Video: PowerPoint native video export + FFmpeg audio composition
- Audio/Video tools: FFmpeg, FFprobe

## Architecture

### Core Modules

1. **pptx_to_video.py** - Main orchestrator
   - Coordinates the full pipeline with FFmpeg composition workflow
   - **Pipeline order**: Parse → Generate scripts → TTS → Export video (no audio) → FFmpeg compose audio
   - Validates resource counts (slides, audio, scripts) before processing
   - Handles CLI arguments and error reporting
   - Checks and auto-installs dependencies on startup

2. **ppt_parser.py** - PPT content extraction
   - Extracts text, titles, notes from PPTX using python-pptx
   - Exports structured text for LLM consumption

3. **ppt_renderer.py** - PowerPoint video rendering with FFmpeg composition
   - **PowerPointRenderer**: Windows + PowerPoint, supports animation
   - Sets slide transition timing based on audio duration
   - Exports video WITHOUT audio using CreateVideo API
   - Uses FFmpeg to precisely compose audio into video
   - Handles CreateVideoStatus monitoring (status 3 may indicate success despite "failed" status)

4. **script_generator.py** - LLM script generation
   - Supports Claude (Anthropic), 智谱AI, DeepSeek, and 通义千问 providers
   - Generates one script per slide
   - Two-phase: analyze PPT structure, then generate per-slide narration scripts
   - Includes exponential backoff retry logic (2s → 4s → 8s)
   - Parses JSON from LLM responses (handles markdown code blocks)
   - Uses OpenAI-compatible API for DeepSeek and 通义千问

5. **tts_service.py** - Text-to-speech synthesis
   - Uses edge-tts for multi-language voice synthesis
   - **Concurrent processing**: generates all MP3s in parallel using asyncio.gather()
   - Generates audio files with slide prefix: slide_001.mp3, slide_002.mp3, etc.
   - Configurable voice, rate, pitch via .env or config.ini

6. **config.py** - Configuration management
   - Loads config.ini (highest priority), then .env, then defaults
   - Defines paths: input/, output/, temp/
   - Model selection based on LLM_PROVIDER
   - Performance settings: concurrent TTS, caching, retry strategy

7. **ffmpeg_utils.py** - FFmpeg/FFprobe utilities
   - **get_audio_duration()**: Extracts audio file duration using FFprobe
   - **find_ffmpeg()**: Locates FFmpeg executable (config.ini → system PATH)
   - **find_ffprobe()**: Locates FFprobe executable
   - **compose_audio_video()**: Composes multiple audio files into video with frame-level precision
   - Used by ppt_renderer.py for audio-video composition

8. **check_dependencies.py** - Dependency checker
   - Auto-checks required packages on startup
   - Prompts user to auto-install missing dependencies
   - Uses importlib.util.find_spec() to avoid importing problematic modules

### FFmpeg Audio Composition Architecture

The system uses **FFmpeg post-composition** for perfect audio-video synchronization:

**Workflow**:
1. **Parse PPT**: Extract slide content using python-pptx
2. **Generate Scripts**: LLM generates narration for each slide
3. **TTS Synthesis**: Edge TTS generates MP3 files (slide_001.mp3, slide_002.mp3, ...)
4. **Set Slide Timing**: PowerPoint slide transitions set to match audio durations
5. **Export Video (No Audio)**: PowerPoint's `CreateVideo()` exports MP4 with animations only
6. **FFmpeg Composition**: 
   - Concatenate all audio files in order
   - Merge concatenated audio with video
   - Frame-level precision (33.3ms at 30fps)

**Key Advantages**:
- ✅ **Perfect audio-video synchronization** (frame-level precision via FFmpeg)
- ✅ **Full animation support** (all PowerPoint animations preserved)
- ✅ **Audio plays immediately** when slide appears (not after animations)
- ✅ **Industry-standard workflow** (same as professional video editing)
- ✅ **Complete control** over audio timing

**Requirements**:
- Windows OS
- Microsoft PowerPoint 2013 or later
- pywin32 (for COM API access)
- FFmpeg and FFprobe

### FFmpeg Composition Details

**Audio Concatenation**:
```python
# Create concat list file
file 'path/to/slide_001.mp3'
file 'path/to/slide_002.mp3'
...

# Concatenate audio
ffmpeg -f concat -safe 0 -i concat_list.txt -c copy merged_audio.mp3
```

**Audio-Video Merge**:
```python
ffmpeg -i video_no_audio.mp4 -i merged_audio.mp3 \
  -c:v copy \      # Copy video stream (no re-encoding)
  -c:a aac \       # Encode audio to AAC
  -b:a 192k \      # Audio bitrate
  -shortest \      # Match shortest stream
  output.mp4
```

**Precision**:
- Frame-level accuracy: 33.3ms at 30fps, 16.7ms at 60fps
- Audio starts exactly when slide appears
- No delay or desynchronization issues

### PowerPoint COM API Details

**Slide Timing**:
```python
slide.SlideShowTransition.AdvanceOnTime = True
slide.SlideShowTransition.AdvanceTime = duration  # seconds
```

**Video Export (No Audio)**:
```python
presentation.CreateVideo(
    filename,
    UseTimingsAndNarrations=True,  # Use slide timings
    DefaultSlideDuration=5,        # Fallback (not used when AdvanceTime is set)
    VerticalResolution=1080,       # 1080p
    FramesPerSecond=30,
    Quality=80                     # 1-100
)
```

**Status Monitoring**:
- `CreateVideoStatus`: 0=未开始, 1=进行中, 2=完成, 3=失败
- **Important**: Status 3 (failed) may still produce a valid video file
- Always check file existence and size (>100KB) even when status shows failure

## Configuration

**config.ini** (highest priority):
```ini
[paths]
ffmpeg_path = D:\Soft\ffmpeg\bin\ffmpeg.exe
ffprobe_path = D:\Soft\ffmpeg\bin\ffprobe.exe

[llm]
provider = claude

[tts]
voice = zh-CN-XiaoxiaoNeural
rate = +0%
pitch = +0Hz

[video]
width = 1920
height = 1080

[performance]
enable_concurrent_tts = true
enable_cache = true
max_retries = 3
base_retry_delay = 2
```

**Environment Variables** (.env):
- `ANTHROPIC_API_KEY`: Claude API key
- `ZHIPUAI_API_KEY`: 智谱AI API key
- `DEEPSEEK_API_KEY`: DeepSeek API key
- `QIANWEN_API_KEY`: 通义千问 API key

## Usage

```bash
# Basic usage (processes all PPTX in input/)
python pptx_to_video.py

# Specify input file
python pptx_to_video.py --input path/to/presentation.pptx

# Specify output path
python pptx_to_video.py --output path/to/output.mp4

# Use specific LLM provider
python pptx_to_video.py --provider claude

# Skip TTS (use existing audio)
python pptx_to_video.py --skip-tts

# Skip video generation (only generate scripts and audio)
python pptx_to_video.py --skip-video

# List available PPTX files
python pptx_to_video.py --list
```

## Development Notes

### Audio-Video Synchronization

The system achieves perfect synchronization using FFmpeg post-composition:
1. PowerPoint exports video with correct slide timing (based on audio durations)
2. FFmpeg concatenates all audio files in order
3. FFmpeg merges audio with video using `-c:v copy` (no video re-encoding)
4. Result: Frame-level precision with zero desynchronization

**Why FFmpeg instead of PowerPoint audio embedding?**
- PowerPoint's `CreateVideo` with embedded audio has limitations with animations
- Audio may play after animations complete, not simultaneously
- FFmpeg provides complete control over audio timing
- Industry-standard approach used in professional video editing

### Error Handling

**PowerPoint CreateVideoStatus**:
- Status 3 (failed) doesn't always mean failure
- Always check if video file exists and has reasonable size (>100KB)
- Common false-positive: status shows "failed" but video is actually generated

**COM API Cleanup**:
- Always close presentation and quit PowerPoint in finally block
- Clean up temporary PPT and video files after composition

**FFmpeg Errors**:
- Check FFmpeg/FFprobe paths in config.ini if composition fails
- Verify audio files exist before calling compose_audio_video()
- FFmpeg errors are captured in stderr and raised as RuntimeError

### Performance

**Concurrent TTS**:
- All audio files generated in parallel using asyncio.gather()
- Significantly faster than sequential generation

**Caching**:
- Scripts cached in output/{name}/scripts.json
- Audio files cached in temp/{name}/slide_*.mp3
- User prompted to use cache when available

## Removed Features

### 2026-05-07: Migrated from Audio Embedding to FFmpeg Composition

**Old workflow** (removed):
- Parse → Scripts → TTS → Embed audio into PPT → Export video with audio

**New workflow** (current):
- Parse → Scripts → TTS → Export video (no audio) → FFmpeg compose audio

**Reason for migration**:
- PowerPoint's audio embedding caused sync issues with animations
- Audio would play after animations completed, not simultaneously
- FFmpeg provides frame-level precision and complete control
- Industry-standard workflow for professional video production

**Removed code**:
- Audio embedding logic using `AddMediaObject2()` and `AnimationSettings`
- Animation sequence manipulation (MainSequence.AddEffect, MoveTo, TriggerType)
- PlayOnEntry and animation synchronization attempts

### 2026-05-07: Removed Frame Extraction Workflow

**Old workflow** (removed earlier):
- Parse → Scripts → TTS → Render video → Extract frames → Re-encode with FFmpeg

**Reason for removal**:
- Frame extraction was complex and error-prone
- Required FFmpeg re-encoding of entire video
- PowerPoint native export is simpler and preserves animations

**Removed code**:
- LibreOfficeRenderer and PillowRenderer classes
- RendererFactory class
- video_creator.py module
- Frame extraction logic

## Known Issues

1. **PowerPoint CreateVideoStatus unreliable**: Status 3 may indicate success - always check file existence
2. **Windows-only**: Requires Windows + PowerPoint (no cross-platform support)
3. **COM API quirks**: Sometimes need to save, close, and reopen PPT for changes to take effect
4. **FFmpeg required**: Must have FFmpeg and FFprobe installed and configured

## Testing

Run the full pipeline:
```bash
python pptx_to_video.py --input test.pptx
```

Test individual components:
- **TTS only**: `python pptx_to_video.py --skip-video`
- **Script generation only**: `python pptx_to_video.py --skip-tts --skip-video`

**Test files**:
- `test_powerpoint_audio_embed.py`: Legacy audio embedding tests
- `test_audio_sync.py`: Audio synchronization experiments
- `check_animation_settings.py`: PowerPoint animation inspection
