"""FFmpeg 工具函数"""
from pathlib import Path
from typing import Optional

try:
    from .config import FFMPEG_PATH, FFPROBE_PATH
except ImportError:
    import sys
    sys.path.insert(0, str(Path(__file__).parent))
    from config import FFMPEG_PATH, FFPROBE_PATH


def find_ffmpeg() -> Optional[str]:
    """
    从 config.ini 读取 FFmpeg 路径

    Returns:
        FFmpeg 路径，如果未配置返回 None
    """
    if FFMPEG_PATH and Path(FFMPEG_PATH).exists():
        return FFMPEG_PATH
    return None


def find_ffprobe() -> Optional[str]:
    """
    从 config.ini 读取 FFprobe 路径

    Returns:
        FFprobe 路径，如果未配置返回 None
    """
    if FFPROBE_PATH and Path(FFPROBE_PATH).exists():
        return FFPROBE_PATH
    return None


def check_ffmpeg() -> bool:
    """
    检查 FFmpeg 是否可用

    Returns:
        True 如果 FFmpeg 和 FFprobe 都可用
    """
    ffmpeg = find_ffmpeg()
    ffprobe = find_ffprobe()

    if not ffmpeg or not ffprobe:
        print("警告: FFmpeg 未配置或路径不存在", flush=True)
        if not ffmpeg:
            print("  - 未找到 ffmpeg", flush=True)
        if not ffprobe:
            print("  - 未找到 ffprobe", flush=True)
        print("请在 config.ini 中配置 FFmpeg 路径", flush=True)
        return False

    print(f"  FFmpeg: {ffmpeg}", flush=True)
    print(f"  FFprobe: {ffprobe}", flush=True)
    return True
