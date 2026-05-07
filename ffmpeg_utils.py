"""FFmpeg 工具函数"""
import subprocess
import json
from pathlib import Path
from typing import Optional

try:
    from .config import FFPROBE_PATH
except ImportError:
    import sys
    sys.path.insert(0, str(Path(__file__).parent))
    from config import FFPROBE_PATH


def find_ffprobe() -> Optional[str]:
    """
    从 config.ini 读取 FFprobe 路径

    Returns:
        FFprobe 路径，如果未配置返回 None
    """
    if FFPROBE_PATH and Path(FFPROBE_PATH).exists():
        return FFPROBE_PATH
    return None


def check_ffprobe() -> bool:
    """
    检查 FFprobe 是否可用

    Returns:
        True 如果 FFprobe 可用
    """
    ffprobe = find_ffprobe()

    if not ffprobe:
        print("警告: FFprobe 未配置或路径不存在", flush=True)
        print("请在 config.ini 中配置 FFprobe 路径", flush=True)
        return False

    print(f"  FFprobe: {ffprobe}", flush=True)
    return True


def get_audio_duration(audio_path: Path) -> float:
    """
    获取音频文件的时长（秒）

    Args:
        audio_path: 音频文件路径

    Returns:
        音频时长（秒），如果获取失败返回 0.0
    """
    ffprobe = find_ffprobe()
    if not ffprobe:
        raise RuntimeError("未找到 FFprobe")

    cmd = [
        ffprobe,
        "-v", "error",
        "-show_entries", "format=duration",
        "-of", "json",
        str(audio_path)
    ]

    try:
        result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
        if result.returncode == 0:
            data = json.loads(result.stdout)
            duration = float(data.get("format", {}).get("duration", 0.0))
            return duration
        else:
            print(f"  ⚠️  无法获取音频时长: {audio_path.name}", flush=True)
            return 0.0
    except Exception as e:
        print(f"  ⚠️  获取音频时长失败: {e}", flush=True)
        return 0.0


def find_ffmpeg() -> Optional[str]:
    """
    查找 FFmpeg 可执行文件

    Returns:
        FFmpeg 路径，如果未找到返回 None
    """
    # 尝试从配置读取
    try:
        from config import FFMPEG_PATH
        if FFMPEG_PATH and Path(FFMPEG_PATH).exists():
            return FFMPEG_PATH
    except (ImportError, AttributeError):
        pass

    # 尝试从系统 PATH 查找
    import shutil
    ffmpeg = shutil.which("ffmpeg")
    if ffmpeg:
        return ffmpeg

    return None


def check_ffmpeg() -> bool:
    """
    检查 FFmpeg 是否可用

    Returns:
        True 如果 FFmpeg 可用
    """
    ffmpeg = find_ffmpeg()

    if not ffmpeg:
        print("警告: FFmpeg 未找到", flush=True)
        print("请安装 FFmpeg 或在 config.ini 中配置路径", flush=True)
        return False

    print(f"  FFmpeg: {ffmpeg}", flush=True)
    return True


def compose_audio_video(video_path: Path, audio_files: list[Path], output_path: Path) -> Path:
    """
    使用 FFmpeg 将多个音频文件精确合成到视频中

    流程：
    1. 计算每个音频的时间轴位置
    2. 使用 FFmpeg concat 拼接所有音频
    3. 将拼接后的音频与视频合成

    Args:
        video_path: 无音频的视频文件路径
        audio_files: 音频文件列表（按幻灯片顺序）
        output_path: 输出视频路径

    Returns:
        输出视频路径
    """
    ffmpeg = find_ffmpeg()
    if not ffmpeg:
        raise RuntimeError("未找到 FFmpeg")

    if not video_path.exists():
        raise FileNotFoundError(f"视频文件不存在: {video_path}")

    if not audio_files:
        raise ValueError("音频文件列表为空")

    # 验证所有音频文件存在
    for audio_file in audio_files:
        if not audio_file.exists():
            raise FileNotFoundError(f"音频文件不存在: {audio_file}")

    print(f"  使用 FFmpeg 合成音视频...", flush=True)
    print(f"    视频: {video_path.name}", flush=True)
    print(f"    音频: {len(audio_files)} 个文件", flush=True)

    # 创建临时文件列表（用于 concat）
    concat_list_path = output_path.parent / "ffmpeg_concat_list.txt"
    with open(concat_list_path, "w", encoding="utf-8") as f:
        for audio_file in audio_files:
            # FFmpeg concat 需要使用绝对路径，并转义特殊字符
            abs_path = audio_file.resolve()
            # Windows 路径需要转换为 Unix 风格并转义
            path_str = str(abs_path).replace("\\", "/").replace("'", "'\\''")
            f.write(f"file '{path_str}'\n")

    try:
        # 方案：先拼接音频，再与视频合成
        temp_audio = output_path.parent / "temp_merged_audio.mp3"

        # 步骤1：拼接所有音频
        print(f"  [1/2] 拼接音频文件...", flush=True)
        concat_cmd = [
            ffmpeg,
            "-f", "concat",
            "-safe", "0",
            "-i", str(concat_list_path),
            "-c", "copy",
            "-y",
            str(temp_audio)
        ]

        result = subprocess.run(
            concat_cmd,
            capture_output=True,
            text=True,
            encoding='utf-8'
        )

        if result.returncode != 0:
            raise RuntimeError(f"音频拼接失败: {result.stderr}")

        print(f"    ✓ 音频已拼接: {temp_audio.name}", flush=True)

        # 步骤2：将拼接后的音频与视频合成
        print(f"  [2/2] 合成音视频...", flush=True)
        merge_cmd = [
            ffmpeg,
            "-i", str(video_path),
            "-i", str(temp_audio),
            "-c:v", "copy",  # 视频流直接复制，不重新编码
            "-c:a", "aac",   # 音频编码为 AAC
            "-b:a", "192k",  # 音频比特率
            "-shortest",     # 以最短的流为准
            "-y",
            str(output_path)
        ]

        result = subprocess.run(
            merge_cmd,
            capture_output=True,
            text=True,
            encoding='utf-8'
        )

        if result.returncode != 0:
            raise RuntimeError(f"音视频合成失败: {result.stderr}")

        print(f"    ✓ 音视频已合成: {output_path.name}", flush=True)

        # 清理临时文件
        if temp_audio.exists():
            temp_audio.unlink()
        if concat_list_path.exists():
            concat_list_path.unlink()

        return output_path

    except Exception as e:
        # 清理临时文件
        if concat_list_path.exists():
            try:
                concat_list_path.unlink()
            except:
                pass
        raise e
