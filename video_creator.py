"""视频合成模块 - 使用 FFmpeg"""
import subprocess
import os
import sys
from pathlib import Path
from typing import List, Dict, Any

try:
    from .config import VIDEO_WIDTH, VIDEO_HEIGHT
except ImportError:
    sys.path.insert(0, str(Path(__file__).parent))
    from config import VIDEO_WIDTH, VIDEO_HEIGHT


class VideoCreator:
    """使用 FFmpeg 合成视频"""

    def __init__(self, output_dir: str, temp_dir: str):
        self.output_dir = Path(output_dir)
        self.temp_dir = Path(temp_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.temp_dir.mkdir(parents=True, exist_ok=True)

    def create_segment(self, image_path: str, audio_path: str, output_path: str, duration: float = None) -> bool:
        """创建单个视频片段（图片 + 音频）"""
        if duration is None:
            # 获取音频时长
            duration = self._get_audio_duration(audio_path)

        cmd = [
            "ffmpeg", "-y",
            "-loop", "1",
            "-i", image_path,
            "-i", audio_path,
            "-c:v", "libx264",
            "-t", str(duration),
            "-pix_fmt", "yuv420p",
            "-vf", f"scale={VIDEO_WIDTH}:{VIDEO_HEIGHT}:force_original_aspect_ratio=decrease,pad={VIDEO_WIDTH}:{VIDEO_HEIGHT}:(ow-iw)/2:(oh-ih)/2",
            "-shortest",  # 使用音频长度
            "-movflags", "+faststart",
            output_path
        ]

        try:
            subprocess.run(cmd, check=True, capture_output=True)
            return True
        except subprocess.CalledProcessError as e:
            # 尝试多种编码解码错误信息
            error_msg = str(e)
            if e.stderr:
                try:
                    error_msg = e.stderr.decode('utf-8')
                except UnicodeDecodeError:
                    try:
                        error_msg = e.stderr.decode('gbk')  # Windows 中文编码
                    except:
                        error_msg = e.stderr.decode('utf-8', errors='ignore')
            print(f"FFmpeg 错误: {error_msg}", flush=True)
            return False

    def _get_audio_duration(self, audio_path: str) -> float:
        """获取音频时长"""
        cmd = [
            "ffprobe", "-v", "error",
            "-show_entries", "format=duration",
            "-of", "default=noprint_wrappers=1:nokey=1",
            audio_path
        ]

        try:
            result = subprocess.run(cmd, capture_output=True, text=True, check=True)
            duration = float(result.stdout.strip())
            if duration <= 0:
                raise ValueError(f"音频时长无效: {duration}")
            return duration
        except Exception as e:
            print(f"错误: 无法获取音频时长 {audio_path}: {e}", flush=True)
            print(f"请检查音频文件是否完整", flush=True)
            raise

    def concatenate_videos(self, video_list: List[str], output_path: str) -> bool:
        """合并多个视频片段"""
        # 创建临时文件列表
        list_file = self.temp_dir / "concat_list.txt"

        # 使用 UTF-8 编码写入，路径使用正斜杠
        with open(list_file, "w", encoding="utf-8") as f:
            for video in video_list:
                # 转换为绝对路径并使用正斜杠（FFmpeg 兼容性更好）
                video_path = str(Path(video).resolve()).replace("\\", "/")
                f.write(f"file '{video_path}'\n")

        cmd = [
            "ffmpeg", "-y",
            "-f", "concat",
            "-safe", "0",
            "-i", str(list_file),
            "-c", "copy",
            output_path
        ]

        try:
            result = subprocess.run(cmd, check=True, capture_output=True)
            return True
        except subprocess.CalledProcessError as e:
            # 尝试多种编码解码错误信息
            error_msg = str(e)
            if e.stderr:
                try:
                    error_msg = e.stderr.decode('utf-8')
                except (UnicodeDecodeError, AttributeError):
                    try:
                        error_msg = e.stderr.decode('gbk') if isinstance(e.stderr, bytes) else e.stderr
                    except:
                        error_msg = str(e.stderr)
            print(f"视频合并错误: {error_msg}", flush=True)
            return False

    def create_video(self, thumbnails: List[str], audio_files: List[str], output_path: str) -> bool:
        """创建完整视频"""
        # 验证输入数量匹配
        if len(thumbnails) != len(audio_files):
            print(f"错误: 幻灯片数量({len(thumbnails)})与音频数量({len(audio_files)})不匹配", flush=True)
            print(f"  幻灯片: {len(thumbnails)} 个", flush=True)
            print(f"  音频文件: {len(audio_files)} 个", flush=True)
            return False

        if len(thumbnails) == 0:
            print("错误: 没有幻灯片可处理", flush=True)
            return False

        print(f"开始合成视频，共 {len(thumbnails)} 个片段...", flush=True)

        video_segments = []

        for idx, (image, audio) in enumerate(zip(thumbnails, audio_files)):
            segment_path = self.temp_dir / f"segment_{idx+1:03d}.mp4"

            print(f"  处理片段 [{idx+1}/{len(thumbnails)}]...", flush=True)

            # 验证文件存在
            if not Path(image).exists():
                print(f"    ❌ 图片文件不存在: {image}", flush=True)
                continue
            if not Path(audio).exists():
                print(f"    ❌ 音频文件不存在: {audio}", flush=True)
                continue

            try:
                if self.create_segment(image, audio, str(segment_path)):
                    video_segments.append(str(segment_path))
                    print(f"    ✓ 片段 {idx+1} 创建完成", flush=True)
                else:
                    print(f"    ❌ 片段 {idx+1} 创建失败，跳过", flush=True)
            except Exception as e:
                print(f"    ❌ 片段 {idx+1} 创建异常: {e}", flush=True)

        if not video_segments:
            print("错误: 没有成功创建任何视频片段", flush=True)
            return False

        # 合并视频
        print("合并视频片段...", flush=True)
        return self.concatenate_videos(video_segments, output_path)
