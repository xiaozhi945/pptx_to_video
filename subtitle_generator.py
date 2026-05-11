"""字幕生成模块"""
import json
from pathlib import Path
from typing import List, Dict
import subprocess
import sys

try:
    from .config import get_config, get_config_int, get_config_bool
except ImportError:
    sys.path.insert(0, str(Path(__file__).parent))
    from config import get_config, get_config_int, get_config_bool


class SubtitleGenerator:
    """字幕生成器"""

    def __init__(self, ffprobe_path: str = None):
        """
        初始化字幕生成器

        Args:
            ffprobe_path: ffprobe 可执行文件路径
        """
        if ffprobe_path is None:
            ffprobe_path = get_config("paths", "ffprobe_path", "FFPROBE_PATH", "ffprobe")
        self.ffprobe_path = ffprobe_path

    def get_audio_duration(self, audio_path: str) -> float:
        """
        获取音频文件时长（秒）

        Args:
            audio_path: 音频文件路径

        Returns:
            音频时长（秒）
        """
        try:
            cmd = [
                self.ffprobe_path,
                "-v", "error",
                "-show_entries", "format=duration",
                "-of", "default=noprint_wrappers=1:nokey=1",
                audio_path
            ]
            result = subprocess.run(cmd, capture_output=True, text=True, check=True)
            return float(result.stdout.strip())
        except Exception as e:
            print(f"⚠️  获取音频时长失败 {audio_path}: {e}")
            return 0.0

    def format_srt_time(self, seconds: float) -> str:
        """
        将秒数转换为 SRT 时间格式 (HH:MM:SS,mmm)

        Args:
            seconds: 秒数

        Returns:
            SRT 格式时间字符串
        """
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        secs = int(seconds % 60)
        millis = int((seconds % 1) * 1000)
        return f"{hours:02d}:{minutes:02d}:{secs:02d},{millis:03d}"

    def generate_srt(
        self,
        scripts: List[Dict[str, str]],
        audio_dir: Path,
        output_path: Path
    ) -> bool:
        """
        生成 SRT 字幕文件

        Args:
            scripts: 脚本列表，每个元素包含 slide_number 和 script
            audio_dir: 音频文件目录
            output_path: 输出 SRT 文件路径

        Returns:
            是否成功生成
        """
        try:
            srt_content = []
            current_time = 0.0
            subtitle_index = 1

            for script_item in scripts:
                # 支持 slide_index（从0开始）或 slide_number（从1开始）
                slide_index = script_item.get("slide_index")
                if slide_index is not None:
                    # slide_index 从 0 开始，音频文件从 001 开始
                    slide_num = slide_index + 1
                else:
                    # 兼容旧格式的 slide_number
                    slide_num = script_item.get("slide_number", 1)

                script_text = script_item.get("script", "").strip()

                # 跳过空脚本
                if not script_text:
                    continue

                # 查找对应的音频文件
                audio_file = audio_dir / f"slide_{slide_num:03d}.mp3"

                if not audio_file.exists():
                    print(f"⚠️  音频文件不存在: {audio_file}")
                    continue

                # 获取音频时长
                duration = self.get_audio_duration(str(audio_file))
                if duration <= 0:
                    continue

                # 生成字幕条目
                start_time = current_time
                end_time = current_time + duration

                # SRT 格式：
                # 序号
                # 开始时间 --> 结束时间
                # 字幕文本
                # 空行
                srt_entry = f"{subtitle_index}\n"
                srt_entry += f"{self.format_srt_time(start_time)} --> {self.format_srt_time(end_time)}\n"
                srt_entry += f"{script_text}\n"

                srt_content.append(srt_entry)

                current_time = end_time
                subtitle_index += 1

            # 写入 SRT 文件
            output_path.parent.mkdir(parents=True, exist_ok=True)
            with open(output_path, "w", encoding="utf-8") as f:
                f.write("\n".join(srt_content))

            print(f"✅ 字幕文件已生成: {output_path}")
            print(f"   共 {len(srt_content)} 条字幕，总时长 {current_time:.1f} 秒")
            return True

        except Exception as e:
            print(f"❌ 生成字幕失败: {e}")
            import traceback
            traceback.print_exc()
            return False

    def generate_word_level_srt(
        self,
        scripts: List[Dict[str, str]],
        audio_dir: Path,
        output_path: Path,
        max_chars_per_line: int = 42
    ) -> bool:
        """
        生成逐词显示的 SRT 字幕文件（基于 TTS 时间戳）

        Args:
            scripts: 脚本列表，每个元素包含 slide_index 和 script
            audio_dir: 音频文件目录（包含 .timing.json 文件）
            output_path: 输出 SRT 文件路径
            max_chars_per_line: 每行最大字符数（用于分组单词）

        Returns:
            是否成功生成
        """
        try:
            srt_content = []
            subtitle_index = 1
            global_offset = 0.0  # 全局时间偏移（秒）

            for script_item in scripts:
                # 支持 slide_index（从0开始）或 slide_number（从1开始）
                slide_index = script_item.get("slide_index")
                if slide_index is not None:
                    slide_num = slide_index + 1
                else:
                    slide_num = script_item.get("slide_number", 1)

                script_text = script_item.get("script", "").strip()

                # 跳过空脚本
                if not script_text:
                    continue

                # 查找对应的时间戳文件
                timing_file = audio_dir / f"slide_{slide_num:03d}.timing.json"

                if not timing_file.exists():
                    print(f"⚠️  时间戳文件不存在: {timing_file}")
                    # 降级为整句字幕
                    audio_file = audio_dir / f"slide_{slide_num:03d}.mp3"
                    if audio_file.exists():
                        duration = self.get_audio_duration(str(audio_file))
                        if duration > 0:
                            srt_entry = f"{subtitle_index}\n"
                            srt_entry += f"{self.format_srt_time(global_offset)} --> {self.format_srt_time(global_offset + duration)}\n"
                            srt_entry += f"{script_text}\n"
                            srt_content.append(srt_entry)
                            subtitle_index += 1
                            global_offset += duration
                    continue

                # 读取时间戳数据
                with open(timing_file, "r", encoding="utf-8") as f:
                    word_timings = json.load(f)

                if not word_timings:
                    print(f"⚠️  时间戳数据为空: {timing_file}")
                    continue

                # 按字符数分组单词，生成多行字幕
                current_group = []
                current_length = 0
                group_start_time = None

                for i, word_data in enumerate(word_timings):
                    word_text = word_data["text"]
                    word_offset = word_data["offset"] / 1000  # 转换为秒
                    word_duration = word_data["duration"] / 1000  # 转换为秒

                    # 设置组的开始时间
                    if group_start_time is None:
                        group_start_time = word_offset

                    # 添加单词到当前组
                    current_group.append(word_text)
                    current_length += len(word_text)

                    # 判断是否需要换行
                    is_last_word = (i == len(word_timings) - 1)
                    should_break = current_length >= max_chars_per_line or is_last_word

                    if should_break and current_group:
                        # 生成字幕条目
                        group_text = " ".join(current_group) if all(ord(c) < 128 for c in "".join(current_group)) else "".join(current_group)
                        group_end_time = word_offset + word_duration

                        srt_entry = f"{subtitle_index}\n"
                        srt_entry += f"{self.format_srt_time(global_offset + group_start_time)} --> {self.format_srt_time(global_offset + group_end_time)}\n"
                        srt_entry += f"{group_text}\n"

                        srt_content.append(srt_entry)
                        subtitle_index += 1

                        # 重置当前组
                        current_group = []
                        current_length = 0
                        group_start_time = None

                # 更新全局偏移
                audio_file = audio_dir / f"slide_{slide_num:03d}.mp3"
                if audio_file.exists():
                    duration = self.get_audio_duration(str(audio_file))
                    global_offset += duration

            # 写入 SRT 文件
            output_path.parent.mkdir(parents=True, exist_ok=True)
            with open(output_path, "w", encoding="utf-8") as f:
                f.write("\n".join(srt_content))

            print(f"✅ 逐词字幕文件已生成: {output_path}")
            print(f"   共 {len(srt_content)} 条字幕，总时长 {global_offset:.1f} 秒")
            return True

        except Exception as e:
            print(f"❌ 生成逐词字幕失败: {e}")
            import traceback
            traceback.print_exc()
            return False

    def burn_subtitles(
        self,
        video_path: Path,
        srt_path: Path,
        output_path: Path,
        ffmpeg_path: str = None,
        font_size: int = None,
        font_name: str = None,
        outline_width: int = None
    ) -> bool:
        """
        将字幕烧录到视频中（硬字幕）

        Args:
            video_path: 输入视频路径
            srt_path: SRT 字幕文件路径
            output_path: 输出视频路径
            ffmpeg_path: ffmpeg 可执行文件路径
            font_size: 字体大小
            font_name: 字体名称
            outline_width: 描边宽度（1-4）

        Returns:
            是否成功烧录
        """
        try:
            if ffmpeg_path is None:
                ffmpeg_path = get_config("paths", "ffmpeg_path", "FFMPEG_PATH", "ffmpeg")
            if font_size is None:
                font_size = get_config_int("subtitle", "font_size", "SUBTITLE_FONT_SIZE", 24)
            if font_name is None:
                font_name = get_config("subtitle", "font_name", "SUBTITLE_FONT_NAME", "Arial")
            if outline_width is None:
                outline_width = get_config_int("subtitle", "outline_width", "SUBTITLE_OUTLINE_WIDTH", 2)

            # Windows 路径需要转义
            srt_path_str = str(srt_path).replace("\\", "/").replace(":", "\\:")

            # 字幕样式：白色字体 + 黑色描边 + 阴影（任何背景都清晰）
            # PrimaryColour: &HFFFFFF = 白色
            # OutlineColour: &H000000 = 黑色
            # Outline: 描边宽度（1-4，推荐 2-3）
            # Shadow: 1 = 阴影深度
            # BackColour: &H80000000 = 半透明黑色背景（可选）
            subtitle_style = (
                f"FontSize={font_size},"
                f"FontName={font_name},"
                f"PrimaryColour=&HFFFFFF,"  # 白色字体
                f"OutlineColour=&H000000,"  # 黑色描边
                f"Outline={outline_width},"  # 描边宽度
                f"Shadow=1"                 # 阴影
            )

            cmd = [
                ffmpeg_path,
                "-i", str(video_path),
                "-vf", f"subtitles='{srt_path_str}':force_style='{subtitle_style}'",
                "-c:a", "copy",  # 音频不重新编码
                "-y",  # 覆盖输出文件
                str(output_path)
            ]

            print(f"🔥 正在烧录字幕到视频...")
            print(f"   字体: {font_name}, 大小: {font_size}")
            print(f"   样式: 白色字体 + 黑色描边（宽度 {outline_width}px）")
            result = subprocess.run(cmd, capture_output=True, text=True)

            if result.returncode == 0:
                print(f"✅ 硬字幕视频已生成: {output_path}")
                return True
            else:
                print(f"❌ 烧录字幕失败: {result.stderr}")
                return False

        except Exception as e:
            print(f"❌ 烧录字幕失败: {e}")
            import traceback
            traceback.print_exc()
            return False

    def embed_soft_subtitles(
        self,
        video_path: Path,
        srt_path: Path,
        output_path: Path,
        ffmpeg_path: str = None,
        language: str = None
    ) -> bool:
        """
        将字幕作为软字幕嵌入视频（可开关）

        Args:
            video_path: 输入视频路径
            srt_path: SRT 字幕文件路径
            output_path: 输出视频路径
            ffmpeg_path: ffmpeg 可执行文件路径
            language: 字幕语言代码（chi=中文, eng=英文, jpn=日文）

        Returns:
            是否成功嵌入
        """
        try:
            if ffmpeg_path is None:
                ffmpeg_path = get_config("paths", "ffmpeg_path", "FFMPEG_PATH", "ffmpeg")
            if language is None:
                language = get_config("subtitle", "language", "SUBTITLE_LANGUAGE", "chi")

            cmd = [
                ffmpeg_path,
                "-i", str(video_path),
                "-i", str(srt_path),
                "-c", "copy",  # 不重新编码
                "-c:s", "mov_text",  # 字幕编码格式（MP4 兼容）
                "-metadata:s:s:0", f"language={language}",
                "-y",  # 覆盖输出文件
                str(output_path)
            ]

            print(f"📝 正在嵌入软字幕...")
            print(f"   语言: {language}")
            result = subprocess.run(cmd, capture_output=True, text=True)

            if result.returncode == 0:
                print(f"✅ 软字幕视频已生成: {output_path}")
                print(f"   提示：播放器中可以开关字幕")
                return True
            else:
                print(f"❌ 嵌入字幕失败: {result.stderr}")
                return False

        except Exception as e:
            print(f"❌ 嵌入字幕失败: {e}")
            import traceback
            traceback.print_exc()
            return False


def main():
    """测试字幕生成功能"""
    if len(sys.argv) < 3:
        print("用法: python subtitle_generator.py <scripts.json> <audio_dir> [output.srt]")
        sys.exit(1)

    scripts_path = Path(sys.argv[1])
    audio_dir = Path(sys.argv[2])
    output_path = Path(sys.argv[3]) if len(sys.argv) > 3 else Path("output.srt")

    # 加载脚本
    with open(scripts_path, "r", encoding="utf-8") as f:
        scripts = json.load(f)

    # 生成字幕
    generator = SubtitleGenerator()
    generator.generate_srt(scripts, audio_dir, output_path)


if __name__ == "__main__":
    main()
