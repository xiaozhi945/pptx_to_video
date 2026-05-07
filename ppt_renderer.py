"""PPT 渲染器 - PowerPoint 音频嵌入方案"""
import platform
import shutil
import time
from pathlib import Path
from typing import List


class PowerPointRenderer:
    """PowerPoint COM API 渲染器（仅 Windows，支持动画）

    使用音频嵌入方案：将 TTS 音频嵌入到幻灯片中，然后导出完整视频
    """

    def __init__(self):
        self._available = None

    def get_name(self) -> str:
        return "PowerPoint"

    def supports_animation(self) -> bool:
        return True

    def is_available(self) -> bool:
        """检查 PowerPoint 是否可用"""
        if self._available is not None:
            return self._available

        # 仅在 Windows 上可用
        if platform.system() != 'Windows':
            self._available = False
            return False

        try:
            import win32com.client
            # 尝试创建 PowerPoint 应用实例
            ppt = win32com.client.Dispatch("PowerPoint.Application")
            ppt.Quit()
            self._available = True
            return True
        except Exception:
            self._available = False
            return False

    def render(self, pptx_path: Path, audio_files: List[Path], output_video: Path, width: int = 1920, height: int = 1080, use_ffmpeg_composition: bool = True) -> Path:
        """使用 PowerPoint COM API 渲染（FFmpeg 后期合成方案）

        流程：
        1. 设置幻灯片切换时间（根据音频时长）
        2. 导出为 MP4 视频（无音频，仅动画）
        3. 使用 FFmpeg 精确合成音频到视频
        4. 返回视频路径

        Args:
            pptx_path: PPT 文件路径
            audio_files: 音频文件列表（按幻灯片顺序）
            output_video: 输出视频路径
            width: 视频宽度
            height: 视频高度
            use_ffmpeg_composition: 是否使用 FFmpeg 后期合成（默认 True）

        Returns:
            视频文件路径
        """
        if not self.is_available():
            raise RuntimeError("PowerPoint 不可用")

        import win32com.client
        from pywintypes import com_error
        from ffmpeg_utils import get_audio_duration, compose_audio_video, check_ffmpeg

        output_video = Path(output_video)
        output_video.parent.mkdir(parents=True, exist_ok=True)

        # 检查 FFmpeg 是否可用
        if use_ffmpeg_composition and not check_ffmpeg():
            print("  ⚠️  FFmpeg 不可用，将使用 PowerPoint 内置音频嵌入方案", flush=True)
            use_ffmpeg_composition = False

        # 创建临时 PPT 副本
        temp_pptx = output_video.parent / f"{pptx_path.stem}_for_video.pptx"
        shutil.copy(pptx_path, temp_pptx)

        ppt = None
        presentation = None

        try:
            print(f"  启动 PowerPoint...", flush=True)
            ppt = win32com.client.Dispatch("PowerPoint.Application")
            ppt.Visible = 1

            print(f"  打开演示文稿: {temp_pptx.name}", flush=True)
            presentation = ppt.Presentations.Open(str(temp_pptx.resolve()))

            slide_count = presentation.Slides.Count
            print(f"  幻灯片数量: {slide_count}", flush=True)

            if len(audio_files) != slide_count:
                raise RuntimeError(f"音频文件数量({len(audio_files)})与幻灯片数量({slide_count})不匹配")

            # 设置幻灯片切换时间
            print(f"  设置幻灯片切换时间...", flush=True)
            for i, audio_file in enumerate(audio_files, start=1):
                slide = presentation.Slides(i)
                audio_path = Path(audio_file).resolve()

                # 获取音频时长
                duration = get_audio_duration(audio_path)

                print(f"    第 {i} 页: {duration:.2f}秒", flush=True)

                # 设置幻灯片切换时间（音频时长）
                slide.SlideShowTransition.AdvanceOnTime = True
                slide.SlideShowTransition.AdvanceTime = duration

            # 保存修改
            print(f"  保存修改后的 PPT...", flush=True)
            presentation.Save()

            # 关闭并重新打开（确保更改生效）
            presentation.Close()
            presentation = ppt.Presentations.Open(str(temp_pptx.resolve()))

            # 导出为 MP4 视频（无音频）
            if use_ffmpeg_composition:
                temp_video = output_video.parent / f"{output_video.stem}_no_audio.mp4"
                print(f"  导出视频（无音频）: {temp_video.name}", flush=True)
                video_output_path = str(temp_video.resolve())
            else:
                print(f"  导出视频: {output_video.name}", flush=True)
                video_output_path = str(output_video.resolve())

            presentation.CreateVideo(
                video_output_path,
                True,  # UseTimingsAndNarrations（使用时间设置）
                5,     # DefaultSlideDuration（不会使用，因为已设置 AdvanceTime）
                height,  # VerticalResolution
                30,    # FramesPerSecond
                80     # Quality
            )

            # 等待视频生成完成
            print(f"  等待视频生成...", flush=True)
            max_wait = 600  # 最多等待 10 分钟
            wait_time = 0
            last_status = -1

            # 确定要检查的视频文件路径
            if use_ffmpeg_composition:
                check_video_path = temp_video
            else:
                check_video_path = output_video

            while wait_time < max_wait:
                time.sleep(2)
                wait_time += 2

                status = presentation.CreateVideoStatus
                if status != last_status:
                    status_names = {0: "未开始", 1: "进行中", 2: "完成", 3: "失败"}
                    print(f"  状态: {status_names.get(status, f'未知({status})')}", flush=True)
                    last_status = status

                if status == 2:  # 完成
                    print(f"  视频生成完成！", flush=True)
                    break
                elif status == 3:  # 失败
                    # 检查文件是否实际存在
                    if check_video_path.exists() and check_video_path.stat().st_size > 1024 * 100:
                        print(f"  ⚠️  状态显示失败，但视频文件已生成 ({check_video_path.stat().st_size / 1024 / 1024:.2f} MB)", flush=True)
                        break
                    else:
                        raise RuntimeError("PowerPoint 视频生成失败")
                elif status == 1 and wait_time % 20 == 0:
                    print(f"  生成中... ({wait_time}s)", flush=True)

                # 检查文件是否存在且大小稳定
                if check_video_path.exists():
                    current_size = check_video_path.stat().st_size
                    if current_size > 0:
                        time.sleep(2)
                        new_size = check_video_path.stat().st_size
                        if new_size == current_size and new_size > 1024 * 100:
                            print(f"  视频文件已稳定 ({new_size / 1024 / 1024:.2f} MB)", flush=True)
                            break

            # 检查视频文件
            if use_ffmpeg_composition:
                if not temp_video.exists():
                    raise RuntimeError(f"视频文件未生成: {temp_video}")
                print(f"  ✓ 无音频视频已生成: {temp_video}", flush=True)

                # 使用 FFmpeg 合成音频
                print(f"\n  使用 FFmpeg 合成音频...", flush=True)
                try:
                    compose_audio_video(temp_video, audio_files, output_video)
                    print(f"  ✓ 音视频合成完成: {output_video}", flush=True)

                    # 清理临时视频
                    if temp_video.exists():
                        temp_video.unlink()
                        print(f"  清理临时文件: {temp_video.name}", flush=True)

                except Exception as e:
                    print(f"  ❌ FFmpeg 合成失败: {e}", flush=True)
                    print(f"  💾 无音频视频已保存: {temp_video}", flush=True)
                    raise RuntimeError(f"FFmpeg 音频合成失败: {e}")
            else:
                if not output_video.exists():
                    raise RuntimeError(f"视频文件未生成: {output_video}")
                print(f"  ✓ 视频已生成: {output_video}", flush=True)

            return output_video

        except com_error as e:
            raise RuntimeError(f"PowerPoint COM 错误: {e}")
        finally:
            # 清理资源
            if presentation:
                try:
                    presentation.Close()
                except:
                    pass
            if ppt:
                try:
                    ppt.Quit()
                except:
                    pass

            # 保留临时文件用于调试
            # TODO: 调试完成后可以取消注释以下代码来清理临时文件
            # if temp_pptx.exists():
            #     try:
            #         temp_pptx.unlink()
            #     except:
            #         pass
            if temp_pptx.exists():
                print(f"  💾 临时 PPT 已保存: {temp_pptx}", flush=True)
