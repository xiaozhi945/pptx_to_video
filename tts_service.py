"""TTS 语音合成服务 - 使用 Edge TTS"""
import asyncio
import sys
import edge_tts
from pathlib import Path
from typing import List, Dict, Any

try:
    from .config import TTS_VOICE, TTS_RATE, TTS_PITCH
except ImportError:
    sys.path.insert(0, str(Path(__file__).parent))
    from config import TTS_VOICE, TTS_RATE, TTS_PITCH


class TTSService:
    """Edge TTS 语音合成服务"""

    def __init__(self, output_dir: str):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

    async def _synthesize_one(self, text: str, output_path: Path, slide_index: int) -> Dict[str, Any]:
        """异步合成单个语音，同时获取单词时间戳"""
        try:
            communicate = edge_tts.Communicate(text, TTS_VOICE, rate=TTS_RATE, pitch=TTS_PITCH, boundary='WordBoundary')

            # 收集单词边界时间戳
            word_timings = []
            audio_offset = 0  # 音频偏移量（毫秒）

            # 使用 stream() 获取音频和时间戳
            with open(output_path, "wb") as audio_file:
                async for chunk in communicate.stream():
                    if chunk["type"] == "audio":
                        audio_file.write(chunk["data"])
                    elif chunk["type"] == "WordBoundary":
                        # 记录单词边界
                        word_timings.append({
                            "text": chunk["text"],
                            "offset": chunk["offset"] / 10000,  # 转换为毫秒
                            "duration": chunk["duration"] / 10000  # 转换为毫秒
                        })

            # 验证文件生成
            if not output_path.exists():
                raise FileNotFoundError(f"音频文件未生成: {output_path}")

            # 验证文件大小
            file_size = output_path.stat().st_size
            if file_size < 1024:  # 小于 1KB 可能有问题
                print(f"  ⚠ 警告: 音频文件过小 ({file_size} bytes): {output_path.name}", flush=True)

            # 保存时间戳数据到 JSON 文件
            timing_path = output_path.with_suffix(".timing.json")
            import json
            with open(timing_path, "w", encoding="utf-8") as f:
                json.dump(word_timings, f, ensure_ascii=False, indent=2)

            return {
                "text": text,
                "output_path": str(output_path),
                "timing_path": str(timing_path),
                "slide_index": slide_index,
                "word_timings": word_timings,
                "success": True
            }
        except Exception as e:
            return {
                "text": text,
                "output_path": str(output_path),
                "slide_index": slide_index,
                "success": False,
                "error": str(e)
            }

    async def _synthesize_all(self, scripts: List[Dict[str, Any]], prefix: str = "slide") -> List[Dict[str, Any]]:
        """并发合成所有语音"""
        tasks = []

        for idx, script_data in enumerate(scripts):
            script_text = script_data.get("script", "").strip()

            if not script_text:
                print(f"  [{idx+1}/{len(scripts)}] ⚠ 警告: 脚本内容为空，跳过", flush=True)
                continue

            output_path = self.output_dir / f"{prefix}_{idx+1:03d}.mp3"
            slide_index = script_data.get("slide_index", idx)

            tasks.append(self._synthesize_one(script_text, output_path, slide_index))

        # 并发执行所有任务
        results = await asyncio.gather(*tasks, return_exceptions=False)
        return results

    def synthesize(self, scripts: List[Dict[str, Any]], prefix: str = "slide") -> List[Dict[str, Any]]:
        """合成所有脚本的语音（并发处理）"""
        print(f"开始 TTS 语音合成，共 {len(scripts)} 条...", flush=True)

        # 使用 asyncio.run 执行并发任务
        all_results = asyncio.run(self._synthesize_all(scripts, prefix))

        # 统计结果
        results = []
        failed_count = 0

        for idx, result in enumerate(all_results):
            if result.get("success"):
                results.append(result)
                print(f"  [{idx+1}/{len(scripts)}] ✓ 语音已生成: {Path(result['output_path']).name}", flush=True)
            else:
                failed_count += 1
                print(f"  [{idx+1}/{len(scripts)}] ❌ 语音生成失败: {result.get('error', 'Unknown error')}", flush=True)

        if failed_count > 0:
            print(f"  ⚠ 共 {failed_count} 个音频生成失败", flush=True)

        return results
