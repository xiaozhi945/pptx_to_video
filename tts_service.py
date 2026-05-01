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

    async def _synthesize(self, text: str, output_path: Path) -> Dict[str, Any]:
        """异步合成语音"""
        communicate = edge_tts.Communicate(text, TTS_VOICE, rate=TTS_RATE, pitch=TTS_PITCH)
        await communicate.save(str(output_path))

        return {
            "text": text,
            "output_path": str(output_path)
        }

    def synthesize(self, scripts: List[Dict[str, Any]], prefix: str = "slide") -> List[Dict[str, Any]]:
        """合成所有脚本的语音"""
        print(f"开始 TTS 语音合成，共 {len(scripts)} 条...", flush=True)

        results = []
        failed_count = 0

        for idx, script_data in enumerate(scripts):
            script_text = script_data.get("script", "").strip()

            if not script_text:
                print(f"  [{idx+1}/{len(scripts)}] ⚠ 警告: 脚本内容为空，跳过", flush=True)
                failed_count += 1
                continue

            output_path = self.output_dir / f"{prefix}_{idx+1:03d}.mp3"

            try:
                # 同步调用异步函数
                result = asyncio.run(
                    self._synthesize(script_text, output_path)
                )
                result["slide_index"] = script_data.get("slide_index", idx)

                # 验证音频文件是否成功生成
                if not output_path.exists():
                    raise FileNotFoundError(f"音频文件未生成: {output_path}")

                # 验证音频文件大小
                file_size = output_path.stat().st_size
                if file_size < 1024:  # 小于 1KB 可能有问题
                    print(f"  [{idx+1}/{len(scripts)}] ⚠ 警告: 音频文件过小 ({file_size} bytes)", flush=True)

                results.append(result)
                print(f"  [{idx+1}/{len(scripts)}] ✓ 语音已生成: {output_path.name}", flush=True)
            except Exception as e:
                print(f"  [{idx+1}/{len(scripts)}] ❌ 语音生成失败: {e}", flush=True)
                failed_count += 1

        if failed_count > 0:
            print(f"  ⚠ 共 {failed_count} 个音频生成失败", flush=True)

        return results
