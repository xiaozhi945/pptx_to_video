#!/usr/bin/env python3
"""
PPTX 智能讲解视频生成器

根据 PPTX 文件自动生成讲解视频
流程: PPT解析 -> LLM分析 -> 脚本生成 -> TTS语音 -> FFmpeg合成视频
"""
import os
import sys
import json
import argparse
import shutil
from pathlib import Path

# Windows 下强制 UTF-8 输出
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

# 添加项目路径，使模块可导入
sys.path.insert(0, str(Path(__file__).parent))

from config import INPUT_DIR, OUTPUT_DIR, TEMP_DIR, PROJECT_ROOT, ZHIPUAI_API_KEY, ANTHROPIC_API_KEY, LLM_PROVIDER, TTS_VOICE
from ppt_parser import PPTParser
from script_generator import ScriptGenerator
from tts_service import TTSService
from video_creator import VideoCreator


def check_ffmpeg():
    """检查 FFmpeg 是否已安装"""
    ffmpeg = shutil.which("ffmpeg")
    ffprobe = shutil.which("ffprobe")
    if not ffmpeg or not ffprobe:
        print("警告: FFmpeg 未安装或未在 PATH 中", flush=True)
        print("请安装 FFmpeg: https://ffmpeg.org/download.html", flush=True)
        return False
    return True


def main():
    parser = argparse.ArgumentParser(description="PPTX 智能讲解视频生成器")
    parser.add_argument("--input", "-i", help=f"输入 PPTX 文件路径（默认: {INPUT_DIR}/*.pptx）")
    parser.add_argument("--output", "-o", help="输出视频路径")
    parser.add_argument("--api-key", "-k", help="LLM API Key（也可通过环境变量设置）")
    parser.add_argument("--provider", "-p", choices=["claude", "zhipu"], help=f"LLM 提供商（默认: {LLM_PROVIDER}）")
    parser.add_argument("--skip-tts", action="store_true", help="跳过 TTS 步骤")
    parser.add_argument("--skip-video", action="store_true", help="跳过视频合成步骤")
    parser.add_argument("--list", "-l", action="store_true", help="列出 input 目录中的 PPTX 文件")

    args = parser.parse_args()

    # 确保目录存在
    for d in [INPUT_DIR, OUTPUT_DIR, TEMP_DIR]:
        Path(d).mkdir(parents=True, exist_ok=True)

    # 初始化各模块
    parser_mod = PPTParser(INPUT_DIR, TEMP_DIR)

    # 列出文件（不需要 API key）
    if args.list:
        pptx_files = parser_mod.find_pptx_files()
        print(f"input 目录中的 PPTX 文件 ({len(pptx_files)} 个):\n", flush=True)
        for f in pptx_files:
            print(f"  - {f.name}", flush=True)
        return

    # 确定使用的 LLM 提供商
    provider = args.provider or LLM_PROVIDER

    # 检查 API Key
    if provider == "claude":
        api_key = args.api_key or ANTHROPIC_API_KEY
        if not api_key:
            print("错误: 请设置 ANTHROPIC_API_KEY 环境变量或使用 --api-key 参数", flush=True)
            print("export ANTHROPIC_API_KEY='your-api-key-here'", flush=True)
            sys.exit(1)
    elif provider == "zhipu":
        api_key = args.api_key or ZHIPUAI_API_KEY
        if not api_key:
            print("错误: 请设置 ZHIPUAI_API_KEY 环境变量或使用 --api-key 参数", flush=True)
            print("export ZHIPUAI_API_KEY='your-api-key-here'", flush=True)
            sys.exit(1)
    else:
        print(f"错误: 不支持的 LLM 提供商: {provider}", flush=True)
        sys.exit(1)

    print(f"使用 LLM 提供商: {provider.upper()}", flush=True)

    # 查找输入文件
    if args.input:
        input_file = Path(args.input)
        if not input_file.exists():
            print(f"错误: 文件不存在 {input_file}", flush=True)
            sys.exit(1)
        pptx_files = [input_file]
    else:
        pptx_files = parser_mod.find_pptx_files()
        if not pptx_files:
            print(f"错误: input 目录中没有找到 PPTX 文件: {INPUT_DIR}", flush=True)
            print("请将 PPTX 文件放入 input 目录，或使用 --input 指定文件路径", flush=True)
            sys.exit(1)

    # 处理每个文件
    for pptx_file in pptx_files:
        print(f"\n{'='*60}", flush=True)
        print(f"处理文件: {pptx_file.name}", flush=True)
        print(f"{'='*60}\n", flush=True)

        # 创建临时目录
        file_temp_dir = Path(TEMP_DIR) / pptx_file.stem
        file_temp_dir.mkdir(parents=True, exist_ok=True)
        file_output_dir = Path(OUTPUT_DIR) / pptx_file.stem
        file_output_dir.mkdir(parents=True, exist_ok=True)

        # 1. 解析 PPT
        print("[1/5] 解析 PPT 文件...", flush=True)
        ppt_data = parser_mod.parse(pptx_file)
        print(f"  共 {ppt_data['total_slides']} 张幻灯片", flush=True)

        # 导出为文本用于 LLM 分析
        text_content_path = file_temp_dir / "content.txt"
        parser_mod.export_to_text(pptx_file, text_content_path)
        ppt_text_content = text_content_path.read_text(encoding="utf-8")

        # 2. 生成讲解脚本
        print("[2/5] 生成讲解脚本...", flush=True)
        prompts_dir = Path(__file__).parent / "prompts"
        generator = ScriptGenerator(api_key, prompts_dir, provider)
        result = generator.generate(ppt_text_content, ppt_data)
        scripts = result["scripts"]

        # 验证脚本数量与幻灯片数量匹配
        expected_count = ppt_data['total_slides']
        actual_count = len(scripts)
        if actual_count != expected_count:
            print(f"  ⚠ 警告: 生成的脚本数量({actual_count})与幻灯片数量({expected_count})不匹配", flush=True)

        # 按 slide_index 排序
        scripts.sort(key=lambda x: x.get("slide_index", 0))

        # 检测索引是从 0 还是从 1 开始，并标准化为从 0 开始
        if scripts and scripts[0].get("slide_index", 0) == 1:
            print(f"  → 检测到索引从 1 开始，自动调整为从 0 开始", flush=True)
            for script in scripts:
                if "slide_index" in script:
                    script["slide_index"] -= 1

        # 验证连续性
        for i, script in enumerate(scripts):
            slide_idx = script.get("slide_index", i)
            if slide_idx != i:
                print(f"  ⚠ 警告: 脚本索引不连续，期望 {i}，实际 {slide_idx}", flush=True)
            # 确保每个脚本都有内容
            if not script.get("script", "").strip():
                print(f"  ⚠ 警告: 第 {i+1} 页脚本内容为空", flush=True)

        # 保存脚本
        script_path = file_output_dir / "scripts.json"
        with open(script_path, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        print(f"  脚本已保存: {script_path}", flush=True)
        print(f"  共生成 {len(scripts)} 条脚本", flush=True)

        # 3. TTS 语音合成
        if not args.skip_tts:
            print("[3/5] 语音合成...", flush=True)
            print(f"  使用语音: {TTS_VOICE}", flush=True)
            tts_service = TTSService(file_temp_dir)
            audio_results = tts_service.synthesize(scripts)
            print(f"  已生成 {len(audio_results)} 个音频文件", flush=True)

            # 验证音频生成结果
            if len(audio_results) != len(scripts):
                print(f"  ⚠ 警告: 音频数量({len(audio_results)})与脚本数量({len(scripts)})不匹配", flush=True)

        # 4. 生成缩略图
        print("[4/5] 生成幻灯片缩略图...", flush=True)
        thumbnails = parser_mod.generate_thumbnail(pptx_file, file_temp_dir)
        print(f"  已生成 {len(thumbnails)} 张缩略图", flush=True)

        # 5. 合成视频
        if not args.skip_video:
            if not check_ffmpeg():
                print("错误: FFmpeg 是合成视频所必需的", flush=True)
                sys.exit(1)
            print("[5/5] 合成视频...", flush=True)

            # 获取音频文件列表
            audio_files = sorted(file_temp_dir.glob("slide_*.mp3"))

            if not audio_files:
                print("  警告: 没有找到音频文件，跳过视频合成", flush=True)
            else:
                # 最终验证：确保幻灯片、音频、脚本数量一致
                print(f"  验证资源数量:", flush=True)
                print(f"    幻灯片: {len(thumbnails)} 张", flush=True)
                print(f"    音频文件: {len(audio_files)} 个", flush=True)
                print(f"    脚本: {len(scripts)} 条", flush=True)

                if len(thumbnails) != len(audio_files):
                    print(f"  ❌ 错误: 幻灯片与音频数量不匹配，无法合成视频", flush=True)
                    print(f"  请检查 TTS 步骤是否完全成功", flush=True)
                else:
                    output_video = args.output or str(file_output_dir / f"{pptx_file.stem}.mp4")
                    video_creator = VideoCreator(file_output_dir, file_temp_dir)

                    if video_creator.create_video(thumbnails, [str(a) for a in audio_files], output_video):
                        print(f"\n✅ 视频生成完成: {output_video}", flush=True)
                    else:
                        print("\n❌ 视频生成失败", flush=True)

        print(f"\n📁 输出目录: {file_output_dir}", flush=True)


if __name__ == "__main__":
    main()
