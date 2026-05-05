"""PPTX 解析模块 - 提取幻灯片内容和缩略图"""
import os
import sys
import json
from pathlib import Path
from typing import List, Dict, Any
from pptx import Presentation

try:
    from .config import (LIBREOFFICE_PATH, FONT_TITLE, FONT_BODY, POPPLER_PATH,
                         RENDER_BACKEND, ENABLE_ANIMATION)
    from .ppt_renderer import RendererFactory
except ImportError:
    sys.path.insert(0, str(Path(__file__).parent))
    from config import (LIBREOFFICE_PATH, FONT_TITLE, FONT_BODY, POPPLER_PATH,
                        RENDER_BACKEND, ENABLE_ANIMATION)
    from ppt_renderer import RendererFactory


class PPTParser:
    """PPTX 文档解析器"""

    def __init__(self, input_dir: str, temp_dir: str):
        self.input_dir = Path(input_dir)
        self.temp_dir = Path(temp_dir)
        self.temp_dir.mkdir(parents=True, exist_ok=True)

        # 创建渲染器
        backend = RENDER_BACKEND if ENABLE_ANIMATION else "auto"
        self.renderer = RendererFactory.create_renderer(
            backend=backend,
            libreoffice_path=LIBREOFFICE_PATH,
            poppler_path=POPPLER_PATH,
            font_title=FONT_TITLE,
            font_body=FONT_BODY
        )

    def find_pptx_files(self) -> List[Path]:
        """查找 input 目录下的所有 pptx 文件"""
        return list(self.input_dir.glob("*.pptx"))

    def parse_slide(self, slide, slide_index: int) -> Dict[str, Any]:
        """解析单张幻灯片内容"""
        content = {
            "index": slide_index,
            "title": "",
            "texts": [],
            "notes": ""
        }

        # 提取标题
        if slide.shapes.title:
            content["title"] = slide.shapes.title.text.strip()

        # 提取所有文本
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    text = paragraph.text.strip()
                    if text and text != content["title"]:
                        content["texts"].append(text)

        # 提取备注
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
            notes_text = slide.notes_slide.notes_text_frame.text.strip()
            content["notes"] = notes_text

        return content

    def parse(self, pptx_path: Path) -> Dict[str, Any]:
        """解析整个 PPTX 文件"""
        prs = Presentation(str(pptx_path))

        result = {
            "filename": pptx_path.name,
            "total_slides": len(prs.slides),
            "slides": []
        }

        for idx, slide in enumerate(prs.slides):
            slide_content = self.parse_slide(slide, idx)
            result["slides"].append(slide_content)

        return result

    def export_to_text(self, pptx_path: Path, output_path: Path):
        """将 PPTX 内容导出为纯文本（用于 LLM 分析）"""
        data = self.parse(pptx_path)

        lines = []
        lines.append(f"# {data['filename']}\n")
        lines.append(f"共 {data['total_slides']} 张幻灯片\n")
        lines.append("=" * 50 + "\n\n")

        for slide in data["slides"]:
            lines.append(f"## 第 {slide['index'] + 1} 页\n")
            if slide["title"]:
                lines.append(f"标题: {slide['title']}\n")
            if slide["texts"]:
                lines.append("内容:\n")
                for text in slide["texts"]:
                    lines.append(f"  - {text}\n")
            if slide["notes"]:
                lines.append(f"备注: {slide['notes']}\n")
            lines.append("\n" + "-" * 30 + "\n\n")

        output_path.write_text("".join(lines), encoding="utf-8")
        return output_path

    def generate_thumbnail(self, pptx_path: Path, output_dir: Path, width: int = 1920, height: int = 1080):
        """生成幻灯片缩略图（支持动画）"""
        output_dir.mkdir(parents=True, exist_ok=True)

        print(f"  使用 {self.renderer.get_name()} 渲染器", flush=True)
        if self.renderer.supports_animation():
            print(f"  ✓ 支持动画渲染", flush=True)
        else:
            print(f"  ⚠️  仅支持静态渲染（不支持动画）", flush=True)

        try:
            frames = self.renderer.render(pptx_path, output_dir, width, height)
            print(f"  ✓ 已生成 {len(frames)} 个渲染帧", flush=True)
            return frames
        except Exception as e:
            print(f"  ✗ 渲染失败: {e}", flush=True)
            raise
