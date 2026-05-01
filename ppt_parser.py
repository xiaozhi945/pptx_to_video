"""PPTX 解析模块 - 提取幻灯片内容和缩略图"""
import os
import sys
import json
from pathlib import Path
from typing import List, Dict, Any
from pptx import Presentation
from pptx.util import Inches, Pt
from PIL import Image
import io


class PPTParser:
    """PPTX 文档解析器"""

    def __init__(self, input_dir: str, temp_dir: str):
        self.input_dir = Path(input_dir)
        self.temp_dir = Path(temp_dir)
        self.temp_dir.mkdir(parents=True, exist_ok=True)

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
        """生成幻灯片缩略图"""
        output_dir.mkdir(parents=True, exist_ok=True)
        base_name = pptx_path.stem

        # 方案 1: 尝试使用 LibreOffice 转换（保留原始样式）
        thumbnails = self._convert_with_libreoffice(pptx_path, output_dir, base_name)

        if thumbnails:
            print(f"  使用 LibreOffice 生成缩略图", flush=True)
            return thumbnails

        # 方案 2: 降级到文本渲染
        print(f"  LibreOffice 不可用，使用文本渲染模式", flush=True)
        return self._generate_text_thumbnail(pptx_path, output_dir, width, height)

    def _convert_with_libreoffice(self, pptx_path: Path, output_dir: Path, base_name: str):
        """使用 LibreOffice 转换 PPT 为图片"""
        import subprocess
        import shutil

        # 检查 LibreOffice 是否安装
        soffice = shutil.which("soffice")
        if not soffice:
            # Windows 常见路径
            possible_paths = [
                r"D:\Soft\LibreOffice\program\soffice.exe",
                r"C:\Program Files\LibreOffice\program\soffice.exe",
                r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            ]
            for path in possible_paths:
                if Path(path).exists():
                    soffice = path
                    break

        if not soffice:
            return None

        try:
            # 创建临时输出目录
            temp_output = output_dir / "libreoffice_temp"
            temp_output.mkdir(exist_ok=True)

            # 转换为 PDF
            pdf_path = temp_output / f"{base_name}.pdf"
            cmd = [
                soffice,
                "--headless",
                "--convert-to", "pdf",
                "--outdir", str(temp_output),
                str(pptx_path.resolve())
            ]
            subprocess.run(cmd, check=True, capture_output=True, timeout=60)

            if not pdf_path.exists():
                return None

            # 使用 pdf2image 转换 PDF 为图片
            try:
                from pdf2image import convert_from_path
                images = convert_from_path(str(pdf_path), dpi=150)

                thumbnails = []
                for idx, img in enumerate(images):
                    output_path = output_dir / f"{base_name}_slide_{idx+1:03d}.png"
                    img.save(output_path, "PNG")
                    thumbnails.append(str(output_path))

                # 清理临时文件
                shutil.rmtree(temp_output, ignore_errors=True)
                return thumbnails

            except ImportError:
                print("  提示: 安装 pdf2image 和 poppler 可获得更好的效果", flush=True)
                print("  pip install pdf2image", flush=True)
                return None

        except Exception as e:
            print(f"  LibreOffice 转换失败: {e}", flush=True)
            return None

    def _generate_text_thumbnail(self, pptx_path: Path, output_dir: Path, width: int = 1920, height: int = 1080):
        """生成文本渲染的缩略图（降级方案）"""
        prs = Presentation(str(pptx_path))
        thumbnails = []
        base_name = pptx_path.stem

        for idx, slide in enumerate(prs.slides):
            # 创建空白背景
            img = Image.new('RGB', (width, height), color='white')

            # 简单渲染：提取文本信息绘制
            from PIL import ImageDraw, ImageFont

            draw = ImageDraw.Draw(img)

            # 使用默认字体
            try:
                # Windows 字体路径
                import platform
                if platform.system() == "Windows":
                    title_font = ImageFont.truetype("C:/Windows/Fonts/simhei.ttf", 48)  # 黑体
                    body_font = ImageFont.truetype("C:/Windows/Fonts/simsun.ttc", 24)   # 宋体
                else:
                    title_font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 48)
                    body_font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 24)
            except:
                title_font = ImageFont.load_default()
                body_font = ImageFont.load_default()

            # 绘制标题
            y_offset = 100
            if slide.shapes.title:
                title = slide.shapes.title.text[:50]  # 截断过长标题
                draw.text((100, y_offset), title, font=title_font, fill='black')
                y_offset += 80

            # 绘制内容（简化渲染）
            text_y = y_offset
            for shape in slide.shapes:
                if shape.has_text_frame and not shape == slide.shapes.title:
                    for para in shape.text_frame.paragraphs:
                        text = para.text[:100].strip()
                        if text and text_y < height - 100:
                            draw.text((100, text_y), text, font=body_font, fill='gray')
                            text_y += 40

            # 保存缩略图
            output_path = output_dir / f"{base_name}_slide_{idx+1:03d}.png"
            img.save(output_path)
            thumbnails.append(str(output_path))

        return thumbnails
