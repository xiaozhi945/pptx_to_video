#!/usr/bin/env python3
"""测试渲染器功能"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from ppt_renderer import RendererFactory, PowerPointRenderer, LibreOfficeRenderer, PillowRenderer
from config import LIBREOFFICE_PATH, POPPLER_PATH, FONT_TITLE, FONT_BODY

def test_renderers():
    """测试所有渲染器的可用性"""
    print("=" * 60)
    print("渲染器可用性测试")
    print("=" * 60)

    # 测试 PowerPoint 渲染器
    print("\n1. PowerPoint 渲染器:")
    ppt_renderer = PowerPointRenderer()
    print(f"   可用: {ppt_renderer.is_available()}")
    print(f"   支持动画: {ppt_renderer.supports_animation()}")

    # 测试 LibreOffice 渲染器
    print("\n2. LibreOffice 渲染器:")
    lo_renderer = LibreOfficeRenderer(LIBREOFFICE_PATH, POPPLER_PATH)
    print(f"   可用: {lo_renderer.is_available()}")
    print(f"   支持动画: {lo_renderer.supports_animation()}")
    if LIBREOFFICE_PATH:
        print(f"   路径: {LIBREOFFICE_PATH}")

    # 测试 Pillow 渲染器
    print("\n3. Pillow 渲染器:")
    pillow_renderer = PillowRenderer(FONT_TITLE, FONT_BODY)
    print(f"   可用: {pillow_renderer.is_available()}")
    print(f"   支持动画: {pillow_renderer.supports_animation()}")

    # 测试自动选择
    print("\n4. 自动选择渲染器:")
    auto_renderer = RendererFactory.create_renderer(
        backend="auto",
        enable_animation=True,
        libreoffice_path=LIBREOFFICE_PATH,
        poppler_path=POPPLER_PATH,
        font_title=FONT_TITLE,
        font_body=FONT_BODY
    )
    print(f"   选择的渲染器: {auto_renderer.get_name()}")
    print(f"   支持动画: {auto_renderer.supports_animation()}")

    print("\n" + "=" * 60)
    print("测试完成")
    print("=" * 60)

if __name__ == "__main__":
    test_renderers()
