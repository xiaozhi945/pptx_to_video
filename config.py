"""配置文件"""
import os
from pathlib import Path
from dotenv import load_dotenv

# 加载 .env 文件（如果存在）
env_path = Path(__file__).parent / ".env"
if env_path.exists():
    load_dotenv(env_path)

# 项目根目录（pptx_to_video 目录本身）
PROJECT_ROOT = Path(__file__).parent.resolve()
INPUT_DIR = PROJECT_ROOT / "input"
OUTPUT_DIR = PROJECT_ROOT / "output"
TEMP_DIR = PROJECT_ROOT / "temp"

# LLM 提供商配置
LLM_PROVIDER = os.environ.get("LLM_PROVIDER", "claude").lower()

# Claude API 配置
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")

# 智谱 AI API 配置（备用）
ZHIPUAI_API_KEY = os.environ.get("ZHIPUAI_API_KEY", "")

# TTS 配置
TTS_VOICE = os.environ.get("TTS_VOICE", "zh-CN-XiaoxiaoNeural")  # 默认中文女声
TTS_RATE = os.environ.get("TTS_RATE", "+0%")  # 语速
TTS_PITCH = os.environ.get("TTS_PITCH", "+0Hz")  # 音调

# 视频配置
VIDEO_WIDTH = 1920
VIDEO_HEIGHT = 1080

# LLM 配置（根据提供商自动选择）
if LLM_PROVIDER == "claude":
    MODEL_NAME = "claude-sonnet-4-6"  # Claude Sonnet 4.6
    MAX_TOKENS = 8192
elif LLM_PROVIDER == "zhipu":
    MODEL_NAME = "glm-4"  # 智谱 GLM-4
    MAX_TOKENS = 4096
else:
    raise ValueError(f"不支持的 LLM 提供商: {LLM_PROVIDER}，支持: claude, zhipu")
