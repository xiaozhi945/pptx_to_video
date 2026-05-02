"""配置文件"""
import os
from pathlib import Path
from dotenv import load_dotenv
import configparser

# 加载 .env 文件（如果存在）
env_path = Path(__file__).parent / ".env"
if env_path.exists():
    load_dotenv(env_path)

# 加载 config.ini 文件（如果存在，优先级高于 .env）
config_ini_path = Path(__file__).parent / "config.ini"
config_parser = configparser.ConfigParser()
if config_ini_path.exists():
    config_parser.read(config_ini_path, encoding='utf-8')

def get_config(section: str, key: str, env_key: str = None, default: str = ""):
    """
    获取配置值，优先级：config.ini > 环境变量 > 默认值

    Args:
        section: config.ini 中的节名
        key: config.ini 中的键名
        env_key: 环境变量名（如果为 None，则使用 key 的大写形式）
        default: 默认值
    """
    # 1. 优先从 config.ini 读取
    if config_parser.has_option(section, key):
        value = config_parser.get(section, key).strip()
        if value:  # 非空值
            return value

    # 2. 从环境变量读取
    if env_key is None:
        env_key = key.upper()
    env_value = os.environ.get(env_key, "").strip()
    if env_value:
        return env_value

    # 3. 返回默认值
    return default

def get_config_bool(section: str, key: str, env_key: str = None, default: bool = False) -> bool:
    """获取布尔类型配置"""
    value = get_config(section, key, env_key, str(default))
    return value.lower() in ('true', 'yes', '1', 'on')

def get_config_int(section: str, key: str, env_key: str = None, default: int = 0) -> int:
    """获取整数类型配置"""
    value = get_config(section, key, env_key, str(default))
    try:
        return int(value)
    except ValueError:
        return default

# 项目根目录（pptx_to_video 目录本身）
PROJECT_ROOT = Path(__file__).parent.resolve()
INPUT_DIR = PROJECT_ROOT / "input"
OUTPUT_DIR = PROJECT_ROOT / "output"
TEMP_DIR = PROJECT_ROOT / "temp"

# LLM 提供商配置
LLM_PROVIDER = get_config("llm", "provider", "LLM_PROVIDER", "claude").lower()

# Claude API 配置
ANTHROPIC_API_KEY = get_config("llm", "anthropic_api_key", "ANTHROPIC_API_KEY", "")

# 智谱 AI API 配置
ZHIPUAI_API_KEY = get_config("llm", "zhipuai_api_key", "ZHIPUAI_API_KEY", "")

# DeepSeek API 配置
DEEPSEEK_API_KEY = get_config("llm", "deepseek_api_key", "DEEPSEEK_API_KEY", "")
DEEPSEEK_BASE_URL = get_config("llm", "deepseek_base_url", "DEEPSEEK_BASE_URL", "https://api.deepseek.com")

# 通义千问 API 配置
QIANWEN_API_KEY = get_config("llm", "qianwen_api_key", "QIANWEN_API_KEY", "")
QIANWEN_BASE_URL = get_config("llm", "qianwen_base_url", "QIANWEN_BASE_URL", "https://dashscope.aliyuncs.com/compatible-mode/v1")

# TTS 配置
TTS_VOICE = get_config("tts", "voice", "TTS_VOICE", "zh-CN-XiaoxiaoNeural")
TTS_RATE = get_config("tts", "rate", "TTS_RATE", "+0%")
TTS_PITCH = get_config("tts", "pitch", "TTS_PITCH", "+0Hz")

# 外部工具路径配置（仅从 config.ini 读取）
LIBREOFFICE_PATH = get_config("paths", "libreoffice_path", None, "")
FFMPEG_PATH = get_config("paths", "ffmpeg_path", None, "")
FFPROBE_PATH = get_config("paths", "ffprobe_path", None, "")
POPPLER_PATH = get_config("paths", "poppler_path", None, "")

# 字体路径配置（仅从 config.ini 读取）
FONT_TITLE = get_config("paths", "font_title", None, "")
FONT_BODY = get_config("paths", "font_body", None, "")

# 视频配置
VIDEO_WIDTH = get_config_int("video", "width", "VIDEO_WIDTH", 1920)
VIDEO_HEIGHT = get_config_int("video", "height", "VIDEO_HEIGHT", 1080)

# 性能配置
ENABLE_CONCURRENT_TTS = get_config_bool("performance", "enable_concurrent_tts", "ENABLE_CONCURRENT_TTS", True)
ENABLE_CACHE = get_config_bool("performance", "enable_cache", "ENABLE_CACHE", True)
MAX_RETRIES = get_config_int("performance", "max_retries", "MAX_RETRIES", 3)
BASE_RETRY_DELAY = get_config_int("performance", "base_retry_delay", "BASE_RETRY_DELAY", 2)

# LLM 配置（根据提供商自动选择）
if LLM_PROVIDER == "claude":
    MODEL_NAME = "claude-sonnet-4-6"  # Claude Sonnet 4.6
    MAX_TOKENS = 8192
elif LLM_PROVIDER == "zhipu":
    MODEL_NAME = "glm-4"  # 智谱 GLM-4
    MAX_TOKENS = 4096
elif LLM_PROVIDER == "deepseek":
    MODEL_NAME = "deepseek-chat"  # DeepSeek Chat
    MAX_TOKENS = 8192
elif LLM_PROVIDER == "qianwen":
    MODEL_NAME = "qwen-plus"  # 通义千问 Plus
    MAX_TOKENS = 8192
else:
    raise ValueError(f"不支持的 LLM 提供商: {LLM_PROVIDER}，支持: claude, zhipu, deepseek, qianwen")

# 导出配置信息（用于调试）
def print_config():
    """打印当前配置"""
    print("=" * 60)
    print("当前配置:")
    print("=" * 60)
    print(f"LLM 提供商: {LLM_PROVIDER}")
    print(f"模型: {MODEL_NAME}")
    print(f"TTS 语音: {TTS_VOICE}")
    print(f"视频分辨率: {VIDEO_WIDTH}x{VIDEO_HEIGHT}")
    print(f"LibreOffice 路径: {LIBREOFFICE_PATH or '(自动检测)'}")
    print(f"标题字体: {FONT_TITLE or '(使用默认)'}")
    print(f"正文字体: {FONT_BODY or '(使用默认)'}")
    print(f"并发 TTS: {ENABLE_CONCURRENT_TTS}")
    print(f"启用缓存: {ENABLE_CACHE}")
    print(f"最大重试次数: {MAX_RETRIES}")
    print(f"基础重试延迟: {BASE_RETRY_DELAY}s")
    print("=" * 60)
