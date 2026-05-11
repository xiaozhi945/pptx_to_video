"""LLM 讲解脚本生成模块"""
import json
import sys
import time
from pathlib import Path
from typing import Dict, Any, List

try:
    from .config import MODEL_NAME, MAX_TOKENS, LLM_PROVIDER, MAX_RETRIES, BASE_RETRY_DELAY, DEEPSEEK_BASE_URL, QIANWEN_BASE_URL, SCRIPT_LANGUAGE
except ImportError:
    sys.path.insert(0, str(Path(__file__).parent))
    from config import MODEL_NAME, MAX_TOKENS, LLM_PROVIDER, MAX_RETRIES, BASE_RETRY_DELAY, DEEPSEEK_BASE_URL, QIANWEN_BASE_URL, SCRIPT_LANGUAGE


class ScriptGenerator:
    """使用 LLM 生成 PPT 讲解脚本（支持 Claude 和智谱AI）"""

    def __init__(self, api_key: str, prompts_dir: str, provider: str = None):
        self.provider = provider or LLM_PROVIDER
        self.prompts_dir = Path(prompts_dir)
        self.max_retries = MAX_RETRIES  # 从配置读取
        self.base_retry_delay = BASE_RETRY_DELAY  # 从配置读取

        # 初始化对应的客户端
        if self.provider == "claude":
            from anthropic import Anthropic
            self.client = Anthropic(api_key=api_key)
        elif self.provider == "zhipu":
            from zhipuai import ZhipuAI
            self.client = ZhipuAI(api_key=api_key)
        elif self.provider == "deepseek":
            from openai import OpenAI
            self.client = OpenAI(api_key=api_key, base_url=DEEPSEEK_BASE_URL)
        elif self.provider == "qianwen":
            from openai import OpenAI
            self.client = OpenAI(api_key=api_key, base_url=QIANWEN_BASE_URL)
        else:
            raise ValueError(f"不支持的 LLM 提供商: {self.provider}")

        # 加载提示词
        self.analyze_prompt = (self.prompts_dir / "analyze_ppt.txt").read_text(encoding="utf-8")
        self.script_prompt = (self.prompts_dir / "generate_script.txt").read_text(encoding="utf-8")

    def _call_llm(self, prompt: str, max_tokens: int = None) -> str:
        """调用 LLM API（统一接口）"""
        if max_tokens is None:
            max_tokens = MAX_TOKENS

        print(f"  → 模型: {MODEL_NAME}", flush=True)

        if self.provider == "claude":
            # 使用流式传输以支持长响应
            full_response = ""
            with self.client.messages.stream(
                model=MODEL_NAME,
                max_tokens=max_tokens,
                messages=[
                    {
                        "role": "user",
                        "content": prompt
                    }
                ]
            ) as stream:
                for text in stream.text_stream:
                    full_response += text
            return full_response

        elif self.provider in ("zhipu", "deepseek", "qianwen"):
            # 智谱、DeepSeek 和通义千问使用 OpenAI 兼容接口
            response = self.client.chat.completions.create(
                model=MODEL_NAME,
                messages=[
                    {
                        "role": "user",
                        "content": prompt
                    }
                ],
                max_tokens=max_tokens,
            )
            if not response.choices:
                error_msg = f"API 返回空 choices（模型={MODEL_NAME}）"
                # 尝试提取更详细的错误信息
                if hasattr(response, 'error') and response.error:
                    error_msg += f": {response.error}"
                raise RuntimeError(error_msg)
            return response.choices[0].message.content

    def analyze_ppt(self, ppt_content: str) -> Dict[str, Any]:
        """分析 PPT 结构和内容（带重试机制）"""
        print(f"正在使用 {self.provider.upper()} 分析 PPT 内容...", flush=True)
        print("  → 发送请求到 API...", flush=True)

        last_error = None
        for attempt in range(self.max_retries):
            try:
                prompt = f"{self.analyze_prompt}\n\nPPT内容:\n{ppt_content}"
                response_text = self._call_llm(prompt)

                print("  → 正在处理响应...", flush=True)
                print("  ✓ 分析完成", flush=True)

                # 尝试解析 JSON
                try:
                    # 提取 JSON 部分
                    if "```json" in response_text:
                        json_str = response_text.split("```json")[1].split("```")[0]
                    elif "```" in response_text:
                        json_str = response_text.split("```")[1].split("```")[0]
                    else:
                        json_str = response_text

                    return json.loads(json_str.strip())
                except json.JSONDecodeError:
                    # 用正则从文本中提取第一个完整的 {...} 块
                    import re
                    match = re.search(r'\{.*\}', response_text, re.DOTALL)
                    if match:
                        try:
                            return json.loads(match.group())
                        except json.JSONDecodeError:
                            pass
                    print("警告: 无法解析 JSON，返回原始文本", flush=True)
                    return {"analysis": response_text, "raw": True}

            except Exception as e:
                last_error = e
                if attempt < self.max_retries - 1:
                    # 指数退避：2^attempt 秒
                    delay = self.base_retry_delay * (2 ** attempt)
                    print(f"  ⚠ API 请求失败 (尝试 {attempt + 1}/{self.max_retries}): {str(e)}", flush=True)
                    print(f"  → {delay} 秒后重试...", flush=True)
                    time.sleep(delay)
                else:
                    print(f"  ❌ API 请求失败，已重试 {self.max_retries} 次", flush=True)

        # 所有重试都失败
        raise Exception(f"API 请求失败，错误: {str(last_error)}")

    def generate_script(self, analysis: Dict[str, Any], slides_content: str) -> List[Dict[str, Any]]:
        """为每张幻灯片生成讲解脚本（带重试机制）"""
        print("正在生成讲解脚本...", flush=True)
        print("  → 准备生成参数...", flush=True)

        prompt = self.script_prompt.format(
            language=SCRIPT_LANGUAGE,
            analysis=json.dumps(analysis, ensure_ascii=False),
            slides=slides_content
        )

        print("  → 发送请求到 API...", flush=True)

        last_error = None
        for attempt in range(self.max_retries):
            try:
                response_text = self._call_llm(prompt, max_tokens=MAX_TOKENS * 2)

                print("  → 正在处理响应...", flush=True)
                print("  ✓ 脚本生成完成", flush=True)

                # 解析脚本
                try:
                    if "```json" in response_text:
                        json_str = response_text.split("```json")[1].split("```")[0]
                    elif "```" in response_text:
                        json_str = response_text.split("```")[1].split("```")[0]
                    else:
                        json_str = response_text

                    scripts = json.loads(json_str.strip())
                    return scripts if isinstance(scripts, list) else [scripts]
                except json.JSONDecodeError:
                    print("警告: 无法解析脚本 JSON，返回备用方案", flush=True)
                    # 返回原始文本作为整体脚本
                    return [{
                        "slide_index": 0,
                        "script": response_text
                    }]

            except Exception as e:
                last_error = e
                if attempt < self.max_retries - 1:
                    # 指数退避：2^attempt 秒
                    delay = self.base_retry_delay * (2 ** attempt)
                    print(f"  ⚠ API 请求失败 (尝试 {attempt + 1}/{self.max_retries}): {str(e)}", flush=True)
                    print(f"  → {delay} 秒后重试...", flush=True)
                    time.sleep(delay)
                else:
                    print(f"  ❌ API 请求失败，已重试 {self.max_retries} 次", flush=True)

        # 所有重试都失败
        raise Exception(f"API 请求失败，错误: {str(last_error)}")

    def generate(self, ppt_text_content: str, ppt_data: Dict[str, Any]) -> Dict[str, Any]:
        """完整流程：分析 PPT 并生成脚本"""
        # 1. 分析 PPT
        analysis = self.analyze_ppt(ppt_text_content)

        # 2. 生成讲解脚本
        scripts = self.generate_script(analysis, ppt_text_content)

        return {
            "analysis": analysis,
            "scripts": scripts
        }

    def generate_with_animation(self, ppt_text_content: str, ppt_data: Dict[str, Any], frames: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """为动画模式生成脚本（每个动画帧一条脚本）"""
        print("  → 动画模式：为每个动画步骤生成讲解", flush=True)

        # 1. 分析 PPT
        analysis = self.analyze_ppt(ppt_text_content)

        # 2. 构建动画帧信息
        frames_info = []
        for idx, frame in enumerate(frames):
            frames_info.append({
                "frame_index": idx,  # 全局帧索引（0-based）
                "slide_index": frame['slide_index'],
                "animation_step": frame['animation_step'],
                "path": frame['path']
            })

        frames_info_text = json.dumps(frames_info, ensure_ascii=False, indent=2)

        # 3. 生成动画脚本
        print(f"  → 准备为 {len(frames)} 个动画帧生成脚本...", flush=True)

        prompt = self.animation_script_prompt.format(
            analysis=json.dumps(analysis, ensure_ascii=False),
            slides=ppt_text_content,
            frames_info=frames_info_text
        )

        print("  → 发送请求到 API...", flush=True)

        last_error = None
        for attempt in range(self.max_retries):
            try:
                response_text = self._call_llm(prompt, max_tokens=MAX_TOKENS * 3)

                print("  → 正在处理响应...", flush=True)

                print("  ✓ 动画脚本生成完成", flush=True)

                # 解析脚本
                try:
                    if "```json" in response_text:
                        json_str = response_text.split("```json")[1].split("```")[0]
                    elif "```" in response_text:
                        json_str = response_text.split("```")[1].split("```")[0]
                    else:
                        json_str = response_text

                    scripts = json.loads(json_str.strip())
                    print(f"  → 解析得到 {len(scripts)} 条脚本", flush=True)

                    # 验证脚本数量
                    if len(scripts) != len(frames):
                        print(f"  ⚠ 警告: 生成的脚本数量({len(scripts)})与动画帧数量({len(frames)})不匹配", flush=True)

                    return scripts if isinstance(scripts, list) else [scripts]

                except json.JSONDecodeError as e:
                    print(f"  ⚠ 警告: 无法解析脚本 JSON: {e}", flush=True)
                    print(f"  → 尝试重新生成...", flush=True)
                    raise

            except Exception as e:
                last_error = e
                if attempt < self.max_retries - 1:
                    delay = self.base_retry_delay * (2 ** attempt)
                    print(f"  ⚠ API 请求失败 (尝试 {attempt + 1}/{self.max_retries}): {str(e)}", flush=True)
                    print(f"  → {delay} 秒后重试...", flush=True)
                    time.sleep(delay)
                else:
                    print(f"  ❌ API 请求失败，已重试 {self.max_retries} 次", flush=True)

        # 所有重试都失败
        raise Exception(f"动画脚本生成失败，错误: {str(last_error)}")
