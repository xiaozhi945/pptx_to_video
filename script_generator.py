"""LLM 讲解脚本生成模块"""
import json
import sys
import time
from pathlib import Path
from typing import Dict, Any, List

try:
    from .config import MODEL_NAME, MAX_TOKENS, LLM_PROVIDER
except ImportError:
    sys.path.insert(0, str(Path(__file__).parent))
    from config import MODEL_NAME, MAX_TOKENS, LLM_PROVIDER


class ScriptGenerator:
    """使用 LLM 生成 PPT 讲解脚本（支持 Claude 和智谱AI）"""

    def __init__(self, api_key: str, prompts_dir: str, provider: str = None):
        self.provider = provider or LLM_PROVIDER
        self.prompts_dir = Path(prompts_dir)
        self.max_retries = 3  # 最大重试次数
        self.retry_delay = 2  # 重试间隔（秒）

        # 初始化对应的客户端
        if self.provider == "claude":
            from anthropic import Anthropic
            self.client = Anthropic(api_key=api_key)
        elif self.provider == "zhipu":
            from zhipuai import ZhipuAI
            self.client = ZhipuAI(api_key=api_key)
        else:
            raise ValueError(f"不支持的 LLM 提供商: {self.provider}")

        # 加载提示词
        self.analyze_prompt = (self.prompts_dir / "analyze_ppt.txt").read_text(encoding="utf-8")
        self.script_prompt = (self.prompts_dir / "generate_script.txt").read_text(encoding="utf-8")

    def _call_llm(self, prompt: str, max_tokens: int = None) -> str:
        """调用 LLM API（统一接口）"""
        if max_tokens is None:
            max_tokens = MAX_TOKENS

        if self.provider == "claude":
            response = self.client.messages.create(
                model=MODEL_NAME,
                max_tokens=max_tokens,
                messages=[
                    {
                        "role": "user",
                        "content": prompt
                    }
                ]
            )
            return response.content[0].text

        elif self.provider == "zhipu":
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
                    print("警告: 无法解析 JSON，返回原始文本", flush=True)
                    return {"analysis": response_text, "raw": True}

            except Exception as e:
                last_error = e
                if attempt < self.max_retries - 1:
                    print(f"  ⚠ API 请求失败 (尝试 {attempt + 1}/{self.max_retries}): {str(e)}", flush=True)
                    print(f"  → {self.retry_delay} 秒后重试...", flush=True)
                    time.sleep(self.retry_delay)
                else:
                    print(f"  ❌ API 请求失败，已重试 {self.max_retries} 次", flush=True)

        # 所有重试都失败
        raise Exception(f"API 请求失败，错误: {str(last_error)}")

    def generate_script(self, analysis: Dict[str, Any], slides_content: str) -> List[Dict[str, Any]]:
        """为每张幻灯片生成讲解脚本（带重试机制）"""
        print("正在生成讲解脚本...", flush=True)
        print("  → 准备生成参数...", flush=True)

        prompt = self.script_prompt.format(
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
                    print(f"  ⚠ API 请求失败 (尝试 {attempt + 1}/{self.max_retries}): {str(e)}", flush=True)
                    print(f"  → {self.retry_delay} 秒后重试...", flush=True)
                    time.sleep(self.retry_delay)
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
