# 多模型支持说明

项目现已支持 **4 种 LLM 提供商**，可根据需求灵活切换。

## 支持的模型

| 提供商 | 模型 | API Key 获取 | 特点 | 推荐场景 |
|--------|------|-------------|------|---------|
| **claude** | Claude Sonnet 4.6 | [console.anthropic.com](https://console.anthropic.com/) | 推理能力强，支持长文本 | 复杂 PPT，需要深度理解 |
| **zhipu** | GLM-4 | [open.bigmodel.cn](https://open.bigmodel.cn/) | 国内访问快，中文友好 | 国内用户，中文内容 |
| **deepseek** | DeepSeek Chat | [platform.deepseek.com](https://platform.deepseek.com/) | 性价比高，推理能力强 | 预算有限，大量处理 |
| **qianwen** | 通义千问 Plus | [dashscope.console.aliyun.com](https://dashscope.console.aliyun.com/) | 阿里云服务，稳定可靠 | 企业用户，需要稳定性 |

## 配置方法

### 方式 1：环境变量（.env 文件）

```bash
# 选择提供商
LLM_PROVIDER=deepseek

# 配置对应的 API Key
DEEPSEEK_API_KEY=sk-xxx
```

### 方式 2：配置文件（config.ini）

```ini
[llm]
provider = qianwen
qianwen_api_key = sk-xxx
```

### 方式 3：命令行参数

```bash
python -m pptx_to_video --provider deepseek
```

## 使用示例

### 使用 Claude（默认）

```bash
# 配置 .env
LLM_PROVIDER=claude
ANTHROPIC_API_KEY=sk-ant-xxx

# 运行
python -m pptx_to_video
```

### 使用 DeepSeek（性价比高）

```bash
# 配置 .env
LLM_PROVIDER=deepseek
DEEPSEEK_API_KEY=sk-xxx

# 运行
python -m pptx_to_video
```

### 使用通义千问（国内稳定）

```bash
# 配置 .env
LLM_PROVIDER=qianwen
QIANWEN_API_KEY=sk-xxx

# 运行
python -m pptx_to_video
```

### 临时切换模型

```bash
# 不修改配置，临时使用其他模型
python -m pptx_to_video --provider zhipu
```

## 技术实现

### API 兼容性

- **Claude**: 使用 Anthropic SDK
- **智谱AI**: 使用 ZhipuAI SDK
- **DeepSeek**: 使用 OpenAI 兼容 API
- **通义千问**: 使用 OpenAI 兼容 API

### 统一接口

所有模型通过 `script_generator.py` 的统一接口调用：

```python
generator = ScriptGenerator(api_key, prompts_dir, provider)
result = generator.generate(ppt_text, ppt_data)
```

内部自动处理不同 API 的差异。

## 模型对比

### 推理能力

1. **Claude Sonnet 4.6** ⭐⭐⭐⭐⭐
   - 最强推理能力
   - 最好的长文本理解
   - 适合复杂 PPT

2. **DeepSeek Chat** ⭐⭐⭐⭐
   - 推理能力接近 Claude
   - 性价比极高
   - 适合大量处理

3. **通义千问 Plus** ⭐⭐⭐⭐
   - 中文理解优秀
   - 稳定性好
   - 适合企业用户

4. **智谱 GLM-4** ⭐⭐⭐
   - 中文友好
   - 国内访问快
   - 适合简单 PPT

### 价格对比（参考）

| 模型 | 输入价格 | 输出价格 | 性价比 |
|------|---------|---------|--------|
| Claude Sonnet 4.6 | 较高 | 较高 | ⭐⭐⭐ |
| DeepSeek Chat | 极低 | 极低 | ⭐⭐⭐⭐⭐ |
| 通义千问 Plus | 中等 | 中等 | ⭐⭐⭐⭐ |
| 智谱 GLM-4 | 中等 | 中等 | ⭐⭐⭐⭐ |

### 访问速度（国内）

1. **智谱 GLM-4** - 最快
2. **通义千问 Plus** - 很快
3. **DeepSeek Chat** - 快
4. **Claude Sonnet 4.6** - 需要代理

## 常见问题

### Q: 如何选择模型？

A: 根据需求选择：
- **追求质量**：Claude Sonnet 4.6
- **追求性价比**：DeepSeek Chat
- **国内用户**：智谱 GLM-4 或通义千问
- **企业用户**：通义千问 Plus

### Q: 可以混合使用吗？

A: 可以。每次运行时通过 `--provider` 参数切换：
```bash
python -m pptx_to_video --provider deepseek --input file1.pptx
python -m pptx_to_video --provider claude --input file2.pptx
```

### Q: 如何测试不同模型的效果？

A: 使用同一个 PPT 测试：
```bash
# 测试 Claude
python -m pptx_to_video --provider claude --input test.pptx --skip-video

# 测试 DeepSeek
python -m pptx_to_video --provider deepseek --input test.pptx --skip-video

# 对比生成的 scripts.json
```

### Q: DeepSeek 和通义千问需要额外配置吗？

A: 不需要。它们使用 OpenAI 兼容 API，项目已自动配置好 base_url：
- DeepSeek: `https://api.deepseek.com`
- 通义千问: `https://dashscope.aliyuncs.com/compatible-mode/v1`

### Q: 如何查看当前使用的模型？

A: 运行时会显示：
```bash
python -m pptx_to_video
# 输出: 使用 LLM 提供商: DEEPSEEK
```

或查看配置：
```bash
python -c "import config; config.print_config()"
```

## 依赖安装

新增了 `openai` 包用于 DeepSeek 和通义千问：

```bash
pip install openai>=1.0.0
```

或直接运行程序，会自动检测并安装。

## 迁移指南

### 从单模型迁移

如果之前只使用 Claude 或智谱AI，现在可以：

1. 保持原有配置不变（向后兼容）
2. 需要时添加新模型的 API Key
3. 通过命令行参数临时切换

### 配置示例

```bash
# .env 文件 - 配置所有模型（按需）
LLM_PROVIDER=claude

ANTHROPIC_API_KEY=sk-ant-xxx
ZHIPUAI_API_KEY=xxx
DEEPSEEK_API_KEY=sk-xxx
QIANWEN_API_KEY=sk-xxx
```

## 总结

✅ **支持 4 种模型**：Claude, 智谱AI, DeepSeek, 通义千问  
✅ **统一接口**：无需修改代码，切换模型只需改配置  
✅ **灵活切换**：支持环境变量、配置文件、命令行参数  
✅ **向后兼容**：原有配置继续有效  
✅ **自动安装**：依赖会自动检测和安装

---

**更新时间**: 2026-05-02  
**支持模型数**: 4 个  
**新增依赖**: openai>=1.0.0
