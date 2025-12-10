#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
图表数据识别模块
使用LLM识别Markdown文本中需要生成饼图的数据
"""

import os
import json
from typing import Dict, List, Any, Optional
from openai import OpenAI


class ChartRecognizer:
    """图表数据识别器"""
    
    def __init__(self, api_key: Optional[str] = None, base_url: Optional[str] = None):
        """初始化识别器
        
        Args:
            api_key: 百炼API Key，如果不提供则从环境变量DASHSCOPE_API_KEY读取
            base_url: API基础URL，默认使用百炼API
        """
        # self.api_key = api_key or os.getenv("DASHSCOPE_API_KEY")
        self.api_key = "sk-8bcc11e0876843dea82b1ff7c9b460ea"
        self.base_url = base_url or "https://dashscope.aliyuncs.com/compatible-mode/v1"
        
        if not self.api_key:
            error_msg = "未提供API Key，请设置环境变量DASHSCOPE_API_KEY或传入api_key参数"
            print(f"错误: {error_msg}")
            raise ValueError(error_msg)
        
        print(f"图表识别器初始化成功，API Key已设置（长度: {len(self.api_key)}）")
        
        self.client = OpenAI(
            api_key=self.api_key,
            base_url=self.base_url,
        )
    
    def recognize(self, markdown_text: str) -> List[Dict[str, Any]]:
        """识别Markdown文本中的饼图数据
        
        Args:
            markdown_text: Markdown文本内容
            
        Returns:
            图表数据列表，格式：[{"type": "pie", "title": "...", "position": "...", "data": {...}}]
        """
        prompt = self._build_prompt(markdown_text)
        
        try:
            print("正在调用LLM API识别图表数据...")
            completion = self.client.chat.completions.create(
                model="qwen3-max",
                messages=[
                    {
                        "role": "system",
                        "content": "你是一个专业的数据分析助手，擅长从文本中识别需要制作图表的数据。"
                    },
                    {
                        "role": "user",
                        "content": prompt
                    }
                ],
                temperature=0.3,
            )
            
            response_text = completion.choices[0].message.content
            print(f"LLM响应长度: {len(response_text)}")
            print(f"LLM响应内容（前500字符）: {response_text[:500]}")
            
            # 解析JSON响应
            charts = self._parse_response(response_text)
            print(f"解析得到 {len(charts)} 个图表")
            return charts
            
        except Exception as e:
            print(f"LLM识别图表数据失败: {e}")
            import traceback
            traceback.print_exc()
            return []
    
    def _build_prompt(self, markdown_text: str) -> str:
        """构建识别Prompt"""
        prompt = f"""请分析以下Markdown文本，识别其中需要制作饼图的数据：

{markdown_text}

## 任务要求

1. **识别饼图数据**：找出文本中包含占比、分布、分类等适合用饼图展示的数据
2. **确定插入位置**：找到数据所在的完整段落，饼图将插入在该段落的后面

## 识别规则

需要识别以下类型的数据：
- 占比数据：如"系统告警占比61.76%，硬件告警占比38.24％"
- 分布数据：如"容量超过阈值告警4起，设备状态异常告警4起，RAID阵列状态异常告警2起"
- 分类统计：如"4级告警29条，5级告警5条"

## 插入位置规则

饼图插入在包含数据的**完整段落**之后。position字段应该包含该段落的**完整文本内容**（从段落开头到结尾的所有文字）。

## 返回格式

返回JSON格式（必须是有效的JSON，不要包含markdown代码块标记）：
{{
  "charts": [
    {{
      "type": "pie",
      "title": "图表标题（简洁描述数据内容）",
      "position": "after:完整段落文本内容",
      "data": {{
        "标签1": 数值1,
        "标签2": 数值2
      }}
    }}
  ]
}}

## 重要要求

1. **type**：必须为"pie"
2. **title**：图表标题，简洁描述数据内容，如"告警类型分布"、"系统告警占比"等
3. **position**：
   - 格式：`"after:完整段落文本"`
   - 必须包含段落的**完整文本内容**（从开头到结尾的所有文字）
   - 段落文本必须与原文**完全一致**，不能修改、省略或添加任何字符
   - 段落以换行符或空行为边界
4. **data**：
   - 键值对格式，键为分类标签，值为数值
   - 数值可以是百分比（如27.6）或具体数量（如4）
   - 确保所有数值加起来为100%（如果是百分比）或总和正确（如果是数量）
5. 如果没有需要制作图表的数据，返回：`{{"charts": []}}`
6. **只返回JSON，不要包含任何其他文字、说明或markdown代码块标记**

## 示例

假设原文有段落：
"本月硬件监控告警共计13条，占总告警量的38.24%。主要告警类型包括：容量超过阈值告警4起、设备状态异常告警4起、RAID阵列状态异常告警2起及物理磁盘健康状况异常告警3起。"

则返回：
{{
  "charts": [
    {{
      "type": "pie",
      "title": "硬件监控告警分布",
      "position": "after:本月硬件监控告警共计13条，占总告警量的38.24%。主要告警类型包括：容量超过阈值告警4起、设备状态异常告警4起、RAID阵列状态异常告警2起及物理磁盘健康状况异常告警3起。",
      "data": {{
        "容量超过阈值告警": 4,
        "设备状态异常告警": 4,
        "RAID阵列状态异常告警": 2,
        "物理磁盘健康状况异常告警": 3
      }}
    }}
  ]
}}"""
        return prompt
    
    def _parse_response(self, response_text: str) -> List[Dict[str, Any]]:
        """解析LLM返回的JSON响应"""
        try:
            # 尝试提取JSON（可能包含markdown代码块）
            text = response_text.strip()
            
            # 如果包含```json或```，提取其中的JSON
            if "```json" in text:
                start = text.find("```json") + 7
                end = text.find("```", start)
                if end != -1:
                    text = text[start:end].strip()
            elif "```" in text:
                start = text.find("```") + 3
                end = text.find("```", start)
                if end != -1:
                    text = text[start:end].strip()
            
            # 解析JSON
            result = json.loads(text)
            
            # 验证格式
            if not isinstance(result, dict) or "charts" not in result:
                return []
            
            charts = result["charts"]
            if not isinstance(charts, list):
                return []
            
            # 验证每个图表数据
            valid_charts = []
            for chart in charts:
                if self._validate_chart(chart):
                    valid_charts.append(chart)
            
            return valid_charts
            
        except json.JSONDecodeError as e:
            print(f"解析JSON失败: {e}")
            print(f"响应内容: {response_text[:500]}")
            return []
        except Exception as e:
            print(f"解析响应失败: {e}")
            return []
    
    def _validate_chart(self, chart: Dict[str, Any]) -> bool:
        """验证图表数据格式"""
        required_fields = ["type", "title", "position", "data"]
        
        # 检查必需字段
        for field in required_fields:
            if field not in chart:
                return False
        
        # 检查type必须为pie
        if chart.get("type") != "pie":
            return False
        
        # 检查position格式
        position = chart.get("position", "")
        if not position.startswith("after:"):
            return False
        
        # 清理position：去除首尾空白（现在使用完整段落，不需要提取）
        keyword = position.replace("after:", "").strip()
        # 确保position格式正确
        chart["position"] = f"after:{keyword}"
        
        # 检查data必须是字典且不为空
        data = chart.get("data")
        if not isinstance(data, dict) or len(data) == 0:
            return False
        
        return True
    
    def _extract_keyword(self, text: str) -> str:
        """提取段落文本（现在直接使用完整段落，不需要提取）
        
        Args:
            text: 段落文本（已经是完整段落）
            
        Returns:
            返回原文本（不做处理，因为现在使用完整段落匹配）
        """
        # 现在使用完整段落匹配，不需要提取关键句
        # 只做基本的清理：去除首尾空白
        return text.strip() if text else text

