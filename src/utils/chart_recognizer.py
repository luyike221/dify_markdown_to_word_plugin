#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
图表数据识别模块
直接解析已处理好的图表数据（JSON格式）
"""

import json
from typing import Dict, List, Any


class ChartRecognizer:
    """图表数据识别器"""
    
    def __init__(self):
        """初始化识别器"""
        print("图表识别器初始化成功")
    
    def recognize(self, chart_data: str) -> List[Dict[str, Any]]:
        """解析图表数据
        
        Args:
            chart_data: 已处理好的图表数据（JSON字符串格式）
            
        Returns:
            图表数据列表，格式：[{"type": "pie"/"bar", "title": "...", "position": "...", "data": {...}}]
        """
        # 如果未提供图表数据，直接返回空数组
        if not chart_data or not chart_data.strip():
            print("未提供图表数据，直接返回空数组")
            return []
        
        try:
            print("正在解析图表数据...")
            # 直接解析传入的JSON数据
            charts = self._parse_response(chart_data)
            print(f"解析得到 {len(charts)} 个图表")
            return charts
            
        except Exception as e:
            print(f"解析图表数据失败: {e}")
            import traceback
            traceback.print_exc()
            return []
    
    def _parse_response(self, response_text: str) -> List[Dict[str, Any]]:
        """解析JSON格式的图表数据（支持包含markdown代码块标记的情况）"""
        try:
            # 尝试提取JSON（可能包含markdown代码块）
            text = response_text.strip()
            
            # 如果包含```json或```，提取其中的JSON
            if "```json" in text:
                start = text.find("```json") + 7
                end = text.find("```", start)
                if end != -1:
                    text = text[start:end].strip()
                    print(f"检测到 ```json 代码块，已提取其中的JSON内容")
                else:
                    # 没有找到结束标记，尝试从开始位置提取到末尾
                    text = text[start:].strip()
                    print(f"检测到 ```json 开始标记，但未找到结束标记，尝试解析剩余内容")
            elif "```" in text:
                start = text.find("```") + 3
                end = text.find("```", start)
                if end != -1:
                    text = text[start:end].strip()
                    print(f"检测到 ``` 代码块，已提取其中的JSON内容")
                else:
                    # 没有找到结束标记，尝试从开始位置提取到末尾
                    text = text[start:].strip()
                    print(f"检测到 ``` 开始标记，但未找到结束标记，尝试解析剩余内容")
            else:
                print("未检测到代码块标记，直接解析整个文本")
            
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
            print(f"原始内容（前500字符）: {response_text[:500]}")
            # 如果原始内容包含代码块标记，提示用户
            if "```" in response_text:
                print("提示: 检测到代码块标记，但JSON解析失败。请确保代码块内的内容是有效的JSON格式。")
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
        
        # 检查type必须为pie、bar或line
        chart_type = chart.get("type")
        if chart_type not in ["pie", "bar", "line"]:
            return False
        
        # 检查position格式
        position = chart.get("position", "")
        if not (position.startswith("after:") or position.startswith("before:")):
            return False
        
        # 清理position：去除首尾空白（现在使用完整段落，不需要提取）
        if position.startswith("after:"):
            keyword = position.replace("after:", "").strip()
            # 确保position格式正确
            chart["position"] = f"after:{keyword}"
        elif position.startswith("before:"):
            keyword = position.replace("before:", "").strip()
            # 确保position格式正确
            chart["position"] = f"before:{keyword}"
        
        # 检查data必须是字典且不为空
        data = chart.get("data")
        if not isinstance(data, dict) or len(data) == 0:
            return False
        
        return True
    

