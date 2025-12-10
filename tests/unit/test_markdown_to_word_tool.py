# -*- coding: utf-8 -*-
"""
测试 MarkdownToWordTool 的 _invoke 方法
"""

import unittest
import sys
import os
from pathlib import Path
from unittest.mock import Mock

# 添加项目根目录到Python路径
project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))
sys.path.insert(0, str(project_root / "src"))

from tools.markdown_to_word import MarkdownToWordTool


class TestMarkdownToWordTool(unittest.TestCase):
    """MarkdownToWordTool 测试类"""
    
    def setUp(self):
        """测试前的准备工作"""
        # 创建工具实例，需要模拟 Tool 基类的初始化
        self.tool = MarkdownToWordTool(
            runtime=Mock(),
            session=Mock()
        )
        
        # 创建测试输出目录
        self.output_dir = project_root / "test_output"
        self.output_dir.mkdir(exist_ok=True)
        
        # 保存原始方法，以便保存文件
        self.original_create_blob_message = self.tool.create_blob_message
        self.original_create_json_message = self.tool.create_json_message
        
        # 存储生成的文件内容
        self.generated_file_content = None
        self.generated_filename = None
    
    def _save_blob_message(self, blob, meta=None):
        """保存 blob 消息中的文件内容"""
        if blob:
            self.generated_file_content = blob
            if meta and 'filename' in meta:
                self.generated_filename = meta['filename']
                # 保存文件到测试输出目录
                output_path = self.output_dir / meta['filename']
                with open(output_path, 'wb') as f:
                    f.write(blob)
                print(f"\n✅ Word文档已保存到: {output_path.absolute()}")
        return self.original_create_blob_message(blob=blob, meta=meta)
    
    def test_invoke_with_custom_parameters(self):
        """测试使用自定义参数的情况"""
        markdown_text = "# 2025年09月份主机告警信息分析报告\n\n## 一、短信告警情况汇总\n\n2025年09月份主机系统运行整体安全平稳，未发生重大故障事件。全月共产生告警34条，日均告警1.13条。从告警类型分布来看，系统监控类告警占比61.76%（21条），硬件监控类告警占比38.24%（13条）。按告警级别划分，4级告警29条（85.29%），5级告警5条（14.71%）。具体告警分类统计如下表所示：\n\n|告警分类|告警子类|**月告警量（条）|\n|---|---|--:|\n|硬件监控告警|容量超过阈值|4|\n|硬件监控告警|设备状态异常|4|\n|硬件监控告警|RAID 阵列状态异常|2|\n|硬件监控告警|硬盘故障|0|\n|系统监控告警|CPU使用率过高|7|\n|系统监控告警|内存使用率过高|2|\n|系统监控告警|磁盘空间不足|3|\n||合计|34|\n\n## 二、异常告警分析\n\n### 1．硬件监控告警\n\n本月硬件监控告警共计13条，占总告警量的38.24%。主要告警类型包括：容量超过阈值（4条）、设备状态异常（4条）、RAID阵列状态异常（2条）及物理磁盘健康状况异常（3条）。所有硬件告警均集中于三个核心业务系统，其中“互联网电子渠道接入系统”占比最高（61.54%），其次为“海量数据处理平台”（30.77%）和“手机盾统一认证系统”（7.69%）。所有告警级别均为4级，尚未出现5级紧急硬件故障。\n\n告警清单及处理进度如下：\n\n|序号|应用系统|告警描述|级别|告警次数|处理进度|\n|--:|---|---|---|--:|---|\n|1|互联网电子渠道接入系统|文件系统vol_qnxd_app已用容量180.07GB（90%）达到阈值|4|1|未处理|\n|2|互联网电子渠道接入系统|存储池StoragePool001已用容量84.64TB（80%）达阈值|4|1|未处理|\n|3|互联网电子渠道接入系统|设备状态异常（机箱＝严重）|4|1|未处理|\n|4|互联网电子渠道接入系统|设备状态异常（HDD＝严重）|4|1|未处理|\n|5|互联网电子渠道接入系统|物理磁盘HDD1健康状况＝严重|4|1|未处理|\n|6|互联网电子渠道接入系统|设备状态异常（重要>0）|4|1|未处理|\n|7|手机盾统一认证系统|文件系统TM_DB_BAK已用容量32.31TB（95%）达阈值|4|1|未处理|\n|8|海量数据处理平台|物理磁盘Disk19预测错误数>1|4|1|未处理|\n|9|海量数据处理平台|Disk19 RAID array is invalid|4|1|未处理|\n|10|海量数据处理平台|逻辑驱动器2 under RAID card 1 is degraded|4|1|未处理|\n|11|海量数据处理平台|设备健康状况＝Major|4|1|未处理|\n\n处理情况描述：所有硬件告警目前均处于“未处理”状态，涉及存储容量预警、设备状态异常及RAID降级风险，需尽快安排资源扩容与硬件检修。\n\n### 2．系统监控告警\n\n#### 2.1 CPU使用率过高\n\n本月CPU相关告警共12条，占总告警量的35.29%，全部属于系统监控告警范畴。涉及6个应用系统，其中“海量数据处理平台”最为突出，共6条（占CPU告警50%），其次为“互联网数字银行金融服务平”（2条）及其他4个系统各1条。告警级别以4级为主（11条），仅1条为5级（来自海量数据处理平台）。时间分布上，告警集中在月初（9月1日）、中旬（9月11-13日）及月底（9月26日），部分告警持续时间较长（如9月1日CPU持续40分钟超90%）。可能诱因包括批处理任务高峰、负载突增或资源调度异常。\n\n告警清单及处理进度如下：\n\n|序号|应用系统|告警描述|级别|告警次数|处理进度|\n|--:|---|---|---|--:|---|\n|1|海量数据处理平台|CPU使用率过高（>90% for 5m），2025.09.11 19:15:11|4|1|未处理|\n|2|海量数据处理平台|CPU使用率过高（>85% for 30m），2025.09.16 02:07:38|4|1|未处理|\n|3|海量数据处理平台|CPU持续40分钟使用率过高（>90%），2025.09.01 16:58:25|5|1|未处理|\n|4|海量数据处理平台|主机CPU负载太高（per CPU load >10 for 5m），2025.09.13 03:08:11|4|2|未处理|\n|5|海量数据处理平台|主机CPU负载太高（per CPU load >3 for 10m），2025.09.06 12:23:24|4|1|未处理|\n|6|互联网数字银行金融服务平|CPU使用率过高（>85% for 5m），2025.09.08 04:33:53|4|2|未处理|\n|7|数据仓库平台|CPU使用率过高（>90% for 5m），2025.09.11 19:15:11|4|1|未处理|\n|8|分布式技术平台|主机CPU负载太高（per CPU load >3 for 5m），2025.09.26 03:21:59|4|1|未处理|\n|9|客户中心|CPU使用率过高（>90% for 5m），2025.09.12 13:12:05|4|1|未处理|\n|10|审计信息系统|CPU使用率过高（>90% for 5m），2025.09.10 19:52:57|4|1|未处理|\n\n#### 2.2 内存使用率过高\n\n本月内存使用率过高告警共2条，占总告警量的5.88%，全部来自“互联网电子渠道接入系统”。其中1条为5级告警（内存使用率91.15%持续10分钟），1条为4级告警。告警发生于9月1日10:21:59，表明该系统在月初存在内存资源紧张问题，可能与启动初期缓存加载或并发请求激增有关。\n\n告警清单及处理进度如下：\n\n|序号|应用系统|告警描述|级别|告警次数|处理进度|\n|--:|---|---|---|--:|---|\n|1|互联网电子渠道接入系统|内存使用率过高（>90% for 10m），当前值91.15%，2025.09.01 10:21:59|5|1|未处理|\n|2|互联网电子渠道接入系统|内存使用率过高（>90% for 10m）|4|1|未处理|\n\n#### 2.3磁盘空间或性能不足\n\n本月磁盘空间不足告警共4条，占总告警量的11.76%。其中“互联网电子渠道接入系统”3条（含1条5级严重告警，F盘使用率>95%），“海量数据处理平台”1条（F盘使用率>90%）。告警时间集中在9月2日（3条）和9月4日、9月17日各1条。高使用率磁盘涉及/home/ydmh/share/resource、F盘等关键路径，存在业务中断风险，需立即清理或扩容。\n\n告警清单及处理进度如下：\n\n|序号|应用系统|告警描述|级别|告警次数|处理进度|\n|--:|---|---|---|--:|---|\n|1|互联网电子渠道接入系统|/home/ydmh/share/resource磁盘空间不足（used>85%），2025.09.17 15:56:44|4|1|未处理|\n|2|互联网电子渠道接入系统|F盘磁盘空间不足（used>85%），2025.09.02 00:10:39|4|1|未处理|\n|3|互联网电子渠道接入系统|F盘磁盘空间严重不足（used>95%），2025.09.02 05:41:57|5|1|未处理|\n|4|海量数据处理平台|F盘磁盘空间不足（used>90%），2025.09.04 21:31:00|4|1|未处理|\n\n#### 2.4 其他告警\n\n无。\n\n## 三、工作计划及优化建议\n\n无。"
        tool_parameters = {
            'markdown_text': markdown_text,
            'templates': '',
            'font_family': '',
            'font_size': 14,
            'line_spacing': 2.0,
            'page_margins': 3.0,
            'paper_size': 'A3',
            'output_file': '',
            'add_page_numbers': True
        }
        
        # 替换 create_blob_message 方法以保存文件
        self.tool.create_blob_message = self._save_blob_message
        
        # 执行 _invoke 方法
        results = list(self.tool._invoke(tool_parameters))
        
        # 验证至少返回了结果
        self.assertGreater(len(results), 0)
        
        # 验证文件已生成
        if self.generated_filename:
            output_path = self.output_dir / self.generated_filename
            self.assertTrue(output_path.exists(), f"Word文档应该已保存到: {output_path}")
            self.assertGreater(os.path.getsize(output_path), 0, "生成的Word文档不应为空")



if __name__ == '__main__':
    # 运行测试
    unittest.main(verbosity=2)

