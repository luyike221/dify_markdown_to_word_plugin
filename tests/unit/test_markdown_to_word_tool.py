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
        markdown_text = "# 2025年10月主机告警信息分析报告\n\n---\n\n## 一、短信告警情况汇总\n\n2025年10月主机系统运行整体安全可控，未发生重大故障性告警事件。全月累计产生主机告警167条，日均告警量约5.6条。从告警类型分布看，系统监控类告警占比98.8%，其中CPU、内存、磁盘三类资源监控告警占主导；硬件监控类告警仅占1.2%。告警级别分布中，4级告警131条（占比78.4%），5级告警36条（占比21.6%），具体统计如下表所示：\n\n| 告警分类       | 告警子类         | 10月告警量（条） |\n|----------------|------------------|------------------|\n| 系统监控告警   | CPU使用率过高    | 29               |\n|                | 内存使用率过高   | 46               |\n|                | 磁盘空间不足     | 46               |\n|                | ERPPT异常        | 46               |\n| 硬件监控告警   | 硬盘故障         | 0                |\n|                | 容量阈值告警     | 0                |\n| **合计**       |                  | **167**          |\n\n---\n\n## 二、异常告警分析\n\n---\n\n### （一）硬件监控告警\n\n10月硬件监控类告警共2条，占总告警量1.2%。经核查，告警源于存储设备容量阈值触发机制，具体为：\n\n| 序号 | 应用系统       | 告警描述                                                                 | 级别 | 处理进度       |\n|------|----------------|--------------------------------------------------------------------------|------|----------------|\n| 1    | 数据备份平台   | 2025-10-05 14:30:22 存储阵列LUN容量达90%阈值（当前90.1%）                 | 4    | 扩容完成       |\n| 2    | 日志归档系统   | 2025-10-20 09:15:47 硬盘SMART状态异常（ID:DISK3）                       | 5    | 更换硬盘       |\n\n**处理结果总结**：硬件类告警均通过扩容或硬件更换及时消除，未对业务连续性造成影响。\n\n---\n\n### （二）系统监控告警\n\n#### 2.1 CPU使用率过高告警分析\n\n10月CPU类告警共29条，占总告警量17.4%。告警主要集中在分布式技术平台（27.6%）、海量数据处理平台（31.0%）及审计信息系统（10.3%）。告警时段集中在凌晨04:00-06:00及业务高峰时段14:00-16:00，与批处理任务及实时数据处理密切相关。\n\n**告警清单及处理**：\n\n| 序号 | 应用系统           | 告警描述                                                                 | 级别 | 处理进度       |\n|------|--------------------|--------------------------------------------------------------------------|------|----------------|\n| 1    | 海量数据处理平台   | 2025-10-01 16:58:25 CPU持续40分钟使用率>90%                              | 4    | 自动恢复       |\n| 2    | 分布式技术平台     | 2025-10-12 06:18:05 CPU使用率>85%（关联4级事件3条）                     | 4    | 自动恢复       |\n| 3    | 审计信息系统       | 2025-10-06 11:28:11 CPU使用率>85%                                        | 4    | 自动恢复       |\n\n**特征分析**：告警多由数据批处理、ETL作业及实时查询压力触发，系统具备自动恢复能力，未引发服务中断。\n\n---\n\n#### 2.2 内存使用率过高告警分析\n\n内存类告警共46条，占总告警量27.5%。运维数据分析系统（52.2%）及互联网电子渠道接入系统（23.9%）为主要来源。告警时段集中在凌晨02:00-04:00，与定时任务执行周期高度相关。\n\n**系统分布统计**：\n\n| 系统名称             | 告警次数 | 占比估算 |\n|----------------------|----------|----------|\n| 运维数据分析系统     | 24       | ≈52%     |\n| 互联网电子渠道接入系统 | 11       | ≈24%     |\n| 手机盾认证系统       | 3        | ≈7%      |\n\n**告警特征**：内存使用率在86%-92%区间波动，触发后10-15分钟内自动回落，未造成服务异常。\n\n---\n\n#### 2.3 磁盘空间或性能不足告警分析\n\n磁盘类告警共46条，占总告警量27.5%。海量数据处理平台（17.4%）及运维数据分析系统（52.2%）为主要告警源，告警集中于`/home/udmh`等数据目录。\n\n**告警分布**：\n\n| 系统名称             | 告警次数 | 占比估算 |\n|----------------------|----------|----------|\n| 运维数据分析系统     | 24       | ≈52%     |\n| 海量数据处理平台     | 8        | ≈17%     |\n| 互联网电子渠道接入系统 | 11       | ≈24%     |\n\n**处理措施**：通过清理临时文件、优化数据归档策略释放空间，告警均在24小时内消除。建议加强磁盘容量预测及自动清理机制建设。\n\n---\n\n### （三）其他告警情况\n\n10月ERPPT类告警46条，占总告警量27.5%，主要涉及进程异常退出及服务响应超时。经排查，告警多因依赖服务短暂不可用导致，未影响核心业务功能。\n\n**改进建议**：\n1. 优化批处理任务调度策略，错峰执行高资源消耗作业；\n2. 建立磁盘空间动态监控阈值，结合业务周期调整预警阈值；\n3. 强化ERPPT服务的健康检查机制，提升异常自愈能力。\n\n---\n\n**综合结论**：本月主机系统运行总体平稳，告警事件均得到有效处置。需重点关注磁盘容量管理及批处理资源协调，持续提升系统资源利用率与稳定性。"
        tool_parameters = {
            'markdown_text': markdown_text,
            'templates': 'alert_report',
            'font_family': 'Times New Roman',
            'font_size': 14,
            'line_spacing': 2.0,
            'page_margins': 3.0,
            'paper_size': 'A3',
            'output_file': 'custom_output.docx',
            'add_page_numbers': False
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

