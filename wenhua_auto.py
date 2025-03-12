#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
文华财经主力合约图表自动化截图与文档生成工具

该脚本可以自动打开文华财经软件，遍历所有主力合约，
对每个合约的小时线、日线和周线图表进行截图，
并将这些截图整合到一个HTML文档中。
"""

import os
import time
import logging
import datetime
import pyautogui
import subprocess
from docx import Document
from docx.shared import Inches
from PIL import Image, ImageGrab
import configparser
import json
import sys
import requests
import pandas as pd
import keyboard  # 添加到文件开头的导入部分
import win32gui
import win32con
# 添加WMI相关导入
import wmi
import ctypes

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("wenhua_auto.log", encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class WenhuaAutoScreenshot:
    """文华财经自动截图类"""
    
    def __init__(self, config_file="config.ini"):
        """初始化"""
        self.config_file = config_file
        self.config = self._load_config()
        self.wenhua_path = self.config.get('paths', 'wenhua_executable')
        self.output_dir = self.config.get('paths', 'output_directory')
        self.screenshot_dir = os.path.join(self.output_dir, 'screenshots')
        self.doc_path = os.path.join(
            self.output_dir, 
            f"主力合约图表截图_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
        )
        
        # 创建输出目录
        os.makedirs(self.output_dir, exist_ok=True)
        os.makedirs(self.screenshot_dir, exist_ok=True)
        
        # 获取用户输入的合约数量
        self.top_n = self._get_top_n_input()
        
        # 加载合约列表
        self.contracts = self._load_contracts()
        
        # 图表周期
        self.periods = ['周线', '日线', '小时线']
        
        # 获取屏幕尺寸
        self.screen_width, self.screen_height = pyautogui.size()
        
        # 文档对象
        self.doc = Document()
        self.doc.add_heading('主力合约图表截图', 0)
        
        # 添加运行状态标志
        self.running = True
        
        # 注册ESC键监听
        keyboard.on_press_key('esc', self._on_esc_press)
        
        # 保存当前屏幕亮度
        self.original_brightness = self.get_brightness()
    
    def _get_top_n_input(self):
        """获取用户输入的合约数量"""
        while True:
            try:
                top_n = input("请输入要获取的成交量前几的合约数量 (默认5): ").strip()
                if not top_n:  # 如果用户直接回车，使用默认值
                    return 5
                top_n = int(top_n)
                if top_n <= 0:
                    print("请输入大于0的整数")
                    continue
                return top_n
            except ValueError:
                print("请输入有效的整数")

    # 添加屏幕亮度控制方法
    def get_brightness(self):
        """获取当前屏幕亮度"""
        try:
            c = wmi.WMI(namespace='wmi')
            methods = c.WmiMonitorBrightnessMethods()[0]
            brightness_info = c.WmiMonitorBrightness()[0]
            current_brightness = brightness_info.CurrentBrightness
            logger.info(f"当前屏幕亮度: {current_brightness}")
            return current_brightness
        except Exception as e:
            logger.error(f"获取屏幕亮度失败: {e}")
            return 100  # 默认返回100%亮度

    def set_brightness(self, brightness):
        """设置屏幕亮度 (0-100)"""
        try:
            c = wmi.WMI(namespace='wmi')
            methods = c.WmiMonitorBrightnessMethods()[0]
            methods.WmiSetBrightness(brightness, 0)
            logger.info(f"已设置屏幕亮度为: {brightness}")
            return True
        except Exception as e:
            logger.error(f"设置屏幕亮度失败: {e}")
            return False

    def _load_config(self):
        """加载配置文件"""
        if not os.path.exists(self.config_file):
            self._create_default_config()
            
        config = configparser.ConfigParser()
        config.read(self.config_file, encoding='utf-8')
        return config
    
    def _create_default_config(self):
        """创建默认配置文件"""
        config = configparser.ConfigParser()
        
        config['paths'] = {
            'wenhua_executable': r'C:\Program Files\文华财经\文华财经V6\qhjy.exe',
            'output_directory': 'output'
        }
        
        config['screenshot'] = {
            'x': '200',
            'y': '200',
            'width': '800',
            'height': '600'
        }
        
        config['hotkeys'] = {
            'contract_list': 'F2',
            'hour_chart': 'F5',
            'daily_chart': 'F6',
            'weekly_chart': 'F7'
        }
        
        with open(self.config_file, 'w', encoding='utf-8') as f:
            config.write(f)
        
        logger.info(f"已创建默认配置文件: {self.config_file}")
    
    def _load_contracts(self):
        """加载合约列表"""
        contracts_file = "contracts.json"
        
        # 添加用户交互
        if os.path.exists(contracts_file):
            update = input(f"是否需要更新主力合约列表？将获取成交量前{self.top_n}的合约 (y/n): ").lower().strip()
            if update != 'y':
                logger.info("使用现有合约列表")
                with open(contracts_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
        
        try:
            # 尝试从openctp.cn获取最新的主力合约信息
            contracts = self.get_main_contracts()
            
            # 保存到本地文件
            with open(contracts_file, 'w', encoding='utf-8') as f:
                json.dump(contracts, f, ensure_ascii=False, indent=4)
            
            logger.info(f"已从openctp.cn获取并保存最新合约列表: {contracts_file}")
            return contracts
            
        except Exception as e:
            logger.warning(f"从openctp.cn获取合约信息失败: {e}，将使用本地合约列表")
            
            if os.path.exists(contracts_file):
                # 如果本地文件存在，则加载本地文件
                with open(contracts_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            else:
                # 使用默认合约列表
                default_contracts = [
                    {"code": "RB", "name": "螺纹钢"},
                    {"code": "HC", "name": "热轧卷板"},
                    {"code": "I", "name": "铁矿石"},
                    {"code": "J", "name": "焦炭"},
                    {"code": "JM", "name": "焦煤"},
                    {"code": "CU", "name": "铜"},
                    {"code": "AL", "name": "铝"},
                    {"code": "ZN", "name": "锌"},
                    {"code": "NI", "name": "镍"},
                    {"code": "SN", "name": "锡"},
                    {"code": "AU", "name": "黄金"},
                    {"code": "AG", "name": "白银"},
                    {"code": "FU", "name": "燃油"},
                    {"code": "BU", "name": "沥青"},
                    {"code": "RU", "name": "橡胶"},
                    {"code": "A", "name": "豆一"},
                    {"code": "M", "name": "豆粕"},
                    {"code": "Y", "name": "豆油"},
                    {"code": "P", "name": "棕榈油"},
                    {"code": "C", "name": "玉米"},
                    {"code": "CS", "name": "淀粉"},
                    {"code": "CF", "name": "棉花"},
                    {"code": "SR", "name": "白糖"},
                    {"code": "TA", "name": "PTA"},
                    {"code": "MA", "name": "甲醇"}
                ]
                
                with open(contracts_file, 'w', encoding='utf-8') as f:
                    json.dump(default_contracts, f, ensure_ascii=False, indent=4)
                
                logger.info(f"已创建默认合约列表文件: {contracts_file}")
                return default_contracts
    
    def get_main_contracts(self):
        """
        从openctp.cn获取主力合约信息，只获取成交量前N的品种
        
        Returns:
            list: 合约列表，格式为 [{"code": 合约代码, "name": 品种名称}, ...]
        """
        # excel文件的下载链接
        url = "http://www.openctp.cn/fees.xls"

        # 下载excel文件
        r = requests.get(url)
        with open('fees.xls', 'wb') as f:
            f.write(r.content)

        # 读取Excel文件
        df = pd.read_excel('fees.xls')

        # 按照'成交量'列降序排序
        df = df.sort_values(by='成交量', ascending=False)

        # 删除重复项，保留成交量最高的项
        df_unique = df.drop_duplicates(subset='品种名称', keep='first')

        # 只保留前N个品种
        df_top_n = df_unique.head(self.top_n)
        
        logger.info(f"已筛选出成交量前{self.top_n}的品种")

        # 创建合约列表
        contracts = []
        for index, row in df_top_n.iterrows():
            合约代码 = row['合约代码']  # 完整合约代码，如 RB2050
            品种名称 = row['品种名称']
            
            # 提取品种代码作为名称标识
            品种代码 = ''.join(filter(lambda x: not x.isdigit(), 合约代码))
            
            if 合约代码:  # 确保代码不为空
                contracts.append({"code": 合约代码, "name": 品种名称})
        
        return contracts
    
    def switch_contract(self, contract):
        """切换到指定合约"""
        try:
            # 按F2打开合约列表
            pyautogui.press(self.config.get('hotkeys', 'contract_list'))
            time.sleep(1)
            
            # 输入完整合约代码
            pyautogui.write(contract['code'])
            time.sleep(1)
            
            # 按回车确认
            pyautogui.press('enter')
            time.sleep(2)
            
            logger.info(f"已切换到合约: {contract['name']}({contract['code']})")
            return True
        except Exception as e:
            logger.error(f"切换合约失败: {e}")
            return False
    
    def switch_period(self, period):
        """切换图表周期"""
        try:
            key = ''
            if period == '小时线':
                key = '7'
            elif period == '日线':
                key = '9'
            elif period == '周线':
                key = '13'
            
            # 输入周期对应的数字
            pyautogui.write(key)
            time.sleep(1)
            
            # 按回车确认
            pyautogui.press('enter')
            time.sleep(2)
            
            logger.info(f"已切换到{period}")
            return True
        except Exception as e:
            logger.error(f"切换周期失败: {e}")
            return False
    
    def take_screenshot(self, contract, period):
        """截取当前图表"""
        try:
            # 截图文件名
            filename = f"{contract['code']}_{contract['name']}_{period}.png"
            filepath = os.path.join(self.screenshot_dir, filename)
            
            # 直接截取全屏
            screenshot = ImageGrab.grab(bbox=(0, 0, self.screen_width, self.screen_height))
            screenshot.save(filepath)
            
            logger.info(f"已保存截图: {filepath}")
            return filepath
        except Exception as e:
            logger.error(f"截图失败: {e}")
            return None
    
    def save_document(self):
        """保存文档为HTML格式"""
        try:
            html_path = os.path.join(
                self.output_dir,
                f"主力合约图表截图_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
            )

            # 极简HTML模板，完全避免CSS问题
            html_head = """<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>主力合约图表截图</title>
</head>
<body>
<table width="100%" border="0">
<tr>
<td width="200" valign="top" style="background-color:#f0f0f0;padding:10px;">
<h2>品种导航</h2>
"""

            html_middle = """
</td>
<td valign="top" style="padding:10px;">
<h1 align="center">主力合约图表截图</h1>
"""

            html_foot = """
</td>
</tr>
</table>
</body>
</html>"""

            # 生成导航内容
            nav_content = ""
            main_content = ""
            
            for contract in self.contracts:
                contract_id = f"contract_{contract['code']}"
                nav_content += f'<p><a href="#{contract_id}">{contract["name"]} ({contract["code"]})</a></p>\n'
                
                main_content += f'<div id="{contract_id}" style="margin-bottom:30px;border-bottom:1px solid #ddd;padding-bottom:20px;">\n'
                main_content += f'<h2>{contract["name"]} ({contract["code"]})</h2>\n'
                main_content += '<table width="100%"><tr>\n'
                
                for period in self.periods:
                    filename = f"{contract['code']}_{contract['name']}_{period}.png"
                    if os.path.exists(os.path.join(self.screenshot_dir, filename)):
                        screenshot_path = os.path.join('screenshots', filename)
                        main_content += f'<td style="padding:10px;vertical-align:top;">\n'
                        main_content += f'<h3>{period}</h3>\n'
                        main_content += f'<img src="{screenshot_path}" alt="{contract["name"]} {period}图表" style="max-width:100%;border:1px solid #ccc;">\n'
                        main_content += '</td>\n'
                
                main_content += '</tr></table>\n'
                main_content += '</div>\n'
            
            # 组合HTML内容
            html_content = html_head + nav_content + html_middle + main_content + html_foot

            # 写入HTML文件
            with open(html_path, 'w', encoding='utf-8') as f:
                f.write(html_content)

            logger.info(f"已生成HTML文档: {html_path}")
            os.startfile(html_path)
            logger.info("已自动打开HTML文档")
            
            return html_path
        except Exception as e:
            import traceback
            logger.error(f"保存HTML文档失败: {e}")
            logger.error(traceback.format_exc())
            return False

    def _on_esc_press(self, _):
        """ESC键按下时的回调函数"""
        self.running = False
        logger.info("检测到ESC键按下，准备停止程序...")

    def start_wenhua(self):
        """启动文华财经软件"""
        try:
            logger.info(f"正在启动文华财经: {self.wenhua_path}")
            subprocess.Popen(self.wenhua_path)

            # 等待软件启动
            logger.info("等待文华财经启动...")
            time.sleep(15)

            logger.info("文华财经已启动")
            return True
        except Exception as e:
            logger.error(f"启动文华财经失败: {e}")
            return False

    def run(self):
        """运行自动截图流程"""
        try:
            # 设置屏幕亮度为最大，以确保截图质量
            self.set_brightness(10)
            
            # 启动文华财经
            if not self.start_wenhua():
                logger.error("启动文华财经失败，程序退出")
                return False
            
            # 遍历所有合约
            for contract in self.contracts:
                if not self.running:
                    logger.info("检测到停止信号，程序中止")
                    break
                    
                logger.info(f"处理合约: {contract['name']}({contract['code']})")
                
                # 切换到当前合约
                if not self.switch_contract(contract):
                    logger.warning(f"切换到合约 {contract['name']} 失败，跳过")
                    continue
                
                # 遍历所有周期
                for period in self.periods:
                    if not self.running:
                        break
                        
                    logger.info(f"处理周期: {period}")
                    
                    # 切换到当前周期
                    if not self.switch_period(period):
                        logger.warning(f"切换到周期 {period} 失败，跳过")
                        continue
                    
                    # 等待图表加载
                    time.sleep(3)
                    
                    # 截图
                    self.take_screenshot(contract, period)
            
            # 生成文档
            if self.running:
                self.save_document()
                logger.info("文档生成完成")
            
            # 恢复屏幕亮度
            self.set_brightness(self.original_brightness)
            
            return True
            
        except Exception as e:
            logger.error(f"运行过程中发生错误: {e}")
            # 恢复屏幕亮度
            self.set_brightness(self.original_brightness)
            return False


def main():
    """主函数"""
    try:
        # 创建自动截图对象
        auto_screenshot = WenhuaAutoScreenshot()
        
        # 运行自动截图流程
        result = auto_screenshot.run()
        
        if result:
            logger.info("程序执行成功")
            return 0
        else:
            logger.error("程序执行失败")
            return 1
            
    except Exception as e:
        logger.error(f"程序执行过程中发生未处理的异常: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())