#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
文华财经主力合约图表自动化截图与文档生成工具

该脚本可以自动打开文华财经软件，遍历所有主力合约，
对每个合约的小时线、日线和周线图表进行截图，
并将这些截图整合到一个Word文档中。
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
            f"主力合约图表截图_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        )
        
        # 创建输出目录
        os.makedirs(self.output_dir, exist_ok=True)
        os.makedirs(self.screenshot_dir, exist_ok=True)
        
        # 加载合约列表
        self.contracts = self._load_contracts()
        
        # 图表周期
        self.periods = ['小时线', '日线', '周线']
        
        # 获取屏幕尺寸
        self.screen_width, self.screen_height = pyautogui.size()
        
        # 文档对象
        self.doc = Document()
        self.doc.add_heading('主力合约图表截图', 0)
        
        # 添加运行状态标志
        self.running = True
        
        # 注册ESC键监听
        keyboard.on_press_key('esc', self._on_esc_press)

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
        从openctp.cn获取主力合约信息，只获取成交量前20的品种
        
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

        # 只保留前20个品种
        df_top20 = df_unique.head(20)
        
        logger.info(f"已筛选出成交量前20的品种")

        # 创建合约列表
        contracts = []
        for index, row in df_top20.iterrows():
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
    
    def add_to_document(self, screenshot_path, contract, period):
        """将截图添加到文档"""
        try:
            if not os.path.exists(screenshot_path):
                logger.error(f"截图文件不存在: {screenshot_path}")
                return False
            
            # 添加标题
            self.doc.add_heading(f"{contract['name']}({contract['code']}) {period}图表", level=1)
            
            # 添加图片
            self.doc.add_picture(screenshot_path, width=Inches(6))
            
            # 添加分隔符
            self.doc.add_paragraph("")
            
            logger.info(f"已将截图添加到文档: {contract['name']}({contract['code']}) {period}")
            return True
        except Exception as e:
            logger.error(f"添加截图到文档失败: {e}")
            return False
    
    def save_document(self):
        """保存文档"""
        try:
            self.doc.save(self.doc_path)
            logger.info(f"已保存文档: {self.doc_path}")
            return True
        except Exception as e:
            logger.error(f"保存文档失败: {e}")
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
    
    def add_to_document(self, screenshot_path, contract, period):
        """将截图添加到文档"""
        try:
            if not os.path.exists(screenshot_path):
                logger.error(f"截图文件不存在: {screenshot_path}")
                return False
            
            # 添加标题
            self.doc.add_heading(f"{contract['name']}({contract['code']}) {period}图表", level=1)
            
            # 添加图片
            self.doc.add_picture(screenshot_path, width=Inches(6))
            
            # 添加分隔符
            self.doc.add_paragraph("")
            
            logger.info(f"已将截图添加到文档: {contract['name']}({contract['code']}) {period}")
            return True
        except Exception as e:
            logger.error(f"添加截图到文档失败: {e}")
            return False
    
    def save_document(self):
        """保存文档"""
        try:
            self.doc.save(self.doc_path)
            logger.info(f"已保存文档: {self.doc_path}")
            return True
        except Exception as e:
            logger.error(f"保存文档失败: {e}")
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
        """运行自动化流程"""
        try:
            # 启动文华财经
            if not self.start_wenhua():
                return False
            
            # 遍历所有合约
            for contract in self.contracts:
                # 检查是否需要中止程序
                if not self.running:
                    logger.info("程序已中止")
                    return True
                
                # 切换合约
                if not self.switch_contract(contract):
                    continue
                
                # 遍历所有周期
                for period in self.periods:
                    # 检查是否需要中止程序
                    if not self.running:
                        logger.info("程序已中止")
                        return True
                    
                    # 切换周期
                    if not self.switch_period(period):
                        continue
                    
                    # 等待图表加载
                    time.sleep(3)
                    
                    # 截图
                    screenshot_path = self.take_screenshot(contract, period)
                    if not screenshot_path:
                        continue
                    
                    # 添加到文档
                    self.add_to_document(screenshot_path, contract, period)
            
            # 保存文档
            return self.save_document()
            
        except Exception as e:
            logger.error(f"自动化流程执行失败: {e}")
            return False


def main():
    """主函数"""
    try:
        # 创建自动化对象
        auto = WenhuaAutoScreenshot()
        
        # 运行自动化流程
        if auto.run():
            logger.info("自动化流程执行成功")
        else:
            logger.error("自动化流程执行失败")
    
    except Exception as e:
        logger.error(f"程序执行出错: {e}")
        return 1
    
    return 0


if __name__ == "__main__":
    sys.exit(main())