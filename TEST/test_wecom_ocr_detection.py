#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
使用OCR检测企业微信聊天窗口消息区域的变化
用于判断AI回复完成的时机
"""

import time
import win32gui
import win32process
import win32ui
import win32con
from PIL import Image
import pytesseract
import os
import hashlib
from wxauto.utils.win32 import GetAllWindows

class WeChatOCRDetector:
    def __init__(self):
        self.last_text_hash = None
        self.last_full_text = ""
        self.stable_count = 0
        self.required_stable_checks = 6  # 需要连续6次(3秒)稳定才认为完成
        
        # 设置tesseract路径（如果需要）
        # pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        
    def find_wecom_chat_windows(self):
        """查找企业微信聊天窗口"""
        chat_windows = []
        all_windows = GetAllWindows()
        
        for hwnd, class_name, window_title in all_windows:
            if self.is_wecom_chat_window(hwnd, class_name, window_title):
                chat_windows.append((hwnd, window_title))
        
        return chat_windows
    
    def is_wecom_chat_window(self, hwnd, class_name, window_title):
        """判断是否是企业微信聊天窗口"""
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return False
            
            # 检查是否是企业微信进程
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            process_name = self.get_process_name(pid)
            
            if "WXWork" not in process_name and "企业微信" not in process_name:
                return False
            
            # 检查窗口类名和标题
            if class_name == "WwStandaloneConversationWnd":
                return True
            
            # 检查主窗口中的对话
            if "企业微信" in window_title and len(window_title) > 3:
                return True
                
            return False
            
        except Exception as e:
            print(f"检查窗口时出错: {e}")
            return False
    
    def get_process_name(self, pid):
        """获取进程名称"""
        try:
            import psutil
            process = psutil.Process(pid)
            return process.name()
        except:
            return ""
    
    def capture_window_region(self, hwnd, region_ratio=(0.1, 0.3, 0.9, 0.8)):
        """
        截取窗口的指定区域
        region_ratio: (left_ratio, top_ratio, right_ratio, bottom_ratio)
        默认截取窗口中间的消息区域，避开标题栏和输入框
        """
        try:
            # 获取窗口矩形
            rect = win32gui.GetWindowRect(hwnd)
            width = rect[2] - rect[0]
            height = rect[3] - rect[1]
            
            # 计算消息区域的实际坐标
            left = int(width * region_ratio[0])
            top = int(height * region_ratio[1])
            right = int(width * region_ratio[2])
            bottom = int(height * region_ratio[3])
            
            crop_width = right - left
            crop_height = bottom - top
            
            # 获取窗口DC
            hwndDC = win32gui.GetWindowDC(hwnd)
            mfcDC = win32gui.CreateCompatibleDC(hwndDC)
            saveBitMap = win32gui.CreateCompatibleBitmap(hwndDC, crop_width, crop_height)
            win32gui.SelectObject(mfcDC, saveBitMap)
            
            # 复制窗口指定区域到内存DC
            result = win32gui.BitBlt(mfcDC, 0, 0, crop_width, crop_height, 
                                   hwndDC, left, top, win32con.SRCCOPY)
            
            if result:
                # 获取位图信息
                bmpinfo = win32gui.GetObject(saveBitMap)
                bmpstr = win32gui.GetBitmapBits(saveBitMap, bmpinfo.bmWidthBytes * bmpinfo.bmHeight)
                
                # 转换为PIL图像
                img = Image.frombuffer(
                    'RGB',
                    (bmpinfo.bmWidth, bmpinfo.bmHeight),
                    bmpstr, 'raw', 'BGRX', 0, 1
                )
                
                # 清理资源
                win32gui.DeleteObject(saveBitMap)
                win32gui.DeleteDC(mfcDC)
                win32gui.ReleaseDC(hwnd, hwndDC)
                
                return img
            else:
                print("截图失败")
                # 清理资源
                win32gui.DeleteObject(saveBitMap)
                win32gui.DeleteDC(mfcDC)
                win32gui.ReleaseDC(hwnd, hwndDC)
                return None
                
        except Exception as e:
            print(f"截取窗口区域失败: {e}")
            return None
    
    def extract_text_from_image(self, image):
        """从图像中提取文本"""
        try:
            # 转换为灰度图像以提高OCR准确性
            gray_image = image.convert('L')
            
            # 使用pytesseract进行OCR
            # 配置OCR参数，适合中文识别
            custom_config = r'--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ，。？！：；""''（）【】《》、·'
            
            # 尝试中文+英文识别
            try:
                text = pytesseract.image_to_string(gray_image, lang='chi_sim+eng', config=custom_config)
            except:
                # 如果中文识别失败，使用英文识别
                text = pytesseract.image_to_string(gray_image, lang='eng')
            
            # 清理文本
            text = text.strip()
            text = '\n'.join(line.strip() for line in text.split('\n') if line.strip())
            
            return text
            
        except Exception as e:
            print(f"OCR识别失败: {e}")
            return ""
    
    def is_text_stable(self, current_text):
        """判断文本是否稳定（AI回复完成）"""
        # 计算文本哈希
        current_hash = hashlib.md5(current_text.encode('utf-8')).hexdigest()
        
        if current_hash == self.last_text_hash:
            self.stable_count += 1
        else:
            self.stable_count = 0
            self.last_text_hash = current_hash
            self.last_full_text = current_text
        
        # 如果连续多次检查文本都没有变化，认为AI回复完成
        return self.stable_count >= self.required_stable_checks
    
    def monitor_ai_reply(self, hwnd, window_title):
        """监控AI回复完成状态"""
        print(f"\n开始监控AI回复: {window_title}")
        print(f"窗口句柄: {hwnd}")
        print("=" * 60)
        
        # 测试截图和OCR
        print("正在测试截图和OCR功能...")
        test_image = self.capture_window_region(hwnd)
        if test_image:
            print("截图成功")
            # 保存测试图片
            test_image.save("test_screenshot.png")
            print("测试截图已保存为: test_screenshot.png")
            
            # 测试OCR
            test_text = self.extract_text_from_image(test_image)
            print(f"OCR测试结果: {test_text[:100]}..." if len(test_text) > 100 else f"OCR测试结果: {test_text}")
        else:
            print("截图失败，请检查窗口状态")
            return
        
        print("\n请在企业微信中发送消息给AI，然后观察检测结果...")
        input("准备好后按回车开始监控...")
        
        print("\n开始实时监控...")
        print("-" * 60)
        
        start_time = time.time()
        check_count = 0
        last_displayed_text = ""
        
        while True:
            check_count += 1
            current_time = time.time()
            elapsed = current_time - start_time
            
            # 截取窗口图像
            image = self.capture_window_region(hwnd)
            if not image:
                print(f"第{check_count}次检查失败：无法截图")
                time.sleep(0.5)
                continue
            
            # OCR识别文本
            current_text = self.extract_text_from_image(image)
            
            # 检查文本是否有变化
            text_changed = current_text != last_displayed_text
            
            if text_changed:
                print(f"\n时间: {elapsed:.1f}s | 检查: #{check_count}")
                print("检测到文本变化:")
                print("-" * 40)
                print(current_text)
                print("-" * 40)
                last_displayed_text = current_text
                
                # 保存变化时的截图
                image.save(f"change_{check_count}.png")
                print(f"截图已保存: change_{check_count}.png")
            else:
                # 简单显示状态
                if check_count % 10 == 0:  # 每10次检查显示一次状态
                    print(f"时间: {elapsed:.1f}s | 检查: #{check_count} | 稳定次数: {self.stable_count}")
            
            # 检查是否稳定
            is_stable = self.is_text_stable(current_text)
            
            if is_stable and len(current_text.strip()) > 0:
                print(f"\n🎉 检测到AI回复完成！")
                print(f"稳定时间: {elapsed:.1f}s")
                print(f"稳定检查次数: {self.stable_count}")
                print("=" * 50)
                print("最终回复内容:")
                print(self.last_full_text)
                print("=" * 50)
                
                # 保存最终截图
                image.save("final_reply.png")
                print("最终截图已保存: final_reply.png")
                break
            
            # 防止无限循环，最多监控5分钟
            if elapsed > 300:
                print(f"\n监控超时（5分钟），停止检测")
                break
            
            time.sleep(0.5)  # 每0.5秒检查一次
        
        return self.last_full_text
    
    def test_screenshot_and_ocr(self, hwnd, window_title):
        """测试截图和OCR功能"""
        print(f"\n测试截图和OCR功能: {window_title}")
        print("=" * 50)
        
        # 截取不同区域进行测试
        regions = {
            "全窗口": (0, 0, 1, 1),
            "上半部分": (0, 0, 1, 0.5),
            "下半部分": (0, 0.5, 1, 1),
            "中间消息区": (0.1, 0.2, 0.9, 0.8),
            "底部区域": (0.1, 0.7, 0.9, 0.95)
        }
        
        for region_name, region_ratio in regions.items():
            print(f"\n测试区域: {region_name}")
            image = self.capture_window_region(hwnd, region_ratio)
            
            if image:
                filename = f"test_{region_name}.png"
                image.save(filename)
                print(f"截图保存: {filename}")
                
                # OCR识别
                text = self.extract_text_from_image(image)
                print(f"OCR结果: {text[:100]}..." if len(text) > 100 else f"OCR结果: {text}")
            else:
                print(f"截图失败")
            
            print("-" * 30)

def main():
    print("企业微信OCR消息检测测试")
    print("=" * 50)
    
    # 检查tesseract是否可用
    try:
        pytesseract.get_tesseract_version()
        print("✓ Tesseract OCR 可用")
    except Exception as e:
        print("✗ Tesseract OCR 不可用")
        print("请安装 Tesseract OCR: https://github.com/tesseract-ocr/tesseract")
        print("或者设置正确的 tesseract_cmd 路径")
        print(f"错误: {e}")
        return
    
    detector = WeChatOCRDetector()
    
    # 查找企业微信聊天窗口
    print("\n正在查找企业微信聊天窗口...")
    chat_windows = detector.find_wecom_chat_windows()
    
    if not chat_windows:
        print("未找到企业微信聊天窗口")
        return
    
    print(f"找到 {len(chat_windows)} 个企业微信聊天窗口:")
    for i, (hwnd, title) in enumerate(chat_windows):
        print(f"  {i+1}. {title} (句柄: {hwnd})")
    
    if len(chat_windows) == 1:
        selected_window = chat_windows[0]
        print(f"\n自动选择窗口: {selected_window[1]}")
    else:
        try:
            choice = int(input(f"\n请选择要监控的窗口 (1-{len(chat_windows)}): ")) - 1
            if 0 <= choice < len(chat_windows):
                selected_window = chat_windows[choice]
            else:
                print("选择无效，使用第一个窗口")
                selected_window = chat_windows[0]
        except ValueError:
            print("输入无效，使用第一个窗口")
            selected_window = chat_windows[0]
    
    hwnd, window_title = selected_window
    
    # 选择测试模式
    print(f"\n请选择测试模式:")
    print("1. 测试截图和OCR功能")
    print("2. 监控AI回复完成")
    
    try:
        mode = input("请选择模式 (1-2，默认2): ").strip() or "2"
    except:
        mode = "2"
    
    if mode == "1":
        detector.test_screenshot_and_ocr(hwnd, window_title)
    else:
        final_text = detector.monitor_ai_reply(hwnd, window_title)
        if final_text:
            print(f"\n最终获取到的完整回复:")
            print(f"'{final_text}'")

if __name__ == "__main__":
    main()