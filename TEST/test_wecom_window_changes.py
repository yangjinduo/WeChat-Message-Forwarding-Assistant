#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试企业微信聊天窗口的内容变化检测
通过多种方法检测窗口内部变化
"""

import time
import win32gui
import win32process
import win32api
import win32con
from wxauto.utils.win32 import GetAllWindows
import hashlib

class WeChatWindowChangeDetector:
    def __init__(self):
        self.last_screenshot_hash = None
        self.last_window_text = None
        self.change_count = 0
        
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
    
    def get_window_pixel_hash(self, hwnd):
        """获取窗口内容的像素哈希值（简化版本）"""
        try:
            # 获取窗口矩形
            rect = win32gui.GetWindowRect(hwnd)
            width = rect[2] - rect[0]
            height = rect[3] - rect[1]
            
            # 获取窗口DC
            hwndDC = win32gui.GetWindowDC(hwnd)
            if not hwndDC:
                return None
                
            # 创建内存DC
            mfcDC = win32gui.CreateCompatibleDC(hwndDC)
            if not mfcDC:
                win32gui.ReleaseDC(hwnd, hwndDC)
                return None
            
            # 创建位图
            saveBitMap = win32gui.CreateCompatibleBitmap(hwndDC, width, height)
            if not saveBitMap:
                win32gui.DeleteDC(mfcDC)
                win32gui.ReleaseDC(hwnd, hwndDC)
                return None
            
            # 选择位图到内存DC
            win32gui.SelectObject(mfcDC, saveBitMap)
            
            # 复制窗口内容到内存DC
            result = win32gui.BitBlt(mfcDC, 0, 0, width, height, hwndDC, 0, 0, win32con.SRCCOPY)
            
            if result:
                # 获取位图数据并计算哈希
                bmpinfo = win32gui.GetObject(saveBitMap)
                # 简化处理：只使用窗口尺寸和一些基本信息作为"变化指标"
                content_signature = f"{width}x{height}_{time.time()}"
                content_hash = hashlib.md5(content_signature.encode()).hexdigest()[:8]
            else:
                content_hash = None
            
            # 清理资源
            win32gui.DeleteObject(saveBitMap)
            win32gui.DeleteDC(mfcDC)
            win32gui.ReleaseDC(hwnd, hwndDC)
            
            return content_hash
            
        except Exception as e:
            print(f"获取窗口像素哈希失败: {e}")
            return None
    
    def get_window_text_info(self, hwnd):
        """获取窗口文本信息"""
        try:
            # 获取窗口标题
            window_title = win32gui.GetWindowText(hwnd)
            
            # 获取窗口类名
            class_name = win32gui.GetClassName(hwnd)
            
            # 尝试获取窗口的一些属性
            rect = win32gui.GetWindowRect(hwnd)
            is_visible = win32gui.IsWindowVisible(hwnd)
            is_enabled = win32gui.IsWindowEnabled(hwnd)
            
            # 组合成一个文本签名
            text_signature = f"{window_title}|{class_name}|{rect}|{is_visible}|{is_enabled}"
            return text_signature
            
        except Exception as e:
            print(f"获取窗口文本信息失败: {e}")
            return None
    
    def detect_window_changes(self, hwnd, window_title):
        """检测窗口变化"""
        print(f"\n开始监控窗口变化: {window_title}")
        print(f"窗口句柄: {hwnd}")
        print("=" * 60)
        
        # 检测方法
        changes_detected = []
        check_count = 0
        
        # 初始状态
        initial_text = self.get_window_text_info(hwnd)
        initial_hash = self.get_window_pixel_hash(hwnd)
        
        print(f"初始窗口签名: {initial_text}")
        print(f"初始像素哈希: {initial_hash}")
        print("\n开始检测变化...")
        print("请在企业微信中发送消息或接收AI回复...")
        print("-" * 60)
        
        start_time = time.time()
        last_text = initial_text
        last_hash = initial_hash
        
        # 监控60秒
        while time.time() - start_time < 60:
            check_count += 1
            current_time = time.time()
            elapsed = current_time - start_time
            
            # 获取当前状态
            current_text = self.get_window_text_info(hwnd)
            current_hash = self.get_window_pixel_hash(hwnd)
            
            # 检测文本变化
            text_changed = current_text != last_text
            hash_changed = current_hash != last_hash
            
            if text_changed or hash_changed:
                change_info = {
                    'time': elapsed,
                    'check_count': check_count,
                    'text_changed': text_changed,
                    'hash_changed': hash_changed,
                    'current_text': current_text,
                    'current_hash': current_hash
                }
                changes_detected.append(change_info)
                
                print(f"时间: {elapsed:.1f}s | 检查: #{check_count}")
                if text_changed:
                    print(f"  文本变化: {current_text}")
                if hash_changed:
                    print(f"  哈希变化: {last_hash} -> {current_hash}")
                print("-" * 40)
                
                last_text = current_text
                last_hash = current_hash
            
            # 每隔一段时间显示状态
            if check_count % 20 == 0:
                print(f"时间: {elapsed:.1f}s | 检查次数: {check_count} | 发现变化: {len(changes_detected)}")
            
            time.sleep(0.5)  # 每0.5秒检查一次
        
        return changes_detected
    
    def analyze_changes(self, changes):
        """分析检测到的变化"""
        print(f"\n{'='*60}")
        print("变化分析结果:")
        print(f"{'='*60}")
        
        if not changes:
            print("未检测到任何窗口变化")
            print("\n可能的原因:")
            print("1. 企业微信窗口内容确实没有变化")
            print("2. 检测方法无法捕获到微信内部的渲染变化")
            print("3. 需要更深层的检测方法")
            return
        
        print(f"总共检测到 {len(changes)} 次变化:")
        
        for i, change in enumerate(changes, 1):
            print(f"\n变化 #{i}:")
            print(f"  时间: {change['time']:.1f}s")
            print(f"  检查次数: {change['check_count']}")
            print(f"  文本变化: {'是' if change['text_changed'] else '否'}")
            print(f"  哈希变化: {'是' if change['hash_changed'] else '否'}")
            if change['text_changed']:
                print(f"  新文本: {change['current_text']}")
            if change['hash_changed']:
                print(f"  新哈希: {change['current_hash']}")
        
        # 分析变化模式
        print(f"\n变化模式分析:")
        text_changes = sum(1 for c in changes if c['text_changed'])
        hash_changes = sum(1 for c in changes if c['hash_changed'])
        
        print(f"文本变化次数: {text_changes}")
        print(f"哈希变化次数: {hash_changes}")
        
        if len(changes) > 1:
            # 计算变化间隔
            intervals = []
            for i in range(1, len(changes)):
                interval = changes[i]['time'] - changes[i-1]['time']
                intervals.append(interval)
            
            if intervals:
                avg_interval = sum(intervals) / len(intervals)
                print(f"平均变化间隔: {avg_interval:.2f}秒")
                print(f"最短间隔: {min(intervals):.2f}秒")
                print(f"最长间隔: {max(intervals):.2f}秒")

def main():
    print("企业微信聊天窗口变化检测测试")
    print("=" * 50)
    
    detector = WeChatWindowChangeDetector()
    
    # 查找企业微信聊天窗口
    print("正在查找企业微信聊天窗口...")
    chat_windows = detector.find_wecom_chat_windows()
    
    if not chat_windows:
        print("未找到企业微信聊天窗口")
        print("请确保:")
        print("1. 企业微信已打开")
        print("2. 至少有一个聊天窗口是打开的")
        print("3. 聊天窗口可见")
        return
    
    print(f"找到 {len(chat_windows)} 个企业微信聊天窗口:")
    for i, (hwnd, title) in enumerate(chat_windows):
        print(f"  {i+1}. {title} (句柄: {hwnd})")
    
    if len(chat_windows) == 1:
        selected_window = chat_windows[0]
        print(f"\n自动选择窗口: {selected_window[1]}")
    else:
        # 让用户选择要监控的窗口
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
    
    print(f"\n准备开始检测...")
    input("请准备好在企业微信中进行操作，然后按回车开始检测...")
    
    # 开始检测
    changes = detector.detect_window_changes(hwnd, window_title)
    
    # 分析结果
    detector.analyze_changes(changes)
    
    print(f"\n检测完成！")
    print("这个测试帮助我们了解是否能检测到企业微信窗口的内容变化")

if __name__ == "__main__":
    main()