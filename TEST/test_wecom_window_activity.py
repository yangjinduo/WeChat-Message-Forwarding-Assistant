#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试企业微信聊天窗口的活动频率
用于检测AI回复完成的时机
"""

import time
import win32gui
import win32process
from wxauto.utils.win32 import GetAllWindows

class WeChatWindowActivityMonitor:
    def __init__(self):
        self.last_update_time = {}
        self.update_count = {}
        self.monitoring = False
        
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
    
    def get_window_update_info(self, hwnd):
        """获取窗口更新信息"""
        try:
            # 获取窗口矩形区域
            rect = win32gui.GetWindowRect(hwnd)
            
            # 获取窗口更新区域信息
            # 这里我们简单使用窗口是否重绘来判断活动
            current_time = time.time()
            
            # 检查窗口是否活跃（可见且在前台）
            is_active = win32gui.GetForegroundWindow() == hwnd
            
            return {
                'rect': rect,
                'time': current_time,
                'is_active': is_active,
                'is_visible': win32gui.IsWindowVisible(hwnd)
            }
        except Exception as e:
            print(f"获取窗口信息失败: {e}")
            return None
    
    def monitor_window_activity(self, hwnd, window_title, duration=60):
        """监控窗口活动频率"""
        print(f"\n开始监控窗口: {window_title}")
        print(f"窗口句柄: {hwnd}")
        print(f"监控时长: {duration} 秒")
        print("-" * 50)
        
        start_time = time.time()
        last_check_time = start_time
        activity_count = 0
        
        # 记录活动变化
        activity_log = []
        
        while time.time() - start_time < duration:
            current_time = time.time()
            
            # 每0.5秒检查一次
            if current_time - last_check_time >= 0.5:
                window_info = self.get_window_update_info(hwnd)
                
                if window_info:
                    activity_count += 1
                    
                    # 检查窗口是否发生变化
                    is_foreground = win32gui.GetForegroundWindow() == hwnd
                    
                    activity_log.append({
                        'time': current_time,
                        'elapsed': current_time - start_time,
                        'is_foreground': is_foreground,
                        'is_visible': window_info['is_visible']
                    })
                    
                    # 实时显示状态
                    elapsed = current_time - start_time
                    status = "前台" if is_foreground else "后台"
                    print(f"时间: {elapsed:.1f}s | 状态: {status} | 检查次数: {activity_count}")
                
                last_check_time = current_time
            
            time.sleep(0.1)  # 短暂休眠避免过度占用CPU
        
        return activity_log
    
    def analyze_activity_pattern(self, activity_log):
        """分析活动模式"""
        print("\n" + "="*50)
        print("活动模式分析:")
        print("="*50)
        
        if not activity_log:
            print("没有活动数据")
            return
        
        # 计算活动频率
        total_time = activity_log[-1]['elapsed']
        total_checks = len(activity_log)
        avg_frequency = total_checks / total_time if total_time > 0 else 0
        
        print(f"总监控时间: {total_time:.1f} 秒")
        print(f"总检查次数: {total_checks}")
        print(f"平均检查频率: {avg_frequency:.2f} 次/秒")
        
        # 分析前台/后台时间
        foreground_time = sum(1 for log in activity_log if log['is_foreground'])
        background_time = total_checks - foreground_time
        
        print(f"前台时间比例: {foreground_time/total_checks*100:.1f}%")
        print(f"后台时间比例: {background_time/total_checks*100:.1f}%")
        
        # 检测活动变化模式
        print("\n近期活动状态变化:")
        last_10_logs = activity_log[-10:] if len(activity_log) >= 10 else activity_log
        for log in last_10_logs:
            status = "前台" if log['is_foreground'] else "后台"
            print(f"  {log['elapsed']:.1f}s: {status}")

def main():
    print("企业微信聊天窗口活动频率监控测试")
    print("=" * 50)
    
    monitor = WeChatWindowActivityMonitor()
    
    # 查找企业微信聊天窗口
    print("正在查找企业微信聊天窗口...")
    chat_windows = monitor.find_wecom_chat_windows()
    
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
    
    # 设置监控时长
    try:
        duration = int(input("\n请输入监控时长（秒，默认30秒）: ") or "30")
    except ValueError:
        duration = 30
    
    print(f"\n请在企业微信窗口 '{window_title}' 中进行一些操作")
    print("比如发送消息、接收AI回复等...")
    print("监控即将开始...")
    
    time.sleep(3)  # 给用户准备时间
    
    # 开始监控
    activity_log = monitor.monitor_window_activity(hwnd, window_title, duration)
    
    # 分析结果
    monitor.analyze_activity_pattern(activity_log)
    
    print(f"\n监控完成！")
    print("分析结果可以帮助我们了解企业微信窗口的活动模式")
    print("特别是AI回复时的活动变化规律")

if __name__ == "__main__":
    main()