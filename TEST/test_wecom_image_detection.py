#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
使用图像差异检测企业微信聊天窗口消息区域的变化
不依赖OCR，通过图像差异判断AI回复完成
"""

import time
import win32gui
import win32process
import win32ui
import win32con
from PIL import Image
import hashlib
import os
from wxauto.utils.win32 import GetAllWindows

class WeChatImageDetector:
    def __init__(self):
        self.last_image_hash = None
        self.stable_count = 0
        self.required_stable_checks = 3  # 需要连续3次(15秒)稳定才认为完成
        self.change_history = []
        
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
    
    def activate_window(self, hwnd):
        """激活指定窗口 - 使用与消息转发器相同的方法"""
        try:
            # 使用与wechat_message_forwarder_fixed相同的激活方法
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
            win32gui.BringWindowToTop(hwnd)
            time.sleep(0.3)
            
            # 验证窗口是否成功激活
            current_foreground = win32gui.GetForegroundWindow()
            if current_foreground == hwnd:
                print(f"✓ 窗口已成功激活")
                return True
            else:
                print(f"⚠ 窗口激活可能失败，当前前台窗口句柄: {current_foreground}")
                return False
            
        except Exception as e:
            print(f"激活窗口失败: {e}")
            return False

    def capture_message_area(self, hwnd, region_ratio=(0.05, 0.15, 0.95, 0.75)):
        """
        截取企业微信的消息显示区域
        region_ratio: (left_ratio, top_ratio, right_ratio, bottom_ratio)
        默认截取窗口中间偏上的消息区域
        """
        try:
            # 检查窗口是否在前台，如果不在则自动激活
            if win32gui.GetForegroundWindow() != hwnd:
                self.activate_window(hwnd)
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
            print(f"截取消息区域失败: {e}")
            return None
    
    def calculate_image_hash(self, image):
        """计算图像的哈希值"""
        try:
            # 缩小图像以提高比较速度
            small_image = image.resize((64, 64))
            
            # 转换为灰度
            gray_image = small_image.convert('L')
            
            # 计算像素数据的哈希
            pixel_data = list(gray_image.getdata())
            pixel_str = ''.join(str(p) for p in pixel_data)
            
            return hashlib.md5(pixel_str.encode()).hexdigest()
        
        except Exception as e:
            print(f"计算图像哈希失败: {e}")
            return None
    
    def calculate_image_difference(self, img1, img2):
        """计算两张图像的差异程度"""
        try:
            if img1.size != img2.size:
                # 调整图像大小到相同尺寸
                img2 = img2.resize(img1.size)
            
            # 转换为灰度图像
            gray1 = img1.convert('L')
            gray2 = img2.convert('L')
            
            # 计算像素差异
            pixels1 = list(gray1.getdata())
            pixels2 = list(gray2.getdata())
            
            total_diff = sum(abs(p1 - p2) for p1, p2 in zip(pixels1, pixels2))
            max_possible_diff = len(pixels1) * 255
            
            # 返回差异百分比
            diff_percentage = (total_diff / max_possible_diff) * 100
            return diff_percentage
        
        except Exception as e:
            print(f"计算图像差异失败: {e}")
            return 0
    
    def is_image_stable(self, current_image):
        """判断图像是否稳定（AI回复完成）"""
        current_hash = self.calculate_image_hash(current_image)
        
        if current_hash == self.last_image_hash:
            self.stable_count += 1
        else:
            self.stable_count = 0
            self.last_image_hash = current_hash
        
        # 如果连续多次检查图像都没有变化，认为AI回复完成
        return self.stable_count >= self.required_stable_checks
    
    def monitor_ai_reply_by_image(self, hwnd, window_title):
        """通过图像差异监控AI回复完成状态"""
        print(f"\n开始通过图像差异监控AI回复: {window_title}")
        print(f"窗口句柄: {hwnd}")
        print("=" * 60)
        
        # 自动激活窗口并测试截图
        print("正在自动激活企业微信窗口...")
        activation_success = self.activate_window(hwnd)
        
        print("正在测试截图功能...")
        test_image = self.capture_message_area(hwnd)
        if test_image:
            print("✓ 截图成功")
            # 保存测试图片
            test_image.save("test_message_area.png")
            print("测试截图已保存为: test_message_area.png")
            print(f"图像尺寸: {test_image.size}")
        else:
            print("✗ 截图失败")
            print("故障排除:")
            print(f"- 窗口激活状态: {'成功' if activation_success else '失败'}")
            print(f"- 当前前台窗口: {win32gui.GetForegroundWindow()}")
            print(f"- 目标窗口句柄: {hwnd}")
            print(f"- 窗口可见性: {win32gui.IsWindowVisible(hwnd)}")
            
            if not activation_success:
                print("\n尝试手动激活:")
                input("请手动点击企业微信聊天窗口，然后按回车重试...")
                # 重试截图
                test_image = self.capture_message_area(hwnd)
                if not test_image:
                    return None
        
        print("\n请在企业微信中发送消息给AI，然后观察检测结果...")
        input("准备好后按回车开始监控...")
        
        print("\n开始实时监控图像变化...")
        print("-" * 60)
        
        start_time = time.time()
        check_count = 0
        last_image = None
        changes_detected = []
        
        # 创建变化记录目录
        if not os.path.exists("changes"):
            os.makedirs("changes")
        
        while True:
            check_count += 1
            current_time = time.time()
            elapsed = current_time - start_time
            
            # 截取当前图像
            current_image = self.capture_message_area(hwnd)
            if not current_image:
                print(f"第{check_count}次检查失败：无法截图")
                time.sleep(0.5)
                continue
            
            # 检查图像是否有变化
            if last_image is not None:
                diff_percentage = self.calculate_image_difference(last_image, current_image)
                
                # 如果差异超过阈值，认为有变化
                if diff_percentage > 0.5:  # 0.5% 的差异阈值
                    change_info = {
                        'time': elapsed,
                        'check_count': check_count,
                        'diff_percentage': diff_percentage
                    }
                    changes_detected.append(change_info)
                    
                    print(f"\n时间: {elapsed:.1f}s | 检查: #{check_count}")
                    print(f"检测到图像变化: {diff_percentage:.2f}%")
                    
                    # 保存变化时的图像
                    filename = f"changes/change_{check_count}_{elapsed:.1f}s.png"
                    current_image.save(filename)
                    print(f"变化截图保存: {filename}")
                    print("-" * 40)
                    
                    # 重置稳定计数
                    self.stable_count = 0
                else:
                    # 没有显著变化
                    if check_count % 4 == 0:  # 每4次检查(20秒)显示一次状态
                        print(f"时间: {elapsed:.1f}s | 检查: #{check_count} | 稳定次数: {self.stable_count} | 差异: {diff_percentage:.3f}%")
            
            # 检查是否稳定
            is_stable = self.is_image_stable(current_image)
            
            if is_stable and len(changes_detected) > 0:
                print(f"\n🎉 检测到AI回复完成！")
                print(f"完成时间: {elapsed:.1f}s")
                print(f"连续稳定次数: {self.stable_count}")
                print(f"总变化次数: {len(changes_detected)}")
                
                # 保存最终图像
                current_image.save("final_reply_image.png")
                print("最终回复截图已保存: final_reply_image.png")
                
                # 显示变化统计
                if changes_detected:
                    print("\n变化记录:")
                    for i, change in enumerate(changes_detected[-5:], 1):  # 只显示最后5次变化
                        print(f"  #{i}: {change['time']:.1f}s - 差异{change['diff_percentage']:.2f}%")
                
                break
            
            # 更新上一张图像
            last_image = current_image.copy()
            
            # 防止无限循环，最多监控5分钟
            if elapsed > 300:
                print(f"\n监控超时（5分钟），停止检测")
                break
            
            time.sleep(5.0)  # 每5秒检查一次
        
        return changes_detected
    
    def test_screenshot_regions(self, hwnd, window_title):
        """测试不同区域的截图效果"""
        print(f"\n测试不同区域的截图效果: {window_title}")
        print("=" * 50)
        
        # 自动激活窗口
        print("正在自动激活企业微信窗口...")
        self.activate_window(hwnd)
        
        # 测试不同区域
        regions = {
            "全窗口": (0, 0, 1, 1),
            "上半部分": (0, 0, 1, 0.5),
            "中间消息区": (0.05, 0.15, 0.95, 0.75),
            "下半部分": (0, 0.5, 1, 1),
            "右侧聊天区": (0.25, 0.1, 0.95, 0.8),
            "底部输入区": (0.05, 0.75, 0.95, 0.95)
        }
        
        for region_name, region_ratio in regions.items():
            print(f"\n测试区域: {region_name} {region_ratio}")
            image = self.capture_message_area(hwnd, region_ratio)
            
            if image:
                filename = f"test_region_{region_name}.png"
                image.save(filename)
                print(f"✓ 截图保存: {filename} (尺寸: {image.size})")
                
                # 计算图像哈希
                img_hash = self.calculate_image_hash(image)
                print(f"  图像哈希: {img_hash[:16]}...")
            else:
                print(f"✗ 截图失败")
            
            print("-" * 30)

def main():
    print("企业微信图像差异检测测试")
    print("=" * 50)
    print("这个版本不需要OCR，通过图像差异检测AI回复完成")
    
    detector = WeChatImageDetector()
    
    # 查找企业微信聊天窗口
    print("\n正在查找企业微信聊天窗口...")
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
    print("1. 测试不同区域的截图效果")
    print("2. 监控AI回复完成（图像差异检测）")
    
    try:
        mode = input("请选择模式 (1-2，默认2): ").strip() or "2"
    except:
        mode = "2"
    
    if mode == "1":
        detector.test_screenshot_regions(hwnd, window_title)
        print(f"\n请查看生成的截图文件，确认哪个区域最适合监控消息变化")
    else:
        changes = detector.monitor_ai_reply_by_image(hwnd, window_title)
        if changes:
            print(f"\n检测完成！共发现 {len(changes)} 次图像变化")
            print("检查 changes/ 目录中的截图文件了解变化过程")

if __name__ == "__main__":
    main()