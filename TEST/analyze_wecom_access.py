#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
企业微信访问权限分析脚本
深入分析企业微信的访问限制问题
"""

import win32api
import win32gui
import win32process
import win32security
import win32con
from wxauto.utils.win32 import FindWindow, GetText
from wxauto import uiautomation as uia
import ctypes
from ctypes import wintypes

def get_window_info(hwnd):
    """获取窗口详细信息"""
    print(f"窗口详细信息 (句柄: {hwnd})")
    print("-" * 50)
    
    try:
        # 基本窗口信息
        class_name = win32gui.GetClassName(hwnd)
        window_text = GetText(hwnd)
        print(f"类名: {class_name}")
        print(f"标题: {window_text}")
        
        # 窗口状态
        is_visible = win32gui.IsWindowVisible(hwnd)
        is_enabled = win32gui.IsWindowEnabled(hwnd)
        print(f"可见: {is_visible}")
        print(f"启用: {is_enabled}")
        
        # 窗口位置和大小
        rect = win32gui.GetWindowRect(hwnd)
        print(f"位置: {rect}")
        
        # 进程信息
        thread_id, process_id = win32process.GetWindowThreadProcessId(hwnd)
        print(f"进程ID: {process_id}")
        print(f"线程ID: {thread_id}")
        
        # 进程名称
        try:
            process_handle = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION, False, process_id)
            process_name = win32process.GetModuleFileNameEx(process_handle, 0)
            print(f"进程路径: {process_name}")
            win32api.CloseHandle(process_handle)
        except Exception as e:
            print(f"获取进程信息失败: {e}")
            
        # 窗口样式
        style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
        ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        print(f"窗口样式: 0x{style:08X}")
        print(f"扩展样式: 0x{ex_style:08X}")
        
    except Exception as e:
        print(f"获取窗口信息失败: {e}")

def test_uia_access_levels(hwnd):
    """测试不同级别的UIAutomation访问"""
    print(f"\n测试UIAutomation访问级别")
    print("-" * 50)
    
    try:
        # 测试基本控件创建
        print("1. 测试基本控件创建...")
        control = uia.ControlFromHandle(hwnd)
        if control:
            print("✓ 基本控件创建成功")
        else:
            print("✗ 基本控件创建失败")
            return
            
        # 测试基本属性访问
        print("2. 测试基本属性访问...")
        try:
            name = control.Name
            class_name = control.ClassName
            control_type = control.ControlTypeName
            print(f"✓ 名称: {name}")
            print(f"✓ 类名: {class_name}")
            print(f"✓ 类型: {control_type}")
        except Exception as e:
            print(f"✗ 基本属性访问失败: {e}")
            
        # 测试边界访问
        print("3. 测试边界访问...")
        try:
            bounds = control.BoundingRectangle
            print(f"✓ 边界: {bounds}")
            if bounds.width() == 0 or bounds.height() == 0:
                print("⚠️  警告: 边界为空，可能表示访问受限")
        except Exception as e:
            print(f"✗ 边界访问失败: {e}")
            
        # 测试状态访问
        print("4. 测试状态访问...")
        try:
            is_enabled = control.IsEnabled
            is_offscreen = control.CurrentIsOffscreen
            print(f"✓ 启用状态: {is_enabled}")
            print(f"✓ 屏幕外状态: {is_offscreen}")
        except Exception as e:
            print(f"✗ 状态访问失败: {e}")
            
        # 测试子控件访问
        print("5. 测试子控件访问...")
        try:
            children = control.GetChildren()
            print(f"✓ 子控件数量: {len(children)}")
            if len(children) == 0:
                print("⚠️  警告: 无子控件，可能表示内容访问受限")
        except Exception as e:
            print(f"✗ 子控件访问失败: {e}")
            
        # 测试模式和模式支持
        print("6. 测试控件模式...")
        try:
            patterns = control.GetSupportedPatterns()
            print(f"✓ 支持的模式数量: {len(patterns)}")
            for pattern in patterns:
                print(f"  - {pattern}")
        except Exception as e:
            print(f"✗ 模式访问失败: {e}")
            
    except Exception as e:
        print(f"UIAutomation测试失败: {e}")

def test_alternative_access_methods(hwnd):
    """测试替代访问方法"""
    print(f"\n测试替代访问方法")
    print("-" * 50)
    
    # 方法1: 直接Win32 API
    print("1. 测试Win32 API直接访问...")
    try:
        def enum_child_proc(child_hwnd, param):
            try:
                child_class = win32gui.GetClassName(child_hwnd)
                child_text = GetText(child_hwnd)
                child_rect = win32gui.GetWindowRect(child_hwnd)
                print(f"  子窗口: {child_hwnd} - {child_class} - {child_text}")
                param.append((child_hwnd, child_class, child_text))
            except:
                pass
            return True
        
        children = []
        win32gui.EnumChildWindows(hwnd, enum_child_proc, children)
        print(f"✓ 通过Win32 API找到 {len(children)} 个子窗口")
        
    except Exception as e:
        print(f"✗ Win32 API访问失败: {e}")
    
    # 方法2: 测试发送消息
    print("2. 测试窗口消息...")
    try:
        # 发送WM_GETTEXT消息获取窗口文本
        result = win32gui.SendMessage(hwnd, win32con.WM_GETTEXT, 1024, ctypes.create_unicode_buffer(1024))
        print(f"✓ 消息访问结果: {result}")
    except Exception as e:
        print(f"✗ 消息访问失败: {e}")

def check_security_context():
    """检查安全上下文"""
    print(f"\n检查安全上下文和权限")
    print("-" * 50)
    
    try:
        # 检查当前用户权限
        import os
        print(f"当前用户: {os.getlogin()}")
        
        # 检查是否以管理员身份运行
        try:
            is_admin = ctypes.windll.shell32.IsUserAnAdmin()
            print(f"管理员权限: {'是' if is_admin else '否'}")
        except:
            print("无法检查管理员权限")
            
        # 检查UIAutomation可用性
        try:
            import uiautomation
            print("✓ UIAutomation模块可用")
        except Exception as e:
            print(f"✗ UIAutomation模块问题: {e}")
            
    except Exception as e:
        print(f"安全上下文检查失败: {e}")

def analyze_wecom_protection():
    """分析企业微信保护机制"""
    print(f"\n分析企业微信保护机制")
    print("-" * 50)
    
    hwnd = FindWindow(classname='WeWorkWindow', name='企业微信')
    if not hwnd:
        print("未找到企业微信窗口")
        return
        
    try:
        # 检查进程保护
        thread_id, process_id = win32process.GetWindowThreadProcessId(hwnd)
        
        try:
            # 尝试打开进程句柄
            process_handle = win32api.OpenProcess(
                win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, 
                False, 
                process_id
            )
            print("✓ 进程可访问")
            win32api.CloseHandle(process_handle)
        except Exception as e:
            print(f"✗ 进程访问受限: {e}")
            
        # 检查窗口层次结构
        parent = win32gui.GetParent(hwnd)
        owner = win32gui.GetWindow(hwnd, win32con.GW_OWNER)
        print(f"父窗口: {parent}")
        print(f"拥有者窗口: {owner}")
        
    except Exception as e:
        print(f"保护机制分析失败: {e}")

def main():
    """主函数"""
    print("企业微信访问权限深度分析")
    print("=" * 80)
    
    # 查找企业微信窗口
    hwnd = FindWindow(classname='WeWorkWindow', name='企业微信')
    if not hwnd:
        print("未找到企业微信窗口")
        return
        
    # 执行各项分析
    get_window_info(hwnd)
    test_uia_access_levels(hwnd)
    test_alternative_access_methods(hwnd)
    check_security_context()
    analyze_wecom_protection()
    
    print("\n" + "=" * 80)
    print("分析结论:")
    print("如果边界为(0,0,0,0)且无子控件，说明企业微信启用了UI访问保护")
    print("这可能是出于安全考虑，防止自动化工具访问敏感的企业数据")
    print("解决方案可能需要:")
    print("1. 管理员权限运行")
    print("2. 企业微信的特殊配置")
    print("3. 使用图像识别等替代方案")
    print("4. 联系企业微信官方了解自动化接口")

if __name__ == "__main__":
    main()
    input("\n按回车键退出...")