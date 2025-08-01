#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试企业微信是否使用自定义渲染
"""

from wxauto.utils.win32 import FindWindow
from wxauto import uiautomation as uia
import win32gui
import win32api
import win32con

def test_custom_rendering():
    """测试企业微信是否使用自定义渲染"""
    print("测试企业微信自定义渲染")
    print("=" * 50)
    
    hwnd = FindWindow(classname='WeWorkWindow', name='企业微信')
    if not hwnd:
        print("未找到企业微信窗口")
        return
        
    try:
        # 获取窗口DC和检查绘制方式
        hdc = win32gui.GetWindowDC(hwnd)
        if hdc:
            print("✓ 可以获取窗口DC")
            win32gui.ReleaseDC(hwnd, hdc)
        
        # 检查窗口类信息
        class_info = win32gui.GetClassLong(hwnd, win32con.GCL_STYLE)
        print(f"窗口类样式: 0x{class_info:08X}")
        
        # 检查是否有子窗口（Win32层面）
        def enum_child_proc(child_hwnd, param):
            try:
                child_class = win32gui.GetClassName(child_hwnd)
                child_rect = win32gui.GetWindowRect(child_hwnd)
                param.append((child_hwnd, child_class, child_rect))
            except:
                pass
            return True
        
        children = []
        win32gui.EnumChildWindows(hwnd, enum_child_proc, children)
        print(f"Win32子窗口数量: {len(children)}")
        
        for child_hwnd, child_class, child_rect in children:
            print(f"  子窗口: {child_class} - {child_rect}")
        
        # 尝试UIAutomation的其他方法
        control = uia.ControlFromHandle(hwnd)
        
        # 尝试不同的查找方法
        print("\n尝试不同的UIAutomation查找方法:")
        
        # 方法1: FindAll
        try:
            all_descendants = control.FindAll(uia.ControlCondition(uia.ControlType.ANY), uia.TreeScope.Descendants)
            print(f"FindAll找到: {len(all_descendants)} 个后代控件")
        except Exception as e:
            print(f"FindAll失败: {e}")
        
        # 方法2: 尝试特定控件类型
        try:
            buttons = control.FindAll(uia.ControlCondition(uia.ControlType.ButtonControl), uia.TreeScope.Descendants)
            print(f"按钮控件: {len(buttons)} 个")
        except Exception as e:
            print(f"查找按钮失败: {e}")
            
        try:
            texts = control.FindAll(uia.ControlCondition(uia.ControlType.TextControl), uia.TreeScope.Descendants)
            print(f"文本控件: {len(texts)} 个")
        except Exception as e:
            print(f"查找文本失败: {e}")
            
        try:
            lists = control.FindAll(uia.ControlCondition(uia.ControlType.ListControl), uia.TreeScope.Descendants)
            print(f"列表控件: {len(lists)} 个")
        except Exception as e:
            print(f"查找列表失败: {e}")
        
    except Exception as e:
        print(f"测试失败: {e}")

def analyze_rendering_technology():
    """分析企业微信使用的渲染技术"""
    print("\n分析渲染技术")
    print("=" * 50)
    
    hwnd = FindWindow(classname='WeWorkWindow', name='企业微信')
    if not hwnd:
        return
        
    try:
        # 检查进程模块
        import win32process
        import win32api
        
        thread_id, process_id = win32process.GetWindowThreadProcessId(hwnd)
        process_handle = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, process_id)
        
        try:
            modules = win32process.EnumProcessModules(process_handle)
            print(f"进程加载的模块数量: {len(modules)}")
            
            # 检查关键的渲染相关DLL
            rendering_dlls = ['d3d9.dll', 'd3d11.dll', 'opengl32.dll', 'gdi32.dll', 'user32.dll', 'dwmapi.dll']
            
            for module in modules[:20]:  # 只检查前20个模块
                try:
                    module_name = win32process.GetModuleFileNameEx(process_handle, module)
                    dll_name = module_name.split('\\')[-1].lower()
                    
                    if any(render_dll in dll_name for render_dll in rendering_dlls):
                        print(f"  渲染相关模块: {dll_name}")
                        
                except:
                    continue
                    
        except Exception as e:
            print(f"模块枚举失败: {e}")
        
        win32api.CloseHandle(process_handle)
        
    except Exception as e:
        print(f"分析渲染技术失败: {e}")

def conclusion():
    """得出结论"""
    print("\n" + "=" * 60)
    print("结论和建议")
    print("=" * 60)
    print()
    print("基于分析结果，企业微信很可能使用了以下技术:")
    print("1. 自定义渲染引擎 (可能基于DirectX/OpenGL)")
    print("2. 非标准Windows控件")
    print("3. 内容完全自绘，不暴露给UIAutomation")
    print()
    print("这解释了为什么:")
    print("- 可以找到主窗口")
    print("- 边界信息正常")
    print("- 但无法访问内部UI元素")
    print()
    print("解决方案建议:")
    print("1. 使用图像识别技术 (OCR + 图像匹配)")
    print("2. 尝试Win32消息模拟")
    print("3. 寻找企业微信官方API")
    print("4. 使用坐标点击的方式")
    print("5. 联系腾讯了解自动化接口")

if __name__ == "__main__":
    print("请确保以管理员权限运行此脚本")
    print()
    
    test_custom_rendering()
    analyze_rendering_technology()
    conclusion()
    
    input("\n按回车键退出...")