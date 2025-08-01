#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
精确查找企业微信聊天窗口
专门寻找"无人机AI助教"等企业微信独立聊天窗口
"""

import win32gui
import win32process
from wxauto.utils.win32 import GetAllWindows, GetText, FindWindow
from wxauto import uiautomation as uia

def find_all_wecom_related_windows():
    """查找所有企业微信相关窗口"""
    print("查找所有企业微信相关窗口...")
    print("=" * 80)
    
    all_windows = GetAllWindows()
    wecom_related = []
    
    for hwnd, class_name, window_title in all_windows:
        # 检查进程是否是企业微信
        try:
            thread_id, process_id = win32process.GetWindowThreadProcessId(hwnd)
            process_handle = win32process.OpenProcess(win32process.PROCESS_QUERY_INFORMATION, False, process_id)
            try:
                process_name = win32process.GetModuleFileNameEx(process_handle, 0)
                if 'WXWork.exe' in process_name or 'WeWork' in process_name:
                    wecom_related.append((hwnd, class_name, window_title, process_name))
                    print(f"企业微信相关窗口:")
                    print(f"  句柄: {hwnd}")
                    print(f"  类名: {class_name}")
                    print(f"  标题: {window_title}")
                    print(f"  进程: {process_name}")
                    
                    # 检查窗口状态
                    is_visible = win32gui.IsWindowVisible(hwnd)
                    is_enabled = win32gui.IsWindowEnabled(hwnd)
                    rect = win32gui.GetWindowRect(hwnd)
                    print(f"  可见: {is_visible}, 启用: {is_enabled}")
                    print(f"  位置: {rect}")
                    print("-" * 50)
            except:
                pass
            finally:
                win32process.CloseHandle(process_handle)
        except:
            pass
    
    return wecom_related

def find_chat_window_by_title(target_title):
    """通过标题查找聊天窗口"""
    print(f"\n直接查找标题为 '{target_title}' 的窗口...")
    print("=" * 80)
    
    # 方法1: 使用FindWindow直接查找
    hwnd = FindWindow(name=target_title)
    if hwnd:
        print(f"✓ 通过FindWindow找到窗口: {hwnd}")
        class_name = win32gui.GetClassName(hwnd)
        print(f"  类名: {class_name}")
        
        # 检查是否属于企业微信进程
        try:
            thread_id, process_id = win32process.GetWindowThreadProcessId(hwnd)
            process_handle = win32process.OpenProcess(win32process.PROCESS_QUERY_INFORMATION, False, process_id)
            process_name = win32process.GetModuleFileNameEx(process_handle, 0)
            win32process.CloseHandle(process_handle)
            print(f"  进程: {process_name}")
            
            if 'WXWork.exe' in process_name:
                print("✅ 确认是企业微信进程")
                return hwnd, class_name
            else:
                print("❌ 不是企业微信进程")
        except Exception as e:
            print(f"检查进程失败: {e}")
    
    # 方法2: 遍历所有窗口查找
    print(f"\n遍历查找标题包含 '{target_title}' 的窗口...")
    all_windows = GetAllWindows()
    
    for hwnd, class_name, window_title in all_windows:
        if target_title in window_title:
            print(f"找到匹配窗口:")
            print(f"  句柄: {hwnd}")
            print(f"  类名: {class_name}")
            print(f"  标题: {window_title}")
            
            # 检查进程
            try:
                thread_id, process_id = win32process.GetWindowThreadProcessId(hwnd)
                process_handle = win32process.OpenProcess(win32process.PROCESS_QUERY_INFORMATION, False, process_id)
                process_name = win32process.GetModuleFileNameEx(process_handle, 0)
                win32process.CloseHandle(process_handle)
                
                if 'WXWork.exe' in process_name:
                    print(f"✅ 确认是企业微信聊天窗口")
                    return hwnd, class_name
                else:
                    print(f"  进程: {process_name} (不是企业微信)")
            except:
                pass
    
    return None, None

def test_chat_window_detailed(hwnd, class_name, title):
    """详细测试聊天窗口"""
    print(f"\n详细测试聊天窗口: {title}")
    print("=" * 80)
    
    try:
        # UIAutomation控件测试
        control = uia.ControlFromHandle(hwnd)
        if not control:
            print("❌ 无法创建UIAutomation控件")
            return False
        
        print("✅ UIAutomation控件创建成功")
        print(f"  名称: {control.Name}")
        print(f"  类名: {control.ClassName}")
        print(f"  类型: {control.ControlTypeName}")
        print(f"  边界: {control.BoundingRectangle}")
        
        # 测试子控件
        children = control.GetChildren()
        print(f"  子控件数量: {len(children)}")
        
        if len(children) == 0:
            print("⚠️  没有子控件，可能使用了自定义渲染")
            
            # 尝试其他UIAutomation方法
            print("尝试其他UIAutomation方法:")
            
            # 方法1: 查找特定控件类型
            try:
                # 查找编辑框
                edit_controls = []
                def find_controls(parent, depth=0):
                    if depth > 5:
                        return
                    try:
                        for child in parent.GetChildren():
                            if child.ControlTypeName in ['EditControl', 'DocumentControl']:
                                edit_controls.append(child)
                            find_controls(child, depth + 1)
                    except:
                        pass
                
                find_controls(control)
                print(f"  找到编辑控件: {len(edit_controls)} 个")
                
                for i, edit_ctrl in enumerate(edit_controls[:3]):
                    print(f"    编辑控件[{i}]: {edit_ctrl.ClassName} - {edit_ctrl.Name}")
                    
            except Exception as e:
                print(f"  查找编辑控件失败: {e}")
            
            return False
        else:
            print("✅ 找到子控件，分析结构:")
            for i, child in enumerate(children[:5]):
                try:
                    print(f"  子控件[{i}]: {child.ClassName} - {child.Name} - {child.ControlTypeName}")
                    print(f"    边界: {child.BoundingRectangle}")
                    
                    # 递归查看子控件
                    grandchildren = child.GetChildren()
                    if len(grandchildren) > 0:
                        print(f"    包含 {len(grandchildren)} 个子元素")
                        for j, gc in enumerate(grandchildren[:3]):
                            print(f"      [{j}] {gc.ClassName} - {gc.Name}")
                except Exception as e:
                    print(f"  子控件[{i}] 分析失败: {e}")
            
            return True
            
    except Exception as e:
        print(f"❌ UIAutomation测试失败: {e}")
        return False

def main():
    """主函数"""
    print("企业微信聊天窗口精确查找工具")
    print("请确保以管理员权限运行")
    print()
    
    # 第一步: 查找所有企业微信相关窗口
    wecom_windows = find_all_wecom_related_windows()
    
    # 第二步: 专门查找"无人机AI助教"窗口
    target_title = "无人机AI助教"
    print(f"\n正在查找 '{target_title}' 聊天窗口...")
    
    hwnd, class_name = find_chat_window_by_title(target_title)
    
    if hwnd:
        print(f"\n🎉 找到目标聊天窗口!")
        print(f"句柄: {hwnd}")
        print(f"类名: {class_name}")
        
        # 详细测试这个窗口
        success = test_chat_window_detailed(hwnd, class_name, target_title)
        
        if success:
            print(f"\n✅ '{target_title}' 聊天窗口可以进行UIAutomation操作!")
            print("消息转发助手应该可以正常工作。")
        else:
            print(f"\n❌ '{target_title}' 聊天窗口无法进行完整的UIAutomation操作")
            print("可能需要使用其他方法（如坐标点击）进行操作。")
    else:
        print(f"\n❌ 未找到 '{target_title}' 聊天窗口")
        print("请检查:")
        print("1. 企业微信是否已启动并登录")
        print("2. 是否已打开'无人机AI助教'的独立聊天窗口")
        print("3. 聊天窗口是否处于可见状态（未最小化）")
        print("4. 窗口标题是否完全匹配'无人机AI助教'")
        
        print(f"\n所有企业微信相关窗口:")
        if wecom_windows:
            for hwnd, class_name, window_title, process_name in wecom_windows:
                print(f"  {window_title} (类名: {class_name})")
        else:
            print("  未找到任何企业微信相关窗口")

if __name__ == "__main__":
    main()
    input("\n按回车键退出...")