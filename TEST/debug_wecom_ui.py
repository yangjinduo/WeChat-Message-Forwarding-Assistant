#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
企业微信UI结构详细分析脚本
用于深入分析企业微信的窗口结构
"""

import win32gui
from wxauto.utils.win32 import FindWindow, GetText
from wxauto import uiautomation as uia

def find_wecom_main_window():
    """查找企业微信主窗口"""
    print("查找企业微信主窗口...")
    
    # 直接查找WeWorkWindow类名的窗口
    hwnd = FindWindow(classname='WeWorkWindow', name='企业微信')
    if hwnd:
        print(f"找到企业微信窗口: 句柄={hwnd}")
        return hwnd
    
    print("未找到企业微信主窗口")
    return None

def analyze_ui_structure(hwnd):
    """深度分析UI结构"""
    print(f"\n详细分析企业微信UI结构 (句柄: {hwnd})")
    print("=" * 80)
    
    try:
        control = uia.ControlFromHandle(hwnd)
        print(f"主窗口信息:")
        print(f"  类名: {control.ClassName}")
        print(f"  名称: {control.Name}")
        print(f"  控件类型: {control.ControlTypeName}")
        print(f"  边界: {control.BoundingRectangle}")
        
        def analyze_children(parent_control, level=0, max_level=4):
            """递归分析子控件"""
            if level > max_level:
                return
                
            indent = "  " * level
            children = parent_control.GetChildren()
            
            print(f"{indent}├─ 子控件数量: {len(children)}")
            
            for i, child in enumerate(children):
                try:
                    print(f"{indent}├─ 子控件 [{i}]:")
                    print(f"{indent}│  类名: {child.ClassName}")
                    print(f"{indent}│  名称: {child.Name}")
                    print(f"{indent}│  类型: {child.ControlTypeName}")
                    print(f"{indent}│  边界: {child.BoundingRectangle}")
                    print(f"{indent}│  可见: {child.IsEnabled}")
                    print(f"{indent}│  状态: {child.CurrentIsOffscreen}")
                    
                    # 如果是重要的控件类型，进一步分析
                    if child.ControlTypeName in ['Pane', 'Window', 'Group']:
                        analyze_children(child, level + 1, max_level)
                    
                    print(f"{indent}│")
                    
                except Exception as e:
                    print(f"{indent}│  分析子控件失败: {e}")
                    
        analyze_children(control)
        
    except Exception as e:
        print(f"分析UI结构失败: {e}")

def test_navigation_methods():
    """测试不同的UI导航方法"""
    print("\n测试不同的UI导航方法")
    print("=" * 50)
    
    hwnd = find_wecom_main_window()
    if not hwnd:
        return
        
    try:
        control = uia.ControlFromHandle(hwnd)
        
        # 方法1: 通过GetChildren获取
        print("方法1: 直接获取所有子控件")
        children = control.GetChildren()
        print(f"直接子控件数量: {len(children)}")
        for i, child in enumerate(children):
            print(f"  子控件[{i}]: {child.ClassName} - {child.Name}")
        
        # 方法2: 通过特定控件类型查找
        print("\n方法2: 查找特定控件类型")
        panes = control.GetChildren(lambda c: c.ControlTypeName == 'Pane')
        print(f"Pane控件数量: {len(panes)}")
        for i, pane in enumerate(panes):
            print(f"  Pane[{i}]: {pane.ClassName} - {pane.Name}")
        
        # 方法3: 尝试查找常见的控件类名
        print("\n方法3: 查找常见控件类名")
        common_classes = ['NavigationPane', 'SessionList', 'ChatBox', 'MessageList']
        for class_name in common_classes:
            try:
                found_controls = control.GetChildren(lambda c: class_name.lower() in c.ClassName.lower())
                if found_controls:
                    print(f"  找到 {class_name} 类型控件: {len(found_controls)} 个")
                    for ctrl in found_controls:
                        print(f"    - {ctrl.ClassName}: {ctrl.Name}")
            except:
                pass
                
    except Exception as e:
        print(f"测试导航方法失败: {e}")

def test_standard_wechat_ui_approach():
    """测试标准微信UI查找方法"""
    print("\n测试标准微信UI查找方法")
    print("=" * 50)
    
    hwnd = find_wecom_main_window()
    if not hwnd:
        return
        
    try:
        control = uia.ControlFromHandle(hwnd)
        
        # 尝试wxauto原始的查找方法
        print("尝试标准wxauto方法:")
        try:
            MainControl1 = [i for i in control.GetChildren() if not i.ClassName][0]
            print(f"  MainControl1找到: {MainControl1.ClassName}")
            
            MainControl2 = MainControl1.GetFirstChildControl()
            print(f"  MainControl2找到: {MainControl2.ClassName}")
            
            children = MainControl2.GetChildren()
            print(f"  MainControl2子控件数量: {len(children)}")
            
            for i, child in enumerate(children):
                print(f"    子控件[{i}]: {child.ClassName} - {child.Name} - 边界:{child.BoundingRectangle}")
                
        except Exception as e:
            print(f"  标准方法失败: {e}")
            
        # 尝试其他可能的方法
        print("\n尝试其他查找方法:")
        try:
            # 查找所有无类名的控件
            no_class_controls = [i for i in control.GetChildren() if not i.ClassName]
            print(f"  无类名控件数量: {len(no_class_controls)}")
            
            # 查找所有Pane类型控件
            pane_controls = [i for i in control.GetChildren() if i.ControlTypeName == 'Pane']
            print(f"  Pane类型控件数量: {len(pane_controls)}")
            
        except Exception as e:
            print(f"  其他方法失败: {e}")
            
    except Exception as e:
        print(f"测试失败: {e}")

def main():
    """主函数"""
    print("企业微信UI结构深度分析工具")
    print("请确保企业微信已启动并登录")
    print()
    
    hwnd = find_wecom_main_window()
    if hwnd:
        analyze_ui_structure(hwnd)
        test_navigation_methods()
        test_standard_wechat_ui_approach()
    else:
        print("无法找到企业微信窗口，请检查:")
        print("1. 企业微信是否已启动")
        print("2. 企业微信是否已登录")
        print("3. 企业微信窗口是否可见")

if __name__ == "__main__":
    main()
    input("\n按回车键退出...")