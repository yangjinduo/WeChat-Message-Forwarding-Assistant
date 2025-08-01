#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
分析企业微信2个子控件的具体结构
"""

from wxauto.utils.win32 import FindWindow
from wxauto import uiautomation as uia

def analyze_two_controls():
    """详细分析2个子控件的结构"""
    print("分析企业微信的2个子控件结构")
    print("=" * 60)
    
    hwnd = FindWindow(classname='WeWorkWindow', name='企业微信')
    if not hwnd:
        print("未找到企业微信窗口")
        return
        
    try:
        control = uia.ControlFromHandle(hwnd)
        children = control.GetChildren()
        
        print(f"找到 {len(children)} 个子控件")
        print()
        
        for i, child in enumerate(children):
            print(f"=== 子控件 [{i}] ===")
            print(f"类名: {child.ClassName}")
            print(f"名称: {child.Name}")
            print(f"类型: {child.ControlTypeName}")
            print(f"边界: {child.BoundingRectangle}")
            print(f"可见: {child.IsEnabled}")
            
            # 分析每个子控件的子元素
            try:
                grandchildren = child.GetChildren()
                print(f"子元素数量: {len(grandchildren)}")
                
                for j, grandchild in enumerate(grandchildren):
                    print(f"  └─ 子元素[{j}]: {grandchild.ClassName} - {grandchild.Name} - {grandchild.ControlTypeName}")
                    print(f"     边界: {grandchild.BoundingRectangle}")
                    
                    # 再深入一层
                    try:
                        great_grandchildren = grandchild.GetChildren()
                        if len(great_grandchildren) > 0:
                            print(f"     └─ 更深层元素数量: {len(great_grandchildren)}")
                            for k, ggchild in enumerate(great_grandchildren[:3]):  # 只显示前3个
                                print(f"        └─ [{k}]: {ggchild.ClassName} - {ggchild.Name}")
                    except:
                        pass
                        
            except Exception as e:
                print(f"分析子元素失败: {e}")
            
            print()
            
    except Exception as e:
        print(f"分析失败: {e}")

def guess_control_roles():
    """推测2个控件的功能角色"""
    print("推测控件功能角色")
    print("=" * 60)
    
    hwnd = FindWindow(classname='WeWorkWindow', name='企业微信')
    if not hwnd:
        print("未找到企业微信窗口")
        return
        
    try:
        control = uia.ControlFromHandle(hwnd)
        children = control.GetChildren()
        
        # 按边界大小和位置推测功能
        for i, child in enumerate(children):
            bounds = child.BoundingRectangle
            width = bounds.width()
            height = bounds.height()
            left = bounds.left
            
            print(f"控件[{i}] 功能推测:")
            print(f"  位置: ({left}, {bounds.top})")
            print(f"  大小: {width} x {height}")
            
            # 根据位置和大小推测
            if width < 400:
                print(f"  推测: 可能是导航栏/侧边栏 (宽度较小)")
            elif width > 800:
                print(f"  推测: 可能是主聊天区域 (宽度较大)")
            else:
                print(f"  推测: 可能是会话列表 (中等宽度)")
                
            if left < 100:
                print(f"  推测: 位于左侧")
            elif left > 500:
                print(f"  推测: 位于右侧或中央")
                
            print()
            
    except Exception as e:
        print(f"推测失败: {e}")

if __name__ == "__main__":
    print("请确保以管理员权限运行此脚本")
    print()
    
    analyze_two_controls()
    guess_control_roles()
    
    input("\n按回车键退出...")