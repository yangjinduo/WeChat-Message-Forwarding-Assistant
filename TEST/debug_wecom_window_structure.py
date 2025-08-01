#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
深度分析企业微信独立聊天窗口的UIAutomation结构
专门用于调试"无人机AI助教"等聊天窗口的控件信息
"""

from wxauto.utils.win32 import FindWindow
from wxauto import uiautomation as uia
import win32gui

def analyze_wecom_window_detailed(window_title="无人机AI助教"):
    """详细分析企业微信聊天窗口结构"""
    print(f"深度分析企业微信聊天窗口: {window_title}")
    print("=" * 80)
    
    # 查找窗口
    hwnd = FindWindow(name=window_title)
    if not hwnd:
        print(f"❌ 未找到窗口: {window_title}")
        print("请确保该聊天窗口已打开且可见")
        return
    
    class_name = win32gui.GetClassName(hwnd)
    rect = win32gui.GetWindowRect(hwnd)
    
    print(f"✅ 找到窗口:")
    print(f"  句柄: {hwnd}")
    print(f"  类名: {class_name}")
    print(f"  标题: {window_title}")
    print(f"  位置: {rect}")
    print(f"  大小: {rect[2]-rect[0]} x {rect[3]-rect[1]}")
    print()
    
    try:
        # 创建UIAutomation控件
        control = uia.ControlFromHandle(hwnd)
        if not control:
            print("❌ 无法创建UIAutomation控件")
            return
        
        print("✅ UIAutomation控件创建成功")
        print(f"  控件名称: {control.Name}")
        print(f"  控件类名: {control.ClassName}")
        print(f"  控件类型: {control.ControlTypeName}")
        print(f"  控件边界: {control.BoundingRectangle}")
        print()
        
        # 递归分析所有子控件
        print("🔍 开始深度分析子控件结构...")
        print("-" * 80)
        analyze_children_recursive(control, depth=0, max_depth=6)
        
    except Exception as e:
        print(f"❌ 分析失败: {e}")

def analyze_children_recursive(parent_control, depth=0, max_depth=6):
    """递归分析子控件"""
    if depth > max_depth:
        return
    
    indent = "  " * depth
    
    try:
        children = parent_control.GetChildren()
        
        if len(children) == 0:
            print(f"{indent}└─ 无子控件")
            return
        
        print(f"{indent}├─ 子控件数量: {len(children)}")
        
        for i, child in enumerate(children):
            try:
                # 获取控件基本信息
                child_name = child.Name
                child_class = child.ClassName
                child_type = child.ControlTypeName
                child_bounds = child.BoundingRectangle
                child_enabled = child.IsEnabled
                
                # 判断是否可能是输入相关控件
                is_input_candidate = False
                input_indicators = []
                
                # 检查控件类型
                if child_type in ['EditControl', 'DocumentControl', 'TextControl']:
                    is_input_candidate = True
                    input_indicators.append(f"类型:{child_type}")
                
                # 检查类名
                if any(keyword in child_class.lower() for keyword in ['edit', 'input', 'text', 'rich']):
                    is_input_candidate = True
                    input_indicators.append(f"类名:{child_class}")
                
                # 检查名称
                if child_name and any(keyword in child_name.lower() for keyword in ['输入', 'input', 'edit', '消息']):
                    is_input_candidate = True
                    input_indicators.append(f"名称:{child_name}")
                
                # 检查边界（输入框通常在窗口底部）
                if child_bounds.height() > 20 and child_bounds.width() > 100:
                    window_rect = parent_control.BoundingRectangle
                    # 如果控件在窗口下半部分
                    if child_bounds.top > window_rect.top + window_rect.height() * 0.6:
                        input_indicators.append("位置:窗口下部")
                
                # 输出控件信息
                prefix = "🎯" if is_input_candidate else "├─"
                print(f"{indent}{prefix} 子控件[{i}]:")
                print(f"{indent}   名称: '{child_name}'")
                print(f"{indent}   类名: {child_class}")
                print(f"{indent}   类型: {child_type}")
                print(f"{indent}   边界: {child_bounds}")
                print(f"{indent}   启用: {child_enabled}")
                
                if is_input_candidate:
                    print(f"{indent}   🔥 可能的输入控件! 原因: {', '.join(input_indicators)}")
                    
                    # 尝试获取更多输入控件属性
                    try:
                        if hasattr(child, 'CurrentValue'):
                            value = child.CurrentValue
                            print(f"{indent}   当前值: '{value}'")
                    except:
                        pass
                    
                    try:
                        if hasattr(child, 'CurrentIsPassword'):
                            is_password = child.CurrentIsPassword
                            print(f"{indent}   是否密码框: {is_password}")
                    except:
                        pass
                
                # 如果是重要的容器控件，继续递归
                if (child_type in ['PaneControl', 'GroupControl', 'WindowControl'] or 
                    len(child.GetChildren()) > 0):
                    analyze_children_recursive(child, depth + 1, max_depth)
                
                print()
                
            except Exception as e:
                print(f"{indent}├─ 子控件[{i}] 分析失败: {e}")
                
    except Exception as e:
        print(f"{indent}获取子控件失败: {e}")

def find_all_input_candidates(window_title="无人机AI助教"):
    """查找所有可能的输入控件候选者"""
    print(f"\n🔍 查找所有可能的输入控件...")
    print("=" * 60)
    
    hwnd = FindWindow(name=window_title)
    if not hwnd:
        print(f"❌ 未找到窗口: {window_title}")
        return []
    
    try:
        control = uia.ControlFromHandle(hwnd)
        input_candidates = []
        
        def search_input_controls(parent, path="root"):
            try:
                children = parent.GetChildren()
                for i, child in enumerate(children):
                    current_path = f"{path}->child[{i}]"
                    
                    # 检查是否是输入控件候选者
                    child_type = child.ControlTypeName
                    child_class = child.ClassName
                    child_name = child.Name
                    
                    is_candidate = (
                        child_type in ['EditControl', 'DocumentControl', 'TextControl'] or
                        any(keyword in child_class.lower() for keyword in ['edit', 'input', 'text', 'rich']) or
                        (child_name and any(keyword in child_name.lower() for keyword in ['输入', 'input', 'edit']))
                    )
                    
                    if is_candidate:
                        input_candidates.append({
                            'control': child,
                            'path': current_path,
                            'type': child_type,
                            'class': child_class,
                            'name': child_name,
                            'bounds': child.BoundingRectangle,
                            'enabled': child.IsEnabled
                        })
                    
                    # 递归搜索
                    search_input_controls(child, current_path)
                    
            except:
                pass
        
        search_input_controls(control)
        
        print(f"找到 {len(input_candidates)} 个输入控件候选者:")
        for i, candidate in enumerate(input_candidates):
            print(f"\n候选者 {i+1}:")
            print(f"  路径: {candidate['path']}")
            print(f"  类型: {candidate['type']}")
            print(f"  类名: {candidate['class']}")
            print(f"  名称: '{candidate['name']}'")
            print(f"  边界: {candidate['bounds']}")
            print(f"  启用: {candidate['enabled']}")
        
        return input_candidates
        
    except Exception as e:
        print(f"搜索输入控件失败: {e}")
        return []

def test_input_methods(window_title="无人机AI助教"):
    """测试不同的输入方法"""
    print(f"\n🧪 测试输入方法...")
    print("=" * 60)
    
    candidates = find_all_input_candidates(window_title)
    
    if not candidates:
        print("❌ 没有找到输入控件候选者")
        return
    
    print(f"将测试 {len(candidates)} 个候选控件:")
    
    test_message = "测试消息"
    
    for i, candidate in enumerate(candidates):
        print(f"\n测试候选者 {i+1}:")
        try:
            ctrl = candidate['control']
            
            # 测试点击
            print("  - 测试点击...")
            ctrl.Click()
            
            # 测试SetValue
            print("  - 测试SetValue...")
            ctrl.SetValue(test_message)
            
            # 测试SendKeys
            print("  - 测试SendKeys...")
            ctrl.SendKeys(test_message)
            
            print("  ✅ 该候选者可以接受输入")
            
        except Exception as e:
            print(f"  ❌ 测试失败: {e}")

def main():
    """主函数"""
    print("企业微信聊天窗口UIAutomation结构分析工具")
    print("请确保'无人机AI助教'聊天窗口已打开且可见")
    print()
    
    window_title = input("请输入窗口标题 (默认: 无人机AI助教): ").strip() or "无人机AI助教"
    
    # 1. 详细分析窗口结构
    analyze_wecom_window_detailed(window_title)
    
    # 2. 查找所有输入控件候选者
    find_all_input_candidates(window_title)
    
    # 3. 询问是否测试输入方法
    test_input = input("\n是否测试输入方法? (y/n): ").strip().lower()
    if test_input == 'y':
        print("⚠️  注意: 测试可能会在聊天窗口中发送测试消息!")
        confirm = input("确认继续? (y/n): ").strip().lower()
        if confirm == 'y':
            test_input_methods(window_title)

if __name__ == "__main__":
    main()
    input("\n按回车键退出...")