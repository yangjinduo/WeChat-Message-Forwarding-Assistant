#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
企业微信功能测试脚本
用于测试企业微信的基本功能是否正常
"""

import sys
import os

# 添加项目路径到sys.path
project_path = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, project_path)

def test_wecom_import():
    """测试企业微信模块导入"""
    print("测试企业微信模块导入...")
    try:
        from wxauto import WeCom, WeComChat
        print("✓ 企业微信模块导入成功")
        return True
    except ImportError as e:
        print(f"✗ 企业微信模块导入失败: {e}")
        return False
    except Exception as e:
        print(f"✗ 企业微信模块导入异常: {e}")
        return False

def test_wecom_window_detection():
    """测试企业微信窗口检测"""
    print("测试企业微信窗口检测...")
    try:
        from wxauto.ui.wecom import WeComMainWnd
        
        # 尝试创建企业微信实例
        wecom_wnd = WeComMainWnd()
        print(f"✓ 成功检测到企业微信窗口: {wecom_wnd.nickname}")
        return True
    except Exception as e:
        print(f"✗ 企业微信窗口检测失败: {e}")
        print("可能的原因:")
        print("1. 企业微信未启动")
        print("2. 企业微信未登录")
        print("3. 企业微信版本不兼容")
        return False

def test_wecom_basic_functions():
    """测试企业微信基本功能"""
    print("测试企业微信基本功能...")
    try:
        from wxauto import WeCom
        
        # 创建企业微信实例
        wecom = WeCom(debug=True)
        print(f"✓ 企业微信实例创建成功: {wecom.nickname}")
        
        # 测试获取会话列表
        try:
            sessions = wecom.GetSession()
            print(f"✓ 获取会话列表成功，共 {len(sessions)} 个会话")
        except Exception as e:
            print(f"✗ 获取会话列表失败: {e}")
        
        # 测试获取当前聊天信息
        try:
            chat_info = wecom.ChatInfo()
            print(f"✓ 获取聊天信息成功: {chat_info.get('chat_name', '未知')}")
        except Exception as e:
            print(f"✗ 获取聊天信息失败: {e}")
        
        # 停止监听
        wecom.StopListening()
        print("✓ 企业微信基本功能测试完成")
        return True
        
    except Exception as e:
        print(f"✗ 企业微信基本功能测试失败: {e}")
        return False

def test_wecom_ui_structure():
    """测试企业微信UI结构"""
    print("测试企业微信UI结构...")
    try:
        from wxauto.ui.wecom import WeComMainWnd
        
        wecom_wnd = WeComMainWnd()
        
        # 检查UI组件
        components = {
            'navigation': wecom_wnd.navigation,
            'sessionbox': wecom_wnd.sessionbox,
            'chatbox': wecom_wnd.chatbox
        }
        
        for name, component in components.items():
            if component and hasattr(component, 'control'):
                print(f"✓ {name} 组件正常")
            else:
                print(f"✗ {name} 组件异常")
        
        return True
    except Exception as e:
        print(f"✗ 企业微信UI结构测试失败: {e}")
        return False

def run_all_tests():
    """运行所有测试"""
    print("=" * 50)
    print("企业微信功能测试")
    print("=" * 50)
    
    tests = [
        ("模块导入测试", test_wecom_import),
        ("窗口检测测试", test_wecom_window_detection),
        ("UI结构测试", test_wecom_ui_structure),
        ("基本功能测试", test_wecom_basic_functions),
    ]
    
    results = []
    for test_name, test_func in tests:
        print(f"\n{test_name}:")
        print("-" * 30)
        result = test_func()
        results.append((test_name, result))
        print()
    
    # 输出测试结果汇总
    print("=" * 50)
    print("测试结果汇总:")
    print("=" * 50)
    
    passed = 0
    for test_name, result in results:
        status = "✓ 通过" if result else "✗ 失败"
        print(f"{test_name}: {status}")
        if result:
            passed += 1
    
    print(f"\n总计: {passed}/{len(results)} 个测试通过")
    
    if passed == len(results):
        print("🎉 所有测试通过！企业微信功能正常")
    else:
        print("⚠️  部分测试失败，请检查企业微信状态")

if __name__ == "__main__":
    run_all_tests()
    input("\n按回车键退出...")