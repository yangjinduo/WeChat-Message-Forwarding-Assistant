#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试微信聊天对象的类型和方法
验证ChatWith返回的对象是否正确
"""

import sys
import os

# 添加当前目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

def test_wechat_chat_object():
    """测试微信聊天对象"""
    print("=" * 60)
    print("微信聊天对象测试")
    print("=" * 60)
    
    try:
        # 导入wxauto
        from wxauto import WeChat
        
        print("✅ 成功导入wxauto库")
        
        # 尝试初始化微信实例
        try:
            wechat = WeChat()
            print("✅ 成功初始化微信实例")
            print(f"微信实例类型: {type(wechat)}")
        except Exception as e:
            print(f"❌ 初始化微信实例失败: {e}")
            print("请确保微信已经打开并登录")
            return
        
        # 测试ChatWith方法
        test_chat_name = "测试聊天"  # 这个可以是任意名称，只是测试对象类型
        
        try:
            print(f"\n🔍 测试ChatWith方法: {test_chat_name}")
            chat_obj = wechat.ChatWith(test_chat_name)
            
            print(f"✅ ChatWith返回对象类型: {type(chat_obj)}")
            print(f"✅ 对象字符串表示: {str(chat_obj)}")
            
            # 检查是否有SendMsg方法
            if hasattr(chat_obj, 'SendMsg'):
                print("✅ 对象具有SendMsg方法")
                print(f"SendMsg方法类型: {type(chat_obj.SendMsg)}")
            else:
                print("❌ 对象没有SendMsg方法")
            
            # 检查其他常用方法
            methods_to_check = ['SendMsg', 'GetAllMessage', 'SendFile', 'GetChatInfo']
            print(f"\n📋 检查常用方法:")
            for method in methods_to_check:
                if hasattr(chat_obj, method):
                    print(f"  ✅ {method}")
                else:
                    print(f"  ❌ {method}")
                    
        except Exception as e:
            print(f"❌ ChatWith方法测试失败: {e}")
            import traceback
            print(f"详细错误: {traceback.format_exc()}")
        
        print(f"\n📊 总结:")
        print("1. 如果ChatWith返回的对象类型不是字符串，那就没问题")
        print("2. 如果对象具有SendMsg方法，那转发功能应该正常")
        print("3. 如果出现错误，可能是微信版本或wxauto版本问题")
        
    except ImportError as e:
        print(f"❌ 导入失败: {e}")
        print("请确保wxauto库已正确安装")
    except Exception as e:
        print(f"❌ 测试过程中出现错误: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_wechat_chat_object()