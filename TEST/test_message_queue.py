#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试消息队列系统
验证消息队列的基本功能
"""

import sys
import os
import time
import json

# 添加当前目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 导入消息转发器
try:
    from wechat_message_forwarder_fixed import MessageQueue
    print("✅ 成功导入 MessageQueue 类")
except ImportError as e:
    print(f"❌ 导入失败: {e}")
    sys.exit(1)

class MockMessage:
    """模拟消息对象"""
    def __init__(self, content, sender="测试用户"):
        self.content = content
        self.sender = sender

class MockChat:
    """模拟聊天对象"""
    def __init__(self, name):
        self.name = name

class MockForwarder:
    """模拟转发器对象"""
    def __init__(self):
        self.logs = []
    
    def log_message(self, message):
        timestamp = time.strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}"
        self.logs.append(log_entry)
        print(log_entry)

def test_message_queue():
    """测试消息队列功能"""
    print("=" * 60)
    print("消息队列系统测试开始")
    print("=" * 60)
    
    # 创建模拟转发器
    mock_forwarder = MockForwarder()
    
    # 创建消息队列
    try:
        queue = MessageQueue(mock_forwarder)
        print("✅ 消息队列创建成功")
    except Exception as e:
        print(f"❌ 消息队列创建失败: {e}")
        return False
    
    # 测试添加消息
    print("\n📝 测试添加消息到队列...")
    test_messages = [
        MockMessage("你好，这是第一条测试消息", "用户A"),
        MockMessage("这是第二条测试消息", "用户B"),
        MockMessage("这是第三条测试消息，内容比较长，用来测试消息队列是否能正确处理较长的消息内容", "用户C"),
    ]
    
    test_chat = MockChat("测试群聊")
    
    for i, msg in enumerate(test_messages, 1):
        try:
            message_item = queue.add_message(msg, msg.sender, test_chat)
            print(f"✅ 第{i}条消息添加成功: ID={message_item['id']}")
        except Exception as e:
            print(f"❌ 第{i}条消息添加失败: {e}")
    
    # 测试队列状态
    print("\n📊 测试队列状态获取...")
    try:
        status = queue.get_queue_status()
        print(f"✅ 队列状态: {status}")
    except Exception as e:
        print(f"❌ 获取队列状态失败: {e}")
    
    # 测试获取下一条消息
    print("\n📤 测试获取下一条消息...")
    try:
        next_msg = queue.get_next_message()
        if next_msg:
            print(f"✅ 获取到消息: {next_msg['content'][:30]}...")
            print(f"   发送者: {next_msg['sender']}")
            print(f"   时间戳: {next_msg['timestamp']}")
            
            # 模拟处理完成
            queue.mark_message_completed(next_msg, "这是模拟的AI回复", success=True)
            print("✅ 消息标记为已完成")
        else:
            print("❌ 未获取到消息")
    except Exception as e:
        print(f"❌ 获取消息失败: {e}")
    
    # 测试失败重试
    print("\n🔄 测试消息处理失败重试...")
    try:
        retry_msg = queue.get_next_message()
        if retry_msg:
            # 模拟处理失败
            queue.mark_message_completed(retry_msg, "模拟处理失败", success=False)
            print("✅ 消息标记为失败（应该重新入队）")
        else:
            print("⚠️ 没有更多消息用于重试测试")
    except Exception as e:
        print(f"❌ 重试测试失败: {e}")
    
    # 最终状态检查
    print("\n📊 最终队列状态:")
    try:
        final_status = queue.get_queue_status()
        print(f"   待处理消息: {final_status['pending_count']}")
        print(f"   正在处理: {final_status['is_processing']}")
        print(f"   已完成消息: {final_status['replied_count']}")
    except Exception as e:
        print(f"❌ 获取最终状态失败: {e}")
    
    # 测试文件持久化
    print("\n💾 测试文件持久化...")
    try:
        if os.path.exists('message_queue.json'):
            with open('message_queue.json', 'r', encoding='utf-8') as f:
                queue_data = json.load(f)
            print(f"✅ 队列文件保存成功，包含 {len(queue_data.get('pending_messages', []))} 条待处理消息")
        else:
            print("⚠️ 队列文件未找到")
        
        if os.path.exists('message_history.json'):
            with open('message_history.json', 'r', encoding='utf-8') as f:
                history_data = json.load(f)
            print(f"✅ 历史文件保存成功，包含 {len(history_data)} 条历史记录")
        else:
            print("⚠️ 历史文件未找到")
            
    except Exception as e:
        print(f"❌ 文件持久化测试失败: {e}")
    
    print("\n" + "=" * 60)
    print("消息队列系统测试完成")
    print("=" * 60)
    
    return True

def cleanup_test_files():
    """清理测试文件"""
    test_files = ['message_queue.json', 'message_history.json']
    for file in test_files:
        if os.path.exists(file):
            try:
                os.remove(file)
                print(f"🗑️ 清理测试文件: {file}")
            except Exception as e:
                print(f"⚠️ 清理文件失败: {file}, 错误: {e}")

if __name__ == "__main__":
    try:
        # 运行测试
        success = test_message_queue()
        
        # 询问是否清理测试文件
        if success:
            print("\n是否清理测试生成的文件？")
            print("1. 是（推荐）")
            print("2. 否（保留文件查看内容）")
            
            try:
                choice = input("请选择 (1/2，默认1): ").strip() or "1"
                if choice == "1":
                    cleanup_test_files()
                    print("✅ 测试文件已清理")
                else:
                    print("📁 测试文件已保留")
            except KeyboardInterrupt:
                print("\n\n📁 测试文件已保留")
        
    except KeyboardInterrupt:
        print("\n\n⚠️ 测试被用户中断")
    except Exception as e:
        print(f"\n❌ 测试出现未预期错误: {e}")
        import traceback
        traceback.print_exc()