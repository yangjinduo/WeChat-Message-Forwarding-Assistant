#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试消息队列UI界面
验证队列状态显示和操作功能
"""

import sys
import os
import time
import json

# 添加当前目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

def create_test_queue_data():
    """创建测试队列数据"""
    import time
    from datetime import datetime
    
    # 创建测试消息队列数据
    test_pending_messages = [
        {
            'id': f"{int(time.time() * 1000)}_1",
            'content': "这是第一条待处理的测试消息",
            'sender': "测试用户1",
            'chat_name': "测试群聊1",
            'timestamp': time.time(),
            'status': 'pending',
            'retry_count': 0,
            'created_time': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        },
        {
            'id': f"{int(time.time() * 1000)}_2",
            'content': "这是第二条待处理的测试消息，内容比较长，用来测试界面显示效果",
            'sender': "测试用户2",
            'chat_name': "测试群聊2",
            'timestamp': time.time(),
            'status': 'pending',
            'retry_count': 1,
            'created_time': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
    ]
    
    test_replied_messages = [
        {
            'id': f"{int(time.time() * 1000)}_3",
            'content': "这是一条已完成的测试消息",
            'sender': "测试用户3",
            'chat_name': "测试群聊3",
            'timestamp': time.time() - 300,  # 5分钟前
            'status': 'replied',
            'retry_count': 0,
            'created_time': (datetime.now()).strftime("%Y-%m-%d %H:%M:%S"),
            'ai_reply': "这是AI的回复内容",
            'completed_time': time.time() - 250
        },
        {
            'id': f"{int(time.time() * 1000)}_4",
            'content': "这是一条失败的测试消息",
            'sender': "测试用户4",
            'chat_name': "测试群聊4",
            'timestamp': time.time() - 600,  # 10分钟前
            'status': 'failed',
            'retry_count': 3,
            'created_time': (datetime.now()).strftime("%Y-%m-%d %H:%M:%S"),
            'last_error': "测试错误：AI回复超时"
        }
    ]
    
    # 创建队列数据
    queue_data = {
        'pending_messages': test_pending_messages,
        'processing_message': None,
        'is_processing': False,
        'last_save_time': time.time(),
        'version': '1.0'
    }
    
    # 保存到文件
    with open('message_queue.json', 'w', encoding='utf-8') as f:
        json.dump(queue_data, f, ensure_ascii=False, indent=2)
    
    with open('message_history.json', 'w', encoding='utf-8') as f:
        json.dump(test_replied_messages, f, ensure_ascii=False, indent=2)
    
    print("已创建测试队列数据文件")
    print(f"   - 待处理消息: {len(test_pending_messages)} 条")
    print(f"   - 历史消息: {len(test_replied_messages)} 条")

def main():
    """主函数"""
    print("=" * 60)
    print("消息队列UI测试")
    print("=" * 60)
    
    try:
        # 创建测试数据
        create_test_queue_data()
        
        print("\n现在可以启动主程序来测试UI界面:")
        print("1. 运行 wechat_message_forwarder_fixed.py")
        print("2. 查看消息队列状态区域的显示效果")
        print("3. 测试各种操作按钮")
        print("4. 程序启动时应该会显示未处理消息警告")
        
        print("\n测试要点:")
        print("- 队列状态表格应该显示消息列表")
        print("- 不同状态的消息应该有不同的背景色")
        print("- 应该显示重启警告对话框")
        print("- 操作按钮应该能正常工作")
        
        print("\n测试完成后，可以删除生成的测试文件:")
        print("- message_queue.json")
        print("- message_history.json")
        
    except Exception as e:
        print(f"创建测试数据失败: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()