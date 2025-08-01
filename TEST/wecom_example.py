#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
企业微信使用示例
演示如何使用wxauto操作企业微信
"""

from wxauto import WeCom, WeComChat
from wxauto.msgs import FriendMessage
import time

def basic_usage_example():
    """基本使用示例"""
    print("=== 企业微信基本使用示例 ===")
    
    try:
        # 初始化企业微信实例
        wecom = WeCom(debug=True)
        print(f"成功连接到企业微信，用户：{wecom.nickname}")
        
        # 发送消息
        result = wecom.SendMsg("你好，这是来自wxauto的测试消息", who="测试联系人")
        if result.success:
            print("消息发送成功")
        else:
            print(f"消息发送失败：{result.message}")
        
        # 获取当前聊天窗口消息
        msgs = wecom.GetAllMessage()
        print(f"当前聊天窗口共有 {len(msgs)} 条消息")
        for msg in msgs[-5:]:  # 显示最后5条消息
            print(f"消息内容: {msg.content}, 消息类型: {msg.type}")
        
        # 获取会话列表
        sessions = wecom.GetSession()
        print(f"当前有 {len(sessions)} 个会话")
        for session in sessions[:5]:  # 显示前5个会话
            print(f"会话: {session.nickname}")
            
    except Exception as e:
        print(f"企业微信初始化失败: {e}")
        print("请检查：")
        print("1. 企业微信是否已启动并登录")
        print("2. 企业微信版本是否支持")

def message_listening_example():
    """消息监听示例"""
    print("\n=== 企业微信消息监听示例 ===")
    
    try:
        wecom = WeCom(debug=True)
        
        # 消息处理函数
        def on_message(msg, chat):
            text = f'[企业微信 {msg.type} {msg.attr}]{chat.who} - {msg.content}'
            print(text)
            
            # 保存消息到文件
            with open('wecom_msgs.txt', 'a', encoding='utf-8') as f:
                f.write(text + '\n')

            # 下载图片和视频
            if msg.type in ('image', 'video'):
                download_path = msg.download()
                print(f"文件已下载到: {download_path}")

            # 自动回复
            if isinstance(msg, FriendMessage):
                time.sleep(1)  # 等待一秒
                msg.quote('收到您的消息')

        # 添加监听，监听指定联系人的消息
        chat_result = wecom.AddListenChat(nickname="测试联系人", callback=on_message)
        if isinstance(chat_result, WeComChat):
            print("成功添加消息监听")
            
            # 保持运行，监听消息
            print("开始监听消息，按Ctrl+C停止...")
            wecom.KeepRunning()
        else:
            print(f"添加监听失败: {chat_result.message}")
            
    except KeyboardInterrupt:
        print("\n停止监听")
    except Exception as e:
        print(f"监听失败: {e}")

def file_sending_example():
    """文件发送示例"""
    print("\n=== 企业微信文件发送示例 ===")
    
    try:
        wecom = WeCom()
        
        # 发送文件
        file_path = "test.txt"  # 替换为实际文件路径
        result = wecom.SendFiles(file_path, who="测试联系人")
        
        if result.success:
            print("文件发送成功")
        else:
            print(f"文件发送失败：{result.message}")
            
    except Exception as e:
        print(f"文件发送失败: {e}")

def chat_window_example():
    """聊天窗口操作示例"""
    print("\n=== 企业微信聊天窗口操作示例 ===")
    
    try:
        wecom = WeCom()
        
        # 切换到指定聊天
        wecom.ChatWith("测试联系人")
        
        # 获取聊天信息
        chat_info = wecom.ChatInfo()
        print(f"当前聊天信息: {chat_info}")
        
        # 如果是群聊，获取群成员
        if chat_info.get('chat_type') == 'group':
            members = wecom.GetGroupMembers()
            print(f"群成员: {members}")
        
        # 加载更多历史消息
        wecom.LoadMoreMessage()
        
        # 获取所有子窗口
        sub_windows = wecom.GetAllSubWindow()
        print(f"当前有 {len(sub_windows)} 个子窗口")
        
    except Exception as e:
        print(f"聊天窗口操作失败: {e}")

def main():
    """主函数"""
    print("企业微信 wxauto 使用示例")
    print("请确保企业微信已启动并登录")
    print()
    
    while True:
        print("请选择示例:")
        print("1. 基本使用示例")
        print("2. 消息监听示例")
        print("3. 文件发送示例")
        print("4. 聊天窗口操作示例")
        print("0. 退出")
        
        choice = input("请输入选择 (0-4): ").strip()
        
        if choice == '1':
            basic_usage_example()
        elif choice == '2':
            message_listening_example()
        elif choice == '3':
            file_sending_example()
        elif choice == '4':
            chat_window_example()
        elif choice == '0':
            print("退出程序")
            break
        else:
            print("无效选择，请重新输入")
        
        print("\n" + "="*50 + "\n")

if __name__ == "__main__":
    main()