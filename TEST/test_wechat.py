#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
普通微信功能测试脚本
包含基本使用、消息监听、收发消息等功能测试
"""

from wxauto import WeChat
from wxauto.msgs import FriendMessage
import time

def test_basic_usage():
    """测试基本使用功能"""
    print("=== 测试基本使用功能 ===")
    
    # 初始化微信实例
    wx = WeChat()
    print(f"当前登录用户: {wx.nickname}")
    
    # 获取会话列表
    sessions = wx.GetSession()
    print(f"当前会话数量: {len(sessions)}")
    for i, session in enumerate(sessions[:5]):  # 只显示前5个
        print(f"  {i+1}. {session.name}")
    
    # 发送消息示例（需要手动指定联系人）
    # wx.SendMsg("测试消息", who="联系人昵称")
    
    return wx

def test_message_listening():
    """测试消息监听功能"""
    print("\n=== 测试消息监听功能 ===")
    
    wx = WeChat()
    
    # 消息处理函数
    def on_message(msg, chat):
        text = f'[{msg.type} {msg.attr}]{chat.who} - {msg.content}'
        print(text)
        
        # 保存消息到文件
        with open('msgs.txt', 'a', encoding='utf-8') as f:
            f.write(text + '\n')
        
        # 处理图片和视频消息
        if msg.type in ('image', 'video'):
            download_path = msg.download()
            print(f"媒体文件已下载: {download_path}")
        
        # 自动回复好友消息
        if isinstance(msg, FriendMessage):
            time.sleep(1)  # 稍等片刻再回复
            msg.quote('收到')
            print(f"已自动回复: {chat.who}")
    
    # 添加监听示例（需要手动指定要监听的联系人）
    print("要开始监听，请取消注释下面的代码并指定联系人昵称：")
    print('# wx.AddListenChat(nickname="联系人昵称", callback=on_message)')
    
    # wx.AddListenChat(nickname="联系人昵称", callback=on_message)
    
    return wx

def test_message_operations():
    """测试消息操作功能"""
    print("\n=== 测试消息操作功能 ===")
    
    wx = WeChat()
    
    # 切换到指定聊天（需要手动指定联系人）
    print("切换聊天示例：")
    print('# wx.ChatWith("联系人昵称")')
    
    # 获取当前聊天窗口所有消息
    print("获取消息示例：")
    print('# msgs = wx.GetAllMessage()')
    print('# for msg in msgs:')
    print('#     print(f"消息内容: {msg.content}, 消息类型: {msg.type}")')
    
    # 发送文件示例
    print("发送文件示例：")
    print('# wx.SendFiles("文件路径", who="联系人昵称")')
    
    return wx

def test_interactive_mode():
    """交互式测试模式"""
    print("\n=== 交互式测试模式 ===")
    
    wx = WeChat()
    
    while True:
        print("\n请选择测试功能:")
        print("1. 查看当前会话列表")
        print("2. 发送消息")
        print("3. 获取消息")
        print("4. 开始监听")
        print("5. 停止监听")
        print("0. 退出")
        
        choice = input("请输入选择 (0-5): ").strip()
        
        if choice == '0':
            print("退出测试...")
            wx.StopListening()
            break
        elif choice == '1':
            sessions = wx.GetSession()
            print(f"\n当前会话列表 (共{len(sessions)}个):")
            for i, session in enumerate(sessions):
                print(f"  {i+1}. {session.name}")
        elif choice == '2':
            who = input("请输入联系人昵称: ").strip()
            msg = input("请输入消息内容: ").strip()
            if who and msg:
                result = wx.SendMsg(msg, who=who)
                print(f"发送结果: {result}")
        elif choice == '3':
            who = input("请输入联系人昵称 (回车获取当前聊天): ").strip()
            if who:
                wx.ChatWith(who)
            msgs = wx.GetAllMessage()
            print(f"\n获取到 {len(msgs)} 条消息:")
            for msg in msgs[-5:]:  # 显示最后5条
                print(f"  [{msg.type}] {msg.sender}: {msg.content}")
        elif choice == '4':
            who = input("请输入要监听的联系人昵称: ").strip()
            if who:
                def simple_callback(msg, chat):
                    print(f"[监听] {chat.who}: {msg.content}")
                result = wx.AddListenChat(nickname=who, callback=simple_callback)
                print(f"监听结果: {result}")
        elif choice == '5':
            who = input("请输入要停止监听的联系人昵称: ").strip()
            if who:
                result = wx.RemoveListenChat(nickname=who)
                print(f"停止监听结果: {result}")
        else:
            print("无效选择，请重新输入")

def main():
    """主函数"""
    print("微信自动化测试脚本")
    print("请确保微信已登录并打开")
    
    try:
        # 基本功能测试
        test_basic_usage()
        
        # 消息操作测试  
        test_message_operations()
        
        # 消息监听测试
        test_message_listening()
        
        # 交互式测试
        use_interactive = input("\n是否进入交互式测试模式? (y/n): ").strip().lower()
        if use_interactive == 'y':
            test_interactive_mode()
            
    except Exception as e:
        print(f"测试过程中出现错误: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()