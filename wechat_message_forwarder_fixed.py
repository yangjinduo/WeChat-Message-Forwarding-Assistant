#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
微信消息转发助手
实现普通微信和企业微信之间的消息转发功能
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import json
import threading
import time
import os
import traceback
import hashlib
from datetime import datetime
from wxauto import WeChat, WeCom
from wxauto.msgs import FriendMessage
from PIL import Image, ImageGrab
import ctypes
from ctypes import windll
import win32gui
import win32con
import win32api
import win32ui

class MessageQueue:
    """消息队列类 - 负责管理消息的存储、处理状态和持久化"""
    
    def __init__(self, forwarder):
        self.forwarder = forwarder  # 引用主应用
        
        # 内存中的消息队列
        self.pending_messages = []      # 待处理消息
        self.processing_message = None   # 当前处理中的消息
        self.replied_messages = []      # 已回复消息（最近100条）
        self.is_processing = False      # 处理状态锁
        
        # 文件路径
        self.queue_file = "message_queue.json"
        self.history_file = "message_history.json"  # 旧的历史文件，用于兼容
        
        # 规则对应的历史文件字典 {rule_id: history_file_path}
        self.rule_history_files = {}
        # 规则对应的历史消息 {rule_id: [messages]}
        self.rule_replied_messages = {}
        
        # 启动时加载历史数据
        self.load_from_file()
    
    def generate_rule_history_filename(self, rule):
        """生成规则对应的历史文件名"""
        try:
            source_type = rule['source']['type']
            source_contact = rule['source']['contact']
            target_type = rule['target']['type']
            target_contact = rule['target']['contact']
            
            # 格式: message_history(源类型源联系人<>目标类型目标联系人).json
            filename = f"message_history({source_type}{source_contact}<>{target_type}{target_contact}).json"
            
            # 处理文件名中的非法字符
            import re
            filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
            
            return filename
        except Exception as e:
            # 如果生成失败，使用规则ID作为后备
            return f"message_history_rule_{rule.get('id', 'unknown')}.json"
    
    def get_rule_history_file(self, rule):
        """获取规则对应的历史文件路径"""
        rule_id = rule['id']
        if rule_id not in self.rule_history_files:
            self.rule_history_files[rule_id] = self.generate_rule_history_filename(rule)
        return self.rule_history_files[rule_id]
    
    def get_rule_replied_messages(self, rule_id):
        """获取规则对应的已回复消息列表"""
        if rule_id not in self.rule_replied_messages:
            self.rule_replied_messages[rule_id] = []
        return self.rule_replied_messages[rule_id]
    
    def add_message(self, msg, sender, chat, source_type):
        """按多规则匹配添加消息到队列"""
        # 提取真实的聊天名称
        if hasattr(chat, 'name'):
            chat_name = chat.name
        elif hasattr(chat, 'nickname'):
            chat_name = chat.nickname
        else:
            # 从字符串中提取聊天名称，处理 <wxauto - Chat object("矿泉水会飞")> 格式
            chat_str = str(chat)
            if '"' in chat_str:
                # 提取引号中的内容
                start = chat_str.find('"') + 1
                end = chat_str.rfind('"')
                if start > 0 and end > start:
                    chat_name = chat_str[start:end]
                else:
                    chat_name = chat_str
            else:
                chat_name = chat_str
        
        # 查找匹配的规则
        matching_rules = self.forwarder.find_matching_rules(msg, chat_name, source_type)
        
        if not matching_rules:
            # 没有匹配的规则，不添加到队列
            self.forwarder.log_message(f"⚠️ 消息未匹配任何规则，跳过: {msg.content[:30]}...")
            return None
        
        # 为每个匹配的规则创建一个消息项
        added_messages = []
        for rule in matching_rules:
            message_item = {
                'id': f"{int(time.time() * 1000)}_{hash(msg.content)}_{rule['id']}",
                'content': msg.content,
                'sender': sender,
                'chat_name': chat_name,
                'source_type': source_type,
                'matched_rule': rule,  # 存储匹配的规则
                'timestamp': time.time(),
                'status': 'pending',
                'created_time': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            
            self.pending_messages.append(message_item)
            added_messages.append(message_item)
            self.forwarder.log_message(f"📝 消息入队[{rule['name']}]: {msg.content[:30]}...", rule['id'])
        
        self.save_to_file()  # 立即保存到文件
        self.forwarder.log_message(f"✅ 共添加 {len(added_messages)} 条消息到队列 (总长度: {len(self.pending_messages)})")
        
        return added_messages[0] if added_messages else None
    
    def get_next_message(self):
        """获取下一条待处理消息"""
        if self.pending_messages and not self.is_processing:
            return self.pending_messages.pop(0)
        return None
    
    def save_to_file(self):
        """保存当前状态到文件"""
        try:
            queue_data = {
                'pending_messages': self.pending_messages,
                'processing_message': self.processing_message,
                'is_processing': self.is_processing,
                'last_save_time': time.time(),
                'version': '1.0'
            }
            
            # 保存队列状态
            with open(self.queue_file, 'w', encoding='utf-8') as f:
                json.dump(queue_data, f, ensure_ascii=False, indent=2)
            
            # 保存规则对应的历史记录（每个规则保留最近100条）
            for rule_id, messages in self.rule_replied_messages.items():
                if len(messages) > 0:
                    # 获取规则信息用于生成文件名
                    rule = self.find_rule_by_id(rule_id)
                    if rule:
                        history_file = self.get_rule_history_file(rule)
                        with open(history_file, 'w', encoding='utf-8') as f:
                            json.dump(messages[-100:], f, ensure_ascii=False, indent=2)
            
            # 为了兼容性，仍然保存一个总的历史文件（将所有规则的消息合并）
            all_replied_messages = []
            for messages in self.rule_replied_messages.values():
                all_replied_messages.extend(messages)
            
            if len(all_replied_messages) > 0:
                # 按时间排序
                all_replied_messages.sort(key=lambda x: x.get('completed_time', 0))
                with open(self.history_file, 'w', encoding='utf-8') as f:
                    json.dump(all_replied_messages[-100:], f, ensure_ascii=False, indent=2)
                    
        except Exception as e:
            self.forwarder.log_message(f"💾 保存消息队列失败: {e}")
    
    def find_rule_by_id(self, rule_id):
        """根据ID查找规则"""
        try:
            for rule in self.forwarder.forwarding_rules:
                if rule['id'] == rule_id:
                    return rule
            return None
        except Exception:
            return None
    
    def load_rule_history_files(self):
        """加载所有规则对应的历史文件"""
        try:
            # 遍历所有存在的message_history文件
            import glob
            history_files = glob.glob("message_history*.json")
            
            for file_path in history_files:
                # 跳过旧的全局历史文件
                if file_path == self.history_file:
                    continue
                
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        messages = json.load(f)
                    
                    # 尝试从消息中提取规则ID
                    if messages and len(messages) > 0:
                        for msg in messages:
                            if 'matched_rule' in msg and msg['matched_rule']:
                                rule_id = msg['matched_rule']['id']
                                if rule_id not in self.rule_replied_messages:
                                    self.rule_replied_messages[rule_id] = []
                                # 避免重复添加
                                if msg not in self.rule_replied_messages[rule_id]:
                                    self.rule_replied_messages[rule_id].append(msg)
                                break  # 找到规则ID后停止遍历
                        
                        # 如果找到了规则ID，更新文件路径映射
                        if messages and 'matched_rule' in messages[0]:
                            rule_id = messages[0]['matched_rule']['id']
                            self.rule_history_files[rule_id] = file_path
                            # 确保所有消息都在正确的规则下
                            if rule_id not in self.rule_replied_messages:
                                self.rule_replied_messages[rule_id] = []
                            self.rule_replied_messages[rule_id] = messages
                                
                except Exception as e:
                    if self.forwarder:
                        self.forwarder.log_message(f"⚠️ 加载历史文件{file_path}失败: {e}")
                        
        except Exception as e:
            if self.forwarder:
                self.forwarder.log_message(f"⚠️ 加载规则历史文件失败: {e}")
    
    def load_from_file(self):
        """从文件加载历史状态"""
        try:
            # 加载队列状态
            if os.path.exists(self.queue_file):
                with open(self.queue_file, 'r', encoding='utf-8') as f:
                    queue_data = json.load(f)
                    self.pending_messages = queue_data.get('pending_messages', [])
                    self.processing_message = queue_data.get('processing_message')
                    # 重启后重置处理状态
                    self.is_processing = False
            
            # 加载规则对应的历史记录
            self.load_rule_history_files()
            
            # 为了兼容性，仍然加载旧的全局历史文件
            all_replied_messages = []
            if os.path.exists(self.history_file):
                with open(self.history_file, 'r', encoding='utf-8') as f:
                    legacy_messages = json.load(f)
                    all_replied_messages.extend(legacy_messages)
            
            # 合并所有规则的历史消息用于显示
            for messages in self.rule_replied_messages.values():
                all_replied_messages.extend(messages)
            
            # 按时间排序并去重
            all_replied_messages.sort(key=lambda x: x.get('completed_time', 0))
            seen_ids = set()
            self.replied_messages = []
            for msg in all_replied_messages:
                msg_id = msg.get('id')
                if msg_id and msg_id not in seen_ids:
                    seen_ids.add(msg_id)
                    self.replied_messages.append(msg)
            
            if self.forwarder:
                total_rule_messages = sum(len(messages) for messages in self.rule_replied_messages.values())
                self.forwarder.log_message(f"📂 加载消息队列: 待处理{len(self.pending_messages)}条, 历史{len(self.replied_messages)}条 (各规则共{total_rule_messages}条)")
                
                # 检查是否有未回复或失败的消息
                failed_messages = []
                for messages in self.rule_replied_messages.values():
                    failed_messages.extend([msg for msg in messages if msg.get('status') == 'failed'])
                
                if self.pending_messages or failed_messages:
                    # 延迟调用警告，等待message_queue属性设置完成
                    self.forwarder.root.after(100, self.forwarder.show_restart_warning)
            
        except Exception as e:
            if self.forwarder:
                self.forwarder.log_message(f"📂 加载消息队列失败: {e}")
    
    def get_queue_status(self):
        """获取队列状态信息"""
        return {
            'pending_count': len(self.pending_messages),
            'processing': self.processing_message is not None,
            'replied_count': len(self.replied_messages),
            'is_processing': self.is_processing
        }
    
    def mark_message_completed(self, message_item, ai_reply, success=True):
        """标记消息处理完成"""
        # 提取规则ID用于日志
        rule_id = None
        if 'matched_rule' in message_item and message_item['matched_rule']:
            rule_id = message_item['matched_rule']['id']
        
        if success:
            message_item['status'] = 'replied'
            message_item['ai_reply'] = ai_reply
            message_item['completed_time'] = time.time()
            
            # 添加到规则对应的历史消息中
            if rule_id:
                rule_messages = self.get_rule_replied_messages(rule_id)
                rule_messages.append(message_item)
            
            # 为了兼容性，仍然保持全局列表
            self.replied_messages.append(message_item)
            self.forwarder.log_message(f"✅ 消息处理完成: {message_item['content'][:30]}...", rule_id)
        else:
            message_item['status'] = 'failed'
            message_item['last_error'] = ai_reply  # 这里ai_reply实际是错误信息
            message_item['failed_time'] = time.time()
            
            # 添加到规则对应的历史消息中
            if rule_id:
                rule_messages = self.get_rule_replied_messages(rule_id)
                rule_messages.append(message_item)
            
            # 为了兼容性，仍然保持全局列表
            self.replied_messages.append(message_item)
            self.forwarder.log_message(f"❌ 消息处理失败: {ai_reply}", rule_id)
        
        self.processing_message = None
        self.is_processing = False
        self.save_to_file()
    
    def trim_queue(self, max_size):
        """修剪队列到指定大小"""
        try:
            total_messages = len(self.pending_messages) + len(self.replied_messages)
            if total_messages > max_size:
                # 首先从已完成的消息中删除最早的
                excess = total_messages - max_size
                if len(self.replied_messages) > excess:
                    self.replied_messages = self.replied_messages[excess:]
                    self.forwarder.log_message(f"🗑️ 已清理 {excess} 条历史消息，保持队列在 {max_size} 条以内")
                else:
                    # 如果历史消息不够删，需要从待处理中删除
                    remaining_excess = excess - len(self.replied_messages)
                    self.replied_messages.clear()
                    if remaining_excess < len(self.pending_messages):
                        self.pending_messages = self.pending_messages[remaining_excess:]
                        self.forwarder.log_message(f"⚠️ 队列满，已删除 {excess} 条消息（包括待处理消息）")
                
                self.save_to_file()
        except Exception as e:
            self.forwarder.log_message(f"❌ 修剪队列失败: {e}")

class WeChatMessageForwarder:
    def __init__(self):
        # 创建主窗口
        self.root = tk.Tk()
        
        # 设置为唯一的主窗口
        self.root.withdraw()  # 初始时隐藏窗口
        
        # 设置窗口基本属性
        self.root.title("微信消息转发助手 V0.1")
        self.root.resizable(True, True)
        
        # 获取屏幕尺寸并设置窗口尺寸
        try:
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            
            # 设置窗口尺寸为屏幕的80%，但不超过最大尺寸
            window_width = min(1600, int(screen_width * 0.8))
            window_height = min(1200, int(screen_height * 0.8))
            
            # 计算居中位置
            x = (screen_width - window_width) // 2
            y = (screen_height - window_height) // 2
            
            self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        except Exception as e:
            # 如果获取屏幕尺寸失败，使用默认尺寸
            self.root.geometry("1400x1000+100+50")
        
        # 设置窗口最小尺寸
        self.root.minsize(1000, 700)
        
        # DPI设置（放在窗口创建之后）
        try:
            from ctypes import windll
            # 设置进程DPI感知
            windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            try:
                windll.user32.SetProcessDPIAware()
            except Exception:
                pass  # 如果都失败就忽略
        
        # 确保窗口显示在最前面
        self.root.lift()
        self.root.attributes('-topmost', True)
        self.root.after(1000, lambda: self.root.attributes('-topmost', False))
        
        # 延迟初始化GUI，确保窗口完全创建后再添加内容
        self.root.after(50, self.delayed_init)
        
        # 转发状态
        self.is_forwarding = False
        self.forward_thread = None
        
        # 微信实例
        self.wechat = None
        self.wecom = None
        
        # 记录最近的AI回复，避免循环转发
        self.recent_ai_replies = []
        self.max_recent_replies = 5  # 最多记录5条最近的AI回复
        
        # 其他设置默认值
        self.log_retention_days = 10
        self.queue_max_size = 600
        
        # 当前微信昵称
        self.current_wechat_nickname = None
        
        # 多规则转发系统
        self.forwarding_rules = []
        self.selected_rule_index = 0
        self.init_default_rule()  # 初始化默认规则
        
        # 消息队列系统（延迟初始化）
        self.message_queue = None
        self.message_processor_thread = None
        
        # 转发配置
        self.config = {
            'source': {
                'type': 'wechat',  # wechat 或 wecom
                'contact': '',
                'filter_type': 'all',  # all, mention_me, mention_range
                'range_start': '@本人',
                'range_end': '@本人并说结束'
            },
            'target': {
                'type': 'wecom',  # wechat 或 wecom
                'contact': ''
            }
        }
        
        # GUI创建和配置加载将在delayed_init中执行
    
    def delayed_init(self):
        """延迟初始化GUI和配置"""
        try:
            self.create_gui()
            
            # 创建GUI后初始化消息队列系统
            self.message_queue = MessageQueue(self)
            
            self.load_config()
            
            # 尝试获取当前微信昵称
            self.root.after(1000, self.refresh_wechat_nickname)  # 延迟1秒后获取昵称
            
            # 确保窗口正确显示
            self.root.update_idletasks()
            self.root.deiconify()
        except Exception as e:
            print(f"GUI初始化失败: {e}")
            import traceback
            traceback.print_exc()
    
    def create_gui(self):
        """创建GUI界面"""
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)  # 左侧区域
        main_frame.columnconfigure(1, weight=1)  # 右侧区域
        main_frame.rowconfigure(1, weight=1)     # 主要内容区域

        # 当前微信昵称显示区域
        nickname_frame = ttk.Frame(main_frame)
        nickname_frame.grid(row=0, column=0, columnspan=2, pady=(0, 20), sticky=(tk.W, tk.E))
        nickname_frame.columnconfigure(1, weight=1)
        
        ttk.Label(nickname_frame, text="当前微信昵称:", font=("Microsoft YaHei", 12)).grid(row=1, column=0, sticky=tk.W, padx=(0, 10))
        self.current_nickname_var = tk.StringVar(value="未获取")
        self.current_nickname_label = ttk.Label(nickname_frame, textvariable=self.current_nickname_var, 
                                              font=("Microsoft YaHei", 12, "bold"), foreground="blue")
        self.current_nickname_label.grid(row=1, column=1, sticky=tk.W)
        
        ttk.Button(nickname_frame, text="刷新昵称", command=self.refresh_wechat_nickname).grid(row=1, column=2, padx=(10, 0))
        
        # 创建左右分栏框架
        left_frame = ttk.Frame(main_frame)
        left_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5))
        left_frame.columnconfigure(0, weight=1)
        
        right_frame = ttk.Frame(main_frame)
        right_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(5, 0))
        right_frame.columnconfigure(0, weight=1)
        right_frame.rowconfigure(0, weight=1)
        
        # 左侧内容
        # 多规则管理区域
        self.create_rules_management_section(left_frame, row=0)
        
        # 其他设置区域
        self.create_advanced_settings(left_frame, row=1)
        
        # 控制按钮区域
        self.create_control_section(left_frame, row=2)
        
        # 日志显示区域
        self.create_log_section(left_frame, row=3)
        
        # 底部说明区域
        self.create_footer_section(main_frame, row=2)
        
        # 右侧内容 - 消息队列状态区域
        self.create_queue_status_section(right_frame, row=0)
    
    def init_default_rule(self):
        """初始化默认转发规则"""
        default_rule = {
            'id': 'rule_1',
            'name': '规列1',
            'enabled': True,
            'source': {
                'type': 'wechat',
                'contact': '',
                'filter_type': 'all',
                'range_start': '',
                'range_end': ''
            },
            'target': {
                'type': 'wecom',
                'contact': ''
            }
        }
        self.forwarding_rules = [default_rule]
    
    def create_rules_management_section(self, parent, row):
        """创建多规则管理区域"""
        rules_frame = ttk.LabelFrame(parent, text="转发规则管理", padding="10")
        rules_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        rules_frame.columnconfigure(1, weight=1)
        rules_frame.rowconfigure(1, weight=1)
        
        # 规则列表和操作按钮
        list_frame = ttk.Frame(rules_frame)
        list_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        # 规则列表
        self.rules_tree = ttk.Treeview(list_frame, columns=('sequence', 'enabled', 'source', 'target', 'filter'), show='headings', height=6)
        self.rules_tree.heading('sequence', text='序号')
        self.rules_tree.heading('enabled', text='状态')
        self.rules_tree.heading('source', text='消息来源')
        self.rules_tree.heading('target', text='转发目标')
        self.rules_tree.heading('filter', text='过滤条件')
        
        self.rules_tree.column('sequence', width=50)
        self.rules_tree.column('enabled', width=60)
        self.rules_tree.column('source', width=150)
        self.rules_tree.column('target', width=150)
        self.rules_tree.column('filter', width=120)
        
        self.rules_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 滚动条
        rules_scroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.rules_tree.yview)
        rules_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.rules_tree.configure(yscrollcommand=rules_scroll.set)
        
        # 操作按钮
        button_frame = ttk.Frame(rules_frame)
        button_frame.grid(row=1, column=0, columnspan=2, pady=(10, 0))
        
        ttk.Button(button_frame, text="添加规则", command=self.add_rule).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="编辑规则", command=self.edit_rule).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="删除规则", command=self.delete_rule).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="启用/禁用", command=self.toggle_rule).pack(side=tk.LEFT, padx=(0, 10))
        
        # 规则详情编辑区域
        detail_frame = ttk.LabelFrame(rules_frame, text="规则详情", padding="10")
        detail_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        detail_frame.columnconfigure(1, weight=1)
        detail_frame.columnconfigure(3, weight=1)
        
        # 规则名称
        ttk.Label(detail_frame, text="规则名称:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.rule_name_var = tk.StringVar()
        ttk.Entry(detail_frame, textvariable=self.rule_name_var, width=20).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 20))
        
        # 启用状态
        self.rule_enabled_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(detail_frame, text="启用该规则", variable=self.rule_enabled_var).grid(row=0, column=2, columnspan=2, sticky=tk.W)
        
        # 消息来源设置
        source_frame = ttk.LabelFrame(detail_frame, text="消息来源", padding="5")
        source_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 5), padx=(0, 10))
        source_frame.columnconfigure(1, weight=1)
        
        ttk.Label(source_frame, text="类型:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.rule_source_type_var = tk.StringVar(value="wechat")
        source_type_combo = ttk.Combobox(source_frame, textvariable=self.rule_source_type_var, values=["wechat", "wecom"], state="readonly", width=10)
        source_type_combo.grid(row=0, column=1, sticky=tk.W, padx=(0, 10))
        source_type_combo.bind('<<ComboboxSelected>>', self.on_rule_source_type_change)
        
        ttk.Label(source_frame, text="联系人:").grid(row=0, column=2, sticky=tk.W, padx=(10, 10))
        self.rule_source_contact_var = tk.StringVar()
        self.rule_source_contact_combo = ttk.Combobox(source_frame, textvariable=self.rule_source_contact_var, width=20)
        self.rule_source_contact_combo.grid(row=0, column=3, sticky=(tk.W, tk.E))
        
        # 过滤条件
        ttk.Label(source_frame, text="过滤:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(5, 0))
        self.rule_filter_type_var = tk.StringVar(value="all")
        filter_combo = ttk.Combobox(source_frame, textvariable=self.rule_filter_type_var, 
                                  values=["all", "range", "at_me"], state="readonly", width=10)
        filter_combo.grid(row=1, column=1, sticky=tk.W, padx=(0, 10), pady=(5, 0))
        filter_combo.bind('<<ComboboxSelected>>', self.on_rule_filter_change)
        
        # 过滤范围
        self.rule_range_frame = ttk.Frame(source_frame)
        self.rule_range_frame.grid(row=1, column=2, columnspan=2, sticky=(tk.W, tk.E), pady=(5, 0), padx=(10, 0))
        
        ttk.Label(self.rule_range_frame, text="从:").pack(side=tk.LEFT)
        self.rule_range_start_var = tk.StringVar()
        ttk.Entry(self.rule_range_frame, textvariable=self.rule_range_start_var, width=8).pack(side=tk.LEFT, padx=(5, 5))
        ttk.Label(self.rule_range_frame, text="到:").pack(side=tk.LEFT)
        self.rule_range_end_var = tk.StringVar()
        ttk.Entry(self.rule_range_frame, textvariable=self.rule_range_end_var, width=8).pack(side=tk.LEFT, padx=(5, 0))
        
        # 转发目标设置
        target_frame = ttk.LabelFrame(detail_frame, text="转发目标", padding="5")
        target_frame.grid(row=1, column=2, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 5), padx=(10, 0))
        target_frame.columnconfigure(1, weight=1)
        
        ttk.Label(target_frame, text="类型:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.rule_target_type_var = tk.StringVar(value="wecom")
        target_type_combo = ttk.Combobox(target_frame, textvariable=self.rule_target_type_var, values=["wechat", "wecom"], state="readonly", width=10)
        target_type_combo.grid(row=0, column=1, sticky=tk.W, padx=(0, 10))
        target_type_combo.bind('<<ComboboxSelected>>', self.on_rule_target_type_change)
        
        ttk.Label(target_frame, text="联系人:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(5, 0))
        self.rule_target_contact_var = tk.StringVar()
        self.rule_target_contact_combo = ttk.Combobox(target_frame, textvariable=self.rule_target_contact_var, width=20)
        self.rule_target_contact_combo.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=(5, 0))
        
        # 保存按钮
        ttk.Button(detail_frame, text="保存规则", command=self.save_current_rule).grid(row=2, column=0, columnspan=4, pady=(10, 0))
        
        # 初始化显示
        self.refresh_rules_display()
        self.on_rule_filter_change()
        
        # 绑定选中事件
        self.rules_tree.bind('<<TreeviewSelect>>', self.on_rule_select)
    
    def create_source_section(self, parent, row):
        """创建源消息设置区域"""
        # 源消息框架
        source_frame = ttk.LabelFrame(parent, text="源消息来源", padding="10")
        source_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        source_frame.columnconfigure(1, weight=1)
        
        # 微信类型选择
        ttk.Label(source_frame, text="微信类型:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.source_type_var = tk.StringVar(value="wechat")
        source_type_frame = ttk.Frame(source_frame)
        source_type_frame.grid(row=0, column=1, sticky=(tk.W, tk.E))
        
        ttk.Radiobutton(source_type_frame, text="普通微信", variable=self.source_type_var, 
                       value="wechat", command=self.on_source_type_change).pack(side=tk.LEFT, padx=(0, 20))
        ttk.Radiobutton(source_type_frame, text="企业微信", variable=self.source_type_var, 
                       value="wecom", command=self.on_source_type_change).pack(side=tk.LEFT)
        
        # 联系人选择
        ttk.Label(source_frame, text="联系人/群组:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))
        contact_frame = ttk.Frame(source_frame)
        contact_frame.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=(10, 0))
        contact_frame.columnconfigure(0, weight=1)
        
        self.source_contact_var = tk.StringVar()
        self.source_contact_combo = ttk.Combobox(contact_frame, textvariable=self.source_contact_var, 
                                               state="readonly")
        self.source_contact_combo.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Button(contact_frame, text="刷新列表", 
                  command=self.refresh_source_contacts).grid(row=0, column=1)
    
    def create_target_section(self, parent, row):
        """创建转发目标设置区域"""
        # 转发目标框架
        target_frame = ttk.LabelFrame(parent, text="转发目标", padding="10")
        target_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        target_frame.columnconfigure(1, weight=1)
        
        # 微信类型选择
        ttk.Label(target_frame, text="微信类型:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.target_type_var = tk.StringVar(value="wecom")
        target_type_frame = ttk.Frame(target_frame)
        target_type_frame.grid(row=0, column=1, sticky=(tk.W, tk.E))
        
        ttk.Radiobutton(target_type_frame, text="普通微信", variable=self.target_type_var, 
                       value="wechat", command=self.on_target_type_change).pack(side=tk.LEFT, padx=(0, 20))
        ttk.Radiobutton(target_type_frame, text="企业微信", variable=self.target_type_var, 
                       value="wecom", command=self.on_target_type_change).pack(side=tk.LEFT)
        
        # 联系人选择
        ttk.Label(target_frame, text="联系人/群组:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))
        target_contact_frame = ttk.Frame(target_frame)
        target_contact_frame.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=(10, 0))
        target_contact_frame.columnconfigure(0, weight=1)
        
        self.target_contact_var = tk.StringVar()
        self.target_contact_combo = ttk.Combobox(target_contact_frame, textvariable=self.target_contact_var, 
                                               state="readonly")
        self.target_contact_combo.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Button(target_contact_frame, text="刷新列表", 
                  command=self.refresh_target_contacts).grid(row=0, column=1)
    
    def create_filter_section(self, parent, row):
        """创建转发条件设置区域"""
        # 转发条件框架
        filter_frame = ttk.LabelFrame(parent, text="转发条件", padding="10")
        filter_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        filter_frame.columnconfigure(1, weight=1)
        
        # 过滤类型选择
        self.filter_type_var = tk.StringVar(value="all")
        
        ttk.Radiobutton(filter_frame, text="转发所有消息", variable=self.filter_type_var, 
                       value="all", command=self.on_filter_type_change).grid(row=0, column=0, columnspan=2, sticky=tk.W)
        
        # @某人的消息设置
        mention_frame = ttk.Frame(filter_frame)
        mention_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(5, 0))
        mention_frame.columnconfigure(2, weight=1)
        
        ttk.Radiobutton(mention_frame, text="仅@某人的消息:", variable=self.filter_type_var, 
                       value="mention_me", command=self.on_filter_type_change).grid(row=0, column=0, sticky=tk.W)
        
        self.mention_name_var = tk.StringVar()
        self.mention_name_entry = ttk.Entry(mention_frame, textvariable=self.mention_name_var, width=15)
        self.mention_name_entry.grid(row=0, column=1, sticky=tk.W, padx=(10, 5))
        
        ttk.Button(mention_frame, text="使用当前昵称", command=self.use_current_nickname).grid(row=0, column=2, sticky=tk.W)
        
        ttk.Radiobutton(filter_frame, text="指定范围的消息", variable=self.filter_type_var, 
                       value="mention_range", command=self.on_filter_type_change).grid(row=2, column=0, columnspan=2, sticky=tk.W)
        
        # 范围设置
        range_frame = ttk.Frame(filter_frame)
        range_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), padx=(20, 0), pady=(5, 0))
        range_frame.columnconfigure(1, weight=1)
        range_frame.columnconfigure(3, weight=1)
        
        ttk.Label(range_frame, text="从:").grid(row=0, column=0, sticky=tk.W)
        self.range_start_var = tk.StringVar(value="@本人")
        ttk.Entry(range_frame, textvariable=self.range_start_var, width=15).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(5, 10))
        
        ttk.Label(range_frame, text="到:").grid(row=0, column=2, sticky=tk.W)
        self.range_end_var = tk.StringVar(value="@本人并说结束")
        ttk.Entry(range_frame, textvariable=self.range_end_var, width=15).grid(row=0, column=3, sticky=(tk.W, tk.E), padx=(5, 0))
        
        self.range_frame = range_frame
        self.on_filter_type_change()  # 初始化状态
        
        # 信息对比延迟设置
        delay_frame = ttk.LabelFrame(filter_frame, text="回复检测设置", padding="10")
        delay_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        delay_frame.columnconfigure(1, weight=1)
        
        ttk.Label(delay_frame, text="首次截图延迟(秒):").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.delay_var = tk.StringVar(value="2")
        delay_entry = ttk.Entry(delay_frame, textvariable=self.delay_var, width=10)
        delay_entry.grid(row=0, column=1, sticky=tk.W)
        
        # 绑定自动保存事件
        def save_delay_setting(*args):
            try:
                delay_value = int(self.delay_var.get())
                if delay_value < 0:
                    delay_value = 2
                    self.delay_var.set("2")
                
                # 自动保存到配置文件
                try:
                    with open('forwarder_config.json', 'r', encoding='utf-8') as f:
                        config = json.load(f)
                except:
                    config = {}
                
                config['detection_delay'] = delay_value
                
                with open('forwarder_config.json', 'w', encoding='utf-8') as f:
                    json.dump(config, f, ensure_ascii=False, indent=2)
                    
            except ValueError:
                # 如果输入无效，恢复默认值
                self.delay_var.set("2")
        
        self.delay_var.trace('w', save_delay_setting)
        
        ttk.Label(delay_frame, text="说明：发送消息到企业微信后等待多少秒开始截图检测", 
                 foreground="gray", font=("Arial", 8)).grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(5, 0))
    
    def create_advanced_settings(self, parent, row):
        """创建其他设置区域"""
        settings_frame = ttk.LabelFrame(parent, text="其他设置", padding="10")
        settings_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # AI回复检测延迟设置
        ttk.Label(settings_frame, text="AI回复检测延迟:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.delay_var = tk.StringVar(value="2")
        delay_entry = ttk.Entry(settings_frame, textvariable=self.delay_var, width=8)
        delay_entry.grid(row=0, column=1, sticky=tk.W)
        
        def save_delay_setting(*args):
            try:
                delay = int(self.delay_var.get())
                if delay > 0:
                    self.save_setting('detection_delay', delay)
            except ValueError:
                self.delay_var.set("2")
        
        self.delay_var.trace('w', save_delay_setting)
        
        ttk.Label(settings_frame, text="秒", foreground="gray").grid(row=0, column=2, sticky=tk.W, padx=(5, 20))
        
        # 日志保存天数设置
        ttk.Label(settings_frame, text="日志保存天数:").grid(row=0, column=3, sticky=tk.W, padx=(0, 10))
        self.log_days_var = tk.StringVar(value="10")
        log_days_entry = ttk.Entry(settings_frame, textvariable=self.log_days_var, width=8)
        log_days_entry.grid(row=0, column=4, sticky=tk.W)
        
        def save_log_days(*args):
            try:
                days = int(self.log_days_var.get())
                if days > 0:
                    # 保存到配置
                    self.save_setting('log_retention_days', days)
            except ValueError:
                self.log_days_var.set("10")
        
        self.log_days_var.trace('w', save_log_days)
        
        ttk.Label(settings_frame, text="天", foreground="gray").grid(row=0, column=5, sticky=tk.W, padx=(5, 20))
        
        # 队列最大数量设置
        ttk.Label(settings_frame, text="队列最大数量:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))
        self.queue_max_var = tk.StringVar(value="600")
        queue_max_entry = ttk.Entry(settings_frame, textvariable=self.queue_max_var, width=8)
        queue_max_entry.grid(row=1, column=1, sticky=tk.W, pady=(10, 0))
        
        def save_queue_max(*args):
            try:
                max_count = int(self.queue_max_var.get())
                if max_count > 0:
                    # 保存到配置
                    self.save_setting('queue_max_size', max_count)
                    # 立即清理超出的队列
                    if hasattr(self, 'message_queue'):
                        self.message_queue.trim_queue(max_count)
            except ValueError:
                self.queue_max_var.set("600")
        
        self.queue_max_var.trace('w', save_queue_max)
        
        ttk.Label(settings_frame, text="条", foreground="gray").grid(row=1, column=2, sticky=tk.W, padx=(5, 0), pady=(10, 0))
        
        # 说明文本
        ttk.Label(settings_frame, text="说明：日志超过指定天数后自动删除，队列超过指定数量后删除最早的记录", 
                 foreground="gray", font=("Arial", 8)).grid(row=2, column=0, columnspan=6, sticky=tk.W, pady=(10, 0))
    
    def create_control_section(self, parent, row):
        """创建控制按钮区域"""
        control_frame = ttk.Frame(parent)
        control_frame.grid(row=row, column=0, columnspan=3, pady=20)
        
        self.start_button = ttk.Button(control_frame, text="开始转发", command=self.start_forwarding)
        self.start_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.stop_button = ttk.Button(control_frame, text="停止转发", command=self.stop_forwarding, state="disabled")
        self.stop_button.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(control_frame, text="设置复制坐标", command=self.setup_copy_coordinates).pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(control_frame, text="保存配置", command=self.save_config).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(control_frame, text="加载配置", command=self.load_config).pack(side=tk.LEFT, padx=(0, 10))
        
        # 状态显示
        self.status_var = tk.StringVar(value="状态: 待机")
        ttk.Label(control_frame, textvariable=self.status_var, font=("Microsoft YaHei", 10), foreground="blue").pack(side=tk.LEFT, padx=(20, 0))
    
    def create_footer_section(self, parent, row):
        """创建底部说明区域"""
        footer_frame = ttk.Frame(parent)
        footer_frame.grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        footer_frame.columnconfigure(0, weight=1)
        
        # 说明文本
        info_text = "此项目为开源项目，基于WXAUTO开发。"
        info_label = ttk.Label(footer_frame, text=info_text, font=("Microsoft YaHei", 10))
        info_label.grid(row=0, column=0, pady=(5, 2))
        
        # 链接区域1
        links_frame1 = ttk.Frame(footer_frame)
        links_frame1.grid(row=1, column=0, pady=2)
        
        ttk.Label(links_frame1, text="wxauto项目：", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        wxauto_link = ttk.Label(links_frame1, text="https://github.com/cluic/wxauto", 
                               font=("Microsoft YaHei", 10), foreground="blue", cursor="hand2")
        wxauto_link.pack(side=tk.LEFT)
        wxauto_link.bind("<Button-1>", lambda e: self.open_link("https://github.com/cluic/wxauto"))
        
        # 链接区域2
        links_frame2 = ttk.Frame(footer_frame)
        links_frame2.grid(row=2, column=0, pady=2)
        
        ttk.Label(links_frame2, text="微信消息转发助手（本项目）：", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        project_link = ttk.Label(links_frame2, text="https://github.com/yangjinduo/WeChat-Message-Forwarding-Assistant", 
                                font=("Microsoft YaHei", 10), foreground="blue", cursor="hand2")
        project_link.pack(side=tk.LEFT)
        project_link.bind("<Button-1>", lambda e: self.open_link("https://github.com/yangjinduo/WeChat-Message-Forwarding-Assistant"))
        
        # 链接区域3
        links_frame3 = ttk.Frame(footer_frame)
        links_frame3.grid(row=3, column=0, pady=2)
        
        ttk.Label(links_frame3, text="作者介绍：喜欢无人机、穿越机、3D打印、AI编程菜鸟，如有相同爱好欢迎关注作者Bilibili：", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        bilibili_link = ttk.Label(links_frame3, text="https://space.bilibili.com/409575364", 
                                 font=("Microsoft YaHei", 10), foreground="blue", cursor="hand2")
        bilibili_link.pack(side=tk.LEFT)
        bilibili_link.bind("<Button-1>", lambda e: self.open_link("https://space.bilibili.com/409575364?spm_id_from=333.33.0.0"))
    
    def open_link(self, url):
        """打开链接"""
        import webbrowser
        webbrowser.open(url)
    
    def create_log_section(self, parent, row):
        """创建日志显示区域"""
        log_frame = ttk.LabelFrame(parent, text="转发日志", padding="10")
        log_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        parent.rowconfigure(row, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, width=70, height=15, wrap=tk.WORD)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 清空日志按钮
        ttk.Button(log_frame, text="清空日志", command=self.clear_log).grid(row=1, column=0, pady=(10, 0))
    
    def update_queue_status(self):
        """更新消息队列状态显示"""
        try:
            if hasattr(self, 'message_queue') and self.message_queue is not None:
                status = self.message_queue.get_queue_status()
                processing_status = "是" if status['is_processing'] else "否"
                status_text = f"待处理:{status['pending_count']} | 处理中:{processing_status} | 已完成:{status['replied_count']}"
                self.queue_status_var.set(status_text)
            else:
                self.queue_status_var.set("待处理:0 | 处理中:否 | 已完成:0")
        except Exception as e:
            self.queue_status_var.set("队列状态获取失败")
        
        # 每2秒更新一次队列状态
        self.root.after(2000, self.update_queue_status)
        # 每10秒更新一次队列显示（减少频率避免影响选中状态）
        self.root.after(10000, self.auto_refresh_queue_display)
    
    def create_queue_status_section(self, parent, row):
        """创建消息队列状态区域"""
        queue_frame = ttk.LabelFrame(parent, text="消息队列状态", padding="10")
        queue_frame.grid(row=row, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        queue_frame.columnconfigure(0, weight=1)
        queue_frame.rowconfigure(1, weight=1)  # 让表格区域扩展
        
        # 顶部状态和过滤显示区域（同一行）
        status_filter_frame = ttk.Frame(queue_frame)
        status_filter_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        status_filter_frame.columnconfigure(1, weight=1)
        status_filter_frame.columnconfigure(3, weight=1)
        
        # 队列状态标签  
        ttk.Label(status_filter_frame, text="队列状态:", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky=tk.W)
        self.queue_status_var = tk.StringVar(value="待处理:0 | 处理中:否 | 已完成:0")
        ttk.Label(status_filter_frame, textvariable=self.queue_status_var, foreground="green", 
                 font=("Arial", 10)).grid(row=0, column=1, sticky=tk.W, padx=(10, 20))
        
        # 队列过滤下拉菜单（同一行）
        ttk.Label(status_filter_frame, text="队列过滤:", font=("Arial", 10, "bold")).grid(row=0, column=2, sticky=tk.W, padx=(0, 10))
        self.queue_filter_var = tk.StringVar(value="全部显示")
        self.queue_filter_combo = ttk.Combobox(status_filter_frame, textvariable=self.queue_filter_var, 
                                             state="readonly", width=25)
        self.queue_filter_combo.grid(row=0, column=3, sticky=(tk.W, tk.E))
        self.queue_filter_combo.bind('<<ComboboxSelected>>', self.on_queue_filter_change)
        
        # 启动队列状态更新
        self.update_queue_status()
        
        # 初始化过滤选项
        self.update_queue_filter_options()
        
        # 创建Treeview显示消息列表，添加队列ID列
        columns = ('队列ID', '时间', '来源', '发送者', '内容', '状态')
        self.queue_tree = ttk.Treeview(queue_frame, columns=columns, show='headings', height=20)
        
        # 设置列标题和宽度
        self.queue_tree.heading('队列ID', text='队列ID')
        self.queue_tree.heading('时间', text='时间')
        self.queue_tree.heading('来源', text='来源')
        self.queue_tree.heading('发送者', text='发送者')
        self.queue_tree.heading('内容', text='消息内容')
        self.queue_tree.heading('状态', text='状态')
        
        self.queue_tree.column('队列ID', width=100)
        self.queue_tree.column('时间', width=120)
        self.queue_tree.column('来源', width=100)
        self.queue_tree.column('发送者', width=100)
        self.queue_tree.column('内容', width=250)
        self.queue_tree.column('状态', width=80)
        
        self.queue_tree.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 添加滚动条
        tree_scroll = ttk.Scrollbar(queue_frame, orient=tk.VERTICAL, command=self.queue_tree.yview)
        tree_scroll.grid(row=1, column=1, sticky=(tk.N, tk.S))
        self.queue_tree.configure(yscrollcommand=tree_scroll.set)
        
        # 操作按钮框架 - 所有按钮在一行
        button_frame = ttk.Frame(queue_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=(10, 0))
        
        # 所有按钮在一行
        ttk.Button(button_frame, text="删除选中", command=self.delete_selected_message).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="清除所有队列", command=self.clear_all_queue, 
                  style="Accent.TButton" if hasattr(ttk, 'Style') else None).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="刷新队列", command=self.refresh_queue_display).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="清除已完成", command=self.clear_completed_messages).pack(side=tk.LEFT, padx=(0, 10))
        
        # 禁用自动刷新导致的选中状态丢失
        self.queue_tree.bind('<<TreeviewSelect>>', self.on_queue_select)
        
        # 启动队列显示更新
        self.refresh_queue_display()
    
    def update_queue_filter_options(self):
        """更新队列过滤选项"""
        try:
            filter_options = ['全部显示']
            
            # 添加每个规则的过滤选项
            for i, rule in enumerate(self.forwarding_rules, 1):
                if rule.get('enabled', True):
                    source_type = rule['source']['type']
                    source_contact = rule['source']['contact']
                    target_type = rule['target']['type']
                    target_contact = rule['target']['contact']
                    
                    filter_text = f"队列{i} {source_type}{source_contact}<>{target_type}{target_contact}"
                    filter_options.append(filter_text)
            
            # 更新下拉菜单选项
            if hasattr(self, 'queue_filter_combo'):
                current_value = self.queue_filter_var.get()
                self.queue_filter_combo['values'] = filter_options
                
                # 如果当前选中的值不在新的选项中，重置为全部显示
                if current_value not in filter_options:
                    self.queue_filter_var.set('全部显示')
        except Exception as e:
            if hasattr(self, 'log_message'):
                self.log_message(f"⚠️ 更新队列过滤选项失败: {e}")
    
    def on_queue_filter_change(self, event=None):
        """队列过滤选项变化事件"""
        self.refresh_queue_display()
    
    def get_queue_id_for_message(self, message_item):
        """获取消息对应的队列ID"""
        try:
            if 'matched_rule' in message_item and message_item['matched_rule']:
                rule_id = message_item['matched_rule']['id']
                # 查找规则在列表中的序号
                for i, rule in enumerate(self.forwarding_rules, 1):
                    if rule['id'] == rule_id:
                        source_type = rule['source']['type']
                        source_contact = rule['source']['contact']
                        target_type = rule['target']['type']
                        target_contact = rule['target']['contact']
                        return f"队列{i} {source_type}{source_contact}<>{target_type}{target_contact}"
            return "未知队列"
        except Exception:
            return "未知队列"
    
    def refresh_queue_display(self):
        """刷新消息队列显示"""
        try:
            # 保存当前选中项
            selected_items = self.queue_tree.selection()
            selected_values = []
            for item in selected_items:
                selected_values.append(self.queue_tree.item(item)['values'])
            
            # 清空现有项目
            for item in self.queue_tree.get_children():
                self.queue_tree.delete(item)
            
            if not hasattr(self, 'message_queue') or self.message_queue is None:
                return
            
            # 获取当前过滤选项
            filter_value = getattr(self, 'queue_filter_var', None)
            current_filter = filter_value.get() if filter_value else '全部显示'
            
            # 准备所有消息列表用于过滤
            all_messages = []
            
            # 添加待处理消息
            for msg in self.message_queue.pending_messages:
                queue_id = self.get_queue_id_for_message(msg)
                status_text = "⏳ 待处理" if msg['status'] == 'pending' else "🔄 处理中"
                if msg['status'] == 'processing':
                    status_text = "🔄 处理中"
                
                all_messages.append({
                    'queue_id': queue_id,
                    'created_time': msg['created_time'],
                    'chat_name': msg['chat_name'],
                    'sender': msg['sender'],
                    'content': msg['content'][:50] + '...' if len(msg['content']) > 50 else msg['content'],
                    'status_text': status_text,
                    'tag': 'pending'
                })
            
            # 添加正在处理的消息
            if self.message_queue.processing_message:
                msg = self.message_queue.processing_message
                queue_id = self.get_queue_id_for_message(msg)
                all_messages.append({
                    'queue_id': queue_id,
                    'created_time': msg['created_time'],
                    'chat_name': msg['chat_name'],
                    'sender': msg['sender'],
                    'content': msg['content'][:50] + '...' if len(msg['content']) > 50 else msg['content'],
                    'status_text': "🔄 处理中",
                    'tag': 'processing'
                })
            
            # 添加最近的已完成/失败消息（最多10条）
            recent_completed = self.message_queue.replied_messages[-10:] if self.message_queue.replied_messages else []
            for msg in recent_completed:
                queue_id = self.get_queue_id_for_message(msg)
                if msg['status'] == 'replied':
                    status_text = "✅ 已完成"
                    tag = 'completed'
                elif msg['status'] == 'failed':
                    status_text = "❌ 失败"
                    tag = 'failed'
                else:
                    status_text = msg['status']
                    tag = 'other'
                
                all_messages.append({
                    'queue_id': queue_id,
                    'created_time': msg['created_time'], 
                    'chat_name': msg['chat_name'],
                    'sender': msg['sender'],
                    'content': msg['content'][:50] + '...' if len(msg['content']) > 50 else msg['content'],
                    'status_text': status_text,
                    'tag': tag
                })
            
            # 应用过滤并显示消息
            for msg_data in all_messages:
                # 检查是否符合过滤条件
                if current_filter == '全部显示' or msg_data['queue_id'] == current_filter:
                    self.queue_tree.insert('', 'end', values=(
                        msg_data['queue_id'],
                        msg_data['created_time'],
                        msg_data['chat_name'],
                        msg_data['sender'],
                        msg_data['content'],
                        msg_data['status_text']
                    ), tags=(msg_data['tag'],))
            
            # 设置颜色标签
            self.queue_tree.tag_configure('pending', background='#fff3cd')
            self.queue_tree.tag_configure('processing', background='#d1ecf1')
            self.queue_tree.tag_configure('completed', background='#d4edda')
            self.queue_tree.tag_configure('failed', background='#f8d7da')
            
            # 恢复之前的选中状态
            if selected_values:
                for child in self.queue_tree.get_children():
                    item_values = self.queue_tree.item(child)['values']
                    if item_values in selected_values:
                        self.queue_tree.selection_add(child)
            
        except Exception as e:
            self.log_message(f"❌ 刷新队列显示失败: {e}")
    
    def clear_completed_messages(self):
        """清除已完成的消息"""
        try:
            if hasattr(self, 'message_queue'):
                # 只保留失败的消息
                failed_messages = [msg for msg in self.message_queue.replied_messages if msg['status'] == 'failed']
                self.message_queue.replied_messages = failed_messages
                self.message_queue.save_to_file()
                self.refresh_queue_display()
                self.log_message("🗑️ 已清除完成的消息")
        except Exception as e:
            self.log_message(f"❌ 清除完成消息失败: {e}")
    
    
    def delete_selected_message(self):
        """删除选中的消息"""
        try:
            selected = self.queue_tree.selection()
            if not selected:
                self.log_message("ℹ️ 请先选择要删除的消息")
                return
                
            # 获取要删除的消息信息（在确认对话框之前获取，避免TreeView状态改变）
            messages_to_delete = []
            try:
                for item in selected:
                    values = self.queue_tree.item(item)['values']
                    if len(values) >= 5:  # 确保有足够的列（时间、来源、发送者、内容、状态）
                        messages_to_delete.append({
                            'created_time': values[0],
                            'chat_name': values[1],
                            'sender': values[2],
                            'content': values[3],
                            'status': values[4]
                        })
            except tk.TclError as e:
                self.log_message(f"❌ 获取选中消息信息失败: {e}")
                return
                
            if not messages_to_delete:
                self.log_message("❌ 未能获取选中的消息信息")
                return
                
            # 确认删除
            result = messagebox.askyesno(
                "确认删除", 
                f"确定要删除选中的 {len(messages_to_delete)} 条消息吗？\n\n此操作不可撤销！"
            )
            
            if result:
                deleted_count = 0
                
                # 从队列中删除匹配的消息
                if hasattr(self, 'message_queue'):
                    # 删除待处理消息
                    original_pending = len(self.message_queue.pending_messages)
                    self.message_queue.pending_messages = [
                        msg for msg in self.message_queue.pending_messages
                        if not any(
                            msg['created_time'] == del_msg['created_time'] and
                            msg['chat_name'] == del_msg['chat_name'] and
                            msg['content'] == del_msg['content']
                            for del_msg in messages_to_delete
                        )
                    ]
                    deleted_count += original_pending - len(self.message_queue.pending_messages)
                    
                    # 删除历史消息
                    original_replied = len(self.message_queue.replied_messages)
                    self.message_queue.replied_messages = [
                        msg for msg in self.message_queue.replied_messages
                        if not any(
                            msg['created_time'] == del_msg['created_time'] and
                            msg['chat_name'] == del_msg['chat_name'] and
                            msg['content'] == del_msg['content']
                            for del_msg in messages_to_delete
                        )
                    ]
                    deleted_count += original_replied - len(self.message_queue.replied_messages)
                    
                    # 检查正在处理的消息
                    if self.message_queue.processing_message:
                        proc_msg = self.message_queue.processing_message
                        if any(
                            proc_msg['created_time'] == del_msg['created_time'] and
                            proc_msg['chat_name'] == del_msg['chat_name'] and
                            proc_msg['content'] == del_msg['content']
                            for del_msg in messages_to_delete
                        ):
                            self.message_queue.processing_message = None
                            self.message_queue.is_processing = False
                            deleted_count += 1
                    
                    # 保存更改
                    self.message_queue.save_to_file()
                    
                    # 刷新显示
                    self.refresh_queue_display()
                    
                    self.log_message(f"🗑️ 已删除 {deleted_count} 条消息")
                else:
                    self.log_message("⚠️ 消息队列未初始化")
            else:
                self.log_message("ℹ️ 取消删除操作")
        except Exception as e:
            self.log_message(f"❌ 删除消息失败: {e}")
    
    def clear_all_queue(self):
        """清除所有队列消息"""
        try:
            if hasattr(self, 'message_queue'):
                # 确认对话框
                result = messagebox.askyesno(
                    "确认清除", 
                    "⚠️ 确定要清除所有队列消息吗？\n\n这将删除：\n• 所有待处理消息\n• 所有历史记录\n• 正在处理的消息\n\n此操作不可撤销！"
                )
                
                if result:
                    # 清除所有消息
                    self.message_queue.pending_messages.clear()
                    self.message_queue.replied_messages.clear()
                    self.message_queue.processing_message = None
                    self.message_queue.is_processing = False
                    
                    # 保存到文件
                    self.message_queue.save_to_file()
                    
                    # 刷新显示
                    self.refresh_queue_display()
                    
                    self.log_message("🗑️ 已清除所有队列消息")
                else:
                    self.log_message("ℹ️ 取消清除操作")
            else:
                self.log_message("⚠️ 消息队列未初始化")
                
        except Exception as e:
            self.log_message(f"❌ 清除所有队列失败: {e}")
    
    def on_queue_select(self, event):
        """队列选中事件处理"""
        try:
            # 这个方法用于处理选中事件，但主要是为了避免事件冲突
            # 实际的选中状态保持在refresh_queue_display中处理
            pass
        except Exception as e:
            pass  # 静默处理选中事件错误
    
    def auto_refresh_queue_display(self):
        """自动刷新队列显示"""
        try:
            if hasattr(self, 'queue_tree'):
                self.refresh_queue_display()
        except Exception as e:
            pass  # 静默处理自动刷新错误
        
        # 继续下一次自动刷新（10秒间隔）
        self.root.after(10000, self.auto_refresh_queue_display)
    
    def show_restart_warning(self):
        """显示重启后的未处理消息警告"""
        try:
            pending_count = len(self.message_queue.pending_messages)
            failed_count = sum(1 for msg in self.message_queue.replied_messages if msg['status'] == 'failed')
            
            if pending_count > 0 or failed_count > 0:
                warning_msg = "⚠️ 检测到未处理的消息！\n\n"
                if pending_count > 0:
                    warning_msg += f"📝 待处理消息: {pending_count} 条\n"
                if failed_count > 0:
                    warning_msg += f"❌ 失败消息: {failed_count} 条\n"
                
                warning_msg += "\n注意：程序重启后不会自动处理这些消息。\n"
                warning_msg += "如需处理，请在消息队列状态区域手动操作：\n"
                warning_msg += "• 点击'重试失败消息'重新处理失败的消息\n"
                warning_msg += "• 点击'清除已完成'清理历史记录\n"
                warning_msg += "• 启动转发功能处理待处理消息"
                
                # 延迟显示警告，等GUI完全加载
                self.root.after(1000, lambda: messagebox.showwarning("未处理消息提醒", warning_msg))
                self.log_message(f"⚠️ 发现未处理消息: 待处理{pending_count}条, 失败{failed_count}条")
                
        except Exception as e:
            self.log_message(f"❌ 显示重启警告失败: {e}")
    
    def on_source_type_change(self):
        """源微信类型变化时的回调"""
        self.refresh_source_contacts()
    
    def on_target_type_change(self):
        """目标微信类型变化时的回调"""
        self.refresh_target_contacts()
    
    def on_filter_type_change(self):
        """过滤类型变化时的回调"""
        if self.filter_type_var.get() == "mention_range":
            # 启用范围设置
            for child in self.range_frame.winfo_children():
                child.configure(state="normal")
        else:
            # 禁用范围设置
            for child in self.range_frame.winfo_children():
                if isinstance(child, ttk.Entry):
                    child.configure(state="disabled")
    
    def refresh_source_contacts(self):
        """刷新源联系人列表"""
        try:
            wechat_type = self.source_type_var.get()
            if wechat_type == "wecom":
                self.log_message("企业微信：正在查找已打开的独立聊天窗口...")
            else:
                self.log_message("普通微信：正在获取会话列表...")
                
            contacts = self.get_contacts(wechat_type)
            self.source_contact_combo['values'] = contacts
            if contacts:
                self.source_contact_combo.set(contacts[0])
                
            if wechat_type == "wecom":
                self.log_message(f"已刷新企业微信聊天窗口列表，共{len(contacts)}个")
                if len(contacts) == 0:
                    messagebox.showwarning("提示", "未找到企业微信聊天窗口！\n请先在企业微信中打开要使用的联系人的独立聊天窗口。")
            else:
                self.log_message(f"已刷新普通微信联系人列表，共{len(contacts)}个")
        except Exception as e:
            self.log_message(f"刷新源联系人列表失败: {e}")
            messagebox.showerror("错误", f"刷新联系人列表失败: {e}")
    
    def refresh_target_contacts(self):
        """刷新目标联系人列表"""
        try:
            wechat_type = self.target_type_var.get()
            if wechat_type == "wecom":
                self.log_message("企业微信：正在查找已打开的独立聊天窗口...")
            else:
                self.log_message("普通微信：正在获取会话列表...")
                
            contacts = self.get_contacts(wechat_type)
            self.target_contact_combo['values'] = contacts
            if contacts:
                self.target_contact_combo.set(contacts[0])
                
            if wechat_type == "wecom":
                self.log_message(f"已刷新企业微信聊天窗口列表，共{len(contacts)}个")
                if len(contacts) == 0:
                    messagebox.showwarning("提示", "未找到企业微信聊天窗口！\n请先在企业微信中打开要使用的联系人的独立聊天窗口。")
            else:
                self.log_message(f"已刷新普通微信联系人列表，共{len(contacts)}个")
        except Exception as e:
            self.log_message(f"刷新目标联系人列表失败: {e}")
            messagebox.showerror("错误", f"刷新联系人列表失败: {e}")
    
    def refresh_wechat_nickname(self):
        """刷新当前微信昵称"""
        try:
            if not self.wechat:
                self.wechat = WeChat()
            
            if hasattr(self.wechat, 'nickname') and self.wechat.nickname:
                self.current_wechat_nickname = self.wechat.nickname
                if hasattr(self, 'current_nickname_var'):
                    self.current_nickname_var.set(self.current_wechat_nickname)
                self.log_message(f"✅ 已获取当前微信昵称: {self.current_wechat_nickname}")
                
                # 如果@某人输入框为空，自动填入当前昵称
                if hasattr(self, 'mention_name_var') and not self.mention_name_var.get():
                    self.mention_name_var.set(self.current_wechat_nickname)
                
                # 自动保存昵称到配置文件
                self.save_nickname_to_config()
                    
            else:
                if hasattr(self, 'current_nickname_var'):
                    self.current_nickname_var.set("获取失败")
                self.log_message("❌ 无法获取微信昵称，请确保微信已打开并登录")
                
        except Exception as e:
            if hasattr(self, 'current_nickname_var'):
                self.current_nickname_var.set("获取失败")
            self.log_message(f"❌ 获取微信昵称失败: {e}")
    
    def use_current_nickname(self):
        """使用当前微信昵称填入@某人输入框"""
        if self.current_wechat_nickname:
            self.mention_name_var.set(self.current_wechat_nickname)
        else:
            # 如果没有当前昵称，尝试获取
            self.refresh_wechat_nickname()
    
    def save_nickname_to_config(self):
        """单独保存昵称到配置文件"""
        try:
            # 读取现有配置
            try:
                with open('forwarder_config.json', 'r', encoding='utf-8') as f:
                    config = json.load(f)
            except FileNotFoundError:
                config = {}
            
            # 更新昵称
            config['wechat_nickname'] = self.current_wechat_nickname or ""
            
            # 写回文件
            with open('forwarder_config.json', 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
                
        except Exception as e:
            self.log_message(f"保存昵称到配置失败: {e}")
    
    def get_contacts(self, wechat_type):
        """获取指定微信类型的联系人列表"""
        try:
            if wechat_type == "wechat":
                # 普通微信：通过会话列表获取联系人
                if not self.wechat:
                    self.wechat = WeChat()
                sessions = self.wechat.GetSession()
                return [session.name for session in sessions]
            else:  # wecom
                # 企业微信：通过查找独立聊天窗口获取联系人
                return self.get_wecom_chat_windows()
        except Exception as e:
            self.log_message(f"获取{wechat_type}联系人失败: {e}")
            return []
    
    def get_wecom_chat_windows(self):
        """获取企业微信独立聊天窗口列表"""
        try:
            import win32gui
            import win32process
            from wxauto.utils.win32 import GetAllWindows
            
            chat_windows = []
            all_windows = GetAllWindows()
            
            for hwnd, class_name, window_title in all_windows:
                # 检查是否是企业微信的独立聊天窗口
                if self.is_wecom_chat_window(hwnd, class_name, window_title):
                    chat_windows.append(window_title)
            
            self.log_message(f"找到 {len(chat_windows)} 个企业微信聊天窗口: {chat_windows}")
            return chat_windows
            
        except Exception as e:
            self.log_message(f"获取企业微信聊天窗口失败: {e}")
            return []
    
    def is_wecom_chat_window(self, hwnd, class_name, window_title):
        """判断是否是企业微信聊天窗口"""
        try:
            # 检查类名是否匹配企业微信聊天窗口
            wecom_chat_classes = [
                'WwStandaloneConversationWnd',  # 企业微信独立对话窗口
                'ChatWnd',                      # 通用聊天窗口
                'WeComChatWnd',                 # 企业微信聊天窗口
                'WorkWeChatChatWnd'             # 工作微信聊天窗口
            ]
            
            if class_name in wecom_chat_classes:
                # 进一步验证：检查窗口标题不为空且不是主窗口标题
                if window_title and window_title not in ['', '企业微信', 'WeCom', 'WeChat Work']:
                    # 验证是否属于企业微信进程
                    try:
                        import win32process
                        import win32api
                        thread_id, process_id = win32process.GetWindowThreadProcessId(hwnd)
                        process_handle = win32api.OpenProcess(win32process.PROCESS_QUERY_INFORMATION, False, process_id)
                        try:
                            process_name = win32process.GetModuleFileNameEx(process_handle, 0)
                            if 'WXWork.exe' in process_name or 'WeWork' in process_name:
                                return True
                        finally:
                            win32api.CloseHandle(process_handle)
                    except:
                        # 如果无法验证进程，但类名匹配，也认为是有效的
                        return True
            
            return False
            
        except Exception as e:
            self.log_message(f"验证企业微信窗口失败: {e}")
            return False
    
    def start_forwarding(self):
        """开始转发"""
        if not self.validate_config():
            return
            
        self.is_forwarding = True
        self.start_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        self.status_var.set("状态: 转发中...")
        
        # 启动转发线程
        self.forward_thread = threading.Thread(target=self.forwarding_loop, daemon=True)
        self.forward_thread.start()
        
        # 启动消息处理器线程
        self.start_message_processor()
        
        self.log_message("开始消息转发")
    
    def stop_forwarding(self):
        """停止转发"""
        self.is_forwarding = False
        self.start_button.configure(state="normal")
        self.stop_button.configure(state="disabled")
        self.status_var.set("状态: 已停止")
        
        self.log_message("停止消息转发")
    
    def start_message_processor(self):
        """启动消息处理器线程"""
        def process_loop():
            self.log_message("🚀 消息处理器已启动")
            while self.is_forwarding:
                try:
                    if not self.message_queue.is_processing and len(self.message_queue.pending_messages) > 0:
                        next_message = self.message_queue.get_next_message()
                        if next_message:
                            self.process_single_message(next_message)
                    time.sleep(2)  # 每2秒检查一次队列
                except Exception as e:
                    self.log_message(f"❌ 消息处理循环错误: {e}")
                    time.sleep(5)  # 出错后等待5秒重试
            self.log_message("🛑 消息处理器已停止")
        
        self.message_processor_thread = threading.Thread(target=process_loop, daemon=True)
        self.message_processor_thread.start()
    
    def process_single_message(self, message_item):
        """处理单条消息的完整流程（支持多规则）"""
        try:
            self.message_queue.is_processing = True
            self.message_queue.processing_message = message_item
            message_item['status'] = 'processing'
            message_item['process_start_time'] = time.time()
            self.message_queue.save_to_file()  # 保存处理状态
            
            # 从消息项中获取匹配的规则
            rule = message_item.get('matched_rule')
            if not rule:
                raise Exception("消息项中未找到匹配的规则")
            
            rule_id = rule.get('id')
            
            self.log_message(f"🔄 开始处理消息: {message_item['content'][:30]}...", rule_id)
            
            target_type = rule['target']['type']
            target_contact = rule['target']['contact']
            
            self.log_message(f"🎯 使用规则: {rule['name']} -> {target_type}:{target_contact}", rule_id)
            
            # 1. 发送消息到目标
            success = self.send_message_to_target(message_item, target_type, target_contact)
            if not success:
                raise Exception("发送到目标失败")
            
            # 2. 如果目标是企业微信，等待AI回复
            if target_type == "wecom":
                ai_reply = self.wait_for_ai_reply_with_timeout(timeout=300, rule_id=rule_id, target_contact=target_contact)  # 5分钟超时
                if not ai_reply:
                    raise Exception("AI回复超时")
                
                # 3. 转发回复到源发送者
                success = self.forward_ai_reply_to_source(ai_reply, message_item)
                if not success:
                    raise Exception("转发回复失败")
                
                # 4. 记录AI回复，避免循环转发
                self.record_ai_reply(ai_reply)
                
                # 5. 检查是否需要复制回复（是否有复制坐标配置）
                needs_copy = self.check_if_needs_copy(target_contact)
                if needs_copy:
                    # 如果需要复制，不要在这里标记完成，等待复制完成后再标记
                    self.log_message("⏳ 等待AI回复复制过程完成...", rule_id)
                    return  # 不标记完成，让复制过程来标记
                else:
                    ai_reply = ai_reply
            else:
                ai_reply = "消息已转发"
            
            # 5. 标记完成（只有不需要复制的情况才会到这里）
            self.message_queue.mark_message_completed(message_item, ai_reply, success=True)
            
        except Exception as e:
            # 处理失败
            error_msg = str(e)
            # 尝试获取规则ID用于日志
            rule_id = None
            if 'matched_rule' in message_item:
                rule_id = message_item['matched_rule'].get('id')
            
            self.log_message(f"❌ 消息处理失败: {error_msg}", rule_id)
            self.message_queue.mark_message_completed(message_item, error_msg, success=False)
    
    def check_if_needs_copy(self, target_contact):
        """检查是否需要复制回复（是否配置了复制坐标）"""
        try:
            with open('forwarder_config.json', 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            copy_coords = config.get('copy_coordinates', {}).get(target_contact, None)
            return copy_coords is not None
            
        except Exception as e:
            self.log_message(f"⚠️ 检查复制配置失败: {e}")
            return False
    
    def send_message_to_target(self, message_item, target_type, target_contact):
        """发送消息到目标"""
        try:
            # 构建转发消息内容
            sender = message_item['sender']
            content = message_item['content']
            chat_name = message_item['chat_name']
            rule_id = message_item.get('matched_rule', {}).get('id')
            
            self.log_message(f"📤 准备发送消息:", rule_id)
            self.log_message(f"   源聊天: {chat_name}", rule_id)
            self.log_message(f"   发送者: {sender}", rule_id)
            self.log_message(f"   目标类型: {target_type}", rule_id)
            self.log_message(f"   目标联系人: {target_contact}", rule_id)
            
            # 构建转发消息
            forward_content = f"[来自 {chat_name}] {sender}: {content}"
            
            if target_type == "wecom":
                # 发送到企业微信窗口
                self.log_message(f"🎯 尝试发送到企业微信: {target_contact}", rule_id)
                success = self.send_to_wecom_window(forward_content, target_contact)
                if success:
                    self.log_message(f"✅ 成功发送到企业微信: {target_contact}", rule_id)
                else:
                    self.log_message(f"❌ 发送到企业微信失败: {target_contact}", rule_id)
                return success
            else:
                # 发送到普通微信
                if self.wechat:
                    try:
                        target_chat = self.wechat.ChatWith(target_contact)
                        target_chat.SendMsg(forward_content)
                        self.log_message(f"✅ 消息已发送到普通微信: {target_contact}", rule_id)
                        return True
                    except Exception as e:
                        self.log_message(f"❌ 发送到普通微信失败: {e}", rule_id)
                        return False
                else:
                    self.log_message("❌ 普通微信实例未初始化", rule_id)
                    return False
                
        except Exception as e:
            rule_id = message_item.get('matched_rule', {}).get('id')
            self.log_message(f"❌ 发送消息到目标失败: {e}", rule_id)
            return False
    
    def wait_for_ai_reply_with_timeout(self, timeout=300, rule_id=None, target_contact=None):
        """等待AI回复完成（带超时）"""
        try:
            self.log_message(f"⏳ 等待AI回复完成（超时: {timeout}秒）...", rule_id)
            
            # 查找企业微信窗口
            if not target_contact:
                # 如果没有提供目标联系人，尝试从规则ID获取
                if rule_id:
                    rule = next((r for r in self.forwarding_rules if r['id'] == rule_id), None)
                    if rule:
                        target_contact = rule['target']['contact']
                
                if not target_contact:
                    self.log_message("❌ 无法确定目标联系人", rule_id)
                    return None
            
            hwnd = self.find_wecom_chat_window(target_contact)
            
            if not hwnd:
                self.log_message(f"❌ 未找到企业微信聊天窗口: {target_contact}", rule_id)
                return None
            
            # 启动同步AI回复检测
            ai_reply = self.start_ai_reply_detection_sync(hwnd, timeout, rule_id)
            
            if ai_reply:
                self.log_message(f"✅ AI回复检测完成: {ai_reply[:50]}...", rule_id)
                return ai_reply
            else:
                self.log_message("⏰ AI回复检测超时", rule_id)
                return None
                
        except Exception as e:
            self.log_message(f"❌ AI回复检测错误: {e}", rule_id)
            return None
    
    def start_ai_reply_detection_sync(self, hwnd, timeout=300, rule_id=None):
        """同步版本的AI回复检测"""
        try:
            self.log_message("🔍 开始同步检测AI回复...", rule_id)
            
            # 读取配置的延迟时间
            try:
                delay_seconds = int(self.delay_var.get()) if hasattr(self, 'delay_var') else 2
            except:
                delay_seconds = 2
            
            self.log_message(f"⏰ 等待 {delay_seconds} 秒后开始截图检测...", rule_id)
            time.sleep(delay_seconds)
            
            # 截取第一张图
            previous_image = self.capture_wecom_area(hwnd)
            if not previous_image:
                self.log_message("❌ 初始截图失败", rule_id)
                return None
            
            self.log_message("📸 初始截图成功，开始循环检测...", rule_id)
            
            start_time = time.time()
            check_count = 0
            
            # 开始5秒间隔的循环截图检测
            while True:
                # 检查超时
                elapsed = time.time() - start_time
                if elapsed >= timeout:
                    self.log_message(f"⏰ AI回复检测超时（{timeout}秒）", rule_id)
                    return None
                
                time.sleep(5)  # 等待5秒
                check_count += 1
                
                # 截取当前图像
                current_image = self.capture_wecom_area(hwnd)
                if not current_image:
                    self.log_message(f"❌ 第{check_count}次截图失败", rule_id)
                    continue
                
                # 比较图像是否相同
                is_identical = self.compare_images(previous_image, current_image)
                
                if is_identical:
                    self.log_message(f"✅ 第{check_count}次截图与上次相同，AI回复完成！", rule_id)
                    
                    # 回复完成，复制AI回复消息
                    # 从规则中获取目标联系人
                    target_contact = None
                    if rule_id:
                        rule = next((r for r in self.forwarding_rules if r['id'] == rule_id), None)
                        if rule:
                            target_contact = rule['target']['contact']
                    
                    ai_reply = self.copy_ai_reply_sync(hwnd, rule_id, target_contact)
                    return ai_reply
                else:
                    self.log_message(f"📸 第{check_count}次截图有变化，继续监控...（已用时{elapsed:.1f}秒）", rule_id)
                    previous_image = current_image
                
                # 检查是否应该继续（转发状态）
                if not self.is_forwarding:
                    self.log_message("🛑 转发已停止，结束AI回复检测", rule_id)
                    return None
                    
        except Exception as e:
            self.log_message(f"❌ 同步AI回复检测出错: {e}", rule_id)
            return None
    
    def copy_ai_reply_sync(self, hwnd, rule_id=None, target_contact=None):
        """同步复制AI回复消息"""
        try:
            self.log_message("📋 开始复制AI回复消息...", rule_id)
            
            # 获取目标联系人名称
            if not target_contact:
                # 如果没有提供目标联系人，尝试从规则ID获取
                if rule_id:
                    rule = next((r for r in self.forwarding_rules if r['id'] == rule_id), None)
                    if rule:
                        target_contact = rule['target']['contact']
                
                if not target_contact:
                    self.log_message("❌ 无法确定目标联系人", rule_id)
                    return None
            
            # 从配置文件加载复制坐标
            try:
                with open('forwarder_config.json', 'r', encoding='utf-8') as f:
                    config = json.load(f)
                
                copy_coords = config.get('copy_coordinates', {}).get(target_contact, None)
                if not copy_coords:
                    self.log_message(f"❌ 未找到 {target_contact} 的复制坐标配置", rule_id)
                    return None
                
                right_click_offset = copy_coords['right_click']
                copy_click_offset = copy_coords['copy_click']
                
            except Exception as e:
                self.log_message(f"❌ 加载复制坐标配置失败: {e}", rule_id)
                return None
            
            # 激活企业微信窗口
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.3)
            
            # 获取窗口位置，计算基准点（左下角）
            window_rect = win32gui.GetWindowRect(hwnd)
            base_x = window_rect[0]  # 窗口左边界
            base_y = window_rect[3]  # 窗口下边界（左下角）
            
            # 计算右键点击的绝对坐标
            right_click_x = base_x + right_click_offset[0]
            right_click_y = base_y - right_click_offset[1]  # Y轴反向
            
            # 移动鼠标并右键点击
            win32api.SetCursorPos((right_click_x, right_click_y))
            time.sleep(0.3)
            
            # 右键点击
            win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
            win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
            time.sleep(0.5)  # 等待右键菜单出现
            
            # 计算复制按钮的绝对坐标
            copy_x = base_x + copy_click_offset[0]
            copy_y = base_y - copy_click_offset[1]  # Y轴反向
            
            # 移动鼠标并点击复制
            win32api.SetCursorPos((copy_x, copy_y))
            time.sleep(0.2)
            
            # 左键点击复制按钮
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
            time.sleep(0.5)  # 等待复制完成
            
            # 获取剪贴板内容
            try:
                import win32clipboard
                win32clipboard.OpenClipboard()
                ai_reply = win32clipboard.GetClipboardData()
                win32clipboard.CloseClipboard()
                
                if ai_reply and ai_reply.strip():
                    self.log_message(f"✅ 成功复制AI回复: {ai_reply[:50]}...", rule_id)
                    return ai_reply.strip()
                else:
                    self.log_message("❌ 剪贴板内容为空", rule_id)
                    return None
                    
            except Exception as e:
                self.log_message(f"❌ 获取剪贴板内容失败: {e}", rule_id)
                return None
                
        except Exception as e:
            self.log_message(f"❌ 同步复制回复失败: {e}", rule_id)
            return None
    
    def find_wecom_chat_window(self, target_contact):
        """查找指定联系人的企业微信聊天窗口句柄"""
        try:
            from wxauto.utils.win32 import FindWindow, GetAllWindows
            
            self.log_message(f"🔍 查找企业微信聊天窗口: {target_contact}")
            
            # 首先尝试直接通过窗口标题查找
            hwnd = FindWindow(name=target_contact)
            if hwnd:
                class_name = win32gui.GetClassName(hwnd)
                self.log_message(f"📋 找到窗口: {target_contact}, 类名: {class_name}, 句柄: {hwnd}")
                if self.is_wecom_chat_window(hwnd, class_name, target_contact):
                    self.log_message(f"✅ 直接找到企业微信窗口: {target_contact}, 句柄: {hwnd}")
                    return hwnd
                else:
                    self.log_message(f"⚠️ 窗口不是企业微信聊天窗口: {class_name}")
            else:
                self.log_message(f"⚠️ 直接查找窗口失败: {target_contact}")
            
            # 如果直接查找失败，遍历所有企业微信窗口
            self.log_message("🔍 开始遍历所有窗口查找企业微信聊天窗口...")
            all_windows = GetAllWindows()
            wecom_windows = []
            
            for hwnd, class_name, window_title in all_windows:
                if self.is_wecom_chat_window(hwnd, class_name, window_title):
                    wecom_windows.append((hwnd, class_name, window_title))
                    self.log_message(f"📋 发现企业微信窗口: {window_title} (类名: {class_name})")
                    
                    # 检查窗口标题是否包含目标联系人名称
                    if target_contact in window_title or window_title == target_contact:
                        self.log_message(f"✅ 遍历找到匹配窗口: {window_title}, 句柄: {hwnd}")
                        return hwnd
            
            if wecom_windows:
                self.log_message(f"📊 找到 {len(wecom_windows)} 个企业微信聊天窗口，但没有匹配的:")
                for hwnd, class_name, window_title in wecom_windows:
                    self.log_message(f"   - {window_title} (类名: {class_name})")
            else:
                self.log_message("❌ 没有找到任何企业微信聊天窗口")
            
            self.log_message(f"❌ 未找到企业微信聊天窗口: {target_contact}")
            return None
            
        except Exception as e:
            self.log_message(f"❌ 查找企业微信聊天窗口失败: {e}")
            import traceback
            self.log_message(f"详细错误: {traceback.format_exc()}")
            return None
    
    def forward_ai_reply_to_source(self, ai_reply, message_item):
        """将AI回复转发到源发送者"""
        try:
            # 从消息项中获取规则和源信息
            rule = message_item.get('matched_rule', {})
            rule_id = rule.get('id')
            source_type = message_item.get('source_type', 'wechat')
            chat_name = message_item['chat_name']
            sender = message_item['sender']
            
            self.log_message(f"📨 准备转发AI回复到源:", rule_id)
            self.log_message(f"   目标聊天: {chat_name}", rule_id)
            self.log_message(f"   原发送者: {sender}", rule_id)
            self.log_message(f"   源类型: {source_type}", rule_id)
            
            if source_type == "wechat" and self.wechat:
                # 转发到普通微信 - 直接查找独立聊天窗口
                try:
                    self.log_message(f"🔍 查找微信独立聊天窗口: {chat_name}")
                    
                    # 使用UIAutomation直接查找独立的聊天窗口
                    success = self.send_to_wechat_window(chat_name, ai_reply, sender)
                    if success:
                        self.log_message(f"✅ AI回复已成功转发到普通微信: {chat_name}")
                        return True
                    else:
                        self.log_message(f"❌ 无法找到或发送到微信聊天窗口: {chat_name}")
                        return False
                    
                except Exception as e:
                    self.log_message(f"❌ 转发AI回复到普通微信失败: {e}")
                    self.log_message(f"详细错误信息: {type(e).__name__}: {str(e)}")
                    import traceback
                    self.log_message(f"错误堆栈: {traceback.format_exc()}")
                    return False
            else:
                # 如果源是企业微信，暂不支持反向转发
                self.log_message("⚠️ 暂不支持转发到企业微信源")
                return True  # 标记为成功，避免重试
            
        except Exception as e:
            self.log_message(f"❌ 转发AI回复失败: {e}")
            return False
    
    def validate_config(self):
        """验证多规则配置"""
        try:
            # 检查是否有规则
            if not self.forwarding_rules:
                messagebox.showerror("错误", "没有转发规则，请先添加规则")
                return False
            
            # 检查是否有启用的规则
            enabled_rules = [rule for rule in self.forwarding_rules if rule.get('enabled', True)]
            if not enabled_rules:
                messagebox.showerror("错误", "没有启用的转发规则，请先启用至少一个规则")
                return False
            
            # 逐个验证启用的规则
            invalid_rules = []
            for i, rule in enumerate(enabled_rules, 1):
                rule_name = rule.get('name', f'规则{i}')
                
                # 检查源联系人
                if not rule['source'].get('contact'):
                    invalid_rules.append(f"{rule_name}: 未设置源联系人")
                    continue
                
                # 检查目标联系人
                if not rule['target'].get('contact'):
                    invalid_rules.append(f"{rule_name}: 未设置目标联系人")
                    continue
                
                # 检查过滤条件
                filter_type = rule['source'].get('filter_type', 'all')
                if filter_type == 'range':
                    if not rule['source'].get('range_start') or not rule['source'].get('range_end'):
                        invalid_rules.append(f"{rule_name}: 范围过滤未设置开始和结束标记")
            
            # 如果有无效规则，显示错误
            if invalid_rules:
                error_msg = "以下规则配置不完整：\n\n" + "\n".join(invalid_rules[:5])
                if len(invalid_rules) > 5:
                    error_msg += f"\n\n...还有{len(invalid_rules)-5}个规则有问题"
                error_msg += "\n\n请先完善规则配置后再开始转发。"
                messagebox.showerror("配置错误", error_msg)
                return False
            
            return True
            
        except Exception as e:
            messagebox.showerror("验证错误", f"配置验证失败: {e}")
            return False
    
    def forwarding_loop(self):
        """多规则转发循环"""
        try:
            # 获取所有启用的规则
            enabled_rules = [rule for rule in self.forwarding_rules if rule['enabled']]
            if not enabled_rules:
                self.log_message("⚠️ 没有启用的转发规则")
                return
            
            # 按源类型分组规则
            wechat_rules = [rule for rule in enabled_rules if rule['source']['type'] == 'wechat']
            wecom_rules = [rule for rule in enabled_rules if rule['source']['type'] == 'wecom']
            
            # 初始化微信实例（如果需要）
            if wechat_rules and not self.wechat:
                self.wechat = WeChat()
            
            if wecom_rules and not self.wecom:
                self.wecom = WeCom()
            
            # 为目标是微信的规则初始化微信实例
            target_wechat_rules = [rule for rule in enabled_rules if rule['target']['type'] == 'wechat']
            if target_wechat_rules and not self.wechat:
                self.wechat = WeChat()
            
            # 创建消息回调函数
            def create_message_callback(source_type):
                def message_callback(msg, chat):
                    # 过滤系统消息和自己的消息
                    if self.is_system_message(msg) or self.is_self_message(msg):
                        return
                    
                    # 将消息加入队列（会自动匹配规则）
                    sender = getattr(msg, 'sender', '未知发送者')
                    self.message_queue.add_message(msg, sender, chat, source_type)
                
                return message_callback
            
            # 为微信规则添加监听
            monitored_wechat_contacts = set()
            for rule in wechat_rules:
                contact = rule['source']['contact']
                if contact and contact not in monitored_wechat_contacts:
                    self.wechat.AddListenChat(nickname=contact, callback=create_message_callback('wechat'))
                    monitored_wechat_contacts.add(contact)
                    self.log_message(f"✅ 开始监听微信: {contact}")
            
            # 为企业微信规则添加监听
            monitored_wecom_contacts = set()
            for rule in wecom_rules:
                contact = rule['source']['contact']
                if contact and contact not in monitored_wecom_contacts:
                    self.wecom.AddListenChat(nickname=contact, callback=create_message_callback('wecom'))
                    monitored_wecom_contacts.add(contact)
                    self.log_message(f"✅ 开始监听企业微信: {contact}")
            
            self.log_message(f"🚀 多规则转发已启动：{len(enabled_rules)}条规则，监听 {len(monitored_wechat_contacts)} 个微信联系人和 {len(monitored_wecom_contacts)} 个企业微信联系人")
            
            # 保持监听状态
            while self.is_forwarding:
                time.sleep(1)
            
            # 停止所有监听
            for contact in monitored_wechat_contacts:
                if self.wechat:
                    self.wechat.RemoveListenChat(nickname=contact)
                    self.log_message(f"停止监听微信: {contact}")
            
            for contact in monitored_wecom_contacts:
                if self.wecom:
                    self.wecom.RemoveListenChat(nickname=contact)
                    self.log_message(f"停止监听企业微信: {contact}")
            
        except Exception as e:
            self.log_message(f"转发循环错误: {e}")
            self.root.after(0, self.stop_forwarding)
    
    def should_forward_message(self, msg, chat):
        """判断是否应该转发消息"""
        # 首先过滤掉系统消息
        if self.is_system_message(msg):
            self.log_message(f"跳过系统消息: {msg.content}")
            return False
        
        # 🔥 关键：过滤掉自己发出的消息，避免循环转发
        try:
            # 检查消息发送者是否是自己
            if hasattr(msg, 'sender') and msg.sender:
                # 如果发送者是自己的昵称，不转发
                if self.wechat and hasattr(self.wechat, 'nickname'):
                    if msg.sender == self.wechat.nickname:
                        self.log_message(f"跳过自己发送的消息: {msg.content[:30]}...")
                        return False
            
            # 检查消息类型，通常自己发送的消息有特定属性
            if hasattr(msg, 'type'):
                # 某些wxauto版本中，自己发送的消息可能有特殊类型标识
                if str(msg.type).lower() in ['sent', 'outgoing', 'self']:
                    self.log_message(f"跳过自己发送的消息（类型判断）: {msg.content[:30]}...")
                    return False
            
            # 检查是否是最近的AI回复内容，避免循环转发
            if self.is_recent_ai_reply(msg.content):
                self.log_message(f"跳过最近的AI回复内容: {msg.content[:30]}...")
                return False
                
        except Exception as e:
            self.log_message(f"消息发送者检查失败: {e}")
        
        # 应用用户设置的过滤条件
        filter_type = self.filter_type_var.get()
        
        if filter_type == "all":
            return True
        elif filter_type == "mention_me":
            # 检查是否@了本人（改进版 - 支持多种@格式）
            return self.is_mentioned_me(msg)
        elif filter_type == "mention_range":
            # 检查是否在指定范围内
            start_keyword = self.range_start_var.get()
            end_keyword = self.range_end_var.get()
            # 这里需要实现范围检测逻辑
            return start_keyword in msg.content or end_keyword in msg.content
        
        return False
    
    def is_self_message(self, msg):
        """检查是否是自己发送的消息（改进版 - 支持群聊区分）"""
        try:
            # 方法1: 检查消息的attr属性 - 这是最可靠的方法
            if hasattr(msg, 'attr') and msg.attr == 'self':
                return True
            
            # 方法2: 检查消息发送者是否是自己的昵称
            if hasattr(msg, 'sender') and msg.sender:
                if self.wechat and hasattr(self.wechat, 'nickname'):
                    if msg.sender == self.wechat.nickname:
                        return True
            
            # 方法3: 检查消息类型
            if hasattr(msg, 'type'):
                if str(msg.type).lower() in ['sent', 'outgoing', 'self']:
                    return True
            
            # 方法4: 检查是否是最近的AI回复内容，避免循环转发
            if self.is_recent_ai_reply(msg.content):
                return True
                
        except Exception as e:
            self.log_message(f"消息发送者检查失败: {e}")
        
        return False
    
    def is_mentioned_me(self, msg):
        """检查消息是否@了指定的人（基于输入框中的昵称匹配）"""
        try:
            # 优先使用输入框中的昵称
            target_nickname = self.mention_name_var.get().strip() if hasattr(self, 'mention_name_var') else ""
            
            # 如果输入框为空，尝试使用当前微信昵称
            if not target_nickname:
                if self.current_wechat_nickname:
                    target_nickname = self.current_wechat_nickname
                elif self.wechat and hasattr(self.wechat, 'nickname'):
                    target_nickname = self.wechat.nickname
                else:
                    self.log_message("⚠️ 未设置@检测目标昵称")
                    return False
            
            content = msg.content
            
            self.log_message(f"🔍 检查@消息: 目标昵称='{target_nickname}', 消息内容='{content}'")
            
            # 主要检查方法：消息中是否包含 @目标昵称
            if f"@{target_nickname}" in content:
                self.log_message(f"✅ 检测到@消息: @{target_nickname}")
                return True
            
            # 备用检查：全角@符号
            if f"＠{target_nickname}" in content:
                self.log_message(f"✅ 检测到@消息(全角): ＠{target_nickname}")
                return True
            
            # 如果没有检测到@消息，记录日志用于调试
            self.log_message(f"❌ 未检测到@消息")
            return False
                
        except Exception as e:
            self.log_message(f"检查@消息失败: {e}")
        
        return False
    
    def find_matching_rules(self, msg, chat_name, source_type):
        """查找匹配的转发规则"""
        matching_rules = []
        
        try:
            for rule in self.forwarding_rules:
                # 检查规则是否启用
                if not rule['enabled']:
                    continue
                
                # 检查源类型是否匹配
                if rule['source']['type'] != source_type:
                    continue
                
                # 检查源联系人是否匹配
                if rule['source']['contact'] and rule['source']['contact'] != chat_name:
                    continue
                
                # 检查消息是否符合过滤条件
                if self.message_matches_filter(msg, rule['source']):
                    matching_rules.append(rule)
                    self.log_message(f"✅ 消息匹配规则: {rule['name']}")
        
        except Exception as e:
            self.log_message(f"❌ 规则匹配失败: {e}")
        
        return matching_rules
    
    def message_matches_filter(self, msg, source_config):
        """检查消息是否符合过滤条件"""
        try:
            filter_type = source_config.get('filter_type', 'all')
            
            if filter_type == "all":
                return True
            elif filter_type == "at_me":
                # 检查是否@了本人
                return "@" in msg.content and (
                    self.wechat.nickname in msg.content if self.wechat and hasattr(self.wechat, 'nickname') else False
                )
            elif filter_type == "range":
                # 检查是否在指定范围内
                start_keyword = source_config.get('range_start', '')
                end_keyword = source_config.get('range_end', '')
                content = msg.content
                
                if start_keyword and end_keyword:
                    # 查找范围内的内容
                    start_pos = content.find(start_keyword)
                    if start_pos != -1:
                        end_pos = content.find(end_keyword, start_pos + len(start_keyword))
                        return end_pos != -1
                elif start_keyword:
                    return start_keyword in content
                elif end_keyword:
                    return end_keyword in content
                
                return False
            
            return True
            
        except Exception as e:
            self.log_message(f"❌ 过滤条件检查失败: {e}")
            return False
    
    def is_system_message(self, msg):
        """判断是否是系统消息"""
        try:
            # 检查消息属性
            if hasattr(msg, 'attr'):
                # 过滤系统消息
                if msg.attr in ['system', 'time', 'tickle']:
                    return True
            
            # 检查发送者
            if hasattr(msg, 'sender'):
                if msg.sender in ['system', 'time']:
                    return True
            
            # 检查消息内容是否是系统提示
            system_keywords = [
                "以下为新消息",
                "以上是历史消息",
                "重新载入聊天记录",
                "消息加载中",
                "网络连接失败",
                "正在重新连接",
                "你撤回了一条消息",
                "对方撤回了一条消息"
            ]
            
            if any(keyword in msg.content for keyword in system_keywords):
                return True
            
            # 检查消息类名（如果可用）
            if hasattr(msg, '__class__'):
                class_name = msg.__class__.__name__
                if class_name in ['SystemMessage', 'TimeMessage', 'TickleMessage']:
                    return True
            
            return False
            
        except Exception as e:
            self.log_message(f"检查系统消息失败: {e}")
            return False
    
    def forward_message(self, msg, chat, target_wx, target_contact, target_type):
        """转发消息"""
        try:
            # 构建转发消息内容
            forward_content = f"[来自{self.source_type_var.get()}:{chat.who}] {msg.content}"
            
            if target_type == "wechat":
                # 普通微信：使用wxauto发送
                target_wx.SendMsg(forward_content, who=target_contact)
            else:
                # 企业微信：直接操作独立聊天窗口
                success = self.send_to_wecom_window(forward_content, target_contact)
                if not success:
                    self.log_message(f"企业微信窗口发送失败，尝试备用方法")
                    return
            
            self.log_message(f"已转发: {chat.who} -> {target_contact}: {msg.content[:50]}...")
            
        except Exception as e:
            self.log_message(f"转发消息失败: {e}")
    
    def send_to_wecom_window(self, message, window_title):
        """通过坐标点击向企业微信聊天窗口发送消息"""
        try:
            import win32gui
            import win32con
            import win32api
            import win32clipboard
            import time
            
            # 使用改进的窗口查找逻辑
            hwnd = self.find_wecom_chat_window(window_title)
            if not hwnd:
                self.log_message(f"❌ 未找到企业微信窗口: {window_title}")
                return False
            
            self.log_message(f"找到企业微信窗口: {window_title} (句柄: {hwnd})")
            
            # 激活窗口
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
            win32gui.BringWindowToTop(hwnd)
            time.sleep(0.3)
            self.log_message(f"已激活窗口: {window_title}")
            
            # 获取窗口位置和大小
            rect = win32gui.GetWindowRect(hwnd)
            width = rect[2] - rect[0]
            height = rect[3] - rect[1]
            
            # 🎯 计算输入区域坐标 - 基于窗口左下角的相对位置
            # 坐标说明:
            # - rect[0] = 窗口左边界的屏幕坐标
            # - rect[3] = 窗口下边界的屏幕坐标  
            # - 窗口左下角 = (rect[0], rect[3])
            # - 从左下角向右50像素，向上50像素就是输入区域
            
            offset_right = 50    # 从左下角向右的偏移量（像素）
            offset_up = 100       # 从左下角向上的偏移量（像素）
            
            input_x = rect[0] + offset_right        # 窗口左边界 + 向右偏移
            input_y = rect[3] - offset_up           # 窗口下边界 - 向上偏移
            
            self.log_message(f"输入区域坐标: ({input_x}, {input_y})")
            self.log_message(f"窗口信息: 位置{rect}, 大小{width}x{height}")
            
            # 将消息放入剪贴板（使用Unicode支持）
            try:
                import win32con
                win32clipboard.OpenClipboard()
                win32clipboard.EmptyClipboard()
                # 使用Unicode格式设置剪贴板
                win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, message)
                win32clipboard.CloseClipboard()
                self.log_message("消息已放入剪贴板（Unicode格式）")
            except Exception as clipboard_error:
                self.log_message(f"❌ 设置剪贴板失败: {clipboard_error}")
                # 备用方法：使用tkinter剪贴板
                try:
                    self.root.clipboard_clear()
                    self.root.clipboard_append(message)
                    self.log_message("使用备用方法设置剪贴板成功")
                except Exception as backup_error:
                    self.log_message(f"❌ 备用剪贴板方法也失败: {backup_error}")
                    raise clipboard_error
            
            # 点击输入区域
            win32api.SetCursorPos((input_x, input_y))
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
            time.sleep(0.2)
            self.log_message(f"已点击输入区域: ({input_x}, {input_y})")
            
            # Ctrl+V 粘贴
            win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
            win32api.keybd_event(ord('V'), 0, 0, 0)
            win32api.keybd_event(ord('V'), 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
            time.sleep(0.3)
            self.log_message("已执行粘贴操作 (Ctrl+V)")
            
            # 回车发送
            win32api.keybd_event(win32con.VK_RETURN, 0, 0, 0)
            win32api.keybd_event(win32con.VK_RETURN, 0, win32con.KEYEVENTF_KEYUP, 0)
            self.log_message("已发送回车键")
            
            self.log_message(f"✅ 通过坐标点击发送消息到: {window_title}")
            
            # 🔄 只有在需要复制时才启动异步回复检测
            needs_copy = self.check_if_needs_copy(window_title)
            if needs_copy:
                self.log_message("🔄 启动回复检测和复制转发...")
                self.start_ai_reply_detection(hwnd, window_title, input_x, input_y)
            else:
                self.log_message("⚪ 未配置复制坐标，跳过异步回复检测")
            
            return True
            
        except Exception as e:
            self.log_message(f"坐标点击发送失败: {e}")
            return False
    
    def find_input_control_in_wecom(self, parent_control, depth=0, max_depth=5):
        """在企业微信窗口中查找输入控件"""
        if depth > max_depth:
            return None
            
        try:
            children = parent_control.GetChildren()
            for child in children:
                # 检查是否是输入相关控件
                if (child.ControlTypeName in ['EditControl', 'DocumentControl'] or
                    'Edit' in child.ClassName or 'Input' in child.ClassName):
                    return child
                
                # 递归搜索
                result = self.find_input_control_in_wecom(child, depth + 1, max_depth)
                if result:
                    return result
        except:
            pass
        
        return None
    
    def save_config(self):
        """保存多规则配置"""
        # 保存当前编辑的规则
        self.save_current_rule()
        
        # 首先读取现有配置，保留复制坐标等其他配置
        try:
            with open('forwarder_config.json', 'r', encoding='utf-8') as f:
                config = json.load(f)
        except:
            config = {}
        
        # 更新配置项
        config.update({
            'forwarding_rules': self.forwarding_rules,
            'detection_delay': int(self.delay_var.get()) if hasattr(self, 'delay_var') and self.delay_var.get().isdigit() else 2,
            'log_retention_days': self.log_retention_days,
            'queue_max_size': self.queue_max_size,
            'wechat_nickname': self.current_wechat_nickname or ""
        })
        
        try:
            with open('forwarder_config.json', 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            self.log_message("配置已保存")
            messagebox.showinfo("成功", "配置已保存")
        except Exception as e:
            self.log_message(f"保存配置失败: {e}")
            messagebox.showerror("错误", f"保存配置失败: {e}")
    
    def load_config(self):
        """加载多规则配置"""
        try:
            with open('forwarder_config.json', 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            # 加载多规则配置
            if 'forwarding_rules' in config:
                self.forwarding_rules = config['forwarding_rules']
            else:
                # 兼容旧版单一配置格式
                if 'source' in config and 'target' in config:
                    self.convert_old_config_to_rules(config)
                else:
                    self.init_default_rule()
            
            # 加载延迟设置
            if 'detection_delay' in config and hasattr(self, 'delay_var'):
                self.delay_var.set(str(config['detection_delay']))
            
            # 加载其他设置
            if 'log_retention_days' in config:
                self.log_retention_days = config['log_retention_days']
                if hasattr(self, 'log_days_var'):
                    self.log_days_var.set(str(self.log_retention_days))
            
            if 'queue_max_size' in config:
                self.queue_max_size = config['queue_max_size']
                if hasattr(self, 'queue_max_var'):
                    self.queue_max_var.set(str(self.queue_max_size))
            
            # 加载昵称设置
            if 'wechat_nickname' in config:
                self.current_wechat_nickname = config['wechat_nickname']
                if hasattr(self, 'current_nickname_var'):
                    self.current_nickname_var.set(self.current_wechat_nickname)
                if hasattr(self, 'mention_name_var') and not self.mention_name_var.get():
                    self.mention_name_var.set(self.current_wechat_nickname)
            
            # 刷新规则显示
            if hasattr(self, 'rules_tree'):
                self.refresh_rules_display()
            
            self.log_message("多规则配置已加载")
            
        except FileNotFoundError:
            self.log_message("配置文件不存在，使用默认规则")
            self.init_default_rule()
        except Exception as e:
            self.log_message(f"加载配置失败: {e}")
            messagebox.showerror("错误", f"加载配置失败: {e}")
            self.init_default_rule()
    
    def convert_old_config_to_rules(self, old_config):
        """将旧版单一配置转换为多规则格式"""
        try:
            converted_rule = {
                'id': 'rule_1',
                'name': '转换的规则',
                'enabled': True,
                'source': {
                    'type': old_config['source']['type'],
                    'contact': old_config['source']['contact'],
                    'filter_type': old_config['source']['filter_type'],
                    'range_start': old_config['source']['range_start'],
                    'range_end': old_config['source']['range_end']
                },
                'target': {
                    'type': old_config['target']['type'],
                    'contact': old_config['target']['contact']
                }
            }
            self.forwarding_rules = [converted_rule]
            self.log_message("✅ 已将旧版配置转换为多规则格式")
        except Exception as e:
            self.log_message(f"❌ 转换旧配置失败: {e}")
            self.init_default_rule()
    
    def save_setting(self, key, value):
        """保存单个设置"""
        try:
            # 读取现有配置
            try:
                with open('forwarder_config.json', 'r', encoding='utf-8') as f:
                    config = json.load(f)
            except FileNotFoundError:
                config = {}
            
            # 更新设置
            config[key] = value
            
            # 保存回文件
            with open('forwarder_config.json', 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
            
            # 更新实例变量
            if key == 'log_retention_days':
                self.log_retention_days = value
                self.cleanup_old_logs()
            elif key == 'queue_max_size':
                self.queue_max_size = value
                
        except Exception as e:
            self.log_message(f"保存设置失败: {e}")
    
    def cleanup_old_logs(self):
        """清理过期的日志文件"""
        try:
            import os
            import glob
            from datetime import datetime, timedelta
            
            # 查找所有日志文件
            log_files = glob.glob('logs/forwarder_*.log')
            
            if not log_files:
                return
            
            # 计算过期日期
            cutoff_date = datetime.now() - timedelta(days=self.log_retention_days)
            
            deleted_count = 0
            for log_file in log_files:
                try:
                    # 从文件名提取日期
                    filename = os.path.basename(log_file)
                    date_str = filename.replace('forwarder_', '').replace('.log', '')
                    file_date = datetime.strptime(date_str, '%Y-%m-%d')
                    
                    if file_date < cutoff_date:
                        os.remove(log_file)
                        deleted_count += 1
                        
                except (ValueError, OSError):
                    continue
            
            if deleted_count > 0:
                self.log_message(f"🗑️ 已清理 {deleted_count} 个过期日志文件")
                
        except Exception as e:
            self.log_message(f"❌ 清理日志失败: {e}")
    
    # 多规则管理方法
    def refresh_rules_display(self):
        """刷新规则列表显示"""
        try:
            # 清空列表
            for item in self.rules_tree.get_children():
                self.rules_tree.delete(item)
            
            # 添加规则
            for i, rule in enumerate(self.forwarding_rules):
                sequence = str(i + 1)  # 序号从1开始
                status = "✅ 启用" if rule['enabled'] else "❌ 禁用"
                source_info = f"{rule['source']['type']}:{rule['source']['contact']}"
                target_info = f"{rule['target']['type']}:{rule['target']['contact']}"
                filter_info = rule['source']['filter_type']
                
                self.rules_tree.insert('', 'end', values=(sequence, status, source_info, target_info, filter_info))
            
            # 选中第一个规则
            if self.forwarding_rules and self.rules_tree.get_children():
                self.rules_tree.selection_set(self.rules_tree.get_children()[0])
                self.load_rule_to_ui(0)
                
        except Exception as e:
            self.log_message(f"❌ 刷新规则列表失败: {e}")
    
    def on_rule_select(self, event):
        """规则选中事件"""
        try:
            selected = self.rules_tree.selection()
            if selected:
                index = self.rules_tree.index(selected[0])
                self.selected_rule_index = index
                self.load_rule_to_ui(index)
        except Exception as e:
            self.log_message(f"❌ 选中规则失败: {e}")
    
    def load_rule_to_ui(self, index):
        """加载规则到UI"""
        try:
            if 0 <= index < len(self.forwarding_rules):
                rule = self.forwarding_rules[index]
                
                # 加载规则基本信息
                self.rule_name_var.set(rule['name'])
                self.rule_enabled_var.set(rule['enabled'])
                
                # 加载源设置
                self.rule_source_type_var.set(rule['source']['type'])
                self.rule_source_contact_var.set(rule['source']['contact'])
                self.rule_filter_type_var.set(rule['source']['filter_type'])
                self.rule_range_start_var.set(rule['source']['range_start'])
                self.rule_range_end_var.set(rule['source']['range_end'])
                
                # 加载目标设置
                self.rule_target_type_var.set(rule['target']['type'])
                self.rule_target_contact_var.set(rule['target']['contact'])
                
                # 更新联系人列表
                self.update_rule_contacts()
                self.on_rule_filter_change()
                
        except Exception as e:
            self.log_message(f"❌ 加载规则失败: {e}")
    
    def add_rule(self):
        """添加新规则"""
        try:
            rule_count = len(self.forwarding_rules)
            new_rule = {
                'id': f'rule_{rule_count + 1}',
                'name': f'规则{rule_count + 1}',
                'enabled': True,
                'source': {
                    'type': 'wechat',
                    'contact': '',
                    'filter_type': 'all',
                    'range_start': '',
                    'range_end': ''
                },
                'target': {
                    'type': 'wecom',
                    'contact': ''
                }
            }
            
            self.forwarding_rules.append(new_rule)
            self.refresh_rules_display()
            self.update_queue_filter_options()  # 更新队列过滤选项
            
            # 选中新添加的规则
            last_item = self.rules_tree.get_children()[-1]
            self.rules_tree.selection_set(last_item)
            self.selected_rule_index = len(self.forwarding_rules) - 1
            self.load_rule_to_ui(self.selected_rule_index)
            
            self.log_message(f"✅ 添加新规则: {new_rule['name']}")
            
        except Exception as e:
            self.log_message(f"❌ 添加规则失败: {e}")
    
    def edit_rule(self):
        """编辑当前选中的规则"""
        self.save_current_rule()
    
    def delete_rule(self):
        """删除选中的规则"""
        try:
            if not self.forwarding_rules:
                self.log_message("⚠️ 没有规则可删除")
                return
            
            if len(self.forwarding_rules) == 1:
                self.log_message("⚠️ 至少需要保留一条规则")
                return
            
            selected = self.rules_tree.selection()
            if selected:
                index = self.rules_tree.index(selected[0])
                rule_name = self.forwarding_rules[index]['name']
                
                result = messagebox.askyesno("确认删除", f"确定要删除规则 '{rule_name}' 吗？")
                if result:
                    del self.forwarding_rules[index]
                    self.refresh_rules_display()
                    self.update_queue_filter_options()  # 更新队列过滤选项
                    self.log_message(f"✅ 已删除规则: {rule_name}")
            else:
                self.log_message("⚠️ 请先选中要删除的规则")
                
        except Exception as e:
            self.log_message(f"❌ 删除规则失败: {e}")
    
    def toggle_rule(self):
        """切换规则的启用/禁用状态"""
        try:
            selected = self.rules_tree.selection()
            if selected:
                index = self.rules_tree.index(selected[0])
                rule = self.forwarding_rules[index]
                rule['enabled'] = not rule['enabled']
                
                status = "启用" if rule['enabled'] else "禁用"
                self.log_message(f"✅ 规则 '{rule['name']}' 已{status}")
                
                self.refresh_rules_display()
                self.update_queue_filter_options()  # 更新队列过滤选项
                # 保持选中状态
                self.rules_tree.selection_set(self.rules_tree.get_children()[index])
                self.load_rule_to_ui(index)
            else:
                self.log_message("⚠️ 请先选中要操作的规则")
                
        except Exception as e:
            self.log_message(f"❌ 切换规则状态失败: {e}")
    
    def save_current_rule(self):
        """保存当前编辑的规则"""
        try:
            if 0 <= self.selected_rule_index < len(self.forwarding_rules):
                rule = self.forwarding_rules[self.selected_rule_index]
                
                # 保存基本信息
                rule['name'] = self.rule_name_var.get() or f'规则{self.selected_rule_index + 1}'
                rule['enabled'] = self.rule_enabled_var.get()
                
                # 保存源设置
                rule['source']['type'] = self.rule_source_type_var.get()
                rule['source']['contact'] = self.rule_source_contact_var.get()
                rule['source']['filter_type'] = self.rule_filter_type_var.get()
                rule['source']['range_start'] = self.rule_range_start_var.get()
                rule['source']['range_end'] = self.rule_range_end_var.get()
                
                # 保存目标设置
                rule['target']['type'] = self.rule_target_type_var.get()
                rule['target']['contact'] = self.rule_target_contact_var.get()
                
                self.refresh_rules_display()
                self.update_queue_filter_options()  # 更新队列过滤选项
                # 保持选中状态
                if self.rules_tree.get_children():
                    self.rules_tree.selection_set(self.rules_tree.get_children()[self.selected_rule_index])
                
                self.log_message(f"✅ 已保存规则: {rule['name']}")
            
        except Exception as e:
            self.log_message(f"❌ 保存规则失败: {e}")
    
    def on_rule_source_type_change(self, event=None):
        """源类型改变事件"""
        self.update_rule_source_contacts()
    
    def on_rule_target_type_change(self, event=None):
        """目标类型改变事件"""
        self.update_rule_target_contacts()
    
    def on_rule_filter_change(self, event=None):
        """过滤类型改变事件"""
        try:
            filter_type = self.rule_filter_type_var.get()
            if filter_type == 'range':
                # 显示范围设置
                for widget in self.rule_range_frame.winfo_children():
                    widget.pack(side=tk.LEFT, padx=(0, 5))
            else:
                # 隐藏范围设置
                for widget in self.rule_range_frame.winfo_children():
                    widget.pack_forget()
        except Exception as e:
            pass
    
    def update_rule_contacts(self):
        """更新规则的联系人列表"""
        self.update_rule_source_contacts()
        self.update_rule_target_contacts()
    
    def update_rule_source_contacts(self):
        """更新源联系人列表"""
        try:
            source_type = self.rule_source_type_var.get()
            contacts = self.get_contacts(source_type)
            self.rule_source_contact_combo['values'] = contacts
        except Exception as e:
            pass
    
    def update_rule_target_contacts(self):
        """更新目标联系人列表"""
        try:
            target_type = self.rule_target_type_var.get()
            contacts = self.get_contacts(target_type)
            self.rule_target_contact_combo['values'] = contacts
        except Exception as e:
            pass
    
    def log_message(self, message, rule_id=None):
        """记录日志消息"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # 如果提供了rule_id，在消息前面添加规则ID
        if rule_id is not None:
            # 查找规则的序号
            rule_sequence = self.get_rule_sequence_by_id(rule_id)
            if rule_sequence:
                message = f"【规则{rule_sequence}】{message}"
        
        log_entry = f"[{timestamp}] {message}\n"
        
        # 如果GUI还没有初始化完成，先打印到控制台
        if not hasattr(self, 'log_text') or self.log_text is None:
            print(log_entry.strip())  # 打印到控制台
            return
        
        # 在GUI线程中更新日志
        try:
            self.root.after(0, lambda: self._update_log(log_entry))
        except Exception as e:
            print(f"Log update failed: {e}")
            print(log_entry.strip())
    
    def get_rule_sequence_by_id(self, rule_id):
        """根据规则ID获取规则序号"""
        try:
            for i, rule in enumerate(self.forwarding_rules):
                if rule.get('id') == rule_id:
                    return str(i + 1)
            return None
        except Exception:
            return None
    
    def _update_log(self, log_entry):
        """更新日志显示"""
        try:
            if hasattr(self, 'log_text') and self.log_text is not None:
                self.log_text.insert(tk.END, log_entry)
                self.log_text.see(tk.END)
        except Exception as e:
            print(f"GUI log update failed: {e}")
            print(log_entry.strip())
    
    def clear_log(self):
        """清空日志"""
        self.log_text.delete(1.0, tk.END)
    
    def start_ai_reply_detection(self, hwnd, window_title, input_x, input_y):
        """启动回复检测和反向转发"""
        def detection_worker():
            try:
                self.log_message("🔍 开始检测回复...")
                
                # 创建temp文件夹用于保存截图
                temp_dir = "temp"
                if not os.path.exists(temp_dir):
                    os.makedirs(temp_dir)
                
                # 读取配置的延迟时间
                try:
                    delay_seconds = int(self.delay_var.get())
                except:
                    delay_seconds = 2
                
                self.log_message(f"⏰ 等待 {delay_seconds} 秒后开始截图检测...")
                time.sleep(delay_seconds)
                
                # 截取第一张图
                screenshot_count = 1
                previous_image = self.capture_wecom_area(hwnd)
                if previous_image:
                    screenshot_path = os.path.join(temp_dir, f"screenshot_{screenshot_count:03d}.png")
                    previous_image.save(screenshot_path)
                    self.log_message(f"📸 保存第{screenshot_count}张截图: {screenshot_path}")
                else:
                    self.log_message("❌ 初始截图失败")
                    return
                
                # 开始5秒间隔的循环截图检测
                while self.is_forwarding:
                    time.sleep(5)  # 等待5秒
                    screenshot_count += 1
                    
                    # 截取当前图像
                    current_image = self.capture_wecom_area(hwnd)
                    if not current_image:
                        self.log_message(f"❌ 第{screenshot_count}张截图失败")
                        continue
                    
                    # 保存截图
                    screenshot_path = os.path.join(temp_dir, f"screenshot_{screenshot_count:03d}.png")
                    current_image.save(screenshot_path)
                    
                    # 比较图像是否相同
                    is_identical = self.compare_images(previous_image, current_image)
                    
                    if is_identical:
                        self.log_message(f"✅ 第{screenshot_count}张截图与上次相同，回复完成！")
                        self.log_message(f"📸 最终截图: {screenshot_path}")
                        
                        # 回复完成，开始复制消息
                        self.copy_ai_reply_and_forward(hwnd, input_x, input_y)
                        break
                    else:
                        self.log_message(f"📸 第{screenshot_count}张截图有变化，继续监控...")
                        previous_image = current_image
                    
                    # 防止无限循环，最多检测5分钟 (60次)
                    if screenshot_count >= 60:
                        self.log_message("⏰ 检测超时（5分钟），停止回复检测")
                        # 超时时也要标记消息完成
                        self.handle_detection_timeout()
                        break
                        
            except Exception as e:
                self.log_message(f"❌ 回复检测出错: {e}")
                # 异常时也要标记消息完成
                self.handle_detection_error(str(e))
        
        # 在新线程中运行检测
        import threading
        detection_thread = threading.Thread(target=detection_worker, daemon=True)
        detection_thread.start()
    
    def handle_detection_timeout(self):
        """处理异步检测超时的情况"""
        try:
            if self.message_queue and self.message_queue.processing_message:
                processing_message = self.message_queue.processing_message
                self.message_queue.mark_message_completed(processing_message, "复制检测超时", success=False)
                self.log_message("⚠️ 异步检测超时，已标记消息完成")
        except Exception as e:
            self.log_message(f"处理检测超时失败: {e}")
    
    def handle_detection_error(self, error_msg):
        """处理异步检测错误的情况"""
        try:
            if self.message_queue and self.message_queue.processing_message:
                processing_message = self.message_queue.processing_message
                self.message_queue.mark_message_completed(processing_message, f"复制检测异常: {error_msg}", success=False)
                self.log_message("⚠️ 异步检测异常，已标记消息完成")
        except Exception as e:
            self.log_message(f"处理检测错误失败: {e}")
    
    def capture_wecom_area(self, hwnd, region_ratio=None):
        """使用屏幕截图方式截取企业微信整个窗口（用于回复完成检测）"""
        try:
            # 设置DPI感知，避免截图缩放问题
            try:
                windll.shcore.SetProcessDpiAwareness(1)  # PROCESS_DPI_AWARE
            except:
                try:
                    windll.user32.SetProcessDPIAware()  # 旧版本Windows
                except:
                    pass  # 如果都失败就忽略
            
            self.log_message(f"📸 开始屏幕截图 - 窗口句柄: {hwnd}")
            
            # 检查窗口状态
            if not win32gui.IsWindow(hwnd):
                self.log_message("❌ 窗口句柄无效")
                return None
                
            if not win32gui.IsWindowVisible(hwnd):
                self.log_message("❌ 窗口不可见")
                return None
            
            # 激活窗口
            try:
                if win32gui.IsIconic(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    time.sleep(0.3)
                
                win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
                time.sleep(0.2)
                
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(0.2)
                
                win32gui.BringWindowToTop(hwnd)
                time.sleep(0.5)
                
                current_foreground = win32gui.GetForegroundWindow()
                if current_foreground == hwnd:
                    self.log_message("✅ 窗口已激活")
                else:
                    self.log_message(f"⚠ 窗口可能未完全激活")
                    
            except Exception as e:
                self.log_message(f"⚠ 窗口激活出错: {e}")
            
            # 使用UIAutomation获取准确的窗口信息
            try:
                import uiautomation as auto
                
                # 通过句柄获取UIAutomation控件
                window_control = auto.ControlFromHandle(hwnd)
                if window_control:
                    # 获取窗口的边界矩形
                    ui_rect = window_control.BoundingRectangle
                    self.log_message(f"📐 UIAutomation窗口边界: ({ui_rect.left},{ui_rect.top},{ui_rect.right},{ui_rect.bottom})")
                    
                    # 使用UIAutomation的边界
                    rect = (ui_rect.left, ui_rect.top, ui_rect.right, ui_rect.bottom)
                    window_width = ui_rect.right - ui_rect.left
                    window_height = ui_rect.bottom - ui_rect.top
                    
                    self.log_message(f"✅ UIAutomation窗口信息: 位置{rect}, 大小{window_width}x{window_height}")
                else:
                    # 如果UIAutomation失败，使用win32gui作为备选
                    rect = win32gui.GetWindowRect(hwnd)
                    window_width = rect[2] - rect[0]
                    window_height = rect[3] - rect[1]
                    self.log_message(f"⚠ UIAutomation失败，使用win32gui: 位置{rect}, 大小{window_width}x{window_height}")
                    
            except Exception as e:
                # 如果导入UIAutomation失败，使用win32gui
                rect = win32gui.GetWindowRect(hwnd)
                window_width = rect[2] - rect[0]
                window_height = rect[3] - rect[1]
                self.log_message(f"⚠ UIAutomation不可用({e})，使用win32gui: 位置{rect}, 大小{window_width}x{window_height}")
            
            if window_width <= 0 or window_height <= 0:
                self.log_message("❌ 窗口大小无效")
                return None
            
            # 直接截取整个窗口（已验证正确）
            full_window_img = ImageGrab.grab(bbox=(rect[0], rect[1], rect[2], rect[3]))
            
            self.log_message(f"✅ 企业微信窗口截图成功: {full_window_img.size}")
            return full_window_img
                
        except Exception as e:
            self.log_message(f"❌ 屏幕截图失败: {e}")
            import traceback
            self.log_message(f"详细错误: {traceback.format_exc()}")
            return None
    
    def compare_images(self, img1, img2):
        """比较两张图像是否完全相同"""
        try:
            if img1.size != img2.size:
                return False
            
            # 转换为灰度图像进行比较
            gray1 = img1.convert('L')
            gray2 = img2.convert('L')
            
            # 计算像素数据的哈希
            pixels1 = list(gray1.getdata())
            pixels2 = list(gray2.getdata())
            
            # 完全相同返回True
            return pixels1 == pixels2
        
        except Exception as e:
            self.log_message(f"图像比较失败: {e}")
            return False
    
    def copy_ai_reply_and_forward(self, hwnd, input_x, input_y):
        """复制AI回复并转发到普通微信（多规则系统适配）"""
        try:
            self.log_message("📋 开始复制回复消息...")
            
            # 获取当前正在处理的消息和规则
            if not self.message_queue or not self.message_queue.processing_message:
                self.log_message("⚠️ 没有正在处理的消息")
                return
            
            processing_message = self.message_queue.processing_message
            rule = processing_message.get('matched_rule')
            if not rule:
                self.log_message("⚠️ 无法获取消息对应的规则")
                return
            
            rule_id = rule['id']
            target_contact = rule['target']['contact']
            
            # 从配置文件加载复制坐标
            try:
                with open('forwarder_config.json', 'r', encoding='utf-8') as f:
                    config = json.load(f)
                
                copy_coords = config.get('copy_coordinates', {}).get(target_contact, None)
                if not copy_coords:
                    self.log_message(f"❌ 未找到 {target_contact} 的复制坐标配置，请先设置")
                    messagebox.showerror("错误", f"未找到 {target_contact} 的复制坐标配置\n请点击'设置复制坐标'按钮进行设置")
                    return
                
                right_click_offset = copy_coords['right_click']
                copy_click_offset = copy_coords['copy_click']
                
                self.log_message(f"📍 加载复制坐标配置:")
                self.log_message(f"   右键偏移: {right_click_offset}")
                self.log_message(f"   复制偏移: {copy_click_offset}")
                
            except Exception as e:
                self.log_message(f"❌ 加载复制坐标配置失败: {e}")
                messagebox.showerror("错误", f"加载复制坐标配置失败: {e}")
                return
            
            # 激活企业微信窗口
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.3)
            
            # 获取窗口位置，计算基准点（左下角）
            window_rect = win32gui.GetWindowRect(hwnd)
            base_x = window_rect[0]  # 窗口左边界
            base_y = window_rect[3]  # 窗口下边界（左下角）
            
            # 计算右键点击的绝对坐标
            right_click_x = base_x + right_click_offset[0]
            right_click_y = base_y - right_click_offset[1]  # Y轴反向
            
            self.log_message(f"🎯 右键点击坐标: ({right_click_x}, {right_click_y})")
            
            # 移动鼠标并右键点击
            win32api.SetCursorPos((right_click_x, right_click_y))
            time.sleep(0.3)
            
            # 右键点击
            win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
            win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
            time.sleep(0.5)  # 等待右键菜单出现
            
            self.log_message(f"🖱 右键菜单已弹出")
            
            # 计算复制按钮的绝对坐标
            copy_x = base_x + copy_click_offset[0]
            copy_y = base_y - copy_click_offset[1]  # Y轴反向
            
            self.log_message(f"📋 复制按钮坐标: ({copy_x}, {copy_y})")
            
            # 移动鼠标并点击复制
            win32api.SetCursorPos((copy_x, copy_y))
            time.sleep(0.2)
            
            # 左键点击复制按钮
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
            time.sleep(0.3)
            
            self.log_message(f"✅ 已点击复制按钮")
            
            # 等待复制完成
            time.sleep(0.5)
            
            # 使用多规则系统转发复制的内容
            self.log_message("📋 复制完成，开始转发到目标联系人...")
            success = self.forward_copied_reply_to_target(rule)
            
            # 标记消息处理完成
            if success:
                self.message_queue.mark_message_completed(processing_message, "AI回复已复制并转发", success=True)
            else:
                self.message_queue.mark_message_completed(processing_message, "复制转发失败", success=False)
            
        except Exception as e:
            self.log_message(f"❌ 复制回复失败: {e}")
            # 标记消息处理失败
            if hasattr(self, 'message_queue') and self.message_queue and self.message_queue.processing_message:
                self.message_queue.mark_message_completed(self.message_queue.processing_message, f"复制回复异常: {e}", success=False)
    
    def forward_copied_reply_to_target(self, rule):
        """将复制的AI回复转发到目标联系人（多规则系统）"""
        try:
            target_type = rule['target']['type']
            target_contact = rule['target']['contact']
            rule_id = rule['id']
            
            self.log_message(f"📤 准备转发到 {target_type}:{target_contact}", rule_id)
            
            if target_type == "wechat":
                # 转发到普通微信
                if not self.wechat:
                    self.log_message("❌ 微信实例不存在", rule_id)
                    return False
                
                # 获取剪贴板内容
                import win32clipboard
                try:
                    win32clipboard.OpenClipboard()
                    clipboard_content = win32clipboard.GetClipboardData()
                    win32clipboard.CloseClipboard()
                    
                    if not clipboard_content:
                        self.log_message("❌ 剪贴板内容为空", rule_id)
                        return False
                    
                    # 发送到普通微信
                    self.wechat.SendMsg(clipboard_content, who=target_contact)
                    self.log_message(f"✅ 已转发AI回复到微信: {target_contact}", rule_id)
                    return True
                    
                except Exception as e:
                    self.log_message(f"❌ 获取剪贴板内容失败: {e}", rule_id)
                    return False
                    
            elif target_type == "wecom":
                # 转发到企业微信
                success = self.send_clipboard_to_wecom_window(target_contact)
                if success:
                    self.log_message(f"✅ 已转发AI回复到企业微信: {target_contact}", rule_id)
                else:
                    self.log_message(f"❌ 转发AI回复到企业微信失败: {target_contact}", rule_id)
                return success
            else:
                self.log_message(f"❌ 不支持的目标类型: {target_type}", rule_id)
                return False
                
        except Exception as e:
            self.log_message(f"❌ 转发复制回复失败: {e}", rule.get('id'))
            return False
    
    def send_clipboard_to_wecom_window(self, window_title):
        """将剪贴板内容发送到企业微信窗口"""
        try:
            import win32gui
            import win32con
            import win32api
            import win32clipboard
            import time
            
            # 找到企业微信窗口
            hwnd = self.find_wecom_chat_window(window_title)
            if not hwnd:
                self.log_message(f"❌ 未找到企业微信窗口: {window_title}")
                return False
            
            # 激活窗口
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.3)
            
            # 直接粘贴剪贴板内容
            win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
            win32api.keybd_event(ord('V'), 0, 0, 0)
            win32api.keybd_event(ord('V'), 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
            time.sleep(0.5)
            
            # 发送消息
            win32api.keybd_event(win32con.VK_RETURN, 0, 0, 0)
            win32api.keybd_event(win32con.VK_RETURN, 0, win32con.KEYEVENTF_KEYUP, 0)
            time.sleep(0.3)
            
            return True
            
        except Exception as e:
            self.log_message(f"❌ 发送剪贴板内容到企业微信失败: {e}")
            return False
    
    def forward_ai_reply_to_wechat(self):
        """旧版单规则系统的方法，已废弃，仅保留兼容性"""
        try:
            self.log_message("⚠️ 旧版转发方法已废弃，请使用新的多规则系统")
            return
            
            # 以下代码仅保留兼容性，不会执行
            # 获取源联系人名称（从规则中获取第一个启用的规则）
            enabled_rules = [rule for rule in self.forwarding_rules if rule.get('enabled', True)]
            if not enabled_rules:
                self.log_message("⚠️ 没有启用的规则")
                return
            
            source_contact = enabled_rules[0]['source']['contact']
            
            # 直接在当前监听的聊天窗口粘贴，因为聊天窗口已经打开
            # 不能使用搜索，因为搜索会清空剪贴板中的AI回复内容
            
            if not self.wechat:
                self.log_message("❌ 微信实例不存在")
                return
            
            self.log_message("📋 直接在聊天窗口粘贴AI回复...")
            
            # 由于正在监听该联系人，聊天窗口应该已经是活跃状态
            # 直接发送粘贴的内容，不调用ChatWith（避免搜索操作）
            try:
                # 获取剪贴板内容
                import win32clipboard
                win32clipboard.OpenClipboard()
                ai_reply_content = win32clipboard.GetClipboardData()
                win32clipboard.CloseClipboard()
                
                self.log_message(f"📄 获取到AI回复内容: {ai_reply_content[:50]}...")
                
                # 记录这条AI回复，避免被再次转发
                self.record_ai_reply(ai_reply_content)
                
                # 直接使用wxauto发送消息到当前聊天窗口
                # 不使用ChatWith方法，避免搜索操作
                self.wechat.SendMsg(ai_reply_content, who=source_contact)
                
                self.log_message(f"✅ AI回复已直接发送到: {source_contact}")
                
            except Exception as clipboard_error:
                self.log_message(f"❌ 剪贴板操作失败: {clipboard_error}，尝试备用方法")
                
                # 备用方法：直接键盘操作粘贴
                # 确保聊天窗口处于活跃状态
                time.sleep(0.3)
                
                # 使用Ctrl+V粘贴
                win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
                win32api.keybd_event(ord('V'), 0, 0, 0)
                win32api.keybd_event(ord('V'), 0, win32con.KEYEVENTF_KEYUP, 0)
                win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
                time.sleep(0.3)
                
                # 发送回车
                win32api.keybd_event(win32con.VK_RETURN, 0, 0, 0)
                win32api.keybd_event(win32con.VK_RETURN, 0, win32con.KEYEVENTF_KEYUP, 0)
                time.sleep(0.2)
                
                self.log_message(f"✅ AI回复已通过键盘粘贴发送到: {source_contact}")
            
        except Exception as e:    
            self.log_message(f"❌ 转发AI回复失败: {e}")
            return False
    
    def send_to_wechat_window(self, chat_name, ai_reply, sender):
        """直接查找并发送消息到微信独立聊天窗口"""
        try:
            import win32gui
            import win32con
            import win32clipboard
            from wxauto.utils.win32 import GetAllWindows
            import time
            
            # 查找所有窗口
            all_windows = GetAllWindows()
            target_hwnd = None
            
            # 查找微信聊天窗口
            for hwnd, class_name, window_title in all_windows:
                # 检查是否是微信聊天窗口
                if self.is_wechat_chat_window(hwnd, class_name, window_title, chat_name):
                    target_hwnd = hwnd
                    self.log_message(f"✅ 找到微信聊天窗口: {window_title} (hwnd: {hwnd})")
                    break
            
            if not target_hwnd:
                self.log_message(f"❌ 未找到微信聊天窗口: {chat_name}")
                return False
            
            # 激活窗口
            win32gui.SetForegroundWindow(target_hwnd)
            time.sleep(0.5)
            
            # 构建回复内容
            if sender != chat_name:  # 群聊情况
                reply_content = f"@{sender} {ai_reply}"
                self.log_message(f"📝 群聊回复内容: @{sender} [AI回复内容]")
            else:
                reply_content = ai_reply
                self.log_message(f"📝 私聊回复内容: [AI回复内容]")
            
            # 将内容复制到剪贴板
            self.set_clipboard_text(reply_content)
            
            # 查找输入框并粘贴内容
            success = self.paste_to_wechat_input(target_hwnd)
            if success:
                # 发送消息 (Enter)
                import win32api
                win32api.keybd_event(win32con.VK_RETURN, 0, 0, 0)
                win32api.keybd_event(win32con.VK_RETURN, 0, win32con.KEYEVENTF_KEYUP, 0)
                time.sleep(0.2)
                
                self.log_message(f"✅ 消息已发送到微信窗口: {chat_name}")
                return True
            else:
                self.log_message(f"❌ 无法在微信窗口中找到输入框")
                return False
                
        except Exception as e:
            self.log_message(f"❌ 发送到微信窗口失败: {e}")
            import traceback
            self.log_message(f"详细错误: {traceback.format_exc()}")
            return False
    
    def is_wechat_chat_window(self, hwnd, class_name, window_title, chat_name):
        """判断是否是微信聊天窗口"""
        try:
            # 检查类名是否匹配微信聊天窗口
            wechat_chat_classes = [
                'ChatWnd',           # 普通微信聊天窗口
                'WeChatMainWnd',     # 微信主窗口
                'WeUIDialog',        # 微信对话框
            ]
            
            if class_name in wechat_chat_classes:
                # 检查窗口标题是否包含聊天名称
                if chat_name in window_title:
                    # 验证是否属于微信进程
                    try:
                        import win32process
                        import win32api
                        thread_id, process_id = win32process.GetWindowThreadProcessId(hwnd)
                        process_handle = win32api.OpenProcess(win32process.PROCESS_QUERY_INFORMATION, False, process_id)
                        try:
                            process_name = win32process.GetModuleFileNameEx(process_handle, 0)
                            if 'WeChat.exe' in process_name or 'wechat.exe' in process_name:
                                return True
                        finally:
                            win32api.CloseHandle(process_handle)
                    except:
                        # 如果无法验证进程，但类名和标题匹配，也认为有效
                        return True
            
            return False
            
        except Exception as e:
            self.log_message(f"验证微信窗口失败: {e}")
            return False
    
    def paste_to_wechat_input(self, hwnd):
        """在微信窗口中找到输入框并粘贴内容"""
        try:
            import win32gui
            import win32con
            import win32api
            import time
            
            # 首先尝试直接使用全局粘贴快捷键
            self.log_message("📝 尝试使用全局粘贴快捷键...")
            
            # 确保窗口激活
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.3)
            
            # 直接使用Ctrl+V在激活窗口中粘贴
            win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
            win32api.keybd_event(ord('V'), 0, 0, 0)
            win32api.keybd_event(ord('V'), 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
            time.sleep(0.3)
            
            self.log_message("✅ 已使用全局粘贴快捷键")
            return True
            
        except Exception as e:
            self.log_message(f"❌ 全局粘贴失败: {e}")
            
            # 备用方法：查找输入框控件
            try:
                self.log_message("🔍 尝试查找输入框控件...")
                
                # 查找输入框控件
                input_controls = []
                
                def enum_child_proc(child_hwnd, lparam):
                    try:
                        class_name = win32gui.GetClassName(child_hwnd)
                        # 扩大微信输入框可能的类名
                        input_classes = ['Edit', 'RichEdit', 'RichEdit20W', 'RichEdit50W', 'RichEdit20A', 'RichEdit20WPT']
                        
                        if class_name in input_classes:
                            # 检查控件是否可见
                            if win32gui.IsWindowVisible(child_hwnd):
                                rect = win32gui.GetWindowRect(child_hwnd)
                                width = rect[2] - rect[0]
                                height = rect[3] - rect[1]
                                
                                # 输入框通常有一定的宽度和高度
                                if width > 50 and height > 15:
                                    input_controls.append((child_hwnd, class_name, width, height))
                                    self.log_message(f"🔍 找到控件: {class_name} ({width}x{height})")
                    except:
                        pass
                    return True
                
                win32gui.EnumChildWindows(hwnd, enum_child_proc, None)
                
                if input_controls:
                    # 选择最大的控件作为输入框
                    target_input = max(input_controls, key=lambda x: x[2] * x[3])[0]
                    
                    self.log_message(f"✅ 选中输入框控件: {target_input}")
                    
                    # 点击输入框获取焦点
                    rect = win32gui.GetWindowRect(target_input)
                    center_x = (rect[0] + rect[2]) // 2
                    center_y = (rect[1] + rect[3]) // 2
                    
                    # 使用鼠标点击
                    win32api.SetCursorPos((center_x, center_y))
                    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
                    time.sleep(0.2)
                    
                    # 粘贴
                    win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
                    win32api.keybd_event(ord('V'), 0, 0, 0)
                    win32api.keybd_event(ord('V'), 0, win32con.KEYEVENTF_KEYUP, 0)
                    win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
                    time.sleep(0.2)
                    
                    self.log_message("✅ 已粘贴内容到找到的输入框")
                    return True
                else:
                    self.log_message("❌ 未找到任何输入框控件")
                    return False
            
            except Exception as e2:
                self.log_message(f"❌ 备用方法也失败: {e2}")
                return False
    
    def set_clipboard_text(self, text):
        """设置剪贴板文本"""
        try:
            import win32clipboard
            import win32con
            
            # 确保文本是字符串类型
            if not isinstance(text, str):
                text = str(text)
            
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            
            # 使用Unicode格式设置剪贴板
            win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, text)
            win32clipboard.CloseClipboard()
            
            self.log_message(f"✅ 已复制内容到剪贴板: {text[:30]}...")
        except Exception as e:
            self.log_message(f"❌ 设置剪贴板失败: {e}")
            # 备用方法：使用tkinter剪贴板
            try:
                self.root.clipboard_clear()
                self.root.clipboard_append(text)
                self.log_message(f"✅ 使用备用方法设置剪贴板成功")
            except Exception as e2:
                self.log_message(f"❌ 备用剪贴板方法也失败: {e2}")
    
    def record_ai_reply(self, content):
        """记录AI回复内容，避免循环转发"""
        try:
            # 添加到最近回复列表
            self.recent_ai_replies.append(content.strip())
            
            # 保持列表大小，移除最旧的记录
            if len(self.recent_ai_replies) > self.max_recent_replies:
                self.recent_ai_replies.pop(0)
                
            self.log_message(f"📝 已记录AI回复内容（当前记录数: {len(self.recent_ai_replies)}）")
            
        except Exception as e:
            self.log_message(f"记录AI回复失败: {e}")
    
    def is_recent_ai_reply(self, content):
        """检查消息是否是最近的AI回复"""
        try:
            content_stripped = content.strip()
            
            # 检查是否与最近的AI回复完全匹配
            for recent_reply in self.recent_ai_replies:
                if content_stripped == recent_reply:
                    return True
            
            # 检查是否与最近的AI回复高度相似（防止格式微调）
            for recent_reply in self.recent_ai_replies:
                if len(content_stripped) > 20 and len(recent_reply) > 20:
                    # 计算相似性（简单的字符匹配）
                    similarity = len(set(content_stripped) & set(recent_reply)) / len(set(content_stripped) | set(recent_reply))
                    if similarity > 0.8:  # 80%相似度
                        return True
            
            return False
            
        except Exception as e:
            self.log_message(f"检查AI回复相似性失败: {e}")
            return False
    
    def get_process_name(self, pid):
        """获取进程名称"""
        try:
            import psutil
            process = psutil.Process(pid)
            return process.name()
        except:
            return ""
    
    def setup_copy_coordinates(self):
        """设置复制坐标功能（多规则系统适配）"""
        try:
            # 获取当前选中的规则
            if not hasattr(self, 'selected_rule_index') or self.selected_rule_index < 0 or self.selected_rule_index >= len(self.forwarding_rules):
                messagebox.showerror("错误", "请先选中一个规则")
                return
            
            rule = self.forwarding_rules[self.selected_rule_index]
            target_contact = rule['target']['contact']
            
            if not target_contact:
                messagebox.showerror("错误", "请先为选中的规则设置企业微信联系人")
                return
            
            if rule['target']['type'] != 'wecom':
                messagebox.showerror("错误", "只能为企业微信目标设置复制坐标")
                return
            
            # 查找企业微信窗口
            from wxauto.utils.win32 import FindWindow, GetAllWindows
            hwnd = FindWindow(name=target_contact)
            if not hwnd:
                self.log_message(f"❌ 未找到企业微信窗口: {target_contact}")
                
                # 列出所有可能的企业微信窗口
                all_windows = GetAllWindows()
                wecom_windows = []
                for window_hwnd, class_name, window_title in all_windows:
                    if target_contact in window_title or "企业微信" in window_title:
                        wecom_windows.append((window_hwnd, class_name, window_title))
                
                if wecom_windows:
                    hwnd, class_name, window_title = wecom_windows[0]
                    self.log_message(f"🎯 使用窗口: {window_title} (句柄:{hwnd})")
                else:
                    messagebox.showerror("错误", f"未找到企业微信窗口: {target_contact}")
                    return
            
            # 激活企业微信窗口
            win32gui.SetForegroundWindow(hwnd)
            win32gui.BringWindowToTop(hwnd)
            time.sleep(0.5)
            
            # 获取窗口左下角坐标作为基准点
            window_rect = win32gui.GetWindowRect(hwnd)
            base_x = window_rect[0]  # 窗口左边界
            base_y = window_rect[3]  # 窗口下边界（左下角）
            
            self.log_message(f"📍 窗口左下角基准点: ({base_x}, {base_y})")
            
            # 启动坐标设置流程
            self.coordinate_setup_window(hwnd, base_x, base_y, target_contact)
            
        except Exception as e:
            self.log_message(f"❌ 设置复制坐标异常: {e}")
            messagebox.showerror("错误", f"设置复制坐标异常: {e}")
    
    def coordinate_setup_window(self, hwnd, base_x, base_y, target_contact):
        """坐标设置窗口"""
        # 创建新窗口
        setup_window = tk.Toplevel(self.root)
        setup_window.title("设置复制坐标")
        setup_window.geometry("450x600")  # 高度增加一倍
        setup_window.transient(self.root)
        setup_window.grab_set()
        
        # 存储坐标
        coordinates = {"right_click": None, "copy_click": None}
        
        # 说明文本
        instruction_frame = ttk.Frame(setup_window)
        instruction_frame.pack(fill=tk.X, padx=20, pady=20)
        
        instruction_text = f"""设置 {target_contact} 的复制坐标
        
窗口左下角基准点: ({base_x}, {base_y})

操作步骤:
1. 点击"开始设置复制坐标"按钮
2. 在企业微信聊天窗口的回复消息上点右键
3. 在弹出的右键菜单中点击"复制"按钮
4. 程序会自动保存坐标并提示成功"""
        
        ttk.Label(instruction_frame, text=instruction_text, justify=tk.LEFT, wraplength=400).pack()
        
        # 状态显示
        status_frame = ttk.Frame(setup_window)
        status_frame.pack(fill=tk.X, padx=20, pady=10)
        
        status_var = tk.StringVar(value="等待开始...")
        status_label = ttk.Label(status_frame, textvariable=status_var, foreground="blue")
        status_label.pack(pady=10)
        
        # 坐标显示
        coord_frame = ttk.Frame(setup_window)
        coord_frame.pack(fill=tk.X, padx=20, pady=10)
        
        right_click_var = tk.StringVar(value="右键坐标: 未设置")
        copy_click_var = tk.StringVar(value="复制坐标: 未设置")
        
        ttk.Label(coord_frame, textvariable=right_click_var).pack(anchor=tk.W, pady=2)
        ttk.Label(coord_frame, textvariable=copy_click_var).pack(anchor=tk.W, pady=2)
        
        def start_coordinate_setup():
            status_var.set("请在回复消息上点右键，然后点击复制...")
            setup_window.withdraw()  # 隐藏设置窗口
            
            def monitor_mouse_clicks():
                import time
                step = 1  # 1=等待右键, 2=等待左键
                
                while step <= 2:
                    time.sleep(0.05)  # 更频繁的检查
                    
                    if step == 1:  # 等待右键点击
                        if win32api.GetAsyncKeyState(win32con.VK_RBUTTON) & 0x8000:
                            # 等待右键释放
                            while win32api.GetAsyncKeyState(win32con.VK_RBUTTON) & 0x8000:
                                time.sleep(0.01)
                            
                            # 获取右键点击位置
                            current_pos = win32gui.GetCursorPos()
                            relative_x = current_pos[0] - base_x
                            relative_y = base_y - current_pos[1]  # Y轴反向
                            
                            coordinates["right_click"] = (relative_x, relative_y)
                            
                            # 更新显示
                            setup_window.after(0, lambda: (
                                right_click_var.set(f"右键坐标: ({relative_x}, {relative_y})"),
                                status_var.set("右键坐标已记录，请点击复制按钮...")
                            ))
                            
                            step = 2  # 进入下一步
                            
                    elif step == 2:  # 等待左键点击（复制按钮）
                        if win32api.GetAsyncKeyState(win32con.VK_LBUTTON) & 0x8000:
                            # 等待左键释放
                            while win32api.GetAsyncKeyState(win32con.VK_LBUTTON) & 0x8000:
                                time.sleep(0.01)
                            
                            # 获取复制按钮点击位置
                            current_pos = win32gui.GetCursorPos()
                            relative_x = current_pos[0] - base_x
                            relative_y = base_y - current_pos[1]  # Y轴反向
                            
                            coordinates["copy_click"] = (relative_x, relative_y)
                            
                            # 更新显示并自动保存
                            setup_window.after(0, lambda: (
                                copy_click_var.set(f"复制坐标: ({relative_x}, {relative_y})"),
                                status_var.set("正在保存坐标..."),
                                setup_window.deiconify()  # 显示窗口
                            ))
                            
                            # 自动保存
                            time.sleep(0.5)  # 稍等一下让用户看到状态
                            setup_window.after(0, auto_save_coordinates)
                            break
                    
                    # ESC键取消
                    if win32api.GetAsyncKeyState(win32con.VK_ESCAPE) & 0x8000:
                        setup_window.after(0, lambda: (
                            setup_window.deiconify(),
                            status_var.set("设置已取消")
                        ))
                        break
            
            import threading
            threading.Thread(target=monitor_mouse_clicks, daemon=True).start()
        
        def auto_save_coordinates():
            """自动保存坐标"""
            if not coordinates["right_click"] or not coordinates["copy_click"]:
                status_var.set("坐标不完整，保存失败")
                return
            
            try:
                # 加载现有配置
                try:
                    with open('forwarder_config.json', 'r', encoding='utf-8') as f:
                        config = json.load(f)
                except:
                    config = {}
                
                # 添加复制坐标配置
                if 'copy_coordinates' not in config:
                    config['copy_coordinates'] = {}
                
                config['copy_coordinates'][target_contact] = {
                    'right_click': coordinates["right_click"],
                    'copy_click': coordinates["copy_click"],
                    'window_class': win32gui.GetClassName(hwnd)
                }
                
                # 保存配置
                with open('forwarder_config.json', 'w', encoding='utf-8') as f:
                    json.dump(config, f, ensure_ascii=False, indent=2)
                
                # 记录日志
                self.log_message(f"✅ 复制坐标已自动保存到配置文件")
                self.log_message(f"   目标联系人: {target_contact}")
                self.log_message(f"   右键坐标: {coordinates['right_click']}")
                self.log_message(f"   复制坐标: {coordinates['copy_click']}")
                
                # 验证保存是否成功
                try:
                    with open('forwarder_config.json', 'r', encoding='utf-8') as f:
                        verify_config = json.load(f)
                    if 'copy_coordinates' in verify_config and target_contact in verify_config['copy_coordinates']:
                        self.log_message(f"✅ 配置验证成功：复制坐标已正确保存")
                    else:
                        self.log_message(f"⚠️ 配置验证失败：复制坐标未找到")
                except Exception as verify_error:
                    self.log_message(f"⚠️ 配置验证失败: {verify_error}")
                
                # 显示成功提示
                status_var.set("坐标保存成功！")
                messagebox.showinfo("保存成功", f"复制坐标已成功保存！\n\n右键坐标: {coordinates['right_click']}\n复制坐标: {coordinates['copy_click']}")
                setup_window.destroy()
                
            except Exception as e:
                self.log_message(f"❌ 自动保存坐标失败: {e}")
                status_var.set(f"保存失败: {e}")
                messagebox.showerror("保存失败", f"保存坐标配置失败: {e}")
        
        # 开始设置按钮
        button_frame = ttk.Frame(setup_window)
        button_frame.pack(fill=tk.X, padx=20, pady=20)
        
        ttk.Button(button_frame, text="开始设置复制坐标", command=start_coordinate_setup).pack(pady=10)
        
        # 取消按钮
        ttk.Button(button_frame, text="取消", command=setup_window.destroy).pack(pady=5)
    
    def run(self):
        """运行程序"""
        self.log_message("微信消息转发助手已启动")
        self.root.mainloop()

def main():
    """主函数"""
    try:
        # 确保只有一个主窗口
        import sys
        import tkinter as tk
        
        # 检查是否已有Tk实例
        try:
            tk._default_root.quit()
            tk._default_root.destroy()
        except:
            pass
        
        # 创建应用程序
        app = WeChatMessageForwarder()
        
        # 确保窗口正确显示
        app.root.update_idletasks()
        app.root.deiconify()  # 确保窗口显示
        app.root.focus_force()  # 强制获取焦点
        
        # 运行程序
        app.run()
        
    except Exception as e:
        import traceback
        print(f"程序启动失败: {e}")
        print(traceback.format_exc())
        input("按回车键退出...")

if __name__ == "__main__":
    main()