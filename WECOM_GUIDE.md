# 企业微信支持 - 快速开始指南

## 概述

wxauto 现在支持企业微信（WeCom）！您可以使用相同的API来操作企业微信，包括发送消息、监听消息、文件传输等功能。

## 安装和设置

### 1. 确保企业微信已安装并登录

- 下载并安装企业微信客户端
- 登录您的企业微信账号
- 确保企业微信窗口可见且处于活动状态

### 2. 检测企业微信窗口

运行检测脚本来确认企业微信窗口可以被正确识别：

```bash
python detect_wecom.py
```

### 3. 运行测试

运行测试脚本来验证所有功能是否正常：

```bash
python test_wecom.py
```

## 基本使用

### 导入模块

```python
from wxauto import WeCom, WeComChat
```

### 初始化企业微信

```python
# 基本初始化
wecom = WeCom()

# 开启调试模式
wecom = WeCom(debug=True)
```

### 发送消息

```python
# 发送文本消息
wecom.SendMsg("你好，这是测试消息", who="联系人姓名")

# 发送文件
wecom.SendFiles("文件路径.txt", who="联系人姓名")
```

### 获取消息

```python
# 获取当前聊天窗口的所有消息
msgs = wecom.GetAllMessage()
for msg in msgs:
    print(f"发送者: {msg.sender}")
    print(f"内容: {msg.content}")
    print(f"类型: {msg.type}")
    print(f"时间: {msg.time}")

# 获取新消息
new_msgs = wecom.GetNewMessage()
```

### 监听消息

```python
def message_handler(msg, chat):
    """消息处理函数"""
    print(f"收到来自 {chat.who} 的消息: {msg.content}")
    
    # 自动回复
    if msg.content == "你好":
        chat.SendMsg("你好！我是自动回复")

# 添加监听
wecom.AddListenChat(nickname="联系人姓名", callback=message_handler)

# 保持运行
wecom.KeepRunning()
```

## 高级功能

### 聊天窗口操作

```python
# 切换到指定聊天
wecom.ChatWith("联系人姓名")

# 获取聊天信息
chat_info = wecom.ChatInfo()
print(f"聊天类型: {chat_info.get('chat_type')}")
print(f"聊天名称: {chat_info.get('chat_name')}")

# 获取群成员（如果是群聊）
if chat_info.get('chat_type') == 'group':
    members = wecom.GetGroupMembers()
    print(f"群成员: {members}")
```

### 会话管理

```python
# 获取所有会话
sessions = wecom.GetSession()
for session in sessions:
    print(f"会话: {session.nickname}")

# 获取子窗口
sub_windows = wecom.GetAllSubWindow()
print(f"当前有 {len(sub_windows)} 个子窗口")
```

### 文件下载

```python
def download_handler(msg, chat):
    if msg.type in ('image', 'video', 'file'):
        # 下载文件
        file_path = msg.download()
        print(f"文件已下载到: {file_path}")

wecom.AddListenChat(nickname="联系人姓名", callback=download_handler)
```

## 与个人微信同时使用

```python
from wxauto import WeChat, WeCom

# 同时初始化
personal_wx = WeChat()
work_wx = WeCom()

# 分别操作
personal_wx.SendMsg("个人消息", who="朋友")
work_wx.SendMsg("工作消息", who="同事")

# 分别监听
def personal_handler(msg, chat):
    print(f"个人微信消息: {msg.content}")

def work_handler(msg, chat):
    print(f"企业微信消息: {msg.content}")

personal_wx.AddListenChat(nickname="朋友", callback=personal_handler)
work_wx.AddListenChat(nickname="同事", callback=work_handler)
```

## 常见问题

### Q: 企业微信窗口无法识别？

A: 请检查：
1. 企业微信是否已启动并登录
2. 企业微信版本是否支持（建议使用最新版本）
3. 运行 `detect_wecom.py` 脚本查看详细信息

### Q: 消息发送失败？

A: 请检查：
1. 联系人姓名是否正确
2. 是否有权限发送消息给该联系人
3. 企业微信是否处于活动状态

### Q: 监听消息不工作？

A: 请检查：
1. 回调函数是否正确定义
2. 是否调用了 `KeepRunning()` 方法
3. 企业微信窗口是否保持打开状态

### Q: 如何调试问题？

A: 
1. 开启调试模式：`WeCom(debug=True)`
2. 查看控制台输出的调试信息
3. 运行测试脚本：`python test_wecom.py`

## 示例项目

查看 `wecom_example.py` 文件获取完整的使用示例，包括：
- 基本消息收发
- 文件传输
- 消息监听
- 聊天窗口操作

## 注意事项

1. **权限要求**: 确保程序有足够的权限操作企业微信窗口
2. **版本兼容**: 建议使用最新版本的企业微信客户端
3. **稳定性**: 企业微信的UI可能会随版本更新而变化，如遇问题请及时反馈
4. **合规使用**: 请遵守企业微信的使用条款和相关法律法规

## 技术支持

如果遇到问题，请：
1. 首先运行测试脚本确认问题
2. 查看调试日志
3. 在项目仓库提交Issue，并提供详细的错误信息和环境信息