from .wx import (
    WeChat, 
    Chat
)
from .wecom import (
    WeCom,
    WeComChat
)
from .param import WxParam
import pythoncom

pythoncom.CoInitialize()

__all__ = [
    'WeChat',
    'Chat',
    'WeCom',
    'WeComChat',
    'WxParam'
]