from .ui.wecom import (
    WeComMainWnd,
    WeComSubWnd
)
from .wx import Chat, Listener
from .param import (
    WxResponse, 
    WxParam, 
    PROJECT_NAME
)
from .logger import wxlog
from typing import (
    Union, 
    List,
    Dict,
    Callable,
    TYPE_CHECKING
)

if TYPE_CHECKING:
    from wxauto.msgs.base import Message
    from wxauto.ui.sessionbox import SessionElement


class WeComChat(Chat):
    """企业微信聊天窗口实例"""

    def __init__(self, core: WeComSubWnd=None):
        self.core = core
        self.who = self.core.nickname

    def __repr__(self):
        return f'<{PROJECT_NAME} - WeCom{self.__class__.__name__} object("{self.core.nickname}")>'


class WeCom(WeComChat, Listener):
    """企业微信主窗口实例"""

    def __init__(
            self, 
            debug: bool=False,
            **kwargs
        ):
        hwnd = None
        if 'hwnd' in kwargs:
            hwnd = kwargs['hwnd']
        self.core = WeComMainWnd(hwnd)
        self.nickname = self.core.nickname
        self.listen = {}
        if debug:
            wxlog.set_debug(True)
            wxlog.debug('WeCom Debug mode is on')
        self._listener_start()
        self.Show()

    def __repr__(self):
        return f'<{PROJECT_NAME} - WeCom object("{self.nickname}")>'

    def _get_listen_messages(self):
        """获取监听消息 - 企业微信版本"""
        try:
            import sys
            sys.stdout.flush()
        except:
            pass
        temp_listen = self.listen.copy()
        for who in temp_listen:
            chat, callback = temp_listen.get(who, (None, None))
            try:
                if chat is None or not chat.core.exists():
                    wxlog.debug(f"企业微信窗口 {who} 已关闭，移除监听")
                    self.RemoveListenChat(who, close_window=False)
                    continue
            except:
                continue
            with self._lock:
                msgs = chat.GetNewMessage()
                for msg in msgs:
                    wxlog.debug(f"[企业微信 {msg.attr} {msg.type}]获取到新消息：{who} - {msg.content}")
                    chat.Show()
                    self._safe_callback(callback, msg, chat)

    def AddListenChat(
            self,
            nickname: str,
            callback: Callable[['Message', str], None],
        ) -> WxResponse:
        """添加监听聊天，将聊天窗口独立出去形成Chat对象子窗口，用于监听
        
        Args:
            nickname (str): 要监听的聊天对象
            callback (Callable[[Message, str], None]): 回调函数，参数为(Message对象, 聊天名称)
        """
        if nickname in self.listen:
            return WxResponse.failure('该聊天已监听')
        subwin = self.core.open_separate_window(nickname)
        if subwin is None:
            return WxResponse.failure('找不到聊天窗口')
        name = subwin.nickname
        chat = WeComChat(subwin)
        self.listen[name] = (chat, callback)
        return chat

    def GetSubWindow(self, nickname: str) -> 'WeComChat':
        """获取子窗口实例
        
        Args:
            nickname (str): 要获取的子窗口的昵称
            
        Returns:
            WeComChat: 子窗口实例
        """
        if subwin := self.core.get_sub_wnd(nickname):
            return WeComChat(subwin)
        
    def GetAllSubWindow(self) -> List['WeComChat']:
        """获取所有子窗口实例
        
        Returns:
            List[WeComChat]: 所有子窗口实例
        """
        return [WeComChat(subwin) for subwin in self.core.get_all_sub_wnds()]

    def open_separate_window(self, keywords: str) -> 'WeComSubWnd':
        """打开独立聊天窗口"""
        return self.core.open_separate_window(keywords)

    def get_all_sub_wnds(self):
        """获取所有企业微信子窗口"""
        return self.core.get_all_sub_wnds()

    def get_sub_wnd(self, who: str, timeout: int=0):
        """获取企业微信子窗口"""
        return self.core.get_sub_wnd(who, timeout)