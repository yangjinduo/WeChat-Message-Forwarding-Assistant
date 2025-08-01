from .main import WeChatMainWnd, WeChatSubWnd
from wxauto.utils.win32 import FindWindow
from wxauto.logger import wxlog
from wxauto import uiautomation as uia
from .chatbox import ChatBox
from .sessionbox import SessionBox
from .navigationbox import NavigationBox
from typing import Union


class WeComMainWnd(WeChatMainWnd):
    """企业微信主窗口类"""
    _ui_cls_name: str = 'WeWorkWindow'  # 企业微信实际窗口类名
    _ui_name: str = '企业微信'
    
    # 企业微信可能的窗口类名列表
    _possible_class_names = [
        'WeWorkWindow',        # 实际的企业微信窗口类名
        'WeChatMainWndForPC',  # 可能与个人微信相同
        'WeComMainWnd',        # 企业微信专用类名
        'WorkWeChatMainWnd',   # 另一种可能的类名
        'WxWorkMainWnd',       # 微信工作版类名
    ]
    
    # 企业微信可能的窗口标题关键词
    _possible_titles = [
        '企业微信',
        'WeCom',
        'WeChat Work',
        '微信工作版'
    ]

    def __init__(self, hwnd: int = None):
        self.root = self
        self.parent = self
        
        if hwnd:
            self._setup_ui(hwnd)
        else:
            hwnd = self._find_wecom_window()
            if not hwnd:
                raise Exception(f'未找到企业微信窗口')
            self._setup_ui(hwnd)
        
        print(f'初始化企业微信成功，获取到已登录窗口：{self.nickname}')

    def _find_wecom_window(self) -> int:
        """查找企业微信窗口"""
        wxlog.debug("正在查找企业微信窗口...")
        
        # 首先尝试通过窗口标题查找
        for title_keyword in self._possible_titles:
            hwnd = FindWindow(name=title_keyword)
            if hwnd:
                wxlog.debug(f"通过标题 '{title_keyword}' 找到企业微信窗口: {hwnd}")
                return hwnd
        
        # 然后尝试通过类名查找
        for class_name in self._possible_class_names:
            hwnd = FindWindow(classname=class_name)
            if hwnd:
                # 验证是否为企业微信窗口
                try:
                    control = uia.ControlFromHandle(hwnd)
                    window_title = control.Name
                    if any(keyword in window_title for keyword in self._possible_titles):
                        wxlog.debug(f"通过类名 '{class_name}' 找到企业微信窗口: {hwnd}")
                        return hwnd
                except:
                    continue
        
        # 最后尝试遍历所有窗口查找
        from wxauto.utils.win32 import GetAllWindows
        all_windows = GetAllWindows()
        
        for hwnd, class_name, window_title in all_windows:
            # 检查窗口标题是否包含企业微信关键词
            if any(keyword in window_title for keyword in self._possible_titles):
                wxlog.debug(f"通过遍历找到企业微信窗口: {hwnd} - {window_title}")
                return hwnd
        
        return None

    def _setup_ui(self, hwnd: int):
        """设置UI控件"""
        try:
            self.HWND = hwnd
            self.control = uia.ControlFromHandle(hwnd)
            
            # 企业微信的UI结构可能与个人微信略有不同
            # 先尝试标准结构
            try:
                MainControl1 = [i for i in self.control.GetChildren() if not i.ClassName][0]
                MainControl2 = MainControl1.GetFirstChildControl()
                children = MainControl2.GetChildren()
                
                if len(children) >= 3:
                    navigation_control, sessionbox_control, chatbox_control = children[:3]
                else:
                    # 如果结构不同，尝试其他方式
                    raise Exception("UI结构不匹配")
                    
            except:
                # 尝试其他可能的UI结构
                wxlog.debug("标准UI结构不匹配，尝试其他结构...")
                children = self.control.GetChildren()
                
                # 根据控件类型和位置推断
                navigation_control = None
                sessionbox_control = None
                chatbox_control = None
                
                for child in children:
                    rect = child.BoundingRectangle
                    # 导航栏通常在左侧且较窄
                    if rect.width() < 100:
                        navigation_control = child
                    # 会话列表通常在左侧中等宽度
                    elif rect.width() < 300:
                        sessionbox_control = child
                    # 聊天框通常占据最大区域
                    else:
                        chatbox_control = child
                
                if not all([navigation_control, sessionbox_control, chatbox_control]):
                    raise Exception("无法识别企业微信UI结构")
            
            # 初始化各个组件
            self.navigation = NavigationBox(navigation_control, self)
            self.sessionbox = SessionBox(sessionbox_control, self)
            self.chatbox = ChatBox(chatbox_control, self)
            
            # 获取用户昵称
            try:
                self.nickname = self.navigation.my_icon.Name
            except:
                # 如果无法从导航栏获取，尝试从窗口标题获取
                self.nickname = self.control.Name
                
            # 设置兼容性属性
            self.NavigationBox = self.navigation.control
            self.SessionBox = self.sessionbox.control
            self.ChatBox = self.chatbox.control
            
        except Exception as e:
            wxlog.debug(f"企业微信UI初始化失败: {e}")
            raise Exception(f"企业微信UI初始化失败: {e}")

    def _lang(self, text: str) -> str:
        """语言本地化，企业微信可能有不同的文本"""
        # 企业微信特有的文本映射
        wecom_texts = {
            '微信': '企业微信',
            'WeChat': 'WeCom',
            # 可以根据需要添加更多映射
        }
        
        if text in wecom_texts:
            return wecom_texts[text]
        
        # 回退到父类的语言处理
        return super()._lang(text)

    def get_all_sub_wnds(self):
        """获取所有企业微信子窗口"""
        from wxauto.utils.win32 import GetAllWindows
        sub_wxs = []
        all_windows = GetAllWindows()
        
        # 查找企业微信子窗口
        for hwnd, class_name, window_title in all_windows:
            if class_name in WeComSubWnd._possible_class_names:
                try:
                    sub_wxs.append((hwnd, class_name))
                except:
                    continue
        
        return [WeComSubWnd(i[0], self) for i in sub_wxs]
    
    def get_sub_wnd(self, who: str, timeout: int=0):
        """获取企业微信子窗口"""
        # 尝试不同的类名查找子窗口
        for class_name in WeComSubWnd._possible_class_names:
            hwnd = FindWindow(classname=class_name, name=who, timeout=timeout)
            if hwnd:
                return WeComSubWnd(hwnd, self)
        return None
        
    def open_separate_window(self, keywords: str) -> 'WeComSubWnd':
        """打开独立聊天窗口"""
        if subwin := self.get_sub_wnd(keywords):
            wxlog.debug(f"{keywords} 获取到已存在的企业微信子窗口: {subwin}")
            return subwin
        self._show()
        if nickname := self.sessionbox.switch_chat(keywords):
            wxlog.debug(f"{keywords} 切换到企业微信聊天窗口: {nickname}")
            if subwin := self.get_sub_wnd(nickname):
                wxlog.debug(f"{nickname} 获取到已存在的企业微信子窗口: {subwin}")
                return subwin
            else:
                keywords = nickname
        if result := self.sessionbox.open_separate_window(keywords):
            find_nickname = result['data'].get('nickname', keywords)
            return WeComSubWnd(find_nickname, self)


class WeComSubWnd(WeChatSubWnd):
    """企业微信子窗口类"""
    _ui_cls_name: str = 'ChatWnd'  # 企业微信子窗口类名可能相同
    
    # 企业微信可能的子窗口类名
    _possible_class_names = [
        'ChatWnd',
        'WeComChatWnd',
        'WorkWeChatChatWnd',
    ]

    def __init__(
            self, 
            key: Union[str, int], 
            parent: 'WeComMainWnd', 
            timeout: int = 3
        ):
        self.root = self
        self.parent = parent
        
        if isinstance(key, str):
            hwnd = self._find_wecom_chat_window(key, timeout)
        else:
            hwnd = key
            
        if not hwnd:
            raise Exception(f"未找到企业微信聊天窗口: {key}")
            
        self.control = uia.ControlFromHandle(hwnd)
        if self.control is not None:
            try:
                chatbox_control = self.control.PaneControl(ClassName='', searchDepth=1)
                self.chatbox = ChatBox(chatbox_control, self)
                self.nickname = self.control.Name
            except Exception as e:
                wxlog.debug(f"企业微信子窗口初始化失败: {e}")
                raise

    def _find_wecom_chat_window(self, name: str, timeout: int) -> int:
        """查找企业微信聊天窗口"""
        from wxauto.utils.win32 import FindWindow
        
        # 尝试不同的类名
        for class_name in self._possible_class_names:
            hwnd = FindWindow(classname=class_name, name=name, timeout=timeout)
            if hwnd:
                return hwnd
        
        return None