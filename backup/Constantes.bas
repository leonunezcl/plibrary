Attribute VB_Name = "MAppCon"
Option Explicit

Public Const C_INI = "PLIBRARY.INI"
Public Const C_RELEASE = "28/07/2002"
Public Const C_WEB_PAGE = "http://www.vbsoftware.cl"
Public Const C_WEB_PAGE_PE = "http://www.vbsoftware.cl/plibrary.html"
Public Const C_EMAIL = "lnunez@vbsoftware.cl"
Public Const EM_GETLINE = &HC4

Public Enum ShowCommands
    SW_HIDE = 0
    SW_SHOWNORMAL = 1
    sw_normal = 1
    SW_SHOWMINIMIZED = 2
    SW_SHOWMAXIMIZED = 3
    SW_MAXIMIZE = 3
    SW_SHOWNOACTIVATE = 4
    SW_SHOW = 5
    SW_MINIMIZE = 6
    SW_SHOWMINNOACTIVE = 7
    SW_SHOWNA = 8
    SW_RESTORE = 9
    SW_SHOWDEFAULT = 10
    SW_MAX = 10
End Enum

