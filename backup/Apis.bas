Attribute VB_Name = "MApis"
Option Explicit

' Point struct for ClientToScreen
Private Type PointAPI
    x As Long
    y As Long
End Type

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "KERNEL32" Alias "GetComputerNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Sub InvalidateRect Lib "user32" (ByVal hWnd As Long, ByVal t As Long, ByVal bErase As Long)
Public Declare Sub ValidateRect Lib "user32" (ByVal hWnd As Long, ByVal t As Long)
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetModuleHandle Lib "KERNEL32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Private Const MF_BYPOSITION = &H400&

'constantes para el richbox
Private Const EM_UNDO = &HC7
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_GETFIRSTVISIBLELINE = &HCE
Private Const EM_CANUNDO = &HC6
Private Const EM_GETLINE = &HC4
Private Const EM_GETMODIFY = &HB8
Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_CHARFROMPOS = &HD7
Private Const EM_LINEINDEX = &HBB
Private Const EM_LINELENGTH = &HC1
Private Const EM_SETMODIFY = &HB9

Public Declare Function DeleteFile Lib "KERNEL32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function MoveFile Lib "KERNEL32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Private Declare Function CreateDirectory Lib "KERNEL32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1
Private Const SM_CYCAPTION = 4

Public Declare Function IsDebuggerPresent Lib "KERNEL32" () As Long
Public Declare Sub ClipCursor Lib "user32" (lpRect As Any)
Public Declare Function LockWindowUpdate& Lib "user32" (ByVal hWndLock&)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd&, _
                                              ByVal wMsg&, ByVal wParam&, lParam As Any) As Long
Public Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hWnd&, _
                                                                          ByVal nIndex&)
Public Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hWnd&, _
                                                        ByVal nIndex&, ByVal dwNewLong&)

Private Declare Function ClientToScreen& Lib "user32" (ByVal hWnd&, lpPoint As PointAPI)
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetFileTime Lib "KERNEL32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "KERNEL32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "KERNEL32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function GetSystemDirectory Lib "KERNEL32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function OpenFile Lib "KERNEL32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetSaveFileName Lib "COMDLG32.DLL" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function GetOpenFileName Lib "COMDLG32.DLL" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, lpCursorName As Any) As Long
Private Declare Function CreateFile Lib "KERNEL32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileSize Lib "KERNEL32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function CloseHandle Lib "KERNEL32" (ByVal hObject As Long) As Long
Private Declare Function GetTempFileName Lib "KERNEL32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function SetFileAttributes Lib "KERNEL32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function GetTempPath Lib "KERNEL32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function ExcludeClipRect Lib "gdi32" (ByVal hdc As Long, _
    ByVal X1 As Long, ByVal y1 As Long, ByVal x2 As Long, _
    ByVal y2 As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, _
    ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" _
    (ByVal hObject As Long) As Long
Public Declare Function GetClipRgn Lib "gdi32" (ByVal hdc As Long, _
    ByVal hRgn As Long) As Long
Public Declare Function OffsetClipRgn Lib "gdi32" (ByVal hdc As Long, _
    ByVal x As Long, ByVal y As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" _
(ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpDefault As String, _
ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long 'Optional parameter
    lpClass As String 'Optional parameter
    hkeyClass As Long 'Optional parameter
    dwHotKey As Long 'Optional parameter
    hIcon As Long 'Optional parameter
    hProcess As Long 'Optional parameter
End Type

Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" _
        (SEI As SHELLEXECUTEINFO) As Long

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

'obtener el nombre del computador
Public Function GetComputer() As String

    Dim ret As String
    
    ret = Space$(255)
    
    Call GetComputerName(ret, 255)
    
    ret = StripNulls(ret)
    
    GetComputer = ret
    
End Function


'obtener el usuario
Public Function GetUser() As String

    Dim ret As String
    
    ret = Space$(255)
    
    Call GetUserName(ret, 255)
    
    ret = StripNulls(ret)
    
    GetUser = ret
    
End Function


Public Function LineLen(CharPos As Long)
    'Returns the number of character of the line that
    'contains the character position specified by CharPos
    LineLen = SendMessage(frmMain.rtbCodigo.hWnd, EM_LINELENGTH, CharPos, 0&)
End Function
'obtener lineas desde richbox
Public Function LineCount() As Long
    'Returns the number of lines in the textbox
    LineCount = SendMessage(frmMain.rtbCodigo.hWnd, EM_GETLINECOUNT, 0&, 0&)
End Function
Public Function TopLine() As Long
    'Returns the zero based line index of the first
    'visible line in a multiline textbox.
    'Or the position of the first visible character
    'in a none multiline textbox
    TopLine = SendMessage(frmMain.rtbCodigo.hWnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)
End Function

Public Function CanUndo() As Boolean
    'Returns True if it's possible to make an Undo
    Dim lngRetVal As Long

    lngRetVal = SendMessage(frmMain.rtbCodigo.hWnd, EM_CANUNDO, 0&, 0&)
    CanUndo = (lngRetVal <> 0)
End Function
Public Sub Undo()
    'Undo the last edit
    SendMessage frmMain.rtbCodigo.hWnd, EM_UNDO, 0&, 0&
End Sub
Public Function GetLine(LineIndex As Long) As String
    'Returns the text contained at the specified line
    Dim bArray() As Byte 'byte array to contain the returned string
    Dim lngLineLen As Long 'the length of the line
    Dim sRetVal As String 'the return value
    
    'Check the LineIndex value
    If LineIndex >= LineCount Then
      GetLine = ""
      Exit Function
    End If
    'get the length of the line
    lngLineLen = LineLen(GetCharFromLine(LineIndex))
    If lngLineLen < 1 Then
      GetLine = ""
      Exit Function
    End If
    ReDim bArray(lngLineLen + 1)
    'The first word of the array must contain
    'the length of the line to return
    bArray(0) = lngLineLen And 255
    bArray(1) = lngLineLen \ 256
    SendMessage frmMain.rtbCodigo.hWnd, EM_GETLINE, LineIndex, bArray(0)
    'convert the byte array into a string
    sRetVal = Left(StrConv(bArray, vbUnicode), lngLineLen)
    'return the string
    GetLine = sRetVal
End Function
Public Function GetCharFromLine(LineIndex As Long)
    'Returns the index of the first character of the line
    'check if LineIndex is valid
    If LineIndex < LineCount Then
        GetCharFromLine = SendMessage(frmMain.rtbCodigo.hWnd, EM_LINEINDEX, LineIndex, 0&)
    End If
End Function
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' This function will return true if we are running in the IDE (development) mode else it returns false.
'
' Great for enableling error interception code, eg:
'   If Not InDevelopmentMode Then On Error GoTo ErrorHandler
'
Function InDevelopmentMode() As Boolean
   InDevelopmentMode = Not CBool(GetModuleHandle(App.EXEName))
End Function


Public Function CreaDirectorio(ByVal Path As String) As Boolean
    
    Dim Security As SECURITY_ATTRIBUTES
    Dim ret As Long
    
    ret = CreateDirectory(Path, Security)
    
    If ret = 0 Then
        CreaDirectorio = False
    Else
        CreaDirectorio = True
    End If
    
End Function




Sub CenterWindow(ByVal hWnd As Long)

    Dim wRect As RECT
    
    Dim x As Integer
    Dim y As Integer

    Dim ret As Long
    
    ret = GetWindowRect(hWnd, wRect)
    
    x = (GetSystemMetrics(SM_CXSCREEN) - (wRect.Right - wRect.Left)) / 2
    y = (GetSystemMetrics(SM_CYSCREEN) - (wRect.Bottom - wRect.Top + GetSystemMetrics(SM_CYCAPTION))) / 2
    
    ret = SetWindowPos(hWnd, vbNull, x, y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER)
    
End Sub

'remueve la x
Public Sub RemoveMenus(frm As Form, _
    remove_restore As Boolean, _
    remove_move As Boolean, _
    remove_size As Boolean, _
    remove_minimize As Boolean, _
    remove_maximize As Boolean, _
    remove_seperator As Boolean, _
    remove_close As Boolean)
Dim hMenu As Long
    
    ' Get the form's system menu handle.
    hMenu = GetSystemMenu(frm.hWnd, False)
    
    If remove_close Then DeleteMenu hMenu, 6, MF_BYPOSITION
    If remove_seperator Then DeleteMenu hMenu, 5, MF_BYPOSITION
    If remove_maximize Then DeleteMenu hMenu, 4, MF_BYPOSITION
    If remove_minimize Then DeleteMenu hMenu, 3, MF_BYPOSITION
    If remove_size Then DeleteMenu hMenu, 2, MF_BYPOSITION
    If remove_move Then DeleteMenu hMenu, 1, MF_BYPOSITION
    If remove_restore Then DeleteMenu hMenu, 0, MF_BYPOSITION
End Sub

Public Function LeeIni(ByVal Seccion As String, ByVal LLave As String, ByVal ArchivoIni As String) As String

    Dim lRet As Long
    Dim ret As String
    
    Dim Buffer As String
    
    Buffer = String$(255, " ")
    
    lRet = GetPrivateProfileString(Seccion, LLave, "", Buffer, Len(Buffer), ArchivoIni)
    
    Buffer = Trim$(Buffer)
    ret = Left$(Buffer, Len(Buffer) - 1)
    
    LeeIni = ret
    
End Function

Public Sub GrabaIni(ByVal ArchivoIni As String, ByVal Seccion As String, ByVal LLave As String, ByVal Valor)

    Dim ret As Long
    
    ret = WritePrivateProfileString(Seccion, LLave, CStr(Valor), ArchivoIni)
    
End Sub


Public Sub Shell_Email()

    On Local Error Resume Next
    ShellExecute frmMain.hWnd, vbNullString, "mailto:lnunez@vbsoftware.cl", vbNullString, "C:\", SW_SHOWMAXIMIZED
    Err = 0
    
End Sub
Public Sub Shell_PaginaWeb()

    On Local Error Resume Next
    ShellExecute frmMain.hWnd, vbNullString, "http://www.vbsoftware.cl/", vbNullString, "C:\", SW_SHOWMAXIMIZED
    Err = 0
    
End Sub


Public Sub Hourglass(hWnd As Long, fOn As Boolean)

    If fOn Then
        Call SetCapture(hWnd)
        Call SetCursor(LoadCursor(0, ByVal IDC_WAIT))
    Else
        Call ReleaseCapture
        Call SetCursor(LoadCursor(0, IDC_ARROW))
    End If
    DoEvents
    
End Sub
Public Function VBOpenFile(ByVal Archivo As String) As Boolean

    Dim ret As Boolean
    Dim lRet As Long
    Dim of As OFSTRUCT
    
    ret = False
    
    lRet = OpenFile(Archivo, of, OF_EXIST)
    
    If of.nErrCode = 0 Then ret = True
    
    VBOpenFile = ret
    
End Function

