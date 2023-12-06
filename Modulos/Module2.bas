Attribute VB_Name = "Module2"
'Global Const WM_USER = 1024
'Global Const wm_cap_driver_connect = WM_USER + 10
'Global Const wm_cap_set_preview = WM_USER + 50
'Global Const WM_CAP_EDIT_COPY = WM_USER + 30
'Global Const COPY = 1054
'Global Const WM_CAP_SET_PREVIEWRATE = WM_USER + 52
'Global Const WM_CAP_DRIVER_DISCONNECT = WM_USER + 11
'Global Const WM_CAP_DLG_VIDEOFORMAT = WM_USER + 41
'Global Const WM_CAP_DLG_VIDEOCONFIG = WM_USER + 42
'Global Const WM_CAP_SET_SCALE = WM_USER + 53
'
''Api de 16 bits
''Declare Function SendMessage Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
''Api para crear la ventana de captura
''Declare Function capCreateCaptureWindow Lib "avicap.dll" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hwndParent As Integer, ByVal nID As Integer) As Long
''Declare Function DestroyWindow Lib "User" (ByVal hndw As Integer) As Integer
''Api para crear la ventana de captura
'
'Public Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal nID As Long) As Long
'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Public Declare Function DestroyWindow Lib "user32" (ByVal hndw As Long) As Boolean
'Public hwdc As Long
'Public startcap As Integer


