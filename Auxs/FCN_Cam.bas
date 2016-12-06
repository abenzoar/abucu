Attribute VB_Name = "FCN_Cam"
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal nID As Long) As Long

Public mCapHwnd As Long

Public Const CONNECT As Long = 1034
Public Const DISCONNECT As Long = 1035
Public Const GET_FRAME As Long = 1084
Public Const COPY As Long = 1054

Dim P() As Long
Dim POn() As Boolean

'Dim inten As Integer
'Dim Tolerance As Integer

Dim i As Integer, j As Integer

Dim Ri As Long, Wo As Long
Dim RealRi As Long

Dim c As Long, C2 As Long

Dim R As Integer, G As Integer, b As Integer
Dim R2 As Integer, G2 As Integer, b2 As Integer

Dim Tppx As Single, Tppy As Single

Dim RealMov As Integer

Dim Counter As Integer

Public Declare Function GetTickCount Lib "kernel32" () As Long
Dim LastTime As Long


