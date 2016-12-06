Attribute VB_Name = "FCN_Transparencia"
Option Explicit
''''Always on top''''''''
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2 '

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
''''Always on top''''''''

'Declaración del Api SetLayeredWindowAttributes que establece _
 la transparencia al form

Private Declare Function SetLayeredWindowAttributes Lib "user32" _
                (ByVal hwnd As Long, _
                 ByVal crKey As Long, _
                 ByVal bAlpha As Byte, _
                 ByVal dwFlags As Long) As Long


'Recupera el estilo de la ventana
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
                (ByVal hwnd As Long, _
                 ByVal nIndex As Long) As Long


'Declaración del Api SetWindowLong necesaria para aplicar un estilo _
 al form antes de usar el Api SetLayeredWindowAttributes

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
               (ByVal hwnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long


Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000
'Función para saber si formulario ya es transparente. _
 Se le pasa el Hwnd del formulario en cuestión

Public Function Is_Transparent(ByVal hwnd As Long) As Boolean
On Error Resume Next

Dim Msg As Long

    Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
       
       If (Msg And WS_EX_LAYERED) = WS_EX_LAYERED Then
          Is_Transparent = True
       Else
          Is_Transparent = False
       End If

    If Err Then
       Is_Transparent = False
    End If

End Function

'Función que aplica la transparencia, se le pasa el hwnd del form y un valor de 0 a 255
Public Function Aplicar_Transparencia(ByVal hwnd As Long, _
                                      Valor As Integer) As Long

Dim Msg As Long

On Error Resume Next

If Valor < 0 Or Valor > 255 Then
   Aplicar_Transparencia = 1
Else
   Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
   Msg = Msg Or WS_EX_LAYERED
   
   SetWindowLong hwnd, GWL_EXSTYLE, Msg
   
   'Establece la transparencia
   SetLayeredWindowAttributes hwnd, 0, Valor, LWA_ALPHA

   Aplicar_Transparencia = 0

End If


If Err Then
   Aplicar_Transparencia = 2
End If

End Function







