Attribute VB_Name = "FCN_FRMCierre"
Option Explicit


'Declaraciones del api
'------------------------------------------------------

' PAra deshabilitar el menú y otros
Private Declare Function DeleteMenu Lib "user32" ( _
    ByVal hMenu As Long, _
    ByVal nPosition As Long, _
    ByVal wFlags As Long) As Long

' Obtiene el Handle al menú del sistema de la ventana
Private Declare Function GetSystemMenu Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal bRevert As Long) As Long


Private Const MF_BYPOSITION = &H400&

Public Sub bloquear_cierre(ByVal El_Formulario As Form, _
                            ByVal Menu_Cerrar As Boolean, _
                            ByVal Redimensionar As Boolean, _
                            ByVal Mover As Boolean)

Dim Hwnd_Menu As Long
    
    ' Obtiene el Hwnd del menú para usar con el Api DeleteMenu
    Hwnd_Menu = GetSystemMenu(El_Formulario.hwnd, False)
    
    ' botón Cerrar
    If Menu_Cerrar Then
       Call DeleteMenu(Hwnd_Menu, 6, MF_BYPOSITION)
    End If
    
    'Hace que la ventana no se pueda cambiar de tamaño
    If Redimensionar Then
       Call DeleteMenu(Hwnd_Menu, 2, MF_BYPOSITION)
    End If
    
    ' No permite que la ventana se pueda mover
    If Mover Then
       Call DeleteMenu(Hwnd_Menu, 1, MF_BYPOSITION)
    End If
End Sub




