Attribute VB_Name = "FCN_MoveMouse"
Option Explicit


' estructura POINTAPI para las coordenadas
Private Type POINTAPI
    x As Long
    y As Long
End Type
Dim Anterior As POINTAPI
' declaración Api GetCursorPos
Private Declare Function GetCursorPos Lib "user32" ( _
    lpPoint As POINTAPI) As Long


Public Function SeMueve() As Boolean
    ' variable para copiar las coordenadas actuales
    Dim Mouse As POINTAPI
    ' obtiene las coordenadas
    GetCursorPos Mouse
    ' compara los valores actuales con los almacenados
    If (Mouse.x <> Anterior.x) Or (Mouse.y <> Anterior.y) Then
        ' Retorna True cuando el mouse se mueve
        Anterior.x = Mouse.x
        Anterior.y = Mouse.y
        SeMueve = True
    Else
        'Retorna False cuando No se mueve
        SeMueve = False
    End If
End Function

