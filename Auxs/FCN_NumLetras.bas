Attribute VB_Name = "FCN_NumLetras"
Option Explicit
Public Sub Numeros2(keyascii As Integer)
    If keyascii >= 48 And keyascii <= 57 Or keyascii = 13 Or keyascii = 8 Or keyascii = 27 Then
    Else
        keyascii = 0
    End If
End Sub

Public Sub Numeros(num)
    If num >= 48 And num <= 57 Or num = 13 Or num = 8 Or num = 27 Then
    Else
        num = 0
    End If
End Sub
Public Sub NumerosPunto(num)
    If num >= 48 And num <= 57 Or num = 13 Or num = 8 Or num = 27 Or num = 46 Then
    Else
        num = 0
    End If
End Sub

Public Sub Mayusculas(num)
    If num >= 97 And num <= 122 Or num = 241 Then
        num = num - 32
    End If
End Sub

Public Sub NumerosPuntoMenos(num)
    If num >= 48 And num <= 57 Or num = 13 Or num = 8 Or num = 27 Or num = 46 Or num = 45 Then
    Else
        num = 0
    End If
End Sub

