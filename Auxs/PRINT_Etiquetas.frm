VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form PRINT_Etiquetas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impresión de etiquetas"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid lista 
      Height          =   3375
      Left            =   240
      TabIndex        =   9
      Top             =   3240
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   5953
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      FormatString    =   "Codigo                     | Producto                                        | Etiquetas | Precio  "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   6720
      ScaleHeight     =   1095
      ScaleWidth      =   4815
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6360
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.ComboBox cmbEtqueta 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1320
      Width           =   3735
   End
   Begin VB.CommandButton cmBoton 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   360
      Picture         =   "PRINT_Etiquetas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton cmBoton 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   2160
      Picture         =   "PRINT_Etiquetas.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6840
      Width           =   1695
   End
   Begin VB.TextBox txtInfo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Para cambiar el número de etiquetas a imprimr posicione sobre la casilla de la columna ""Etiquetas"" y escriba el valor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1095
      Left            =   360
      TabIndex        =   10
      Top             =   2040
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Información adicional:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4320
      TabIndex        =   7
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   4320
      TabIndex        =   6
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de etiqueta:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Etiquetas a imprimir: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "PRINT_Etiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL1 As String
Dim RES1 As Recordset

Private Sub cmbEtqueta_Click()
    SQL1 = "SELECT ETQ_NOMBRE, IDETIQUETA, ETQ_VALORX, ETQ_VALORY, ETQ_LARGO, ETQ_ALTO, ETQ_VERTCL, " & _
    "ETQ_HORZT, ETQ_VALORX2, ETQ_VALORY2, if(ETQ_NOMPROD=1, 'SI', 'NO') ETQ_NOMPROD, " & _
    "IF(ETQ_CODIGO=1, 'SI', 'NO') ETQ_CODIGO, IF(ETQ_PRECIO=1, 'SI', 'NO') ETQ_PRECIO, " & _
    "IF(ETQ_LOGO=1, 'SI', 'NO') ETQ_LOGO, IF(ETQ_RAZONSOCIAL=1, 'SI', 'NO') ETQ_RAZONSOCIAL, " & _
    "IF(ETQ_SUCURSAL=1, 'SI', 'NO') ETQ_SUCURSAL FROM CAT_ETIQUETAS " & _
    "WHERE IDETIQUETA = " & cmbEtqueta.ItemData(cmbEtqueta.ListIndex) & ""
    'MsgBox SQL1
    Set RES1 = con.Execute(SQL1)
    
    If Not RES1.EOF Then
       Label2.Caption = "Nombre: " & RES1.Fields("ETQ_NOMPROD") & vbCrLf & "Código: " & RES1.Fields("ETQ_CODIGO") & _
        vbCrLf & "Precio: " & RES1.Fields("ETQ_PRECIO") & vbCrLf & "Logo: " & RES1.Fields("ETQ_LOGO") & _
        vbCrLf & "Razon social: " & RES1.Fields("ETQ_RAZONSOCIAL") & vbCrLf & "Sucursal: " & RES1.Fields("ETQ_SUCURSAL") & _
        vbCrLf & "Horizontales: " & RES1.Fields("ETQ_HORZT") & vbCrLf & "Verticales: " & RES1.Fields("ETQ_VERTCL")
    Else
        Label2.Caption = "Sin información"
    End If

End Sub

Private Sub cmBoton_Click(Index As Integer)
If cmbEtqueta.Text <> "" And Val(txtInfo(1).Text) > 0 Then
    If Index = 0 Then
        imprimirEtiqueta
    Else
        Unload Me
    End If
Else
    MsgBox "Seleccion un tipo de etiqueta para imprimir y verifique que el número de etiquetas sea mayor a 0.", vbInformation
End If
End Sub
Private Sub imprimirEtiqueta()
    SQL1 = "SELECT ETQ_NOMBRE, IDETIQUETA, ETQ_VALORX, ETQ_VALORY, ETQ_LARGO, ETQ_ALTO, ETQ_VERTCL, " & _
    "ETQ_HORZT, ETQ_VALORX2, ETQ_VALORY2, if(ETQ_NOMPROD=1, 'SI', 'NO') ETQ_NOMPROD, " & _
    "IF(ETQ_CODIGO=1, 'SI', 'NO') ETQ_CODIGO, IF(ETQ_PRECIO=1, 'SI', 'NO') ETQ_PRECIO, " & _
    "IF(ETQ_LOGO=1, 'SI', 'NO') ETQ_LOGO, IF(ETQ_RAZONSOCIAL=1, 'SI', 'NO') ETQ_RAZONSOCIAL, " & _
    "IF(ETQ_SUCURSAL=1, 'SI', 'NO') ETQ_SUCURSAL FROM CAT_ETIQUETAS " & _
    "WHERE IDETIQUETA = " & cmbEtqueta.ItemData(cmbEtqueta.ListIndex) & ""
    Set RES1 = con.Execute(SQL1)
    If Not RES1.EOF Then

        Dim ques As String
        Dim valorx As Long
        Dim valory As Long
        Dim Alto As Long, Ancho As Long
        Dim num As Long
        Dim num1 As Long

        valorx = RES1.Fields("ETQ_VALORX")
        valory = RES1.Fields("ETQ_VALORY")
        num = 0
        num1 = 0
        Printer.KillDoc
        With lista
            For b1 = 1 To .Rows - 1
'                    num1 = num1 + 1
                    For c1 = 1 To Val(.TextMatrix(b1, 2))
                        
                        num1 = num1 + 1
                        ''''Para el nombre prod/ser
                        If RES1.Fields("ETQ_NOMPROD") = "SI" Then
                            Printer.Font = "ARIAL"
                            Printer.FontSize = 8
                            Printer.FontBold = False
                            Printer.CurrentX = valorx + 150
                            Printer.CurrentY = valory
                            If Len(.TextMatrix(b1, 1)) > 25 Then
                                Printer.Print Left(.TextMatrix(b1, 1), 25)
                            Else
                                Printer.Print .TextMatrix(b1, 1)
                            End If
                        End If
                        '''''Para el precio
                        If RES1.Fields("ETQ_PRECIO") = "SI" Then
                            'Printer.Font = "Courier New"
                            Printer.Font = "ARIAL"
                            Printer.FontSize = 8
                            Printer.FontBold = True
                            Printer.CurrentX = valorx + 150
                            'Printer.CurrentY = valory + 450 + 800
                            Printer.CurrentY = valory + 200 + 800
                            'If Len(.TextMatrix(b1, 1)) > 25 Then
                             '   Printer.Print Left(.TextMatrix(b1, 5), 25)
                            'Else
                                Printer.Print .TextMatrix(b1, 3)
                            'End If
                        End If
                        If RES1.Fields("ETQ_CODIGO") = "SI" Then
                            'Printer.Font = "Courier New"
                            Printer.Font = "ARIAL"
                            Printer.FontSize = 8
                            Printer.FontBold = True
                            Printer.CurrentX = valorx + 150
                            'Printer.CurrentY = valory + 250 + 800
                            Printer.CurrentY = valory + 0 + 800
                            'If Len(.TextMatrix(b1, 1)) > 25 Then
                             '   Printer.Print Left(.TextMatrix(b1, 5), 25)
                            'Else
                                Printer.Print .TextMatrix(b1, 0)
                            'End If
                        End If
                        '''''Para el código
                        num = num + 1
                        Alto = RES1.Fields("ETQ_ALTO")
                        Ancho = RES1.Fields("ETQ_LARGO")
                        Picture1.Picture = Picture1.Image
                        Picture1.height = Alto
                        Picture1.width = Ancho
                        Call DrawBarcode(.TextMatrix(b1, 0), Picture1)
                        Picture1.Picture = Picture1.Image

                        Printer.PaintPicture Picture1, 150 + valorx, 100 + valory + 120, 2800, 600
                        
                        If num1 >= RES1.Fields("ETQ_VERTCL") Then
                            Printer.NewPage
                            num = 0
                            valorx = RES1.Fields("ETQ_VALORX")
                            valory = RES1.Fields("ETQ_VALORY")
                            num1 = 0
                        Else
                            If num = RES1.Fields("ETQ_HORZT") Then
                                num = 0
                                valorx = RES1.Fields("ETQ_VALORX")
                                valory = valory + RES1.Fields("ETQ_VALORY2")
                            Else
                                valorx = valorx + RES1.Fields("ETQ_VALORX2")
                            End If
                        End If
                    Next c1
                'End If
            Next b1
        End With
        Printer.EndDoc
        MsgBox "Imprimiendo.", vbInformation
    End If
    

End Sub
Private Sub Form_Load()
    cargaEtiqueta
    cargaProductos
End Sub
Private Sub cargaProductos()
    lista.Rows = 1
    With FRM_Productos.ListaSel
        For b1 = 1 To .Rows - 1
            'If .TextMatrix(b1, 14) = Chr(254) Then
                lista.AddItem ""
                lista.TextMatrix(lista.Rows - 1, 0) = .TextMatrix(b1, 0)
                lista.TextMatrix(lista.Rows - 1, 1) = .TextMatrix(b1, 1)
                lista.TextMatrix(lista.Rows - 1, 2) = .TextMatrix(b1, 4)
                lista.TextMatrix(lista.Rows - 1, 3) = .TextMatrix(b1, 5)
            'End If
        Next b1
    End With
End Sub
Private Sub cargaEtiqueta()
    SQL1 = "select idetiqueta, etq_nombre from cat_Etiquetas"
    Set RES1 = con.Execute(SQL1)
    
    cmbEtqueta.Clear
    Do While Not RES1.EOF
        cmbEtqueta.AddItem RES1.Fields("etq_nombre")
        cmbEtqueta.ItemData(cmbEtqueta.ListCount - 1) = RES1.Fields("IdEtiqueta")
        RES1.MoveNext
    Loop
End Sub

Private Sub lista_KeyPress(KeyAscii As Integer)
    If lista.Col = 2 Then
        If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 13 Then
            lista.Text = lista.Text & Chr(KeyAscii)
            lista.Text = Val(lista.Text)
            calculaNum
        End If
    End If
End Sub
Private Sub calculaNum()
    'On Error Resume Next
    Dim valor As Long
    Dim b1 As Long
    valor = 0
    For b1 = 1 To lista.Rows - 1
        valor = valor + Val(lista.TextMatrix(b1, 2))
    Next b1
    
    txtInfo(1).Text = valor
End Sub
Private Sub lista_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
            lista.Text = "0"
        Case vbKeyBack
            If Len(lista.Text) > 0 Then
                lista.Text = Val(Left(lista.Text, Len(lista.Text) - 1))
                If lista.Text = "" Then
                    lista.Text = "0"
                End If
            End If
    End Select
    calculaNum
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    MsgBox KeyAscii
End Sub
