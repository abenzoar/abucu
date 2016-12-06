VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FRM_Seguimiento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Segumiento de productos vendidos para servicio"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   19140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   19140
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAccion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exportar lista"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   11400
      Picture         =   "FRM_Seguimiento.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   240
      Width           =   2535
   End
   Begin VB.ComboBox cmbUser 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Left            =   16560
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox textBus 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   3
      Left            =   9120
      TabIndex        =   6
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox textBus 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   2
      Left            =   5760
      TabIndex        =   5
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox textBus 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   3480
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox textBus 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   16800
      Top             =   120
   End
   Begin MSFlexGridLib.MSFlexGrid lista 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   18975
      _ExtentX        =   33470
      _ExtentY        =   13150
      _Version        =   393216
      Cols            =   15
      FixedCols       =   0
      BackColorFixed  =   9520683
      ForeColorFixed  =   16777215
      BackColorBkg    =   15329769
      GridColor       =   16711680
      AllowUserResizing=   1
      FormatString    =   $"FRM_Seguimiento.frx":058A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H000080FF&
      BorderWidth     =   4
      Height          =   435
      Index           =   2
      Left            =   9120
      Top             =   480
      Width           =   1965
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H000080FF&
      BorderWidth     =   4
      Height          =   435
      Index           =   1
      Left            =   5760
      Top             =   480
      Width           =   3045
   End
   Begin VB.Label lBus 
      BackStyle       =   0  'Transparent
      Caption         =   "Producto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   5760
      TabIndex        =   8
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lBus 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo producto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   9120
      TabIndex        =   7
      Top             =   240
      Width           =   1935
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H000080FF&
      BorderWidth     =   4
      Height          =   435
      Index           =   0
      Left            =   3480
      Top             =   480
      Width           =   1965
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H000080FF&
      BorderWidth     =   4
      Height          =   435
      Index           =   16
      Left            =   120
      Top             =   480
      Width           =   3045
   End
   Begin VB.Label lBus 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lBus 
      BackStyle       =   0  'Transparent
      Caption         =   "Folio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3480
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
   Begin VB.Menu mn_Menu 
      Caption         =   "Menu"
      Begin VB.Menu mn_Export 
         Caption         =   "Exportar"
      End
      Begin VB.Menu mn_Salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "FRM_Seguimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL1 As String
Dim RES1 As Recordset

Private Sub cmbUser_LostFocus()
    cmbUser.Visible = False
End Sub

Private Sub cmdAccion_Click(Index As Integer)
    mn_Export_Click
    
End Sub

Private Sub Form_Load()
    carga_lista
    cargaProveedor
End Sub

Private Sub Lista_DblClick()

    Select Case lista.Col
        Case 6:
            b1 = lista.Row
            lista.Row = b1
            lista.Col = 6
            If lista.TextMatrix(b1, 6) = Chr(168) Then
                Call actualizaDatos("S", "SALIDA")
                lista.TextMatrix(b1, 6) = Chr(254)
            Else
                Call actualizaDatos("N", "SALIDA")
                lista.TextMatrix(b1, 6) = Chr(254)
                'Call actualizaDatos(lista.TextMatrix(b1, 6), lista.TextMatrix(b1, 7), "0", "ACCESO")
                lista.TextMatrix(b1, 6) = Chr(168)
            End If
        Case 9:
            b1 = lista.Row
            lista.Row = b1
            lista.Col = 9
            If lista.TextMatrix(b1, 6) = Chr(168) Then
                MsgBox "No se puede registrar la llegada si no tiene asignada una salida.", vbExclamation
            Else
                If lista.TextMatrix(b1, 9) = Chr(168) Then
                    Call actualizaDatos("S", "LLEGADA")
                    lista.TextMatrix(b1, 9) = Chr(254)
                Else
                    Call actualizaDatos("N", "LLEGADA")
                    lista.TextMatrix(b1, 9) = Chr(168)
                End If
            End If
        Case 8:
            'cargaProveedor
            b1 = lista.Row
            lista.Row = b1
            lista.Col = 8

            cmbUser.Top = lista.CellTop + lista.Top
            cmbUser.Left = lista.CellLeft + lista.Left
            'cmbUser.Height = lista.CellHeight
            cmbUser.width = lista.CellWidth
            'cmbUser.Text = lista.TextMatrix(lista.Row, lista.Col)
            cmbUser.Visible = True
            cmbUser.SetFocus
            
    End Select
End Sub

Private Sub cmbUser_Click()
    lista.TextMatrix(lista.Row, lista.Col) = cmbUser.Text
    cmbUser.Visible = False
    
    SQL1 = "SELECT T4.PERTP_PER_ID, T4.PERTP_TIPO_ID " & _
    "FROM PER_tIPO T4 " & _
    "WHERE concat(T4.PERTP_PER_ID, T4.PERTP_TIPO_ID) = '" & cmbUser.ItemData(cmbUser.ListIndex) & "'"
    'MsgBox SQL1
    Set RES1 = con.Execute(SQL1)
        
    If Not RES1.EOF Then
        lista.TextMatrix(lista.Row, 13) = RES1.Fields("PERTP_TIPO_ID")
        lista.TextMatrix(lista.Row, 14) = RES1.Fields("PERTP_PER_ID")
        Call actualizaDatos("S", "PROVEEDOR")
    End If

    
End Sub
Private Sub actualizaDatos(valor As String, tipo As String)
    
    Select Case tipo
        Case "SALIDA":
            If valor = "S" Then
                SQL1 = "UPDATE SEGUIMIENTO_PRODVENTA SET SEG_SALIDA = '" & valor & "', SEG_SALIDAFECHA = now() WHERE seg_folio = '" & lista.TextMatrix(lista.Row, 0) & "' AND SEG_PRODUCTOID = '" & lista.TextMatrix(lista.Row, 12) & "'"
            Else
                SQL1 = "UPDATE SEGUIMIENTO_PRODVENTA SET SEG_SALIDA = '" & valor & "', SEG_SALIDAFECHA = null WHERE seg_folio = '" & lista.TextMatrix(lista.Row, 0) & "' AND SEG_PRODUCTOID = '" & lista.TextMatrix(lista.Row, 12) & "'"
            End If
            con.Execute (SQL1)
        Case "LLEGADA":
            If valor = "S" Then
                SQL1 = "UPDATE SEGUIMIENTO_PRODVENTA SET SEG_LLEGADA = '" & valor & "', SEG_LLEGADAFECHA = now() WHERE seg_folio = '" & lista.TextMatrix(lista.Row, 0) & "' AND SEG_PRODUCTOID = '" & lista.TextMatrix(lista.Row, 12) & "'"
            Else
                SQL1 = "UPDATE SEGUIMIENTO_PRODVENTA SET SEG_LLEGADA = '" & valor & "', SEG_LLEGADAFECHA = NULL WHERE seg_folio = '" & lista.TextMatrix(lista.Row, 0) & "' AND SEG_PRODUCTOID = '" & lista.TextMatrix(lista.Row, 12) & "'"
            End If
             con.Execute (SQL1)
        Case "PROVEEDOR":
            If valor = "S" Then
                SQL1 = "UPDATE SEGUIMIENTO_PRODVENTA SET seg_salidaperid = '" & lista.TextMatrix(lista.Row, 14) & "', seg_salidaTipoId = '" & lista.TextMatrix(lista.Row, 13) & "', seg_salidaTipo = 'V' WHERE seg_folio = '" & lista.TextMatrix(lista.Row, 0) & "' AND SEG_PRODUCTOID = '" & lista.TextMatrix(lista.Row, 12) & "'"
'            Else
'                SQL1 = "UPDATE SEGUIMIENTO_PRODVENTA SET SEG_LLEGADA = '" & valor & "', SEG_LLEGADAFECHA = NULL WHERE seg_folio = '" & lista.TextMatrix(lista.Row, 0) & "' AND SEG_PRODUCTOID = '" & lista.TextMatrix(lista.Row, 12) & "'"
            End If
            'MsgBox SQL1
            con.Execute (SQL1)
    End Select
    
    carga_lista
End Sub

Private Sub cargaProveedor()

    SQL1 = "SELECT concat(T2.PERTP_PER_ID, T2.PERTP_TIPO_ID) USERID, CONCAT(PER_ALIAS, ' - ', PER_NOMBRE, ' ', PER_PATERNO, ' ', PER_MATERNO) PROVEEDOR " & _
    "FROM PERSONA T1, PER_TIPO T2 " & _
    "WHERE T1.PER_ID = T2.PERTP_PER_ID AND T2.PERTP_PER_TIPO = 'V'  "
    Set RES1 = con.Execute(SQL1)
    
    cmbUser.Clear
    Do While Not RES1.EOF
        cmbUser.AddItem RES1.Fields("PROVEEDOR")
        cmbUser.ItemData(cmbUser.ListCount - 1) = RES1.Fields("USERID")
        RES1.MoveNext
    Loop
    If cmbUser.ListCount > 0 Then
        cmbUser.ListIndex = 0
    End If

End Sub

Private Sub mn_Export_Click()
            ques = MsgBox("¿Exportar la lista a excel? ", vbYesNo + vbQuestion)
            If ques = vbYes Then
                Call exportExcel(lista)
            End If
End Sub

Private Sub mn_Salir_Click()
    Unload Me
End Sub

Private Sub textBus_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        carga_lista
    End If
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    lista.width = Me.width - 500
    lista.height = Me.height - 2500

End Sub

Private Sub carga_lista()

    Dim texto1 As String
    texto1 = ""
    
    If textBus(1).Text <> "" Then
        texto1 = texto1 & " AND DIAS_RESTANTES <= " & Val(textBus(1).Text) & " AND DIAS_RESTANTES > 0 "
    End If
    
    texto1 = texto1 & " ORDER BY FECHA_VENTA DESC "
    

    SQL1 = "SELECT * FROM VIEW_SEGUIMIENTO_VENDET WHERE UPPER(CLIENTE) LIKE UPPER('%" & textBus(0).Text & "%') AND " & _
    "UPPER(FOLIO) LIKE UPPER('%" & textBus(1).Text & "%') AND UPPER(PRODUCTO) LIKE UPPER('%" & textBus(2).Text & "%') AND UPPER(CODIGO) LIKE UPPER('%" & textBus(3).Text & "%')"
    Set RES1 = con.Execute(SQL1)
    
    lista.Rows = 1
    
    lista.ColWidth(12) = 0
    lista.ColWidth(13) = 0
    lista.ColWidth(14) = 0
    Do While Not RES1.EOF
        lista.AddItem ""
        lista.TextMatrix(lista.Rows - 1, 0) = RES1.Fields("FOLIO")
        lista.TextMatrix(lista.Rows - 1, 1) = RES1.Fields("CLIENTE")
        lista.TextMatrix(lista.Rows - 1, 2) = RES1.Fields("PRODUCTO")
        lista.TextMatrix(lista.Rows - 1, 3) = RES1.Fields("CODIGO")
        lista.TextMatrix(lista.Rows - 1, 4) = RES1.Fields("FECHA_VENTA")
        lista.TextMatrix(lista.Rows - 1, 7) = RES1.Fields("FECHA_SALIDA") & ""
        lista.TextMatrix(lista.Rows - 1, 8) = RES1.Fields("PROVEEDOR") & ""
        lista.TextMatrix(lista.Rows - 1, 10) = RES1.Fields("FECHA_LLEGADA") & ""
        lista.TextMatrix(lista.Rows - 1, 11) = RES1.Fields("COSTO") & ""
        lista.TextMatrix(lista.Rows - 1, 12) = RES1.Fields("PRODID") & ""
        
        lista.Row = lista.Rows - 1
        lista.Col = 5
        lista.CellAlignment = 4
        lista.CellFontName = "Wingdings"
        lista.CellFontBold = True
        lista.CellFontSize = 14
        lista.TextMatrix(lista.Rows - 1, 5) = Chr(254)
        lista.Row = lista.Rows - 1
        lista.Col = 6
        lista.CellAlignment = 4
        lista.CellFontName = "Wingdings"
        lista.CellFontBold = True
        lista.CellFontSize = 14
        If RES1.Fields("SALIDA") = "S" Then
            lista.TextMatrix(lista.Rows - 1, 6) = Chr(254)
        Else
            lista.TextMatrix(lista.Rows - 1, 6) = Chr(168)
        End If
        lista.Row = lista.Rows - 1
        lista.Col = 9
        lista.CellAlignment = 4
        lista.CellFontName = "Wingdings"
        lista.CellFontBold = True
        lista.CellFontSize = 14
        If RES1.Fields("LLEGADA") = "S" Then
            lista.TextMatrix(lista.Rows - 1, 9) = Chr(254)
        Else
            lista.TextMatrix(lista.Rows - 1, 9) = Chr(168)
        End If
                      
        RES1.MoveNext
    Loop
End Sub

