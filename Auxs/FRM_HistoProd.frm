VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRM_HistoProd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Historial de Ventas"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   18555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   18555
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   375
      Left            =   13680
      TabIndex        =   14
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton cmdAccion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exportar detalle"
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
      Index           =   7
      Left            =   15840
      Picture         =   "FRM_HistoProd.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Width           =   2655
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
      Left            =   8280
      TabIndex        =   7
      Top             =   360
      Width           =   1455
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
      Left            =   5520
      TabIndex        =   5
      Top             =   360
      Width           =   2535
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
      Left            =   2760
      TabIndex        =   3
      Top             =   360
      Width           =   2535
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
      Top             =   360
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   16200
      Top             =   240
   End
   Begin MSFlexGridLib.MSFlexGrid Lista 
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   17175
      _ExtentX        =   30295
      _ExtentY        =   14631
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   $"FRM_HistoProd.frx":058A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtFecha1 
      Height          =   375
      Index           =   0
      Left            =   9960
      TabIndex        =   10
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   111411201
      CurrentDate     =   40829
   End
   Begin MSComCtl2.DTPicker dtFecha1 
      Height          =   375
      Index           =   1
      Left            =   11880
      TabIndex        =   11
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   112721921
      CurrentDate     =   40829
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H000080FF&
      BorderWidth     =   4
      Height          =   435
      Index           =   4
      Left            =   11880
      Top             =   360
      Width           =   1725
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H000080FF&
      BorderWidth     =   4
      Height          =   435
      Index           =   3
      Left            =   9960
      Top             =   360
      Width           =   1725
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "De"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   9960
      TabIndex        =   13
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11880
      TabIndex        =   12
      Top             =   120
      Width           =   735
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H000080FF&
      BorderWidth     =   4
      Height          =   435
      Index           =   2
      Left            =   8280
      Top             =   360
      Width           =   1485
   End
   Begin VB.Label lBus 
      BackStyle       =   0  'Transparent
      Caption         =   "Folio venta"
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
      Left            =   8280
      TabIndex        =   8
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lBus 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor"
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
      Index           =   2
      Left            =   5520
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H000080FF&
      BorderWidth     =   4
      Height          =   435
      Index           =   1
      Left            =   5520
      Top             =   360
      Width           =   2565
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
      Index           =   0
      Left            =   2760
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H000080FF&
      BorderWidth     =   4
      Height          =   435
      Index           =   0
      Left            =   2760
      Top             =   360
      Width           =   2565
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
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H000080FF&
      BorderWidth     =   4
      Height          =   435
      Index           =   16
      Left            =   120
      Top             =   360
      Width           =   2445
   End
End
Attribute VB_Name = "FRM_HistoProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql1 As String
Dim RES1 As Recordset

Private Sub cmdAccion_Click(Index As Integer)
    ques = MsgBox("¿Exportar la lista a excel? ", vbYesNo + vbQuestion)
    If ques = vbYes Then
        Call exportExcel(Lista)
    End If
End Sub

Private Sub Command1_Click()
    cargaLista
End Sub

Private Sub Form_Load()
    cargaLista
    dtFecha1(0) = Date
    dtFecha1(1) = Date
End Sub
Private Sub cargaLista()
    
    Lista.Rows = 1
    
    sql1 = "SELECT * fROM VIEW_VENTASDETALLE " & _
    "WHERE UPPER(PRODUCTO) LIKE upper('%" & textBus(0).Text & "%') " & _
    "AND upper(CLIENTE) LIKE upper('%" & textBus(1).Text & "%') AND upper(USUARIO) LIKE upper('%" & textBus(2).Text & "%') AND upper(FOLIO) LIKE upper('%" & textBus(3).Text & "%') " & _
    "AND date_format(FECHAHORA_PROD, '%Y-%m-%d') >= '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND date_format(FECHAHORA_PROD, '%Y-%m-%d') <= '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' ORDER BY FECHA_HORA DESC"
    Set RES1 = con.Execute(sql1)
    Lista.Redraw = False
    Do While Not RES1.EOF
        Lista.AddItem ""
        Lista.TextMatrix(Lista.Rows - 1, 0) = RES1.Fields("codigo")
        Lista.TextMatrix(Lista.Rows - 1, 1) = RES1.Fields("producto")
        Lista.TextMatrix(Lista.Rows - 1, 2) = RES1.Fields("cantidad")
        Lista.TextMatrix(Lista.Rows - 1, 3) = FormatCurrency(RES1.Fields("precio"))
        Lista.TextMatrix(Lista.Rows - 1, 4) = FormatCurrency(RES1.Fields("descuento"))
        Lista.TextMatrix(Lista.Rows - 1, 5) = FormatCurrency(RES1.Fields("total"))
        Lista.TextMatrix(Lista.Rows - 1, 6) = RES1.Fields("fechahora_prod")
        Lista.TextMatrix(Lista.Rows - 1, 7) = RES1.Fields("folio")
        Lista.TextMatrix(Lista.Rows - 1, 8) = RES1.Fields("dias")
        Lista.TextMatrix(Lista.Rows - 1, 9) = RES1.Fields("cliente")
        Lista.TextMatrix(Lista.Rows - 1, 10) = RES1.Fields("usuario")
        If RES1.Fields("status") = "C" Then
            Lista.TextMatrix(Lista.Rows - 1, 11) = "Cancelado"
            Lista.Row = Lista.Rows - 1
            Lista.Col = 0
            Lista.CellForeColor = vbRed
            Lista.Col = 1
            Lista.CellForeColor = vbRed
            Lista.Col = 11
            Lista.CellForeColor = vbRed
        Else
            Lista.TextMatrix(Lista.Rows - 1, 11) = "En venta"
            Lista.Row = Lista.Rows - 1
            Lista.Col = 0
            Lista.CellForeColor = vbBlack
            Lista.Col = 1
            Lista.CellForeColor = vbBlack
            Lista.Col = 11
            Lista.CellForeColor = vbBlack
        End If
        
        RES1.MoveNext
    Loop
    Lista.Redraw = True

End Sub

Private Sub Lista_DblClick()
    Call ordenarLista(Lista)
End Sub

Private Sub textBus_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        cargaLista
    End If
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Lista.width = Me.width - 250
    Lista.height = Me.height - 1500
End Sub
