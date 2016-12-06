VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FRM_HistoClie 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Historial de ventas por cliente"
   ClientHeight    =   9675
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9675
   ScaleWidth      =   15660
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAccion 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   5280
      Picture         =   "FRM_HistoClie.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   15240
      Top             =   0
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
      Left            =   12720
      Picture         =   "FRM_HistoClie.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   2655
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
      FormatString    =   $"FRM_HistoClie.frx":0E54
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
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   2175
      _ExtentX        =   3836
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
      Format          =   64290817
      CurrentDate     =   40829
   End
   Begin MSComCtl2.DTPicker dtFecha1 
      Height          =   375
      Index           =   1
      Left            =   2640
      TabIndex        =   4
      Top             =   360
      Width           =   2175
      _ExtentX        =   3836
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
      Format          =   64028673
      CurrentDate     =   40829
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
      Left            =   120
      TabIndex        =   6
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
      Left            =   2640
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "FRM_HistoClie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL1 As String
Dim RES1 As Recordset

Private Sub cmdAccion_Click(Index As Integer)
    If Index = 0 Then
        cargaLista
    Else
        ques = MsgBox("¿Exportar la lista a excel? ", vbYesNo + vbQuestion)
        If ques = vbYes Then
            Call exportExcel(Lista)
        End If
    End If
End Sub

Private Sub Form_Load()
 dtFecha1(0) = Date
dtFecha1(1) = Date
    cargaLista
End Sub
Private Sub cargaLista()
    
    Lista.Rows = 1
    SQL1 = "SELECT T1.CLIENTE,   COUNT(*) PRODUCTOS, SUM(T1.PRECIO) TOTAL_PRECIO,  SUM(T1.dESCUENTO) DESCUENTO, SUM(T1.CANTIDAD) CANTIDAD, SUM((T1.PRECIO - T1.DESCUENTO)*T1.CANTIDAD) TOTAL_VENTA, T1.CLIE_EMAIL, T1.CLIE_TEL1, T1.CLIE_TEL2, " & _
    "(SELECT COUNT(T2.FOLIO) FROM VIEW_VENTAS T2 WHERE T2.CLIE_pERID = T1.CLIE_PERID and  date_format(T2.FECHAHORA, '%Y-%m-%d') BETWEEN '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' ) FOLIOS  " & _
    "From VIEW_VENTASDETALLE T1, VIEW_VENTAS T3 " & _
    "WHERE   T1.FOLIO = T3.FOLIO AND T3.ID_STATUS = 'P' and  date_format(T3.FECHAHORA, '%Y-%m-%d') BETWEEN '" & Format(dtFecha1(0), "yyyy-MM-dd") & "' AND '" & Format(dtFecha1(1), "yyyy-MM-dd") & "' " & _
    "GROUP BY T1.CLIENTE order by SUM(T1.PRECIO) desc"
    Set RES1 = con.Execute(SQL1)
    Lista.Redraw = False
    Do While Not RES1.EOF
        Lista.AddItem ""
        Lista.TextMatrix(Lista.Rows - 1, 0) = RES1.Fields("CLIENTE")
        Lista.TextMatrix(Lista.Rows - 1, 1) = RES1.Fields("PRODUCTOS")
        Lista.TextMatrix(Lista.Rows - 1, 2) = RES1.Fields("FOLIOS")
        Lista.TextMatrix(Lista.Rows - 1, 3) = FormatCurrency(RES1.Fields("TOTAL_PRECIO"))
        Lista.TextMatrix(Lista.Rows - 1, 4) = FormatCurrency(RES1.Fields("DESCUENTO"))
        Lista.TextMatrix(Lista.Rows - 1, 5) = FormatCurrency(RES1.Fields("CANTIDAD"))
        Lista.TextMatrix(Lista.Rows - 1, 6) = FormatCurrency(RES1.Fields("TOTAL_VENTA"))
        Lista.TextMatrix(Lista.Rows - 1, 7) = Format(dtFecha1(0), "Short Date")
        Lista.TextMatrix(Lista.Rows - 1, 8) = Format(dtFecha1(1), "Short Date")
        Lista.TextMatrix(Lista.Rows - 1, 9) = RES1.Fields("CLIE_EMAIL")
        Lista.TextMatrix(Lista.Rows - 1, 10) = RES1.Fields("CLIE_TEL1") & ""
        Lista.TextMatrix(Lista.Rows - 1, 11) = RES1.Fields("CLIE_TEL2") & ""
        
        RES1.MoveNext
    Loop
    Lista.Redraw = True

End Sub



Private Sub Lista_DblClick()
    If Lista.MouseRow = 0 Then
        Call ordenarLista(Lista)
    End If
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Lista.width = Me.width - 250
    Lista.height = Me.height - 1500
End Sub

