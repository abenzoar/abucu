VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRM_HistoVentas 
   Caption         =   "Resumen de Ventas"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16725
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   16725
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   7695
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   16575
      _ExtentX        =   29236
      _ExtentY        =   13573
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Total "
      TabPicture(0)   =   "FRM_HistoVentas.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lista"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Resumen por mes"
      TabPicture(1)   =   "FRM_HistoVentas.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lista2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin MSFlexGridLib.MSFlexGrid lista 
         Height          =   6855
         Left            =   -74880
         TabIndex        =   9
         Top             =   360
         Width           =   16455
         _ExtentX        =   29025
         _ExtentY        =   12091
         _Version        =   393216
         Cols            =   10
         FixedCols       =   0
         FormatString    =   $"FRM_HistoVentas.frx":0038
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid lista2 
         Height          =   6855
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   16455
         _ExtentX        =   29025
         _ExtentY        =   12091
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         FormatString    =   "Año           | Mes                               | Total                           "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   13200
      Top             =   360
   End
   Begin VB.CommandButton cmdAccion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exportar "
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
      Left            =   9360
      Picture         =   "FRM_HistoVentas.frx":011E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   375
      Left            =   8160
      TabIndex        =   6
      Top             =   360
      Width           =   495
   End
   Begin VB.ComboBox cmbProd 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   4800
      Style           =   2  'Dropdown List
      TabIndex        =   4
      ToolTipText     =   "Selecciona la marca a la que pertenece el producto, o agrega o edita las existentes"
      Top             =   360
      Width           =   3015
   End
   Begin VB.ComboBox cmbProd 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Selecciona la marca a la que pertenece el producto, o agrega o edita las existentes"
      Top             =   360
      Width           =   1335
   End
   Begin VB.ComboBox cmbProd 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Selecciona la marca a la que pertenece el producto, o agrega o edita las existentes"
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label lBus 
      BackStyle       =   0  'Transparent
      Caption         =   "Mes"
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
      Left            =   4800
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lBus 
      BackStyle       =   0  'Transparent
      Caption         =   "Año"
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
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lBus 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
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
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "FRM_HistoVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql1 As String
Dim RES1 As Recordset



Private Sub cmdAccion_Click(Index As Integer)
    ques = MsgBox("¿Exportar la lista a excel? ", vbYesNo + vbQuestion)
    If ques = vbYes Then
        If SSTab1.Tab = 0 Then
            Call exportExcel(lista)
            Else
            Call exportExcel(lista2)
        End If
    End If
End Sub

Private Sub Command1_Click()
carga_Lista
End Sub

Private Sub Form_Load()
SSTab1.Tab = 0

cmbProd(0).Clear
cmbProd(0).AddItem "Todo"
cmbProd(0).AddItem "Membresia"
cmbProd(0).AddItem "Producto"

cmbProd(2).Clear
cmbProd(2).AddItem "Enero"
cmbProd(2).AddItem "Febrero"
cmbProd(2).AddItem "Marzo"
cmbProd(2).AddItem "Abril"
cmbProd(2).AddItem "Mayo"
cmbProd(2).AddItem "Junio"
cmbProd(2).AddItem "Julio"
cmbProd(2).AddItem "Agosto"
cmbProd(2).AddItem "Septiembre"
cmbProd(2).AddItem "Octubre"
cmbProd(2).AddItem "Noviembre"
cmbProd(2).AddItem "Diciembre"

sql1 = "SELECT DISTINCT(ANIO) ANIO FROM view_REPORTEVENTAS ORDER BY ANIO DESC"
Set RES1 = con.Execute(sql1)
cmbProd(1).Clear
Do While Not RES1.EOF
    cmbProd(1).AddItem RES1.Fields("anio")
    RES1.MoveNext
Loop

carga_Lista

End Sub

Private Sub carga_Lista()
    Dim texto As String
    Dim totCantidad As Double
    Dim totTotal As Double
    Dim tipo As String
    Dim fila As Integer
    Dim fila2 As Integer
    Dim totMes As Double
    Dim mes As String
    totCantidad = 0
    totTotal = 0
    tipo = ""
    fila = 1
    fila2 = 1
    totMes = 0
    
    texto = " where anio > 0 "
    If cmbProd(0).Text <> "Todo" And cmbProd(0).Text <> "" Then
        If cmbProd(0).Text = "Membresia" Then
            texto = texto & " and upper(TIPO) = 'MEMBRESIA'"
        Else
            texto = texto & " and UPPER(TIPO) <> UPPER('MEMBRESIA') "
        End If
    End If

    If cmbProd(1).Text <> "" Then
        texto = texto & " and anio = '" & (cmbProd(1).Text) & "' "
    End If

    If cmbProd(2).Text <> "" Then
        texto = texto & " and mes = '" & (cmbProd(2).ListIndex + 1) & "' "
    End If

    sql1 = "SELECT * FROM view_REPORTEVENTAS " & texto & " order by anio desc, mes desc, tipo, producto"
    Set RES1 = con.Execute(sql1)
    
    lista.Rows = 1
    lista2.Rows = 1
    
    lista.MergeCells = flexMergeRestrictColumns
            
    Do While Not RES1.EOF
        
        lista.AddItem ""
        
        
        
        If RES1.Fields("tipo") <> tipo And lista.Rows > 1 Then
            For b1 = fila To lista.Rows - 1
                lista.TextMatrix(b1, 7) = totCantidad
                lista.TextMatrix(b1, 8) = FormatCurrency(totTotal)
            Next b1
                'lista.TextMatrix(lista.Rows - 1, 7) = totCantidad
                'lista.TextMatrix(lista.Rows - 1, 8) = FormatCurrency(totTotal)
            totCantidad = 0
            totTotal = 0
            fila = lista.Rows - 1
        End If
        If RES1.Fields("mes_2") <> mes And lista.Rows > 1 Then
            For b1 = fila2 To lista.Rows - 1
                lista.TextMatrix(b1, 9) = FormatCurrency(totMes)
            Next b1
            If lista.TextMatrix(lista.Rows - 2, 1) <> "Mes" Then
                lista2.AddItem ""
                lista2.TextMatrix(lista2.Rows - 1, 0) = lista.TextMatrix(lista.Rows - 2, 0)
                lista2.TextMatrix(lista2.Rows - 1, 1) = lista.TextMatrix(lista.Rows - 2, 1)
                lista2.TextMatrix(lista2.Rows - 1, 2) = FormatCurrency(totMes)
            End If
            totMes = 0
            fila2 = lista.Rows - 1
        End If
        
        tipo = RES1.Fields("TIPO")
        totCantidad = totCantidad + Val(RES1.Fields("cantidad"))
        totTotal = totTotal + Val(RES1.Fields("TOTAL"))
        totMes = totMes + Val(RES1.Fields("TOTAL"))
        mes = RES1.Fields("MES_2")
        
        lista.TextMatrix(lista.Rows - 1, 0) = RES1.Fields("ANIO")
        lista.TextMatrix(lista.Rows - 1, 1) = RES1.Fields("MES_2")
        lista.TextMatrix(lista.Rows - 1, 2) = RES1.Fields("TIPO")
        lista.TextMatrix(lista.Rows - 1, 3) = RES1.Fields("PRODUCTO")
        lista.TextMatrix(lista.Rows - 1, 4) = RES1.Fields("CANTIDAD")
        lista.TextMatrix(lista.Rows - 1, 5) = FormatCurrency(RES1.Fields("PRECIO"))
        lista.TextMatrix(lista.Rows - 1, 6) = FormatCurrency(RES1.Fields("TOTAL"))
        
        RES1.MoveNext
    
    Loop
    lista.MergeCol(7) = True
    lista.MergeCol(8) = True
    lista.MergeCol(9) = True


End Sub

Private Sub lista_DblClick()
    Call ordenarLista(lista)
End Sub

Private Sub lista_GotFocus()
    ConScroll lista
End Sub

Private Sub lista_LostFocus()
    SinScroll lista
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    lista.width = Me.width - 500
    lista.height = Me.height - 1700
    SSTab1.height = Me.height - 1500
    SSTab1.width = Me.width - 450
    
End Sub
