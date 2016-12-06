VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form BUSQ_ProdSer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Búsqueda"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15465
   Icon            =   "BUSQ_ProdSer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   15465
   StartUpPosition =   1  'CenterOwner
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
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   3015
   End
   Begin VB.TextBox textBus 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   3480
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
   Begin VB.ComboBox cmbTipo 
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
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11640
      TabIndex        =   3
      Text            =   "30"
      Top             =   720
      Width           =   615
   End
   Begin MSFlexGridLib.MSFlexGrid lista 
      Height          =   5175
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   9128
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      AllowUserResizing=   1
      FormatString    =   $"BUSQ_ProdSer.frx":058A
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
      Left            =   240
      TabIndex        =   11
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lBus 
      BackStyle       =   0  'Transparent
      Caption         =   "Clave producto"
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
      Left            =   3480
      TabIndex        =   10
      Top             =   360
      Width           =   1335
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H000080FF&
      BorderWidth     =   4
      Height          =   435
      Index           =   15
      Left            =   3480
      Top             =   720
      Width           =   1965
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H000080FF&
      BorderWidth     =   4
      Height          =   435
      Index           =   16
      Left            =   240
      Top             =   720
      Width           =   3045
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H000080FF&
      BorderWidth     =   4
      Height          =   435
      Index           =   0
      Left            =   5640
      Top             =   720
      Width           =   3645
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   3
      Left            =   12360
      TabIndex        =   9
      Top             =   4680
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Imagen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   12720
      TabIndex        =   8
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Image imgFoto 
      BorderStyle     =   1  'Fixed Single
      Height          =   2295
      Index           =   0
      Left            =   12720
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sub tipo "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5640
      TabIndex        =   7
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Registros en la lista: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   6600
      Width           =   4455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Núm registros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11640
      TabIndex        =   5
      Top             =   360
      Width           =   1455
   End
   Begin VB.Image Image2 
      Height          =   9855
      Index           =   1
      Left            =   0
      Picture         =   "BUSQ_ProdSer.frx":065D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   17655
   End
End
Attribute VB_Name = "BUSQ_ProdSer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim resTipo As Recordset
Dim RES1 As Recordset
Dim SQL1 As String

Private Sub cmbTipo_Click()
    buscarProducto
End Sub

Private Sub Form_Load()
    Lista.Rows = 1
    Lista.ColWidth(7) = 0
    cargaTipo
    buscarProducto
End Sub
Private Sub cargaTipo()
    SQL1 = "SELECT CTPT_TIPO tipo, ctpt_id FROM CAT_TIPO WHERE CTPT_SUBTIPO IN ('P', 'S')"
    Set resTipo = con.Execute(SQL1)
                
    cmbTipo.Clear
    
    cmbTipo.AddItem "TODOS"
    
    Do While Not resTipo.EOF
        cmbTipo.AddItem resTipo.Fields("tipo")
        cmbTipo.ItemData(cmbTipo.ListCount - 1) = resTipo.Fields("ctpt_id")
        resTipo.MoveNext
    Loop

    cmbTipo.ListIndex = 0
End Sub

Private Sub lista_Click()
    Label1(3).Caption = "Descripción: " & Lista.TextMatrix(Lista.Row, 7)
    SQL1 = "SELECT PROD_FOTO FROM PRODUCTOS WHERE PROD_CODIGO = '" & Lista.TextMatrix(Lista.Row, 1) & "'"
    Set RES1 = con.Execute(SQL1)
    
    If IsNull(RES1.Fields("PROD_fOTO")) = False Then
        Dim Imagen1 As Stream
        Set Imagen1 = New Stream
        Imagen1.Type = adTypeBinary
        checarCarpetaTemp
        Imagen1.Open
        Imagen1.Write RES1.Fields("PROD_FOTO")
        Imagen1.SaveToFile direccionSistema & "\Temp\TempProd.dat", adSaveCreateOverWrite
        Imagen1.Close
        imgFoto(0).Picture = LoadPicture(direccionSistema & "\Temp\TempProd.dat")
    Else
        imgFoto(0).Picture = LoadPicture("")
    End If
    
End Sub

Private Sub lista_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Lista_DblClick
    End If
End Sub

Private Sub lista_SelChange()
    lista_Click
End Sub
Private Sub Lista_DblClick()
'    MsgBox FrmFocus
    If modBusqueda = "Operaciones" Then
        If FrmFocus.txtClave(0).Text = "MD" Then
            FrmFocus.txtClave(0).Text = FrmFocus.txtClave(0).Text & Lista.TextMatrix(Lista.Row, 1)
        Else
            FrmFocus.txtClave(0).Text = Lista.TextMatrix(Lista.Row, 1)
        End If
        tipoBusqueda = Left(Lista.TextMatrix(Lista.Row, 3), 1)
        Unload Me
        FrmFocus.cmdOperCheck_Click (0)
    Else
        If modBusqueda = "ConsumoInterno" Then
            FRM_ConsumoInterno.txtClave(0).Text = Lista.TextMatrix(Lista.Row, 1)
            Unload Me
            FRM_ConsumoInterno.checkProdCI
        Else
            If modBusqueda = "Apartado" Then
                FRM_Apartados.txtClave(0).Text = Lista.TextMatrix(Lista.Row, 1)
                Unload Me
                FRM_Apartados.aprt_checkProducto
            End If
        End If
    End If

End Sub

Private Sub lista_GotFocus()
    ConScroll Lista
End Sub

Private Sub lista_LostFocus()
    SinScroll Lista
End Sub

Private Sub buscarProducto()
    Dim SQL1 As String
    Dim RES1 As Recordset
    Dim tipo, Operador As String
    
    tipo = cmbTipo.Text
    
    If tipo = "TODOS" Then
        tipo = "is not null"
    Else
        tipo = " = '" & tipo & "'"
    End If


    Dim texto1 As String
    
    texto1 = ""
    If cmbTipo.Text <> "TODOS" Then
        texto1 = texto1 & "AND upper(TIPO) LIKE upper('%" & cmbTipo.Text & "%') "
    End If
    
    If modBusqueda = "ConsumoInterno" Or modBusqueda = "Apartado" Then
        texto1 = texto1 & " AND SUBTIPO = 'PRODUCTO' Limit 0, " & Val(Text2.Text) & ""
    Else
        texto1 = texto1 & "Limit 0, " & Val(Text2.Text) & ""
    End If
    
    SQL1 = "SELECT * FROM VIEW_PRODUCTOS_INVENTARIO WHERE " & _
    "CODIGO LIKE '%" & textBus(0).Text & "%' " & _
    "AND upper(NOMBRE) LIKE upper('%" & textBus(1).Text & "%') " & texto1

    'MsgBox sql1
    Set RES1 = con.Execute(SQL1)
        Lista.Rows = 1
    Do While Not RES1.EOF
        Lista.AddItem ""
        Lista.TextMatrix(Lista.Rows - 1, 0) = Lista.Rows - 1
        Lista.TextMatrix(Lista.Rows - 1, 1) = "" & RES1.Fields("CODIGO")
        Lista.TextMatrix(Lista.Rows - 1, 2) = RES1.Fields("NOMBRE")
        Lista.TextMatrix(Lista.Rows - 1, 3) = RES1.Fields("SUBTIPO")
        Lista.TextMatrix(Lista.Rows - 1, 4) = RES1.Fields("TIPO")
        Lista.TextMatrix(Lista.Rows - 1, 5) = FormatCurrency(RES1.Fields("PRECIO_VENTA"))
        Lista.TextMatrix(Lista.Rows - 1, 6) = RES1.Fields("CANTIDAD")
        Lista.TextMatrix(Lista.Rows - 1, 7) = RES1.Fields("DESCRIPCION")
        
        If RES1.Fields("ID_STATUS") = "I" Or RES1.Fields("CANTIDAD") <= 0 Then
            If RES1.Fields("subtipo") = "PRODUCTO" Then
                Lista.Row = Lista.Rows - 1
                For b1 = 0 To Lista.Cols - 1
                    Lista.Col = b1
                    Lista.CellForeColor = vbRed
                Next b1
            End If
        End If
        
        RES1.MoveNext
    Loop
Label3.Caption = "Registros en la lista: " & Lista.Rows - 1
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    buscarProducto
Else
    Call Numeros(KeyAscii)
End If
End Sub

Private Sub textBus_Change(index As Integer)
'        buscarProducto
End Sub

Private Sub textBus_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
    If KeyAscii = 13 Then
        buscarProducto
    End If


End Sub
