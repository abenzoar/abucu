VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form BUSQ_ProdTouch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Búsqueda de Productos"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12765
   Icon            =   "BUSQ_ProdTouch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   12765
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Frame Cont1 
      Caption         =   "Datos cuenta"
      Height          =   3735
      Index           =   2
      Left            =   9840
      TabIndex        =   21
      Top             =   3960
      Width           =   2775
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3210
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Cont1 
      Caption         =   "Datos pedido"
      Height          =   3735
      Index           =   1
      Left            =   6960
      TabIndex        =   20
      Top             =   3960
      Width           =   2775
      Begin VB.ListBox List2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3210
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Cont1 
      Caption         =   "Datos producto"
      Height          =   3735
      Index           =   0
      Left            =   2760
      TabIndex        =   6
      Top             =   3960
      Width           =   4095
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         Picture         =   "BUSQ_ProdTouch.frx":24B2
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         Picture         =   "BUSQ_ProdTouch.frx":2D7C
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton cmd_Mas 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         Picture         =   "BUSQ_ProdTouch.frx":3646
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cmd_Menos 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         Picture         =   "BUSQ_ProdTouch.frx":3F10
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtDescripcion 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   1200
         MaxLength       =   2450
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   1800
         Width           =   2775
      End
      Begin VB.TextBox txt_Cantidad 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txt_Tiempo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2160
         TabIndex        =   8
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txt_asiento 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   3120
         TabIndex        =   7
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción del producto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   15
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Producto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   6255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   12
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tiempo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   2160
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Asiento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   3120
         TabIndex        =   10
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.CommandButton Up1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   2775
      Picture         =   "BUSQ_ProdTouch.frx":47DA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   920
   End
   Begin VB.CommandButton Down1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   2775
      Picture         =   "BUSQ_ProdTouch.frx":50A4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   920
   End
   Begin VB.CommandButton Up1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   2775
      Picture         =   "BUSQ_ProdTouch.frx":596E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   920
   End
   Begin VB.CommandButton Down1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   2775
      Picture         =   "BUSQ_ProdTouch.frx":6238
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   920
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   14880
      Top             =   -120
   End
   Begin MSFlexGridLib.MSFlexGrid ListProd1 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   13150
      _Version        =   393216
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   12632256
      ForeColor       =   0
      GridColor       =   16777215
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
   Begin MSFlexGridLib.MSFlexGrid listprod2 
      Height          =   3855
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   6800
      _Version        =   393216
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   12632256
      ForeColor       =   0
      GridColor       =   16777215
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
   Begin VB.Image Image1 
      Height          =   495
      Left            =   2760
      Picture         =   "BUSQ_ProdTouch.frx":6B02
      Stretch         =   -1  'True
      Top             =   0
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   3000
      Picture         =   "BUSQ_ProdTouch.frx":73CC
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   615
   End
End
Attribute VB_Name = "BUSQ_ProdTouch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim SQL1 As String
    Dim RES1 As Recordset
    Dim RES2 As Recordset
    Dim RES3 As Recordset
    Dim RESTIPO_PROD As Recordset
    Dim RES_PROD As Recordset
    Dim checkError As Boolean
    Dim prodId As String
    Dim Id As Long
    Dim save As Boolean
    Dim mayus As Boolean
    Dim activaSeleccion As Boolean
    Dim tipoId(90, 3)
    Dim tipoValor(90, 3)
    Dim prodImgId(90, 5)
    Dim cajaTexto As Integer
    
Private Sub cmd_Mas_Click()
    Select Case cajaTexto
        Case 1: txt_Cantidad.Text = Val(txt_Cantidad) + 1
        Case 2: txt_Tiempo.Text = Val(txt_Tiempo) + 1
        Case 3: txt_asiento.Text = Val(txt_asiento) + 1

    End Select

End Sub

Private Sub cmd_Menos_Click()
    Select Case cajaTexto
        Case 1:
            If Val(txt_Cantidad.Text) >= 1 Then
                txt_Cantidad.Text = Val(txt_Cantidad) - 1
            End If
        
        Case 2:
            If Val(txt_Tiempo.Text) >= 1 Then
                txt_Tiempo.Text = Val(txt_Tiempo) - 1
            End If
        
        Case 3:
            If Val(txt_asiento.Text) >= 1 Then
                txt_asiento.Text = Val(txt_asiento) - 1
            End If


    End Select

End Sub

Private Sub Command1_Click()
    If Label2.Caption <> "" Then
        FRM_OperTouch.lblProdId(0).Caption = prodImgId(listprod2.Row, listprod2.Col)
        FRM_OperTouch.lblProdId(1).Caption = Val(txt_Cantidad.Text)
'        FRM_OperTouch.tiempo = 0
'        FRM_OperTouch.Timer_tiempo.Enabled = True
        
        FRM_OperTouch.add_Touch_Producto
        txt_Cantidad.Text = "1"
        Label2.Caption = ""
        'listprod2.Rows = 0
'        Unload Me
    Else
        MsgBox "No ha seleccionado un producto. Verifique.", vbInformation
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Down1_Click(Index As Integer)
    On Error Resume Next
    
    Select Case Index
        Case 0:
            ListProd1.TopRow = ListProd1.TopRow + 5
        Case 1:
            listprod2.TopRow = listprod2.TopRow + 2
    End Select
End Sub

Private Sub Form_Load()
    cajaTexto = 1
    txt_Cantidad.Text = "1"
    Label2.Caption = ""
    listprod2.Rows = 0
    txtDescripcion.Text = ""
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
        FRM_OperTouch.tiempo = 0
        FRM_OperTouch.Timer_tiempo.Enabled = True
        If salida = True Then
            Cancel = 0
            End
        End If
End Sub

Private Sub ListProd1_Click()
    cargaLista_ProdImagen
    Label2.Caption = ""
    Command1.Enabled = False

End Sub
Private Sub cargaLista_ProdImagen()
On Error Resume Next
Dim Ancho As Long, Alto As Long
Dim contaFila As Long
Dim contaCasillas As Long
Dim contaTipos As Long
Dim columnas As Long
Dim Imagen1 As Stream
Set Imagen1 = New Stream

'Ancho = 2100
'Alto = 2415

Ancho = 1500
Alto = 1800

listprod2.Redraw = False
columnas = Int((listprod2.width) / (Ancho))

SQL1 = "SELECT * fROM VIEW_PRODUCTOS_INVENTARIO WHERE TIPO_ID = '" & tipoId(ListProd1.Row, ListProd1.Col) & "' AND STATUS <> 'INACTIVO' ORDER BY NOMBRE ASC"
Set RES_PROD = con.Execute(SQL1)

listprod2.Rows = 0

If RES_PROD.RecordCount >= columnas Then
    listprod2.Cols = columnas
    For b1 = 1 To columnas
        listprod2.ColWidth(b1 - 1) = Ancho
    Next b1
Else
    listprod2.Cols = RES_PROD.RecordCount
    For b1 = 1 To RES_PROD.RecordCount
        listprod2.ColWidth(b1 - 1) = Ancho
    Next b1
    
End If

contaFila = 0
contaTipos = 0
contaCasillas = columnas

Do While Not RES_PROD.EOF
    If contaCasillas = columnas Then
        listprod2.AddItem ""
        listprod2.RowHeight(listprod2.Rows - 1) = Alto
        contaCasillas = 0
    End If

    If RES_PROD.Fields("FOTO_SN") = "SI" Then
        If IsNull(RES_PROD.Fields("FOTO")) = False Then
            Imagen1.Type = adTypeBinary
            checarCarpetaTemp
            Imagen1.Open
            Imagen1.Write RES_PROD.Fields("FOTO")
            Imagen1.SaveToFile direccionSistema & "\Temp\Prod" & contaTipos & ".jpg", adSaveCreateOverWrite
            Imagen1.Close

            listprod2.Row = listprod2.Rows - 1
            listprod2.Col = contaCasillas
            Set listprod2.CellPicture = LoadPicture(direccionSistema & "\Temp\Prod" & contaTipos & ".jpg")
            listprod2.CellAlignment = 8
            listprod2.TextMatrix(listprod2.Rows - 1, contaCasillas) = RES_PROD.Fields("NOMBRE") & " " & FormatCurrency(RES_PROD.Fields("PRECIO_VENTA"))
            prodImgId(listprod2.Row, listprod2.Col) = RES_PROD.Fields("PROD_ID")
        End If
    Else
        listprod2.Row = listprod2.Rows - 1
        listprod2.Col = contaCasillas
        listprod2.CellAlignment = 8
        listprod2.TextMatrix(listprod2.Rows - 1, contaCasillas) = RES_PROD.Fields("NOMBRE") & " " & FormatCurrency(RES_PROD.Fields("PRECIO_VENTA"))
        prodImgId(listprod2.Row, listprod2.Col) = RES_PROD.Fields("PROD_ID")
    End If
'    tipoId(ListProd1.Rows - 1, contaCasillas) = RESTIPO_PROD.Fields("CLAVE")
'    tipoValor(ListProd1.Rows - 1, contaCasillas) = RESTIPO_PROD.Fields("TIPO")
    contaCasillas = contaCasillas + 1
    contaTipos = contaTipos + 1

    RES_PROD.MoveNext
Loop
'
listprod2.WordWrap = True
listprod2.Redraw = True

End Sub

Private Sub ListProd1_GotFocus()
    ConScroll ListProd1
End Sub

Private Sub ListProd1_LostFocus()
    SinScroll ListProd1
End Sub

Private Sub listprod2_Click()
    If listprod2.Rows > 0 Then
        Label2.Caption = listprod2.TextMatrix(listprod2.Row, listprod2.Col)
        Command1.Enabled = True
    End If
End Sub

Private Sub listprod2_DblClick()

    If listprod2.Rows > 0 Then
        FRM_OperTouch.lblProdId(0).Caption = prodImgId(listprod2.Row, listprod2.Col)
        FRM_OperTouch.lblProdId(1).Caption = Val(txt_Cantidad.Text)
        FRM_OperTouch.tiempo = 0
        FRM_OperTouch.Timer_tiempo.Enabled = True
        
        FRM_OperTouch.add_Touch_Producto
        Unload Me
    End If
End Sub

Private Sub listprod2_GotFocus()
    ConScroll listprod2
End Sub

Private Sub listprod2_LostFocus()
    SinScroll listprod2
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    If Me.height >= 8500 Then
        
        ListProd1.height = ListProd1.height + (Me.height - 8500 + 250)
        listprod2.height = listprod2.height + (Me.height - 8500 + 250)
        Cont1(0).Top = Cont1(0).Top + (Me.height - 8500 + 250)
        Cont1(1).Top = Cont1(1).Top + (Me.height - 8500 + 250)
        Cont1(2).Top = Cont1(2).Top + (Me.height - 8500 + 250)
        
    End If
    
    cargaLista_TipoImagen
    

End Sub

Private Sub cargaLista_TipoImagen()
Dim Ancho As Long, Alto As Long
Dim contaFila As Long
Dim contaCasillas As Long
Dim contaTipos As Long

Dim Imagen1 As Stream
Set Imagen1 = New Stream


'Ancho = 2175
'Alto = 2415

Ancho = 1100
Alto = 1100

SQL1 = "SELECT * fROM VIEW_TIPOPRODUCTOS ORDER BY ORDEN ASC"
Set RESTIPO_PROD = con.Execute(SQL1)

ListProd1.Redraw = False
ListProd1.Rows = 0
ListProd1.ColWidth(0) = Ancho
ListProd1.ColWidth(1) = Ancho
'ListProd1.ColWidth(2) = Ancho
contaFila = 0
contaTipos = 0
contaCasillas = 2

Do While Not RESTIPO_PROD.EOF
    If contaCasillas = 2 Then
        ListProd1.AddItem ""
        ListProd1.RowHeight(ListProd1.Rows - 1) = Alto
        contaCasillas = 0
    End If
    
    If RESTIPO_PROD.Fields("FOTO_SN") = "SI" Then
        If IsNull(RESTIPO_PROD.Fields("fOTO")) = False Then
            Imagen1.Type = adTypeBinary
            checarCarpetaTemp
            Imagen1.Open
            Imagen1.Write RESTIPO_PROD.Fields("FOTO")
            Imagen1.SaveToFile direccionSistema & "\Temp\" & contaTipos & ".jpg", adSaveCreateOverWrite
            Imagen1.Close

            ListProd1.Row = ListProd1.Rows - 1
            ListProd1.Col = contaCasillas
            Set ListProd1.CellPicture = LoadPicture(direccionSistema & "\Temp\" & contaTipos & ".jpg")
            ListProd1.CellAlignment = 8
            ListProd1.TextMatrix(ListProd1.Rows - 1, contaCasillas) = RESTIPO_PROD.Fields("TIPO")
            
        End If
    Else
        
        If IsNull(RESTIPO_PROD.Fields("color")) = False Then
            ListProd1.Row = ListProd1.Rows - 1
            ListProd1.Col = contaCasillas
            ListProd1.CellBackColor = RESTIPO_PROD.Fields("COLOR")
            ListProd1.CellAlignment = 8
            ListProd1.TextMatrix(ListProd1.Rows - 1, contaCasillas) = RESTIPO_PROD.Fields("TIPO")
        Else
            ListProd1.Row = ListProd1.Rows - 1
            ListProd1.Col = contaCasillas
            ListProd1.CellAlignment = 8
            ListProd1.TextMatrix(ListProd1.Rows - 1, contaCasillas) = RESTIPO_PROD.Fields("TIPO")
        End If
    End If
    tipoId(ListProd1.Rows - 1, contaCasillas) = RESTIPO_PROD.Fields("CLAVE")
    tipoValor(ListProd1.Rows - 1, contaCasillas) = RESTIPO_PROD.Fields("TIPO")
    contaCasillas = contaCasillas + 1
    contaTipos = contaTipos + 1
    
    RESTIPO_PROD.MoveNext
Loop

ListProd1.WordWrap = True
ListProd1.Redraw = True

End Sub
Private Sub txt_asiento_DblClick()
    Shell "osk.exe"
End Sub

Private Sub txt_asiento_GotFocus()
    cajaTexto = 3
End Sub

Private Sub txt_asiento_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 27 Then
    Else
        KeyAscii = 0
    End If

    
End Sub

Private Sub txt_Cantidad_GotFocus()
    cajaTexto = 1
End Sub

Private Sub txt_Cantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 27 Then
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub txt_Tiempo_DblClick()
'    Shell "osk.exe"


End Sub
Private Sub txt_Tiempo_GotFocus()
    txt_Tiempo.SelStart = 0
    txt_Tiempo.SelLength = Len(txt_Tiempo.Text)
    cajaTexto = 2
End Sub
Private Sub txt_Tiempo_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 27 Then
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub txtDescripcion_DblClick()
'On Error Resume Next
'    Shell "osk.exe"
    
    
    Set formDescripcion = BUSQ_ProdTouch
    teclado = "Desc_touch1"
    FRM_Teclado.Show
End Sub

Private Sub Up1_Click(Index As Integer)
    'On Error Resume Next
    'MsgBox ListProd1.TopRow
    Select Case Index
        Case 0:
            If ListProd1.TopRow < 5 Then
                ListProd1.TopRow = 0
            Else
                ListProd1.TopRow = ListProd1.TopRow - 5
            End If
        Case 1:
            If listprod2.TopRow < 2 Then
                listprod2.TopRow = 0
            Else
                listprod2.TopRow = listprod2.TopRow - 2
            End If
    End Select
End Sub
