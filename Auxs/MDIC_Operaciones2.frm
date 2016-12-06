VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form MDIC_Operaciones2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Operaciones"
   ClientHeight    =   10215
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16395
   Icon            =   "MDIC_Operaciones2.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   16395
   WindowState     =   2  'Maximized
   Begin VB.Timer Time_listaRapida 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   14160
      Top             =   360
   End
   Begin VB.ListBox lista_rapida 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   7680
      TabIndex        =   3
      Top             =   -5000
      Width           =   8175
   End
   Begin VB.TextBox txtClave 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tekton Pro Ext"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
   End
   Begin MSFlexGridLib.MSFlexGrid lista 
      Height          =   6975
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   12303
      _Version        =   393216
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      ForeColorSel    =   -2147483640
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      GridLines       =   0
      GridLinesFixed  =   0
      ScrollBars      =   2
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tekton Pro Ext"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid ListaTotal 
      Height          =   2535
      Left            =   0
      TabIndex        =   2
      Top             =   7680
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4471
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483635
      ForeColor       =   16777215
      ForeColorFixed  =   16777215
      BackColorBkg    =   -2147483635
      WordWrap        =   -1  'True
      GridLines       =   0
      GridLinesFixed  =   0
      BorderStyle     =   0
      FormatString    =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tekton Pro Ext"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgFoto 
      Height          =   3855
      Index           =   0
      Left            =   6840
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   3615
   End
   Begin VB.Label lblDatos 
      BackStyle       =   0  'Transparent
      Caption         =   "Ninguno"
      BeginProperty Font 
         Name            =   "Tekton Pro Ext"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   0
      Left            =   6720
      TabIndex        =   8
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   255
      Index           =   0
      Left            =   8520
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00808000&
      Height          =   255
      Index           =   7
      Left            =   6720
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tekton Pro Ext"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   6720
      TabIndex        =   5
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Abierto"
      BeginProperty Font 
         Name            =   "Tekton Pro Ext"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Index           =   2
      Left            =   8520
      TabIndex        =   4
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "MDIC_Operaciones2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    Call cargaEjemplo("123", "AGUA BONAFONT TAMAÑO FAMILIAR 1.5 LTS", "7.50", "2", "Ninguna", "1.5", "BEBIDAS", "BONAFONT")
End Sub

Private Sub formato_Total()
    ListaTotal.Rows = 0
    ListaTotal.Cols = 4
    ListaTotal.ColWidth(0) = 1562
    ListaTotal.ColWidth(1) = 1562
    ListaTotal.ColWidth(2) = 1563
    ListaTotal.ColWidth(3) = 1563
       
    ListaTotal.MergeCells = flexMergeRestrictRows
       
    ListaTotal.AddItem ""
    ListaTotal.RowHeight(ListaTotal.Rows - 1) = 350
    ListaTotal.TextMatrix(ListaTotal.Rows - 1, 0) = "SUBTOTAL"
    ListaTotal.TextMatrix(ListaTotal.Rows - 1, 1) = "SUBTOTAL"
    ListaTotal.TextMatrix(ListaTotal.Rows - 1, 2) = FormatCurrency(0)
    ListaTotal.TextMatrix(ListaTotal.Rows - 1, 3) = FormatCurrency(0)
    ListaTotal.MergeRow(0) = True
    
    ListaTotal.AddItem ""
    ListaTotal.RowHeight(ListaTotal.Rows - 1) = 350
    ListaTotal.TextMatrix(ListaTotal.Rows - 1, 0) = "DESCUENTO"
    ListaTotal.TextMatrix(ListaTotal.Rows - 1, 1) = "DESCUENTO"
    ListaTotal.TextMatrix(ListaTotal.Rows - 1, 2) = FormatCurrency(0)
    ListaTotal.TextMatrix(ListaTotal.Rows - 1, 3) = FormatCurrency(0)
    ListaTotal.MergeRow(1) = True
    ListaTotal.AddItem ""
    ListaTotal.RowHeight(ListaTotal.Rows - 1) = 550
    ListaTotal.TextMatrix(ListaTotal.Rows - 1, 0) = ""
    ListaTotal.TextMatrix(ListaTotal.Rows - 1, 1) = ""
    ListaTotal.TextMatrix(ListaTotal.Rows - 1, 2) = FormatCurrency(0)
    ListaTotal.TextMatrix(ListaTotal.Rows - 1, 3) = FormatCurrency(0)
    ListaTotal.MergeRow(2) = True
    
    ListaTotal.Row = ListaTotal.Rows - 1
    ListaTotal.Col = 1
    Set ListaTotal.CellPicture = LoadPicture(direccionSistema & "\Com\total.jpg")

    
'    lista.Row = ListaTotal.Rows - 1
    For b1 = 0 To 2
        For c1 = 0 To 3
            ListaTotal.Row = b1
            ListaTotal.Col = c1
            ListaTotal.CellAlignment = 7
            If b1 = 2 Then
                ListaTotal.CellFontSize = 16
            End If
        Next c1
    Next b1


End Sub

Private Sub formato_fila(fila As Integer, columna As Long, color_Fondo As String, color_Letra As String, tam_Letra As Long, negrita As Boolean, alinear As Long, todasColumnas As Boolean, ajustar As Boolean)

    If todasColumnas = True Then
        lista.Row = fila
        For b1 = 0 To 5
            lista.Col = b1
            lista.CellAlignment = alinear
            lista.CellBackColor = color_Fondo
            lista.CellForeColor = color_Letra
            
            lista.CellFontSize = tam_Letra
            lista.CellFontBold = negrita
            lista.CellAlignment = alinear
        Next b1
    Else
        lista.Row = fila
        lista.Col = columna
        lista.WordWrap = ajustar
        lista.CellAlignment = alinear
        lista.CellBackColor = color_Fondo
        lista.CellForeColor = color_Letra
            
        lista.CellFontSize = tam_Letra
        lista.CellFontBold = negrita
        lista.CellAlignment = alinear
    End If
    
    lista.MergeCells = flexMergeRestrictRows
    lista.MergeRow(fila) = True
        
End Sub

Private Sub cargaEjemplo(clave As String, producto As String, Precio As Double, cantidad As Double, Descripcion As String, DESCUENTO As Double, tipo As String, Marca As String)
    Dim SUBTOTAL As Double
    SUBTOTAL = ((cantidad * Precio))
    DESCUENTO = DESCUENTO * (-1)
    lista.Redraw = False
'''''''Primer fila
    
    lista.AddItem ""
    lista.RowHeight(lista.Rows - 1) = 850
    lista.TextMatrix(lista.Rows - 1, 0) = producto
    lista.TextMatrix(lista.Rows - 1, 1) = producto
    lista.TextMatrix(lista.Rows - 1, 2) = producto
    lista.TextMatrix(lista.Rows - 1, 3) = producto
    lista.TextMatrix(lista.Rows - 1, 4) = producto
    lista.TextMatrix(lista.Rows - 1, 5) = producto
    Call formato_fila(lista.Rows - 1, "0", "&HFFFFFF", "0", 14, True, "1", False, True)
    
    lista.AddItem ""
    lista.RowHeight(lista.Rows - 1) = 250
    lista.TextMatrix(lista.Rows - 1, 0) = tipo & "  " & Marca
    lista.TextMatrix(lista.Rows - 1, 1) = tipo & "  " & Marca
    lista.TextMatrix(lista.Rows - 1, 2) = tipo & "  " & Marca
    lista.TextMatrix(lista.Rows - 1, 3) = ""
    lista.TextMatrix(lista.Rows - 1, 4) = ""
    lista.TextMatrix(lista.Rows - 1, 5) = ""
    Call formato_fila(lista.Rows - 1, "0", "&HFFFFFF", "&H808080", 10, False, 1, False, True)
    
    lista.AddItem ""
    lista.RowHeight(lista.Rows - 1) = 350
    lista.TextMatrix(lista.Rows - 1, 0) = clave
    lista.TextMatrix(lista.Rows - 1, 1) = clave
    lista.TextMatrix(lista.Rows - 1, 2) = clave
    lista.TextMatrix(lista.Rows - 1, 3) = ""
    lista.TextMatrix(lista.Rows - 1, 4) = ""
    lista.TextMatrix(lista.Rows - 1, 5) = ""
    Call formato_fila(lista.Rows - 1, "0", "&HFFFFFF", "&H808080", 10, True, 1, False, True)
    
    lista.AddItem ""
    lista.RowHeight(lista.Rows - 1) = 450
    lista.TextMatrix(lista.Rows - 1, 0) = cantidad
    lista.TextMatrix(lista.Rows - 1, 1) = FormatCurrency(Precio)
    lista.TextMatrix(lista.Rows - 1, 2) = FormatCurrency(Precio)
    lista.TextMatrix(lista.Rows - 1, 3) = ""
    lista.TextMatrix(lista.Rows - 1, 4) = FormatCurrency(SUBTOTAL)
    lista.TextMatrix(lista.Rows - 1, 5) = FormatCurrency(SUBTOTAL)
    Call formato_fila(lista.Rows - 1, "0", "-2147483643", "&H808080", 12, True, 4, True, True)
    lista.Row = lista.Rows - 1
    lista.Col = 3
    Set lista.CellPicture = LoadPicture(direccionSistema & "\Com\descuento.jpg")
    
    lista.AddItem ""
    lista.RowHeight(lista.Rows - 1) = 450
    lista.TextMatrix(lista.Rows - 1, 0) = ""
    lista.TextMatrix(lista.Rows - 1, 1) = ""
    lista.TextMatrix(lista.Rows - 1, 2) = ""
    lista.TextMatrix(lista.Rows - 1, 3) = ""
    lista.TextMatrix(lista.Rows - 1, 4) = FormatCurrency(DESCUENTO)
    lista.TextMatrix(lista.Rows - 1, 5) = FormatCurrency(DESCUENTO)
    Call formato_fila(lista.Rows - 1, "0", "-2147483643", "&H8000000D", 14, True, 4, True, True)
        
    lista.Redraw = True

    
End Sub

Private Sub Form_Load()
    formato_Lista
    formato_Total
    cargaDatosInicial
End Sub
Private Sub cargaDatosInicial()
    lista_rapida.Visible = False
    
    'checkDatos
    Set FrmFocus = Me
    numFrmOper = numFrmOper + 1
'    lista.ColWidth(6) = 0
'    lista.ColWidth(7) = 0
'    lblDatos(3).Caption = ""
    
'    SSTab1.Tab = 1
    
'    lista.ColWidth(9) = 0
'    lista.ColWidth(10) = 0
'    lista.ColWidth(13) = 0
'    lInfo(0).Caption = "0"
    
    lista.Rows = 1
'    txtCant.Visible = False
'    textDesc.Visible = False
'    descGral = False
'    txtObservacion.Locked = False
        
        
    
'    cmbEstado.Clear
'    cmbEstado.AddItem "SIN ATENDER"
'    cmbEstado.AddItem "ATENDIDO"
'    cmbEstado.AddItem "RECIBIDO"
'    cmbEstado.AddItem "NINGUNO"
        
    If tikcet = False Then
'        txtClave(1).Text = FRM_Menu.menuBarra2.Panels(7).Text
'        txtClave(2).Text = "CLTE"
'        checkUsuario
'        checkCliente
        crearFolio
    Else
        If tikcet = True Then
            tikcet = False
            lInfo(1).Caption = folioTicket
            'cargaTicket
        End If
    End If

    
    
'    If FRM_Menu.menuBarra2.Panels(14).Text = "A" Then
'        cmbEstado.Visible = True
'        Label1(12).Visible = True
'        Line1(12).Visible = True
'    Else
'        cmbEstado.Visible = False
'        Label1(12).Visible = False
'        Line1(12).Visible = False
'    End If
    
    
    
'    Me.Caption = "Operación Ticket " & folioTicket & " Clte: " & lblDatos(2).Caption

End Sub
Private Sub crearFolio()
'    SQL1 = "INSERT INTO VENTAS (VENT_FECHAHORA, VENT_STATUS, VENT_VENDPERID, VENT_VENDTIPOID, VENT_VENDTIPO, " & _
'    "VENT_CLIEPERID, VENT_CLIETIPOID, VENT_CLIETIPO) VALUES " & _
'    "('" & Format(Date, "yyyy-MM-dd") & " " & Format(Time, "HH:MM:SS") & "', 'G', '" & FRM_Menu.menuBarra2.Panels(7).Text & "', '" & FRM_Menu.menuBarra2.Panels(8).Text & "', 'U', " & _
'    "'" & lblClieId(0).Caption & "', '" & lblClieId(1).Caption & "', '" & lblClieId(2).Caption & "')"

'    con.Execute (SQL1)
    
    
    SQL1 = "select last_insert_id() folioId"
    Set RES1 = con.Execute(SQL1)
    If Not RES1.EOF Then
        folio = RES1.Fields("folioId")
    End If

    lInfo(1).Caption = folio
    lInfo(2).Caption = "Abierto"
    
End Sub
Private Sub formato_Lista()
    lista.Rows = 0
    lista.Cols = 6
    lista.ColWidth(0) = 1000 '(Fila)
    lista.ColWidth(1) = 1550 'Producto
    lista.ColWidth(2) = 700 'Cantidad
    lista.ColWidth(3) = 1000 'Pecio
    lista.ColWidth(4) = 600 'descuento
    lista.ColWidth(5) = 1400 'subtotal

End Sub
Private Sub cargaLista_General(Index As Integer)
    On Error Resume Next
    Dim textoLista As String
    Dim idTexto As String
    
    If Len(txtClave(0).Text) > 0 Then
        If Index = 0 Then
            tipoBusqueda = "P"

            SQL1 = "SELECT concat(PROD_CODIGO, '          ', PROD_NOMBRE, '         ', '$', ROUND(PROD_PRECIO, 2)) PRODUCTOS, PROD_ID FROM PRODUCTOS " & _
            "WHERE UPPER(concat(PROD_CODIGO, ' ', PROD_NOMBRE)) LIKE UPPER('%" & txtClave(0).Text & "%') LIMIT 8 " & _
            " Union All " & _
            "SELECT CONCAT('* ', CTMB_NOMBRE, '          DIAS: ', CTMB_DIAS, '          ', CTMB_PRECIO) PRODUCTOS, CTMB_ID PROD_ID FROM CAT_MEMBRESIAS " & _
            "WHERE UPPER(CONCAT(CTMB_NOMBRE)) LIKE UPPER('%" & txtClave(0).Text & "%') LIMIT 8"
            
            textoLista = "Productos"
            idTexto = "Prod_id"
'            MsgBox txtClave(0).Text
'            MsgBox SQL1
            Set RES1 = con.Execute(SQL1)
        Else
            If Index = 1 Then
                tipoBusqueda = "U"
                SQL1 = "SELECT T4.PERTP_CODIGO_MEMBRESIA,  " & _
                "CONCAT(T2.PER_NOMBRE, '  ', T2.PER_PATERNO, '  ', T2.PER_MATERNO) USUARIO, T2.PER_ID " & _
                "FROM PERSONA T2, CAT_TIPO T3, PER_tIPO T4 " & _
                "WHERE T4.PERTP_TIPO_ID = T3.CTPT_ID AND T4.PERTP_PER_TIPO = T3.CTPT_SUBTIPO AND T2.PER_ID = T4.PERTP_PER_ID " & _
                "AND upper(concat(T2.PER_NOMBRE, ' ', T2.PER_PATERNO, ' ', T2.PER_MATERNO)) LIKE UPPER('%" & txtClave(0).Text & "%') " & _
                "AND T4.PERTP_PER_TIPO = 'U' AND T4.PERTP_STATUS = 'A'" & _
                "ORDER BY T2.PER_NOMBRE ASC"
                textoLista = "Usuario"
                idTexto = "Per_Id"
                'MsgBox SQL1
                Set RES1 = con.Execute(SQL1)
            Else
                If Index = 2 Then
                    tipoBusqueda = "C"
                    SQL1 = "SELECT T4.PERTP_CODIGO_MEMBRESIA,  " & _
                    "CONCAT(T2.PER_NOMBRE, '  ', T2.PER_PATERNO, '  ', T2.PER_MATERNO) CLIENTE, T2.PER_ID " & _
                    "FROM PERSONA T2, CAT_TIPO T3, PER_tIPO T4 " & _
                    "WHERE T4.PERTP_TIPO_ID = T3.CTPT_ID AND T4.PERTP_PER_TIPO = T3.CTPT_SUBTIPO AND T2.PER_ID = T4.PERTP_PER_ID " & _
                    "AND upper(concat(T2.PER_NOMBRE, ' ', T2.PER_PATERNO, ' ', T2.PER_MATERNO)) LIKE UPPER('%" & txtClave(0).Text & "%') " & _
                    "AND T4.PERTP_PER_TIPO = 'C' AND T4.PERTP_STATUS = 'A'" & _
                    "ORDER BY T2.PER_NOMBRE ASC"
                    textoLista = "Cliente"
                    idTexto = "Per_Id"
                    'MsgBox SQL1
                    Set RES1 = con.Execute(SQL1)
                End If
            End If
        End If
        
                                           
        lista_rapida.Clear
        Do While Not RES1.EOF
            lista_rapida.AddItem RES1.Fields(textoLista)
            lista_rapida.ItemData(lista_rapida.NewIndex) = RES1.Fields(idTexto)
            RES1.MoveNext
        Loop
    End If
End Sub
Private Sub lista_rapida_DblClick()
    cargaFrom_ListaRapida
End Sub
Private Sub lista_rapida_GotFocus()
    Time_listaRapida.Enabled = False
End Sub

Private Sub lista_rapida_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cargaFrom_ListaRapida
    End If
End Sub

Private Sub lista_rapida_LostFocus()
    lista_rapida.Visible = False
End Sub
Private Sub checkProducto()

    On Error Resume Next
        
    lista_rapida.Visible = False
     monedero = False
    If UCase(lInfo(2).Caption) <> UCase("Abierto") Then
        MsgBox "No se puede realizar la acción. Verfique.", vbExclamation
        Exit Sub
    End If
    
'    If Left(txtClave(0).Text, 2) = "MD" Then
'        If Val(Format(lblDatos(6).Caption, "General Number")) > 0 Then
'            monedero = True
'            'addMonedero
'            txtClave(0).Text = Right(txtClave(0).Text, (Len(txtClave(0).Text) - 2))
'        Else
'            MsgBox "No se puede asignar monedero a la cuenta del cliente seleccionado. Verifique.", vbInformation
'            Exit Sub
'        End If
'    End If
    
    SQL1 = "SELECT PROD_CODIGO, PROD_NOMBRE, PROD_DESCRIPCION, CTMR_MARCA, " & _
    "if(PROD_STATUS= 'A', 'ACTIVO', 'INACTIVO') STATUS, PROD_PRECIO, PROD_CANT, " & _
    "CTPT_TIPO, PROD_MARCA, PROD_TIPO, PROD_PRESENTACION, PROD_UNIMED_PRESENT,  " & _
    "PROD_FOTO, PROD_STOCK_MIN, PROD_STOCK_MAX, T4.CTPS_NOMBRE, PROD_STATUS, " & _
    "if(PROD_SERV= 'P', 'PRODUCTO', 'SERVICIO') TIPO_PROD, PROD_SERV, PROD_ID, PROD_DEPENDIENTE " & _
    "FROM PRODUCTOS T1, CAT_MARCA T2, CAT_TIPO T3, CAT_PRESENTACION T4 " & _
    "WHERE T1.PROD_MARCA = T2.CTMR_ID AND T1.PROD_TIPO = T3.CTPT_ID AND T1.PROD_SUBTIPO = T3.CTPT_SUBTIPO " & _
    "AND (T1.PROD_UNIMED_PRESENT = T4.CTPS_ID OR T1.PROD_UNIMED_PRESENT IS NULL) AND " & _
    "PROD_CODIGO = '" & txtClave(0).Text & "' AND PROD_STATUS = 'A'"
    Set RES1 = con.Execute(SQL1)
    Dim b1 As Long
    If Not RES1.EOF Then
        lblDatos(0).Caption = RES1.Fields("PROD_NOMBRE")
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
        
        'addLista
        
'        If monedero = True Then
'            Call addMonedero(0, 0)
'        End If
    Else
        'checkServicio
    End If
    
    
End Sub

Private Sub cargaFrom_ListaRapida()
    On Error Resume Next
    
    If tipoBusqueda = "P" Then
        If Left(lista_rapida.Text, 1) = "*" Then
            txtClave(0).Text = lista_rapida.ItemData(lista_rapida.ListIndex)
            lista_rapida.Visible = False
            'checkMembresia
        Else
            SQL1 = "select prod_codigo from productos where prod_id = '" & lista_rapida.ItemData(lista_rapida.ListIndex) & "'"
            Set RES1 = con.Execute(SQL1)
            
            txtClave(0).Text = RES1.Fields("PROD_CODIGO")
            lista_rapida.Visible = False
            checkProducto
        End If
    Else
        If tipoBusqueda = "U" Then
            'txtClave(1).Text = a
            SQL1 = "select PERTP_CODIGO_MEMBRESIA from PER_TIPO where PERTP_PER_ID = '" & lista_rapida.ItemData(lista_rapida.ListIndex) & "'"
            Set RES1 = con.Execute(SQL1)
            
            txtClave(1).Text = RES1.Fields("PERTP_CODIGO_MEMBRESIA")
            lista_rapida.Visible = False
            'checkUsuario
        Else
            If tipoBusqueda = "C" Then
                'txtClave(2).Text = a
                SQL1 = "select PERTP_CODIGO_MEMBRESIA from PER_TIPO where PERTP_PER_ID = '" & lista_rapida.ItemData(lista_rapida.ListIndex) & "'"
                Set RES1 = con.Execute(SQL1)
                
                txtClave(2).Text = RES1.Fields("PERTP_CODIGO_MEMBRESIA")
                lista_rapida.Visible = False
                'checkCliente
            End If
        End If
    End If
    
    
    tipoBusqueda = ""

End Sub



Private Sub Time_listaRapida_Timer()
    Time_listaRapida.Enabled = False
    lista_rapida.Visible = False
End Sub

Private Sub txtClave_Change(Index As Integer)
    If Len(txtClave(0).Text) > 0 Then
        lista_rapida.Top = txtClave(0).Top + 375
        lista_rapida.Left = txtClave(0).Left
        lista_rapida.Visible = True
        Time_listaRapida.Enabled = False
        Time_listaRapida.Enabled = True
        
        cargaLista_General (0)
    Else
        If Len(txtClave(0).Text) <= 0 Then
            lista_rapida.Visible = False
        End If
    End If
End Sub

Private Sub txtClave_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        If lista_rapida.Visible = True Then
            lista_rapida.SetFocus
        End If
    End If
End Sub
