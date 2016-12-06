VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRM_OperTouch 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comandas / Operaciones"
   ClientHeight    =   10785
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15465
   Icon            =   "FRM_OperTouch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10785
   ScaleWidth      =   15465
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
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
      Height          =   855
      Index           =   1
      Left            =   3720
      Picture         =   "FRM_OperTouch.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5880
      Width           =   975
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
      Height          =   855
      Index           =   1
      Left            =   3720
      Picture         =   "FRM_OperTouch.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4920
      Width           =   975
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
      Height          =   855
      Index           =   0
      Left            =   3720
      Picture         =   "FRM_OperTouch.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2280
      Width           =   975
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
      Height          =   855
      Index           =   0
      Left            =   3720
      Picture         =   "FRM_OperTouch.frx":2328
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmd_Ticket 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Comanda"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      Picture         =   "FRM_OperTouch.frx":2BF2
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Timer Timer_tiempo 
      Interval        =   1000
      Left            =   720
      Top             =   7560
   End
   Begin VB.CommandButton cmd_Nota 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Observación"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      Picture         =   "FRM_OperTouch.frx":34BC
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmd_Bloquear 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bloquear"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      Picture         =   "FRM_OperTouch.frx":596E
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmd_delProd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quitar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      Picture         =   "FRM_OperTouch.frx":6238
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmd_AddProd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Agregar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      Picture         =   "FRM_OperTouch.frx":6B02
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmd_cerrar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      Picture         =   "FRM_OperTouch.frx":73CC
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmd_Abrir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Abrir"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      Picture         =   "FRM_OperTouch.frx":7C96
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   14160
      Top             =   2760
   End
   Begin MSFlexGridLib.MSFlexGrid lista_Mesa 
      Height          =   7815
      Left            =   1320
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   13785
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   16777215
      BackColorBkg    =   16777215
      FormatString    =   "Mesas   |Mesas   "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid lista_Producto 
      Height          =   4695
      Left            =   4800
      TabIndex        =   1
      Top             =   3120
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   8281
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   16777215
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      AllowUserResizing=   1
      FormatString    =   $"FRM_OperTouch.frx":8560
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
   Begin MSFlexGridLib.MSFlexGrid lista_detalle 
      Height          =   2775
      Left            =   4800
      TabIndex        =   2
      Top             =   0
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   4895
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   16777215
      BackColorBkg    =   16777215
      AllowUserResizing=   1
      FormatString    =   $"FRM_OperTouch.frx":861C
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
   Begin VB.Image Image2 
      Height          =   495
      Left            =   4080
      Picture         =   "FRM_OperTouch.frx":86F9
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   3840
      Picture         =   "FRM_OperTouch.frx":8FC3
      Stretch         =   -1  'True
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblProdId 
      Caption         =   "Label10"
      Height          =   135
      Index           =   1
      Left            =   3600
      TabIndex        =   20
      Top             =   8400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Tiempo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Label lblProdId 
      Caption         =   "Label10"
      Height          =   135
      Index           =   0
      Left            =   3600
      TabIndex        =   16
      Top             =   8040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblUserId 
      Caption         =   "Label10"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   -5000
      Width           =   1095
   End
   Begin VB.Label lblUserId 
      Caption         =   "Label10"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   -5000
      Width           =   1095
   End
   Begin VB.Label lblUserId 
      Caption         =   "Label10"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   -5000
      Width           =   1095
   End
   Begin VB.Label lblClieId 
      Caption         =   "Label10"
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   11
      Top             =   -5000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblClieId 
      Caption         =   "Label10"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   10
      Top             =   -5000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblClieId 
      Caption         =   "Label10"
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   9
      Top             =   -5000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta/Mesa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "FRM_OperTouch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RES_MESA As Recordset
Dim RES_DETALLE As Recordset
Dim RES_PRODUCTOS As Recordset
Dim RES_VENTA As Recordset
Dim RES1 As Recordset
Dim sql1 As String
Dim mesaId(50, 2)
Dim mesaStatus(50, 2)
Dim cierre As String
Dim tiempoPass As Integer
'Public protector As Boolean
Public tiempo As Integer

Public Sub add_Touch_Producto()
        
    Dim cantidad As Long
    Dim encuentra As Boolean
    
    cantidad = lblProdId(1).Caption
    encuentra = False
    If lista_detalle.Rows > 1 Then
        sql1 = "SELECT PROD_CODIGO, PROD_NOMBRE, PROD_DESCRIPCION, CTMR_MARCA, " & _
        "if(PROD_STATUS= 'A', 'ACTIVO', 'INACTIVO') STATUS, PROD_PRECIO, PROD_CANT, " & _
        "CTPT_TIPO, PROD_MARCA, PROD_TIPO, PROD_PRESENTACION, PROD_UNIMED_PRESENT,  " & _
        "PROD_FOTO, PROD_STOCK_MIN, PROD_STOCK_MAX, T4.CTPS_NOMBRE, PROD_STATUS, " & _
        "if(PROD_SERV= 'P', 'PRODUCTO', 'SERVICIO') TIPO_PROD, PROD_SERV, PROD_ID, PROD_DEPENDIENTE " & _
        "FROM PRODUCTOS T1, CAT_MARCA T2, CAT_TIPO T3, CAT_PRESENTACION T4 " & _
        "WHERE T1.PROD_MARCA = T2.CTMR_ID AND T1.PROD_TIPO = T3.CTPT_ID AND T1.PROD_SUBTIPO = T3.CTPT_SUBTIPO " & _
        "AND (T1.PROD_UNIMED_PRESENT = T4.CTPS_ID OR T1.PROD_UNIMED_PRESENT IS NULL) AND " & _
        "PROD_ID = '" & lblProdId(0).Caption & "' AND PROD_STATUS = 'A'"
        Set RES_PRODUCTOS = con.Execute(sql1)
    
        If Not RES_PRODUCTOS.EOF Then
                                
            If encuentra = False Then
                Dim idLast As Long
                sql1 = "SELECT (COUNT(*) + 1) IdLast FROM VENTA_DETALLE"
                Set RES1 = con.Execute(sql1)
                
                idLast = RES1.Fields("IDLAST")
                
                sql1 = "INSERT INTO VENTA_DETALLE (VENDET_FOLIO, VENDET_PRODUCTOID, VENDET_PRODSERV, VENDET_PRODUCTONOMBRE, " & _
                "VENDET_CANTIDAD, VENDET_PRECIO, VENDET_TIPO, VENDET_PRODCODIGO, VENDET_VENDPERID, VENDET_VENDTIPOID, VENDET_VENDTIPO, venDet_Descuento, vendet_descripcion, vendet_id, vendet_tiempo, vendet_asiento, vendet_prodDepen, vendet_FechaHora) " & _
                "VALUES (" & _
                "'" & lista_detalle.TextMatrix(1, 1) & "', '" & RES_PRODUCTOS.Fields("PROD_ID") & "', '" & RES_PRODUCTOS.Fields("PROD_SERV") & "', " & _
                "'" & RES_PRODUCTOS.Fields("PROD_NOMBRE") & "', '" & cantidad & "', " & _
                "'" & RES_PRODUCTOS.Fields("PROD_PRECIO") & "', 'V', '" & RES_PRODUCTOS.Fields("PROD_CODIGO") & "', " & _
                "'" & lblUserId(0).Caption & "', '" & lblUserId(1).Caption & "', 'U', '0', '" & BUSQ_ProdTouch.txtDescripcion.Text & "', '" & idLast & "', '" & BUSQ_ProdTouch.txt_Tiempo.Text & "', '" & Val(BUSQ_ProdTouch.txt_asiento.Text) & "', '" & RES_PRODUCTOS.Fields("PROD_DEPENDIENTE") & "', now())"
                con.Execute (sql1)
                           
                BUSQ_ProdTouch.List2.AddItem RES_PRODUCTOS.Fields("PROD_NOMBRE") & "   " & cantidad
            End If
            
            lista_detalle.Row = 1
            lista_detalle.Col = 1
            lista_detalle_Click
            lista_Mesa_Click
            BUSQ_ProdTouch.List1.Clear
            BUSQ_ProdTouch.txtDescripcion.Text = ""

            sql1 = "SELECT * FROM VIEW_VENTASDETALLE WHERE FOLIO = '" & lista_detalle.TextMatrix(1, 1) & "' order by  Producto asc"
            Set RES_DETALLE = con.Execute(sql1)
            Do While Not RES_DETALLE.EOF
                BUSQ_ProdTouch.List1.AddItem RES_DETALLE.Fields("PRODUCTO") & "    " & RES_DETALLE.Fields("CANTIDAD")
                RES_DETALLE.MoveNext
            Loop

        End If
    Else
        MsgBox "No hay mesa o cuenta seleccionada. Verifique. ", vbInformation
    End If

End Sub

Private Sub cmd_Abrir_Click()

    If mesaStatus(lista_Mesa.Row, lista_Mesa.Col) = "ABIERTO" Then
        MsgBox "Selección incorrecta. Verifique.", vbInformation
    Else
        If mesaStatus(lista_Mesa.Row, lista_Mesa.Col) = "DISPONIBLE" Then
            cmd_Abrir.Enabled = False
            cmd_AddProd.Enabled = True
            lista_Mesa.CellBackColor = vbRed
            abrir_Mesa
        End If
    End If
End Sub
Private Sub abrir_Mesa()
    Dim fila, columna As Integer
    
    sql1 = "SELECT PERTP_USUARIO, IF(PERTP_MEMBRESIA ='S', 'SI', 'NO') MEMBRESIA, PERTP_CODIGO_MEMBRESIA, PER_NOMBRE, PER_PATERNO, PER_MATERNO, PERTP_PER_TIPO, PERTP_TIPO_ID, CTPT_TIPO, T1.PER_ID, PER_FOTO, PER_EMAIL, t2.TEMP_MONEDERO, (SELECT T4.TOTAL FROM VIEW_MONEDERO_CLIENTES T4 WHERE T1.PER_ID = T4.PER_ID) TOTAL " & _
    "FROM PERSONA T1, PER_TIPO T2, CAT_TIPO T3 " & _
    "WHERE T1.PER_ID = T2.PERTP_PER_ID AND T2.PERTP_STATUS = 'A' AND T2.PERTP_PER_TIPO = 'C' " & _
    "AND T2.PERTP_TIPO_ID = T3.CTPT_ID AND T3.CTPT_SUBTIPO = 'C' AND " & _
    "T2.PERTP_CODIGO_MEMBRESIA = 'CLTE'"
    Set RES_VENTA = con.Execute(sql1)
    
    If Not RES_VENTA.EOF Then
    
        lblClieId(0).Caption = RES_VENTA.Fields("PER_ID")
        lblClieId(1).Caption = RES_VENTA.Fields("PERTP_TIPO_ID")
        lblClieId(2).Caption = RES_VENTA.Fields("PERTP_PER_TIPO")
    
        sql1 = "INSERT INTO VENTAS (VENT_FECHAHORA, VENT_STATUS, VENT_VENDPERID, VENT_VENDTIPOID, VENT_VENDTIPO, " & _
        "VENT_CLIEPERID, VENT_CLIETIPOID, VENT_CLIETIPO, VENT_MESA) VALUES " & _
        "('" & Format(Date, "yyyy-MM-dd") & " " & Format(Time, "HH:MM:SS") & "', 'G', '" & lblUserId(0).Caption & "', '" & lblUserId(1).Caption & "', 'U', " & _
        "'" & lblClieId(0).Caption & "', '" & lblClieId(1).Caption & "', '" & lblClieId(2).Caption & "', '" & mesaId(lista_Mesa.Row, lista_Mesa.Col) & "')"
        con.Execute (sql1)
    
    
        fila = lista_Mesa.Row
        columna = lista_Mesa.Col
        mesaStatus(fila, columna) = "ABIERTO"
        
        'carga_mesas

        lista_Mesa.Row = fila
        lista_Mesa.Col = columna
        lista_Mesa.CellFontBold = True
        lista_Mesa.CellFontSize = 16
        lista_Mesa.CellForeColor = vbBlue
        lista_Mesa_Click
    
    Else
        MsgBox "No se puede realizar la apertura. No existe el cliente general. Verifique. ", vbInformation
    End If
    
End Sub

Private Sub cmd_AddProd_Click()
        
    If Val(lista_detalle.TextMatrix(lista_detalle.Row, 4)) > 0 Then
        'MsgBox Val(lista_detalle.TextMatrix(lista_detalle.Row, 3))
        Timer_tiempo.Enabled = False
        BUSQ_ProdTouch.List1.Clear
        sql1 = "SELECT * FROM VIEW_VENTASDETALLE WHERE FOLIO = '" & lista_detalle.TextMatrix(1, 1) & "' order by  Producto asc"
        Set RES_DETALLE = con.Execute(sql1)
        Do While Not RES_DETALLE.EOF
            BUSQ_ProdTouch.List1.AddItem RES_DETALLE.Fields("PRODUCTO") & "    " & RES_DETALLE.Fields("CANTIDAD")
            RES_DETALLE.MoveNext
        Loop
        
        BUSQ_ProdTouch.Show 'vbModal
    Else
        MsgBox "Por favor indique el número de personas para la mesa. ", vbInformation
    End If
End Sub

Private Sub cmd_Bloquear_Click()
    Timer_tiempo.Enabled = False
    FRM_Identificador.Show vbModal
    cargaInicial
End Sub

Private Sub cmd_cerrar_Click()
'''Cobrar
    Timer_tiempo.Enabled = False

    tipoCobro = "OPERACIONES_TOUCH"
    'FRM_Cobro.txtTot.Text = FrmFocus.txtTotal.Text
    FRM_Cobro.txtTot.Text = lista_detalle.TextMatrix(1, 3)
    FRM_Cobro.Show vbModal

End Sub

Private Sub cmd_delProd_Click()
    If lista_Producto.Row > 0 Then
        quitar_Producto
    End If
End Sub
Private Sub quitar_Producto()
    Dim ques As String
    
    If lista_Producto.TextMatrix(lista_Producto.Row, 5) = "SI" Then
        MsgBox "El producto ya fue enviado a cocina. No se puede quitar. " & vbCrLf & vbCrLf & "Verifique", vbInformation
    Else
        ques = MsgBox("¿Quitar " & lista_Producto.TextMatrix(lista_Producto.Row, 0) & " mesa " & lista_detalle.TextMatrix(lista_detalle.Row, 0) & "?", vbYesNo + vbQuestion)
        If ques = vbYes Then
            sql1 = "DELETE FROM VENTA_DETALLE " & _
            "WHERE VENDET_FOLIO = '" & lista_detalle.TextMatrix(1, 1) & "' " & _
            "AND VENDET_PRODUCTOID = '" & lista_Producto.TextMatrix(lista_Producto.Row, 8) & "' AND VENDET_ID = '" & lista_Producto.TextMatrix(lista_Producto.Row, 10) & "'"
            'MsgBox SQL1
            con.Execute (sql1)
            
            If lista_Producto.Rows > 2 Then
                lista_Producto.RemoveItem (lista_Producto.Row)
            Else
                lista_Producto.Rows = 1
            End If
        End If
    End If
End Sub

Private Sub cmd_Nota_Click()
    FRM_NotaProducto.txtDescripcion.Text = lista_detalle.TextMatrix(1, 1) & vbCrLf & vbCrLf
    FRM_NotaProducto.txtDescripcion.SelStart = Len(FRM_NotaProducto.txtDescripcion.Text)
    FRM_NotaProducto.Show vbModal

End Sub



Private Sub cmd_Ticket_Click()
    generaTicket
End Sub
Private Sub generaTicket()
    Dim ticketSi As Boolean
'    Call nota_Mesa(lista_detalle.TextMatrix(lista_detalle.Row, 0))

    ticketSi = False
    For b1 = 1 To lista_Producto.Rows - 1
        If lista_Producto.TextMatrix(b1, 5) = "NO" Then
            ticketSi = True
            Exit For
        End If
    Next b1
    
    If ticketSi = True Then
        Call nota_Cocina(lista_detalle.TextMatrix(lista_detalle.Row, 1), "GENERAL")
            
        sql1 = "UPDATE VENTA_DETALLE SET VENDET_NOTAMESA = 'A', vendet_FechaHoraImpresion = now(), vendet_Impresiones = (vendet_Impresiones + 1)" & _
        "WHERE VENDET_FOLIO = '" & lista_detalle.TextMatrix(lista_detalle.Row, 1) & "' AND (VENDET_NOTAMESA <> 'A' OR VENDET_NOTAMESA IS NULL) "
        con.Execute (sql1)
        
        lista_detalle.Row = 1
        lista_detalle.Col = 1
        lista_detalle_Click
    End If
End Sub

Private Sub Command3_Click()
'    Timer_tiempo.Enabled = False
'    FRM_Identificador.Show vbModal
End Sub

Private Sub Command4_Click()

End Sub



Private Sub Down1_Click(Index As Integer)
    On Error Resume Next
    
    Select Case Index
        Case 0:
            lista_Mesa.TopRow = lista_Mesa.TopRow + 10
        Case 1:
            lista_Producto.TopRow = lista_Mesa.TopRow + 7
    End Select

End Sub

Private Sub Form_Load()
    'FRM_Identificador.Show vbModal
    'protector = True
    tiempo = 1
    checar_Cierre
    cargaInicial
    lista_Producto.ColWidth(10) = 0
End Sub
Private Sub checar_Cierre()
    sql1 = "SELECT suc_cierretouch, SUC_TEMP_PASSTOUCH from SUCURSAL "
    Set RES1 = con.Execute(sql1)
    
    If Not RES1.EOF Then
        If RES1.Fields("SUC_CIERRETOUCH") = "S" Then
            cierre = "SI"
        Else
            cierre = "NO"
        End If
        If RES1.Fields("SUC_TEMP_PASSTOUCH") >= 0 Then
            tiempoPass = RES1.Fields("SUC_TEMP_PASSTOUCH")
        Else
            tiempoPass = 0
        End If
    End If

End Sub

Private Sub bloqueo()
    
    cmd_Abrir.Enabled = False
    cmd_cerrar.Enabled = False
    cmd_Ticket.Enabled = False
    cmd_AddProd.Enabled = False
    cmd_delProd.Enabled = False
    cmd_Nota.Enabled = False
    
    
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tiempo = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If tipo_AccesoTouch = "Indentificador de usuario - Menu" Then
    Else
        Timer_tiempo.Enabled = False
        Unload Me
        End
    End If
End Sub

Private Sub lista_detalle_Click()
    carga_Producto (lista_detalle.TextMatrix(lista_detalle.Row, 1))

    cmd_delProd.Enabled = False
    If lista_detalle.Col = 4 Then
        FRM_CambiaCantidad.Caption = "Agregar personas"
        FRM_CambiaCantidad.txt_Cantidad.Text = lista_detalle.TextMatrix(lista_detalle.Row, 4)
        FRM_CambiaCantidad.Show vbModal
    End If
        
End Sub
Private Sub carga_Producto(folio As Long)
    
    sql1 = "SELECT * FROM VIEW_VENTASDETALLE WHERE FOLIO = '" & folio & "' AND STATUS = 'A' order by  NOTA, TIEMPO ASC "
    Set RES_DETALLE = con.Execute(sql1)
    
    With lista_Producto
        .Redraw = False
        .Rows = 1
        Do While Not RES_DETALLE.EOF
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = RES_DETALLE.Fields("PRODUCTO")
            .TextMatrix(.Rows - 1, 1) = RES_DETALLE.Fields("CANTIDAD")
            
            .TextMatrix(.Rows - 1, 2) = RES_DETALLE.Fields("TIEMPO") & ""
            .TextMatrix(.Rows - 1, 3) = RES_DETALLE.Fields("ASIENTO") & ""
            
            .TextMatrix(.Rows - 1, 4) = RES_DETALLE.Fields("DESCRIPCION") & ""
            .TextMatrix(.Rows - 1, 5) = RES_DETALLE.Fields("NOTA")
            
            .TextMatrix(.Rows - 1, 6) = FormatCurrency(RES_DETALLE.Fields("TOTAL"))
            .TextMatrix(.Rows - 1, 7) = FormatCurrency(RES_DETALLE.Fields("PRECIO"))
            
            .TextMatrix(.Rows - 1, 8) = RES_DETALLE.Fields("PROD_ID")
            .TextMatrix(.Rows - 1, 9) = RES_DETALLE.Fields("CODIGO")
            .TextMatrix(.Rows - 1, 10) = RES_DETALLE.Fields("VENDET_ID")
            .TextMatrix(.Rows - 1, 11) = RES_DETALLE.Fields("fechahora_prod") & ""
            
            .RowHeight(.Rows - 1) = 550
            RES_DETALLE.MoveNext

        Loop

        .Redraw = True
        .WordWrap = True
        .ColAlignment(1) = 4
        .ColAlignment(2) = 4
        .ColAlignment(3) = 4
        .ColAlignment(4) = 4
        .ColAlignment(5) = 4
        .ColAlignment(6) = 4
        .ColAlignment(7) = 4
        .ColAlignment(8) = 4
    End With
    
    
End Sub
Public Sub lista_Mesa_Click()
   ' On Error Resume Next

    carga_Detalle (lista_Mesa.TextMatrix(lista_Mesa.Row, lista_Mesa.Col))
    cmd_delProd.Enabled = False
    
End Sub
Public Sub carga_Detalle(MESA As Integer)
'On Error Resume Next
    If mesaStatus(lista_Mesa.Row, lista_Mesa.Col) = "ABIERTO" Then
        sql1 = "SELECT * FROM VIEW_VENTAS WHERE mesa = '" & MESA & "' AND STATUS = 'ABIERTO' ORDER BY STATUS ASC"
        Set RES_DETALLE = con.Execute(sql1)
        cmd_AddProd.Enabled = True
        cmd_Ticket.Enabled = True
        cmd_Abrir.Enabled = False
    Else
        If mesaStatus(lista_Mesa.Row, lista_Mesa.Col) = "DISPONIBLE" Then
'            sql1 = "SELECT * FROM VIEW_VENTAS WHERE date_format(FechaHora_GENERA, '%Y-%m-%d') = date_format(now(),'%Y-%m-%d') " & _
'            "and mesa = '" & MESA & "' ORDER BY STATUS ASC"
            lista_detalle.Rows = 1
            lista_Producto.Rows = 1
            
            cmd_AddProd.Enabled = False
            cmd_Ticket.Enabled = False
            cmd_Abrir.Enabled = True
            Exit Sub
        End If
    End If
    
    With lista_detalle
    .Redraw = False
    .Rows = 1
    lista_Producto.Rows = 1
    Do While Not RES_DETALLE.EOF
        .AddItem ""
        .TextMatrix(.Rows - 1, 0) = RES_DETALLE.Fields("MESA")
        .TextMatrix(.Rows - 1, 1) = RES_DETALLE.Fields("FOLIO")
        .TextMatrix(.Rows - 1, 2) = RES_DETALLE.Fields("CLIENTE")
        .TextMatrix(.Rows - 1, 3) = FormatCurrency(RES_DETALLE.Fields("TOTAL"))
        .TextMatrix(.Rows - 1, 4) = RES_DETALLE.Fields("PERSONAS") & ""
        
        .TextMatrix(.Rows - 1, 5) = Format(RES_DETALLE.Fields("FECHAHORA_GENERA"), "DDDD-DD-MMM  HH:MM")
        .TextMatrix(.Rows - 1, 7) = RES_DETALLE.Fields("STATUS")
'        .TextMatrix(.Rows - 1, 7) = RES_DETALLE.Fields("MESA")
        .TextMatrix(.Rows - 1, 8) = RES_DETALLE.Fields("USUARIO")
        .TextMatrix(.Rows - 1, 9) = RES_DETALLE.Fields("OBSERVACIONES") & ""
        If RES_DETALLE.Fields("status") = "CERRADO" Then
            .Row = .Rows - 1
            For b1 = 0 To .Cols - 1
                .Col = b1
                .CellBackColor = &H80FF80
            Next b1
        .TextMatrix(.Rows - 1, 6) = Format(RES_DETALLE.Fields("FECHAHORA_DOS"), "DDDD-DD-MMM  HH:MM")
        
        Else
            If RES_DETALLE.Fields("status") = "ABIERTO" Then
                .Row = .Rows - 1
                For b1 = 0 To .Cols - 1
                    .Col = b1
                    .CellBackColor = &HC0C0FF
                Next b1
            '.TextMatrix(.Rows - 1, 7) = ""
            
            End If
        
        End If
        
        .Row = lista_detalle.Rows - 1
        .RowHeight(lista_detalle.Rows - 1) = 650
        .ColAlignment(0) = 4
        .ColAlignment(1) = 4
        .ColAlignment(4) = 4
        RES_DETALLE.MoveNext
    Loop
        .Redraw = True
        If .Rows > 1 Then
            .Row = 1
            .Col = 1
            lista_detalle_Click
        End If
    .WordWrap = True
    
    End With

    If lista_Producto.Rows > 1 Then
        If lista_detalle.TextMatrix(1, 7) = "ABIERTO" Then
            If cierre = "SI" Then
                cmd_cerrar.Enabled = True
            Else
                cmd_cerrar.Enabled = False
            End If
        End If
    Else
        cmd_cerrar.Enabled = False
    End If

End Sub

Private Sub actualizaCantidad()

    sql1 = "UPDATE VENTA_DETALLE SET VENDET_CANTIDAD = '" & lista_Producto.TextMatrix(b1, 1) & "' " & _
    "WHERE VENDET_FOLIO = '" & lista_detalle.TextMatrix(1, 1) & "' AND VENDET_PRODUCTOID = '" & lista_Producto.TextMatrix(b1, 5) & "' "
    con.Execute (sql1)

End Sub

Private Sub lista_Mesa_GotFocus()
    ConScroll lista_Mesa
End Sub

Private Sub lista_Mesa_LostFocus()
    SinScroll lista_Mesa
End Sub

Private Sub lista_Producto_Click()
    If lista_Producto.Rows > 1 Then
'        If lista_Producto.Col = 1 Then
'            'FRM_CambiaCantidad.txt_Cantidad.Text = lista_Producto.TextMatrix(lista_Producto.Row, lista_Producto.Col)
'            FRM_CambiaCantidad.txt_Cantidad.Text = "1"
'            FRM_CambiaCantidad.Label1.Caption = lista_Producto.TextMatrix(lista_Producto.Row, 0)
'            FRM_CambiaCantidad.Show vbModal
'        End If
        If lista_Producto.TextMatrix(lista_Producto.Row, 5) = "NO" Then
            If lista_Producto.Col = 1 Then
                FRM_CambiaCantidad.txt_Cantidad.Text = lista_Producto.TextMatrix(lista_Producto.Row, 1)
                'FRM_CambiaCantidad.txt_Cantidad.Text = "1"
                'FRM_CambiaCantidad.Label1.Caption = lista_Producto.TextMatrix(lista_Producto.Row, 0)
                FRM_CambiaCantidad.Show vbModal
            Else
                If lista_Producto.Col = 2 Then
                    FRM_CambiaCantidad.txt_Cantidad.Text = lista_Producto.TextMatrix(lista_Producto.Row, 2)
                    FRM_CambiaCantidad.Caption = "Agregar tiempo"
                    'FRM_CambiaCantidad.txt_Cantidad.Text = "1"
                    'FRM_CambiaCantidad.Label1.Caption = lista_Producto.TextMatrix(lista_Producto.Row, 0)
                    FRM_CambiaCantidad.Show vbModal
                Else
                    If lista_Producto.Col = 3 Then
                        FRM_CambiaCantidad.txt_Cantidad.Text = lista_Producto.TextMatrix(lista_Producto.Row, 3)
                        FRM_CambiaCantidad.Caption = "Agregar asiento"
                        'FRM_CambiaCantidad.txt_Cantidad.Text = "1"
                        'FRM_CambiaCantidad.Label1.Caption = lista_Producto.TextMatrix(lista_Producto.Row, 0)
                        FRM_CambiaCantidad.Show vbModal
                    Else
                        If lista_Producto.Col = 4 Then
                            'FRM_CambiaCantidad.txt_Cantidad.Text = lista_Producto.TextMatrix(lista_Producto.Row, lista_Producto.Col)
                            FRM_NotaProducto.txtDescripcion.Text = lista_Producto.TextMatrix(lista_Producto.Row, 4)
                            'FRM_NotaProducto.Label1.Caption = lista_Producto.TextMatrix(lista_Producto.Row, 0)
                            FRM_NotaProducto.Show
                        End If
                    End If
                End If
            End If
            cmd_delProd.Enabled = True
        Else
            cmd_delProd.Enabled = False
        End If
    End If
End Sub

Private Sub lista_Producto_DblClick()
    If lista_Producto.Col = 5 Then
        If lista_Producto.TextMatrix(lista_Producto.Row, 5) = "SI" Then
        
            cancelarMotivo = "TICKET"
            FRM_Cancelar.Show vbModal
       
        
'            sql1 = "UPDATE VENTA_DETALLE SET VENDET_NOTAMESA = NULL WHERE VENDET_FOLIO = '" & lista_detalle.TextMatrix(lista_detalle.Row, 1) & "' AND VENDET_NOTAMESA = 'A' AND VENDET_PRODUCTOID = '" & lista_Producto.TextMatrix(lista_Producto.Row, 8) & "' AND VENDET_ID = '" & lista_Producto.TextMatrix(lista_Producto.Row, 10) & "'"
'            con.Execute (sql1)
            'MsgBox SQL1
'            lista_Producto.TextMatrix(lista_Producto.Row, 5) = "NO"
        End If
    End If

'    If lista_Producto.TextMatrix(lista_Producto.Row, 6) = "No" Then
'        If lista_Producto.Col = 1 Then
'            'FRM_CambiaCantidad.txt_Cantidad.Text = lista_Producto.TextMatrix(lista_Producto.Row, lista_Producto.Col)
'            FRM_CambiaCantidad.txt_Cantidad.Text = "1"
'            FRM_CambiaCantidad.Label1.Caption = lista_Producto.TextMatrix(lista_Producto.Row, 0)
'            FRM_CambiaCantidad.Show vbModal
'        Else
'            If lista_Producto.Col = 3 Then
'                'FRM_CambiaCantidad.txt_Cantidad.Text = lista_Producto.TextMatrix(lista_Producto.Row, lista_Producto.Col)
'                FRM_NotaProducto.txtNota.Text = "1"
'                FRM_NotaProducto.Label1.Caption = lista_Producto.TextMatrix(lista_Producto.Row, 0)
'                FRM_NotaProducto.Show vbModal
'            End If
'        End If
'    End If
End Sub

Private Sub lista_Producto_GotFocus()
    ConScroll lista_Producto
End Sub

Private Sub lista_Producto_LostFocus()
    SinScroll lista_Producto
End Sub

Private Sub Timer_tiempo_Timer()
    If tiempoPass > 0 Then
        
        If SeMueve Then
            tiempo = 0
        Else
            tiempo = tiempo + 1
        End If
        
        Label3.Caption = tiempo
        Label3.Refresh
        
        If tiempo = tiempoPass Then
            Timer_tiempo.Enabled = False
            cargaInicial
            Unload BUSQ_ProdTouch
            FRM_Identificador.Show vbModal
        End If
    Else
         Timer_tiempo.Enabled = False
         
    End If
End Sub

Private Sub Timer1_Timer()
    
    Timer1.Enabled = False
    lista_Mesa.height = Me.height - 500
    lista_Producto.width = Me.width - 5500
    lista_Producto.height = Me.height - 3600
    lista_detalle.width = Me.width - 5500
    
End Sub

Private Sub cargaInicial()
    
    bloqueo
    lista_Mesa.Rows = 1
    lista_detalle.Rows = 1
    lista_Producto.Rows = 1
    lista_Producto.ColWidth(8) = 0
    
    lista_Mesa.ColAlignment(0) = 3
    lista_Mesa.MergeCol(0) = True
    
    carga_mesas
    check_asiento
End Sub
Private Sub check_asiento()
    sql1 = "SELECT SUC_ASIENTOS FROM SUCURSAL "
    Set RES1 = con.Execute(sql1)
    If Not RES1.EOF Then
        If RES1.Fields("SUC_ASIENTOS") = "N" Then
            'lista_detalle.ColWidth(3) = 0
            lista_Producto.ColWidth(3) = 0
        End If
        
    End If
    
    
End Sub
Public Sub carga_mesas()

    On Error Resume Next
    sql1 = "select * From view_mesas_estado order by mesa_id"
    Set RES_MESA = con.Execute(sql1)

    Dim num_mesa As Integer
    Dim filas As Integer
    lista_Mesa.Redraw = False
    
    lista_Mesa.MergeCells = flexMergeRestrictRows
    lista_Mesa.MergeRow(0) = True

'     lista_Mesa.Rows = 10
     
     filas = Round((RES_MESA.RecordCount / 2), 0)
     lista_Mesa.Rows = filas + 1
     lista_Mesa.Cols = 2
     
     num_mesa = 0
     
'    For b1 = 1 To filas
    b1 = 0
    Do While Not RES_MESA.EOF
        b1 = b1 + 1
        lista_Mesa.Row = b1
        lista_Mesa.RowHeight(b1) = 650
        
        lista_Mesa.TextMatrix(b1, 0) = RES_MESA.Fields("MESA_ID")
        mesaId(b1, 0) = RES_MESA.Fields("MESA_ID")
        
        lista_Mesa.Col = 0
        lista_Mesa.CellAlignment = 4
        lista_Mesa.CellFontSize = 14
        lista_Mesa.CellFontBold = True
        lista_Mesa.CellForeColor = vbBlack
        
        If RES_MESA.Fields("estado") = "DISPONIBLE" Then
            lista_Mesa.CellBackColor = &H80FF80
            mesaStatus(b1, 0) = "DISPONIBLE"
            
        Else
            If RES_MESA.Fields("estado") = "ABIERTO" Then
                lista_Mesa.CellBackColor = &HC0C0FF
                mesaStatus(b1, 0) = "ABIERTO"
            End If
        End If
        
        RES_MESA.MoveNext
        lista_Mesa.TextMatrix(b1, 1) = RES_MESA.Fields("MESA_ID")
        mesaId(b1, 1) = RES_MESA.Fields("MESA_ID")
        lista_Mesa.Col = 1
        lista_Mesa.CellAlignment = 4
        lista_Mesa.CellFontSize = 14
        lista_Mesa.CellFontBold = True
        lista_Mesa.CellForeColor = vbBlack
        If RES_MESA.Fields("estado") = "DISPONIBLE" Then
            lista_Mesa.CellBackColor = &H80FF80
            mesaStatus(b1, 1) = "DISPONIBLE"
        Else
            If RES_MESA.Fields("estado") = "ABIERTO" Then
                lista_Mesa.CellBackColor = &HC0C0FF
                mesaStatus(b1, 1) = "ABIERTO"
            End If
        End If
        
        num_mesa = num_mesa + 2
        RES_MESA.MoveNext
    Loop
'    Next b1
    lista_Mesa.Redraw = True
    
    
End Sub

Private Sub Up1_Click(Index As Integer)
    On Error Resume Next
    
    Select Case Index
        Case 0:
            If lista_Mesa.TopRow < 6 Then
                lista_Mesa.TopRow = 1
            Else
                lista_Mesa.TopRow = lista_Mesa.TopRow - 11
            End If
        Case 1:
            If lista_Producto.TopRow < 4 Then
                lista_Producto.TopRow = 1
            Else
                lista_Producto.TopRow = lista_Producto.TopRow - 4
            End If
    End Select

End Sub
