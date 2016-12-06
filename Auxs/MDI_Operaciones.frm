VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm MDI_Operaciones 
   BackColor       =   &H8000000C&
   Caption         =   "Operaciones"
   ClientHeight    =   9915
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   16920
   Icon            =   "MDI_Operaciones.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar barraOper 
      Align           =   3  'Align Left
      Height          =   9540
      Left            =   0
      TabIndex        =   0
      Top             =   375
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   16828
      Appearance      =   1
      _Version        =   327682
      Begin VB.Image imgBtn 
         Height          =   975
         Index           =   2
         Left            =   120
         Top             =   360
         Width           =   1695
      End
      Begin VB.Image imgBtn 
         Height          =   975
         Index           =   21
         Left            =   120
         Top             =   7920
         Width           =   1695
      End
      Begin VB.Image imgBtn 
         Height          =   720
         Index           =   22
         Left            =   120
         Picture         =   "MDI_Operaciones.frx":058A
         Stretch         =   -1  'True
         Top             =   8040
         Width           =   1635
      End
      Begin VB.Image imgBtn 
         Height          =   975
         Index           =   18
         Left            =   120
         Top             =   6840
         Width           =   1695
      End
      Begin VB.Image imgBtn 
         Height          =   720
         Index           =   19
         Left            =   120
         Picture         =   "MDI_Operaciones.frx":1777
         Stretch         =   -1  'True
         Top             =   6960
         Width           =   1635
      End
      Begin VB.Image imgBtn 
         Height          =   975
         Index           =   17
         Left            =   90
         Top             =   5760
         Width           =   1695
      End
      Begin VB.Image imgBtn 
         Height          =   720
         Index           =   16
         Left            =   120
         Picture         =   "MDI_Operaciones.frx":27F7
         Stretch         =   -1  'True
         Top             =   5880
         Width           =   1635
      End
      Begin VB.Image imgBtn 
         Height          =   720
         Index           =   15
         Left            =   120
         Picture         =   "MDI_Operaciones.frx":3DC8
         Stretch         =   -1  'True
         Top             =   5880
         Width           =   1635
      End
      Begin VB.Image imgBtn 
         Height          =   975
         Index           =   14
         Left            =   90
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Image imgBtn 
         Height          =   975
         Index           =   13
         Left            =   90
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Image imgBtn 
         Height          =   720
         Index           =   12
         Left            =   120
         Picture         =   "MDI_Operaciones.frx":54BA
         Stretch         =   -1  'True
         Top             =   4800
         Width           =   1635
      End
      Begin VB.Image imgBtn 
         Height          =   720
         Index           =   11
         Left            =   90
         Picture         =   "MDI_Operaciones.frx":6A73
         Stretch         =   -1  'True
         Top             =   4800
         Width           =   1635
      End
      Begin VB.Image imgBtn 
         Height          =   720
         Index           =   10
         Left            =   120
         Picture         =   "MDI_Operaciones.frx":816D
         Stretch         =   -1  'True
         Top             =   3720
         Width           =   1635
      End
      Begin VB.Image imgBtn 
         Height          =   720
         Index           =   9
         Left            =   120
         Picture         =   "MDI_Operaciones.frx":9808
         Stretch         =   -1  'True
         Top             =   3720
         Width           =   1635
      End
      Begin VB.Image imgBtn 
         Height          =   975
         Index           =   8
         Left            =   90
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Image imgBtn 
         Height          =   720
         Index           =   7
         Left            =   120
         Picture         =   "MDI_Operaciones.frx":AFDD
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   1635
      End
      Begin VB.Image imgBtn 
         Height          =   720
         Index           =   6
         Left            =   120
         Picture         =   "MDI_Operaciones.frx":BF31
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   1635
      End
      Begin VB.Image imgBtn 
         Height          =   720
         Index           =   1
         Left            =   120
         Picture         =   "MDI_Operaciones.frx":D0E9
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1635
      End
      Begin VB.Image imgBtn 
         Height          =   975
         Index           =   5
         Left            =   90
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Image imgBtn 
         Height          =   720
         Index           =   4
         Left            =   120
         Picture         =   "MDI_Operaciones.frx":E4A6
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   1635
      End
      Begin VB.Image imgBtn 
         Height          =   720
         Index           =   3
         Left            =   120
         Picture         =   "MDI_Operaciones.frx":F2C6
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   1635
      End
      Begin VB.Image imgBtn 
         Height          =   720
         Index           =   0
         Left            =   120
         Picture         =   "MDI_Operaciones.frx":1034F
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1635
      End
      Begin VB.Image imgBtn 
         Height          =   720
         Index           =   20
         Left            =   150
         Picture         =   "MDI_Operaciones.frx":118AF
         Stretch         =   -1  'True
         Top             =   6960
         Width           =   1635
      End
      Begin VB.Image imgBtn 
         Height          =   720
         Index           =   23
         Left            =   120
         Picture         =   "MDI_Operaciones.frx":12B5D
         Stretch         =   -1  'True
         Top             =   8040
         Width           =   1635
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   16920
      _ExtentX        =   29845
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   22093
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            AutoSize        =   2
            Bevel           =   0
            TextSave        =   "08:54 p.m."
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Bevel           =   0
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
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
   Begin VB.Menu mn_Oper 
      Caption         =   "Operaciones"
      Begin VB.Menu mn_NewOper2 
         Caption         =   "Nueva operación (2)"
         Visible         =   0   'False
      End
      Begin VB.Menu mn_NewOper 
         Caption         =   "Nueva operacion"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mn_Tickets 
         Caption         =   "Tickets del día"
      End
      Begin VB.Menu mn_PrintPreTicket2 
         Caption         =   "Imprimir preticket"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mn_OpenCajon 
         Caption         =   "Abrir cajon"
         Shortcut        =   ^O
      End
      Begin VB.Menu mn_Line1 
         Caption         =   "-"
      End
      Begin VB.Menu mn_Cobrar 
         Caption         =   "Cobrar"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mn_CancelarOper 
         Caption         =   "Cancelar"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mn_Guardar 
         Caption         =   "Guardar"
         Enabled         =   0   'False
      End
      Begin VB.Menu mn_Line2 
         Caption         =   "-"
      End
      Begin VB.Menu mn_Exit 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mn_Negocio 
      Caption         =   "Negocio"
      Begin VB.Menu mn_Membresia 
         Caption         =   "Membresias"
      End
      Begin VB.Menu mn_line4 
         Caption         =   "-"
      End
      Begin VB.Menu mn_ConsumoInterno 
         Caption         =   "Consumo interno"
      End
      Begin VB.Menu mn_Line7 
         Caption         =   "-"
      End
      Begin VB.Menu mn_Asistencias 
         Caption         =   "Asistencias"
      End
      Begin VB.Menu mn_line5 
         Caption         =   "-"
      End
      Begin VB.Menu mn_Productos 
         Caption         =   "Productos"
      End
      Begin VB.Menu mn_CLientes 
         Caption         =   "Clientes"
      End
      Begin VB.Menu mn_AgenCalen 
         Caption         =   "Agenda / Calendario"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mn_Servicios 
         Caption         =   "Servicios"
      End
      Begin VB.Menu mn_Line3 
         Caption         =   "-"
      End
      Begin VB.Menu mn_GastEgre 
         Caption         =   "Gastos"
      End
      Begin VB.Menu mn_Anticipos 
         Caption         =   "Apartados"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mn_Devoluciones 
         Caption         =   "Cambios - Devoluciones"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mn_Credito 
         Caption         =   "Crédito"
      End
   End
   Begin VB.Menu mn_Búsqueda 
      Caption         =   "Búsqueda"
      Begin VB.Menu mn_BusqProdu 
         Caption         =   "Búsqueda de Productos"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mn_BusqClientes 
         Caption         =   "Búsqueda de Clientes"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mn_BusqUsuario 
         Caption         =   "Búsqueda de Usuarios"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mn_Config 
      Caption         =   "Configuración"
      Begin VB.Menu mn_Proteger 
         Caption         =   "Proteger"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mn_Impresora 
         Caption         =   "Impresora"
      End
   End
   Begin VB.Menu mn_Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mn_Cancel 
         Caption         =   "Cancelar"
      End
      Begin VB.Menu mn_CancelAll 
         Caption         =   "Cancelar todo"
      End
      Begin VB.Menu mn_line6 
         Caption         =   "-"
      End
      Begin VB.Menu mn_AddServUsu 
         Caption         =   "Agregar servicio con otro usuario"
      End
      Begin VB.Menu mn_Line8 
         Caption         =   "-"
      End
      Begin VB.Menu mn_AddDesc_Other 
         Caption         =   "Agregar este valor de descuento a los demas registros"
      End
      Begin VB.Menu mn_AddDesc_OtherProcen 
         Caption         =   "Agregar este valor de descuento a los demas registros"
      End
   End
   Begin VB.Menu mn_Windows 
      Caption         =   "Ventanas operaciones"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mn_TicketsPrint 
      Caption         =   "Menu_Tickets"
      Visible         =   0   'False
      Begin VB.Menu mn_PrintTicket 
         Caption         =   "Imprimir ticket"
      End
      Begin VB.Menu mn_PrintPreTicket 
         Caption         =   "Imprimir pre ticket"
         Visible         =   0   'False
      End
      Begin VB.Menu mn_LineCancel 
         Caption         =   "-"
      End
      Begin VB.Menu mn_CancelOperTicket 
         Caption         =   "Cancelar operación"
      End
   End
End
Attribute VB_Name = "MDI_Operaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql1 As String
Dim RES1 As Recordset
Dim resProd As Recordset

Private Sub barraOper_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imgBtn(1).Visible = True Then
        imgBtn(1).Visible = False
        imgBtn(0).Visible = True
    End If
    If imgBtn(4).Visible = True Then
        imgBtn(4).Visible = False
        imgBtn(3).Visible = True
    End If
    If imgBtn(7).Visible = True Then
        imgBtn(7).Visible = False
        imgBtn(6).Visible = True
    End If
    If imgBtn(10).Visible = True Then
        imgBtn(10).Visible = False
        imgBtn(9).Visible = True
    End If
    If imgBtn(12).Visible = True Then
        imgBtn(12).Visible = False
        imgBtn(11).Visible = True
    End If
    If imgBtn(16).Visible = True Then
        imgBtn(16).Visible = False
        imgBtn(15).Visible = True
    End If
    If imgBtn(19).Visible = True Then
        imgBtn(19).Visible = False
        imgBtn(20).Visible = True
    End If
    If imgBtn(22).Visible = True Then
        imgBtn(22).Visible = False
        imgBtn(23).Visible = True
    End If


End Sub

Private Sub imgBtn_Click(Index As Integer)
    Select Case Index
        Case 2: mn_NewOper_Click
        Case 5: mn_Cobrar_Click
        Case 13: mn_BusqProdu_Click
        Case 14: mn_BusqClientes_Click
        Case 17: mn_BusqUsuario_Click
        Case 8: mn_CancelAll_Click
        Case 18: mn_Anticipos_Click
        Case 21: mn_Devoluciones_Click
    End Select
    
End Sub

Private Sub imgBtn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 2: imgBtn(0).Visible = False
                imgBtn(1).Visible = True
        Case 5: imgBtn(3).Visible = False
                imgBtn(4).Visible = True
        Case 8: imgBtn(6).Visible = False
                imgBtn(7).Visible = True
        Case 13: imgBtn(9).Visible = False
                imgBtn(10).Visible = True
        Case 14: imgBtn(11).Visible = False
                imgBtn(12).Visible = True
        Case 17: imgBtn(15).Visible = False
                imgBtn(16).Visible = True
        Case 18: imgBtn(20).Visible = False
                imgBtn(19).Visible = True
        Case 21: imgBtn(23).Visible = False
                imgBtn(22).Visible = True
        
    End Select
    
    
End Sub

Private Sub MDIForm_Load()
'    mn_NewOper_Click
    numFrmTicket = 0
    numFrmOper = 0
    tikcet = False
    StatusBar1.Panels(1).Text = Format(Date, "Long Date")
    StatusBar1.Panels(3).Text = FRM_Menu.menuBarra2.Panels(10).Text
    MDI_Operaciones.StatusBar1.Panels(4).Text = "0"
    mn_tickets_Click
    'mn_Cobrar.Enabled = False
    checkImags


End Sub

Private Sub checkImags()
    If imgBtn(1).Visible = True Then
        imgBtn(1).Visible = False
        imgBtn(0).Visible = True
    End If
    If imgBtn(4).Visible = True Then
        imgBtn(4).Visible = False
        imgBtn(3).Visible = True
    End If
    If imgBtn(7).Visible = True Then
        imgBtn(7).Visible = False
        imgBtn(6).Visible = True
    End If
    If imgBtn(10).Visible = True Then
        imgBtn(10).Visible = False
        imgBtn(9).Visible = True
    End If
    If imgBtn(12).Visible = True Then
        imgBtn(12).Visible = False
        imgBtn(11).Visible = True
    End If
    If imgBtn(16).Visible = True Then
        imgBtn(16).Visible = False
        imgBtn(15).Visible = True
    End If
    If imgBtn(19).Visible = True Then
        imgBtn(19).Visible = False
        imgBtn(20).Visible = True
    End If
    If imgBtn(22).Visible = True Then
        imgBtn(22).Visible = False
        imgBtn(23).Visible = True
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    numFrmOper = 0
End Sub

Private Sub mn_AddDesc_Other_Click()
    ''''
    Dim valDesc As String
    
    valDesc = FrmFocus.lista.TextMatrix(FrmFocus.lista.Row, 11)
    For b1 = 1 To FrmFocus.lista.Rows - 1
        FrmFocus.textDesc.Text = Format(valDesc, "General Number")
        FrmFocus.lista.Col = 11
        FrmFocus.lista.Row = b1
        FrmFocus.lista.TextMatrix(b1, 11) = valDesc
        FrmFocus.textDesc_KeyPress (13)
        
    Next b1
    'MsgBox "Ok"
End Sub

Private Sub mn_AddDesc_OtherProcen_Click()
    Dim valDesc As String
    
    valDesc = FrmFocus.lista.TextMatrix(FrmFocus.lista.Row, 12)
    
    
    
    For b1 = 1 To FrmFocus.lista.Rows - 1
        FrmFocus.textDesc.Visible = True
        FrmFocus.textDesc.Text = Format(valDesc, "General Number")
        FrmFocus.lista.Col = 12
        FrmFocus.lista.Row = b1
        FrmFocus.lista.TextMatrix(b1, 12) = valDesc
        FrmFocus.textDesc_KeyPress (13)
        
    Next b1
        FrmFocus.textDesc.Visible = False

End Sub

Private Sub mn_AddServUsu_Click()
    Dim question As String
    question = MsgBox("¿Agregar: " & FrmFocus.lista.TextMatrix(FrmFocus.lista.Row, 0) & "  " & FrmFocus.lista.TextMatrix(FrmFocus.lista.Row, 2) & " con otro usuario? ", vbYesNo + vbQuestion)
    If question = vbYes Then
        aregarFilaUsuario (FrmFocus.lista.Row)
    End If
End Sub

Private Sub aregarFilaUsuario(fila As Long)
    
    sql1 = "SELECT T4.PERTP_PER_ID, T4.PERTP_TIPO_ID, concat(T4.PERTP_PER_ID, T4.PERTP_TIPO_ID) USERID,  " & _
    "CONCAT(T2.PER_NOMBRE, ' ', T2.PER_PATERNO, ' ', T2.PER_MATERNO) USUARIO " & _
    "FROM PERSONA T2, PER_tIPO T4 " & _
    "WHERE T2.PER_ID = T4.PERTP_PER_ID AND T4.PERTP_STATUS = 'A' AND T4.PERTP_PER_TIPO = 'U' " & _
    "AND T4.PERTP_PER_ID <> '" & FrmFocus.lista.TextMatrix(FrmFocus.lista.Row, 10) & "'"
    Set RES1 = con.Execute(sql1)
    
    If Not RES1.EOF Then
        FrmFocus.lista.AddItem ""
        FrmFocus.lista.TextMatrix(FrmFocus.lista.Rows - 1, 0) = FrmFocus.lista.TextMatrix(fila, 0)
        FrmFocus.lista.TextMatrix(FrmFocus.lista.Rows - 1, 1) = FrmFocus.lista.TextMatrix(fila, 1)
        FrmFocus.lista.TextMatrix(FrmFocus.lista.Rows - 1, 2) = FrmFocus.lista.TextMatrix(fila, 2)
        FrmFocus.lista.TextMatrix(FrmFocus.lista.Rows - 1, 3) = FrmFocus.lista.TextMatrix(fila, 3)
        FrmFocus.lista.TextMatrix(FrmFocus.lista.Rows - 1, 4) = FrmFocus.lista.TextMatrix(fila, 4)
        FrmFocus.lista.TextMatrix(FrmFocus.lista.Rows - 1, 5) = FrmFocus.lista.TextMatrix(fila, 5)
        FrmFocus.lista.TextMatrix(FrmFocus.lista.Rows - 1, 6) = FrmFocus.lista.TextMatrix(fila, 6)
        FrmFocus.lista.TextMatrix(FrmFocus.lista.Rows - 1, 7) = FrmFocus.lista.TextMatrix(fila, 7)
        FrmFocus.lista.TextMatrix(FrmFocus.lista.Rows - 1, 8) = RES1.Fields("USUARIO")
        FrmFocus.lista.TextMatrix(FrmFocus.lista.Rows - 1, 9) = RES1.Fields("PERTP_TIPO_ID")
        FrmFocus.lista.TextMatrix(FrmFocus.lista.Rows - 1, 10) = RES1.Fields("PERTP_PER_ID")
        FrmFocus.lista.TextMatrix(FrmFocus.lista.Rows - 1, 11) = FrmFocus.lista.TextMatrix(fila, 11)
        FrmFocus.lista.TextMatrix(FrmFocus.lista.Rows - 1, 12) = FrmFocus.lista.TextMatrix(fila, 12)
        FrmFocus.lista.TextMatrix(FrmFocus.lista.Rows - 1, 13) = FrmFocus.lista.TextMatrix(fila, 13)
        FrmFocus.lista.TextMatrix(FrmFocus.lista.Rows - 1, 14) = FrmFocus.lista.TextMatrix(fila, 14)
        FrmFocus.lista.TextMatrix(FrmFocus.lista.Rows - 1, 15) = FrmFocus.lista.TextMatrix(fila, 15)
        FrmFocus.lista.TextMatrix(FrmFocus.lista.Rows - 1, 16) = FrmFocus.lista.TextMatrix(fila, 16)
        FrmFocus.lista.TextMatrix(FrmFocus.lista.Rows - 1, 17) = FrmFocus.lista.TextMatrix(fila, 17)
        FrmFocus.lista.TextMatrix(FrmFocus.lista.Rows - 1, 18) = FrmFocus.lista.TextMatrix(fila, 18)
        FrmFocus.lista.TextMatrix(FrmFocus.lista.Rows - 1, 19) = FrmFocus.lista.TextMatrix(fila, 19)
        
        
'    lista.AddItem ""
'    lista.TextMatrix(lista.Rows - 1, 0) = RES1.Fields("TIPO_PROD")
'    lista.TextMatrix(lista.Rows - 1, 1) = RES1.Fields("PROD_CODIGO")
'    lista.TextMatrix(lista.Rows - 1, 2) = RES1.Fields("PROD_NOMBRE")
'    lista.TextMatrix(lista.Rows - 1, 3) = "1"
'    lista.TextMatrix(lista.Rows - 1, 4) = FormatCurrency(RES1.Fields("PROD_PRECIO"))
'    lista.TextMatrix(lista.Rows - 1, 6) = RES1.Fields("PROD_SERV")
'    lista.TextMatrix(lista.Rows - 1, 7) = RES1.Fields("PROD_ID")
'    lista.TextMatrix(lista.Rows - 1, 8) = lblDatos(1).Caption
'    lista.TextMatrix(lista.Rows - 1, 9) = lblUserId(1).Caption
'    lista.TextMatrix(lista.Rows - 1, 10) = lblUserId(0).Caption
'    If RES1.Fields("PROD_APLICADESC") = "S" Then
'        lista.TextMatrix(lista.Rows - 1, 11) = FormatCurrency(Val(RES1.Fields("PROD_PRECIO")) - (Val(RES1.Fields("PROD_PRECIODESC"))))
'        valor = Val(Format(lista.TextMatrix(lista.Rows - 1, 11), "General number")) * ((100) / (Val(Format(lista.TextMatrix(lista.Rows - 1, 4), "General Number"))))
'        lista.TextMatrix(lista.Rows - 1, 12) = Round(valor, 2)
''        lista.TextMatrix(lista.Rows - 1, 14) = FormatCurrency(RES1.Fields("PROD_PRECIO"))
'
'    Else
'        lista.TextMatrix(lista.Rows - 1, 11) = FormatCurrency(0)
'        lista.TextMatrix(lista.Rows - 1, 12) = Round(0, 2)
'        lista.TextMatrix(lista.Rows - 1, 14) = FormatCurrency(RES1.Fields("PROD_PRECIO"))
'    End If
'    lista.TextMatrix(lista.Rows - 1, 13) = RES1.Fields("prod_DESCRIPCION")
'    If RES1.Fields("PROD_SERV") = "P" Then
'        lista.TextMatrix(lista.Rows - 1, 18) = RES1.Fields("PROD_DEPENDIENTE")
'        lista.TextMatrix(lista.Rows - 1, 19) = RES1.Fields("SUBTIPO")
'    Else
'        lista.TextMatrix(lista.Rows - 1, 18) = "N"
'        lista.TextMatrix(lista.Rows - 1, 19) = "N"
'    End If
'
'
'
'    mone = Val(Format(lblDatos(6).Caption, "General Number"))
'    moneus = Val(Format(lblDatos(7).Caption, "General Number"))
'
'    If monedero = True Then
'        If mone - moneus > 0 Then
'            lista.TextMatrix(lista.Rows - 1, 15) = "MND"
'        Else
'            lista.TextMatrix(lista.Rows - 1, 15) = ""
'        End If
'    End If
'
'    lista.Row = lista.Rows - 1
'    lista.Col = 16
'    lista.CellFontName = "Wingdings"
'    lista.CellFontBold = True
'    lista.CellFontSize = 16
'    lista.TextMatrix(lista.Rows - 1, 16) = Chr(168)
''    lista.TextMatrix(lista.Rows - 1, 14) = Chr(254)
'
'    checkPrecio (lista.Rows - 1)
'    checkDescuentoInd
'    addVentDet
'
        

        
        FrmFocus.addVentDet
        'FrmFocus.checkPrecio (FrmFocus.lista.Rows - 1)
        'FrmFocus.checkDescuentoInd
        FrmFocus.checkPrecioFinal

    FrmFocus.lista.TextMatrix(FrmFocus.lista.Rows - 1, 17) = vendetId

    Else
        MsgBox "No se puede agregar otro servicio por que no hay usuarios disónibles. Verifique.", vbInformation
    End If
End Sub

Private Sub mn_AgenCalen_Click()
    FRM_Agenda.Show
    FRM_Agenda.WindowState = 2
End Sub

Private Sub mn_Anticipos_Click()
    
    Call checarPermisos("FRM_APARTADOS", FRM_Menu.menuBarra2.Panels(8).Text)
    
    If permAcceso = "SI" Then
        tipoAprt = "APRT"
        FRM_Apartados.Show vbModal
    Else
        MsgBox "Opción no disponible.", vbInformation
    End If
    
    
End Sub

Private Sub mn_Asistencias_Click()
    FRM_Asistencias.Show
End Sub

Private Sub mn_BusqClientes_Click()
    If numFrmOper > 0 And UCase(FrmFocus.lInfo(2).Caption) = "ABIERTO" Then
        tipoBusqueda = "C"
        BUSQ_Usuarios.Caption = "Búsqueda de clientes."
        modBusqueda = "Operaciones"
        BUSQ_Usuarios.Show vbModal
    End If
End Sub

Private Sub mn_BusqProdu_Click()
    If numFrmOper > 0 And UCase(FrmFocus.lInfo(2).Caption) = "ABIERTO" Then
        modBusqueda = "Operaciones"
        BUSQ_ProdSer.Show vbModal
    End If

End Sub

Private Sub mn_BusqUsuario_Click()
    If numFrmOper > 0 And UCase(FrmFocus.lInfo(2).Caption) = "ABIERTO" Then
        tipoBusqueda = "U"
        modBusqueda = "Operaciones"
        BUSQ_Usuarios.Caption = "Búsqueda de usuarios."
        BUSQ_Usuarios.Show vbModal
    End If
    
End Sub

Private Sub mn_Cancel_Click()
    Dim question As String
    question = MsgBox("Cancelar: " & FrmFocus.lista.TextMatrix(FrmFocus.lista.Row, 0) & "  " & FrmFocus.lista.TextMatrix(FrmFocus.lista.Row, 2), vbYesNo + vbQuestion)
    If question = vbYes Then
        cancelarMotivo = "OPERACION"
        FRM_Cancelar.Show vbModal
        'cancelFila (FrmFocus.lista.Row)
    End If
End Sub
Public Sub cancelFila(fila As Long)
    ''''
    Dim monedero As Boolean
    
    monedero = False
    If FrmFocus.lista.TextMatrix(fila, 15) = "MND" Then
        monedero = True
    End If
        
    FrmFocus.deleteVentDet (fila)
    If FrmFocus.lista.Rows = 2 And fila = 1 Then
        FrmFocus.lista.Rows = 1
    Else
        If FrmFocus.lista.Rows > 2 Then
            FrmFocus.lista.RemoveItem (fila)
        End If
    End If
    
    
    If monedero = True Then
        FrmFocus.addMonedero
    End If
    
    FrmFocus.checkPrecioFinal
    
End Sub

Public Sub canelarOperacion(folioOper As Long)
    
    sql1 = "UPDATE VENTAS SET VENT_STATUS = 'C', " & _
    "VENT_FECHAHORA_COBRO = NOW() " & _
    "WHERE VENT_IDFOLIO = '" & folioOper & "' "
    con.Execute (sql1)
    
    sql1 = "UPDATE VENTA_DETALLE " & _
    "SET vendet_Status = 'C', vendet_CancelMotivo = '" & FRM_Cancelar.txtMotivo.Text & "', vendet_FechaHoraCancel = now(), " & _
    "vendet_AutorizaPerId = '" & FRM_Cancelar.lblAutoriza(0).Caption & "', vendet_AutorizaTipoId = '" & FRM_Cancelar.lblAutoriza(1).Caption & "', vendet_AutorizaTipo = '" & FRM_Cancelar.lblAutoriza(2).Caption & "'  " & _
    "WHERE VENDET_FOLIO = '" & folioOper & "' AND vendet_Status <> 'C' "
    con.Execute (sql1)
    
    
    
    sql1 = "SELECT VENDET_PRODUCTOID, VENDET_CANTIDAD FROM VENTA_DETALLE WHERE VENDET_FOLIO = '" & folioOper & "'"
    Set RES1 = con.Execute(sql1)

    Do While Not RES1.EOF
        sql1 = "UPDATE PRODUCTOS SET PROD_CANT = PROD_CANT + " & Val(RES1.Fields("vendet_cantidad")) & "  " & _
        " WHERE PROD_ID = '" & RES1.Fields("vendet_productoid") & "' "
        con.Execute (sql1)
        
        
        RES1.MoveNext
    Loop
    
    
    
    MsgBox "Operación cancelada", vbInformation

End Sub

Private Sub mn_CancelAll_Click()
    Dim ques As String
    ques = MsgBox("¿Cancelar todo en la lista?", vbYesNo + vbQuestion)
    If ques = vbYes Then
        
        FRM_Cancelar.Show vbModal
        
        'FrmFocus.lista.Rows = 1
        'FrmFocus.deleteVentDetAll
        'FrmFocus.checkPrecioFinal
    End If
    
End Sub

Private Sub mn_CancelOperTicket_Click()
    Dim ques As String
        
    Call checarPermisos("MDIC_OperTickets", FRM_Menu.menuBarra2.Panels(8).Text)

    MsgBox "aqui es"
    If permAcceso = "SI" Then
        ques = MsgBox("¿Cancelar operación con folio: " & FrmFocus.lista.TextMatrix(FrmFocus.lista.Row, 0) & vbCrLf & vbclrf & "Fecha y hora: " & FrmFocus.lista.TextMatrix(FrmFocus.lista.Row, 1) & _
        vbCrLf & vbCrLf & "Verifique muy bien la cancelación. " & vbCrLf & "Una vez realizada la acción no podrá deshacerse", vbYesNo + vbQuestion)
        If ques = vbYes Then
            cancelarMotivo = "OPERACION_ALL"
            FRM_Cancelar.Show vbModal
'            canelarOperacion (FrmFocus.lista.TextMatrix(FrmFocus.lista.Row, 0))
'            MDIC_OperTickets.cargaTickets
        End If
    Else
        MsgBox "Opción no disponible. Verifique", vbInformation
    End If
        
        
End Sub

Private Sub mn_Cobrar_Click()
 '   On Error Resume Next
    Dim usuarios(40) As String
    Dim ListUsuarios As String
    Dim encuentra As Boolean
        If numFrmOper > 0 Then
            If Val(Format(FrmFocus.txtSub.Text, "General number")) >= 0 And FrmFocus.lista.Rows > 1 And UCase(FrmFocus.lInfo(2).Caption) = "ABIERTO" Then
                ListUsuarios = ""
                encuentra = False
                                    
                For b1 = 1 To FrmFocus.lista.Rows - 1
                    For c1 = 0 To b1 - 1
                        If usuarios(c1) = FrmFocus.lista.TextMatrix(b1, 8) Then
                            encuentra = True
                            Exit For
                        End If
                    Next c1
                    If encuentra = False Then
                        usuarios(b1 - 1) = FrmFocus.lista.TextMatrix(b1, 8)
                        ListUsuarios = ListUsuarios & vbCrLf & FrmFocus.lista.TextMatrix(b1, 8)
                    End If
                    encuentra = False
                Next b1
                MsgBox "Usuarios asignados en la venta: " & vbCrLf & ListUsuarios, vbInformation, "Usuarios en venta"
                
                tipoCobro = "OPERACIONES"
                FRM_Cobro.txtTot.Text = FrmFocus.txtTotal.Text
                FRM_Cobro.Show vbModal
            Else
                MsgBox "Opcion no disponible", vbInformation
            End If
        Else
            MsgBox "Opción no disponible", vbInformation
        End If
End Sub

Private Sub mn_ConsumoInterno_Click()
    FRM_ConsumoInterno.Show vbModal
End Sub

Private Sub mn_Credito_Click()
    tipoAprt = "CRED"
    FRM_Apartados.Show vbModal
End Sub

Private Sub mn_Devoluciones_Click()
    FRM_Cambios.Show vbModal
End Sub

Private Sub mn_GastEgre_Click()
    FRM_Gastos.Show vbModal
End Sub

Private Sub mn_Impresora_Click()
    PRINT_Impresora.Show vbModal
End Sub

Private Sub mn_Membresia_Click()

MsgBox "Para poder asignar mebresias solo escriba el nombre de la membresia desde el cuadro de clave de producto." & vbCrLf & vbCrLf & _
"A continuación se presentará la pantalla de consultas de membresias", vbInformation

FRM_Membresias.Show

'On Error Resume Next
'        If numFrmOper > 0 Then
'            If FrmFocus.lblDatos(2).Caption <> "" Then
'                FRM_AsignMembresia.Show vbModal
'            Else
'                MsgBox "Debe de seleccionar un cliente para poder asignar membresía.", vbInformation
'            End If
'        Else
'            MsgBox "Opción no disponible", vbInformation
'        End If

End Sub

Private Sub mn_NewOper_Click()
    
'    MDIC_OperTickets.cargaTickets
    
    Set FrmOper = New MDIC_Operaciones
    numFrmOper = numFrmOper + 1
    'FrmOper.Caption = sCaption & numFrmOper
    FrmOper.Show
    
    

End Sub

Private Sub nm_Busqueda_Click()

End Sub

Private Sub mn_NewOper2_Click()
    
    
    Set FrmOper2 = New MDIC_Operaciones2
    
    numFrmOper2 = numFrmOper2 + 1
    'FrmOper.Caption = sCaption & numFrmOper
    FrmOper2.Show

End Sub

Private Sub mn_OpenCajon_Click()
 Call abrirCajon
End Sub

Private Sub mn_PrintPreTicket_Click()
'Dim ques As String
    
'ques = MsgBox("Imprimir ticket folio: " & FrmFocus.Lista.TextMatrix(FrmFocus.Lista.Row, 0) & "?", vbYesNo + vbQuestion)
'If ques = vbYes Then
    If mesas = True Then
        notaPreTicket (FrmFocus.lista.TextMatrix(FrmFocus.lista.Row, 0))
    Else
        nota (FrmFocus.lista.TextMatrix(FrmFocus.lista.Row, 0))
    End If
'End If
End Sub

Private Sub mn_PrintPreTicket2_Click()
    notaPreTicket (FrmFocus.lInfo(1).Caption)
End Sub

Private Sub mn_PrintTicket_Click()
Dim ques As String
    
ques = MsgBox("Imprimir ticket folio: " & FrmFocus.lista.TextMatrix(FrmFocus.lista.Row, 0) & "?", vbYesNo + vbQuestion)
If ques = vbYes Then
    nota (FrmFocus.lista.TextMatrix(FrmFocus.lista.Row, 0))
End If

End Sub

Private Sub mn_Productos_Click()
    FRM_Productos.Show vbModal

End Sub

Private Sub mn_Servicios_Click()
    FRM_Servicios.Show vbModal
End Sub

Private Sub mn_tickets_Click()
    If Val(MDI_Operaciones.StatusBar1.Panels(4).Text) = 0 Then
        Set FrmTickets = New MDIC_OperTickets
        FrmTickets.Show
    Else
        If Val(MDI_Operaciones.StatusBar1.Panels(4).Text) >= 1 Then
            Dim formularios As Form
            For Each formularios In Forms
                If formularios.Name = "MDIC_OperTickets" Then
                    formularios.Show
                    formularios.cargaTickets
                End If
                Exit For
            Next
        End If
    End If
    
'    If numFrmTicket = 0 Then
'        Set FrmTickets = New MDIC_OperTickets
'        FrmTickets.Show
'    Else
'        FrmTickets.Show
'    End If

End Sub

