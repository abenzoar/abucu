VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form FRM_Menu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu principal          AuxsSis          ABUCU"
   ClientHeight    =   10155
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   19590
   Icon            =   "FRM_Menu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "FRM_Menu.frx":08CA
   ScaleHeight     =   10155
   ScaleWidth      =   19590
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar menuBarra2 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   9780
      Width           =   19590
      _ExtentX        =   34555
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   15
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   2646
            Picture         =   "FRM_Menu.frx":411C6
            TextSave        =   "08/03/2016"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            AutoSize        =   2
            Bevel           =   0
            TextSave        =   "12:52 p. m."
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Bevel           =   0
            Picture         =   "FRM_Menu.frx":414E0
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Bevel           =   0
            Picture         =   "FRM_Menu.frx":417FA
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Bevel           =   0
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel8 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel9 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel10 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Bevel           =   0
            Picture         =   "FRM_Menu.frx":41B14
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel11 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel12 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel13 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel14 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel15 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Object.Tag             =   ""
         EndProperty
      EndProperty
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
   Begin VB.Timer tmr_Asistencias 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   960
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3500
      Left            =   240
      Top             =   720
   End
   Begin VB.Image imgMenu 
      Height          =   1695
      Index           =   33
      Left            =   10920
      Top             =   5760
      Width           =   2895
   End
   Begin VB.Image imgMenu 
      Height          =   1695
      Index           =   23
      Left            =   10920
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Image imgMenu 
      Height          =   1695
      Index           =   20
      Left            =   7680
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Image imgMenu 
      Height          =   1695
      Index           =   30
      Left            =   7680
      Top             =   5760
      Width           =   2895
   End
   Begin VB.Image imgMenu 
      Height          =   1695
      Index           =   29
      Left            =   4440
      Top             =   5760
      Width           =   2895
   End
   Begin VB.Image imgMenu 
      Height          =   1695
      Index           =   26
      Left            =   1200
      Top             =   5760
      Width           =   2895
   End
   Begin VB.Image imgMenu 
      Height          =   1695
      Index           =   14
      Left            =   1200
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Image imgMenu 
      Height          =   1695
      Index           =   17
      Left            =   4440
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Image imgMenu 
      Height          =   1695
      Index           =   2
      Left            =   1200
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Image imgMenu 
      Height          =   1695
      Index           =   5
      Left            =   4440
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Image imgMenu 
      Height          =   1695
      Index           =   8
      Left            =   7680
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Image imgMenu 
      Height          =   1695
      Index           =   11
      Left            =   10920
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Image imgMenu 
      Height          =   1455
      Index           =   34
      Left            =   10920
      Picture         =   "FRM_Menu.frx":41E2E
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   2895
   End
   Begin VB.Image imgMenu 
      Height          =   1425
      Index           =   35
      Left            =   10920
      Picture         =   "FRM_Menu.frx":42A30
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   2835
   End
   Begin VB.Image imgMenu 
      Height          =   1455
      Index           =   31
      Left            =   7680
      Picture         =   "FRM_Menu.frx":4387D
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   2895
   End
   Begin VB.Image imgMenu 
      Height          =   1455
      Index           =   32
      Left            =   7680
      Picture         =   "FRM_Menu.frx":445C3
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   2895
   End
   Begin VB.Image imgMenu 
      Height          =   1455
      Index           =   28
      Left            =   4440
      Picture         =   "FRM_Menu.frx":45566
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   2895
   End
   Begin VB.Image imgMenu 
      Height          =   1455
      Index           =   27
      Left            =   4440
      Picture         =   "FRM_Menu.frx":4683F
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   2895
   End
   Begin VB.Image imgMenu 
      Height          =   1455
      Index           =   25
      Left            =   1200
      Picture         =   "FRM_Menu.frx":47C80
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   2895
   End
   Begin VB.Image imgMenu 
      Height          =   1455
      Index           =   24
      Left            =   1200
      Picture         =   "FRM_Menu.frx":48BD6
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   2895
   End
   Begin VB.Image imgMenu 
      Height          =   1455
      Index           =   22
      Left            =   10920
      Picture         =   "FRM_Menu.frx":49D3A
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   2895
   End
   Begin VB.Image imgMenu 
      Height          =   1455
      Index           =   21
      Left            =   10920
      Picture         =   "FRM_Menu.frx":4A8DA
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   2895
   End
   Begin VB.Image imgMenu 
      Height          =   1455
      Index           =   19
      Left            =   7680
      Picture         =   "FRM_Menu.frx":4B727
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   2895
   End
   Begin VB.Image imgMenu 
      Height          =   1455
      Index           =   18
      Left            =   7680
      Picture         =   "FRM_Menu.frx":4C4EA
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   2895
   End
   Begin VB.Image imgMenu 
      Height          =   1455
      Index           =   16
      Left            =   4440
      Picture         =   "FRM_Menu.frx":4D50A
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   2895
   End
   Begin VB.Image imgMenu 
      Height          =   1455
      Index           =   15
      Left            =   4440
      Picture         =   "FRM_Menu.frx":4E709
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   2895
   End
   Begin VB.Image imgMenu 
      Height          =   1455
      Index           =   13
      Left            =   1200
      Picture         =   "FRM_Menu.frx":4FAB8
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   2895
   End
   Begin VB.Image imgMenu 
      Height          =   1455
      Index           =   12
      Left            =   1200
      Picture         =   "FRM_Menu.frx":50943
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   2895
   End
   Begin VB.Image imgMenu 
      Height          =   1455
      Index           =   10
      Left            =   10920
      Picture         =   "FRM_Menu.frx":519F8
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Image imgMenu 
      Height          =   1455
      Index           =   7
      Left            =   7680
      Picture         =   "FRM_Menu.frx":527B2
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Image imgMenu 
      Height          =   1455
      Index           =   4
      Left            =   4440
      Picture         =   "FRM_Menu.frx":53509
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Image imgMenu 
      Height          =   1455
      Index           =   3
      Left            =   4440
      Picture         =   "FRM_Menu.frx":54242
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Licencia"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   14880
      TabIndex        =   5
      Top             =   8040
      Width           =   1335
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "2011-2012"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   4
      Left            =   14880
      TabIndex        =   4
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   2
      X1              =   14760
      X2              =   18720
      Y1              =   7800
      Y2              =   7800
   End
   Begin VB.Image imgInfo 
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Index           =   1
      Left            =   15600
      Picture         =   "FRM_Menu.frx":551F2
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   1
      X1              =   14760
      X2              =   18720
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Matriz"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   3
      Left            =   15000
      TabIndex        =   3
      Top             =   5400
      Width           =   3495
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Sucursal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   15000
      TabIndex        =   2
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   0
      X1              =   14760
      X2              =   18720
      Y1              =   4860
      Y2              =   4860
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de usuario"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   1
      Left            =   15000
      TabIndex        =   1
      Top             =   4440
      Width           =   3495
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de usuario"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   15000
      TabIndex        =   0
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Image imgInfo 
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Index           =   0
      Left            =   15600
      Picture         =   "FRM_Menu.frx":57886
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Image imgMenu 
      Height          =   1455
      Index           =   1
      Left            =   1200
      Picture         =   "FRM_Menu.frx":59F1A
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Image imgMenu 
      Height          =   1455
      Index           =   0
      Left            =   1200
      Picture         =   "FRM_Menu.frx":5AD09
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Image imgMenu 
      Height          =   1455
      Index           =   6
      Left            =   7680
      Picture         =   "FRM_Menu.frx":5BD40
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Image imgMenu 
      Height          =   1455
      Index           =   9
      Left            =   10920
      Picture         =   "FRM_Menu.frx":5CD17
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Image imgFondo 
      Height          =   15000
      Left            =   -240
      Picture         =   "FRM_Menu.frx":5DD35
      Stretch         =   -1  'True
      Top             =   0
      Width           =   21000
   End
   Begin VB.Image imgFondo2 
      Height          =   17280
      Left            =   -1320
      Picture         =   "FRM_Menu.frx":2FD531
      Top             =   0
      Width           =   29760
   End
   Begin VB.Menu mn_Oper 
      Caption         =   "Operaciones"
      Begin VB.Menu mn_Operaciones 
         Caption         =   "Operaciones"
      End
      Begin VB.Menu mn_Apartados 
         Caption         =   "Apartados"
      End
      Begin VB.Menu mn_VentTouch 
         Caption         =   "Venta touch"
      End
      Begin VB.Menu mn_Line2 
         Caption         =   "-"
      End
      Begin VB.Menu mn_Caja 
         Caption         =   "Caja"
      End
      Begin VB.Menu mn_Line3 
         Caption         =   "-"
      End
      Begin VB.Menu mn_PagUser 
         Caption         =   "Pagos a usuarios"
      End
      Begin VB.Menu mn_CSI 
         Caption         =   "Consumo interno"
      End
      Begin VB.Menu mn_Gastos 
         Caption         =   "Gastos"
      End
   End
   Begin VB.Menu mn_Negocio 
      Caption         =   "Negocio"
      Begin VB.Menu mn_Usuario 
         Caption         =   "Usuarios"
      End
      Begin VB.Menu mn_ProdPadre 
         Caption         =   "Productos"
         Begin VB.Menu mn_Productos 
            Caption         =   "Productos"
         End
         Begin VB.Menu mn_linProd1 
            Caption         =   "-"
         End
         Begin VB.Menu mn_Invent 
            Caption         =   "Inventario"
         End
         Begin VB.Menu mn_Traslado 
            Caption         =   "Traslado"
         End
         Begin VB.Menu mn_Pedidos 
            Caption         =   "Pedidos/Entrada a almacen"
         End
         Begin VB.Menu mn_linProd2 
            Caption         =   "-"
         End
         Begin VB.Menu mn_SegServ 
            Caption         =   "Seguimiento para servicio"
         End
      End
      Begin VB.Menu mn_Servicios 
         Caption         =   "Servicios"
      End
      Begin VB.Menu mn_Clientes 
         Caption         =   "Clientes"
      End
      Begin VB.Menu mn_HistorialVent 
         Caption         =   "Historial"
         Begin VB.Menu mn_HistoriVent 
            Caption         =   "Ventas general"
         End
         Begin VB.Menu mn_HistoClieResumen 
            Caption         =   "Clientes/Compras"
         End
         Begin VB.Menu mn_ventResumen 
            Caption         =   "Ventas mensual"
         End
      End
      Begin VB.Menu mn_LineNeg1 
         Caption         =   "-"
      End
      Begin VB.Menu mn_Indicadores 
         Caption         =   "Indicadores"
      End
      Begin VB.Menu mn_membresiasClts 
         Caption         =   "Membresias"
      End
      Begin VB.Menu mn_Monedero 
         Caption         =   "Monedero"
      End
   End
   Begin VB.Menu mn_Agen 
      Caption         =   "Agenda"
      Begin VB.Menu mn_Agenda 
         Caption         =   "Agenda"
      End
   End
   Begin VB.Menu mn_Cat 
      Caption         =   "Catálogos"
      Begin VB.Menu mn_Periodos 
         Caption         =   "Periodos de tiempo"
      End
      Begin VB.Menu mn_Etiquetas 
         Caption         =   "Etiquetas"
      End
      Begin VB.Menu mn_PagosCom 
         Caption         =   "Pagos - Comisiones"
      End
      Begin VB.Menu mn_CatMembresias 
         Caption         =   "Membresias"
      End
      Begin VB.Menu mn_PuntosMone 
         Caption         =   "Puntos - Monedero"
      End
   End
   Begin VB.Menu mn_AuxSis 
      Caption         =   "AuxsSis"
      Begin VB.Menu mn_Asistencias 
         Caption         =   "Asistencias"
      End
      Begin VB.Menu mn_DatosSuc 
         Caption         =   "Datos de la sucursal"
      End
      Begin VB.Menu mn_Perm 
         Caption         =   "Permisos y accesos"
      End
      Begin VB.Menu mn_UserPagos 
         Caption         =   "Asiganción pagos de usuarios"
      End
      Begin VB.Menu mn_Msjs 
         Caption         =   "Mensajes"
      End
   End
End
Attribute VB_Name = "FRM_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim producto As Boolean
Dim sql1 As String
Dim RES1 As Recordset
Dim ResAst As Recordset
Dim ResAst2 As Recordset
Dim SQL2 As String
Dim RES2 As Recordset
Dim numAsistencias As Long
Dim tipoAcceso As Integer


Private Sub Form_Load()
    numAsistencias = 0
    check_InfoAistencia
    tipoAcceso = 0
End Sub
Private Sub check_InfoAistencia()
    sql1 = "SELECT SUC_NOTIFICAR_ASISTENCIA FROM SUCURSAL"
    Set ResAst = con.Execute(sql1)
    If Not ResAst.EOF Then
        If ResAst.Fields("SUC_NOTIFICAR_ASISTENCIA") = "1" Then
            tmr_Asistencias.Enabled = True
        Else
            tmr_Asistencias.Enabled = False
        End If
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    If numFrmOper > 0 Then
        MsgBox "La ventana de operaciones está ejecutándose.", vbInformation
        MDI_Operaciones.Show
        Cancel = 1
    Else
        Cancel = 0
        con.Close
        End
    End If
End Sub

Private Sub Hs1_Change()
    Call Aplicar_Transparencia(Me.hWnd, CByte(Hs1.value))

End Sub

Private Sub imgFondo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If tipoAcceso = 0 Then
    cambiaAccesos (0)
End If


End Sub



Private Sub imgFondo2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If tipoAcceso = 0 Then
    cambiaAccesos (0)
End If

End Sub

Private Sub imgMenu_Click(Index As Integer)
    Select Case Index
        Case 2: mn_Productos_Click
        Case 5: mn_Clientes_Click
        Case 8: mn_Usuario_Click
        Case 11: mn_Servicios_Click
        Case 14: mn_Asistencias_Click
        Case 17: mn_DatosSuc_Click
        ''Case 20: paquete
        Case 23: mn_Caja_Click
        Case 26: mn_Operaciones_Click
        Case 29: mn_UserPagos_Click
        Case 30: mn_Agenda_Click
        Case 33: tipoAcceso = 1
        'cambiaAccesos
    
    End Select

End Sub
Private Sub cambiaAccesos(tipoAcceso As Integer)
    
    If tipoAcceso = 0 Then
    ''Productos
        If imgMenu(1).Visible = True Then
            imgMenu(0).Visible = True
            imgMenu(1).Visible = False
        End If
    ''Clientes
        If imgMenu(4).Visible = True Then
            imgMenu(3).Visible = True
            imgMenu(4).Visible = False
        End If
        
        ''Usuario
        If imgMenu(8).Visible = True Then
            imgMenu(6).Visible = True
            imgMenu(7).Visible = False
        End If
        
        ''Servicios
        If imgMenu(11).Visible = True Then
            imgMenu(9).Visible = True
            imgMenu(10).Visible = False
        End If
        
        ''Asistencia
        If imgMenu(14).Visible = True Then
            imgMenu(12).Visible = True
            imgMenu(13).Visible = False
         End If
            ''Datos
        If imgMenu(17).Visible = True Then
            imgMenu(15).Visible = True
            imgMenu(16).Visible = False
        End If
        
         ''paquetes
        If imgMenu(20).Visible = True Then
            imgMenu(18).Visible = True
            imgMenu(19).Visible = False
        End If
        
         ''caja
        If imgMenu(23).Visible = True Then
            imgMenu(21).Visible = True
            imgMenu(22).Visible = False
        End If
        ''operacione
        If imgMenu(26).Visible = True Then
            imgMenu(24).Visible = True
            imgMenu(25).Visible = False
        End If
        ''PagosUserAsig
        If imgMenu(29).Visible = True Then
            imgMenu(27).Visible = True
            imgMenu(28).Visible = False
        End If
        ''Agenda
        If imgMenu(30).Visible = True Then
            imgMenu(32).Visible = True
            imgMenu(31).Visible = False
        End If
        If imgMenu(33).Visible = True Then
            imgMenu(35).Visible = True
            imgMenu(34).Visible = False
        End If
    Else
         If tipoAcceso = 1 Then
                imgMenu(0).Visible = False
                imgMenu(1).Visible = False
                imgMenu(3).Visible = False
                imgMenu(4).Visible = False
                imgMenu(6).Visible = False
                imgMenu(7).Visible = False
                imgMenu(9).Visible = False
                imgMenu(10).Visible = False
                imgMenu(12).Visible = False
                imgMenu(13).Visible = False
                imgMenu(15).Visible = False
                imgMenu(16).Visible = False
                imgMenu(18).Visible = False
                imgMenu(19).Visible = False
                imgMenu(21).Visible = False
                imgMenu(22).Visible = False
                imgMenu(24).Visible = False
                imgMenu(25).Visible = False
                imgMenu(27).Visible = False
                imgMenu(28).Visible = False
                imgMenu(32).Visible = False
                imgMenu(31).Visible = False
'                imgMenu(35).Visible = False
'                imgMenu(34).Visible = False
'    Esto va
'CREDITOS
'ASISTENCIAS
'AGENDA
'gastos
'PEDIDOS
'INDICADORES
'HSTORIAL VENTAS
'HISTPRIAL CLIENTES
'CREDITOS
         End If
    End If


End Sub
Private Sub imgMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tipoAcceso = 0 Then
        Select Case Index
            Case 2: imgMenu(1).Visible = True
                    imgMenu(0).Visible = False
                    
            Case 5: imgMenu(4).Visible = True
                    imgMenu(3).Visible = False
                    
            Case 8: imgMenu(7).Visible = True
                    imgMenu(6).Visible = False
                    
            Case 11: imgMenu(10).Visible = True
                     imgMenu(9).Visible = False
                     
            Case 14: imgMenu(13).Visible = True
                     imgMenu(12).Visible = False
                     
            Case 17: imgMenu(16).Visible = True
                     imgMenu(15).Visible = False
                     
            Case 20: imgMenu(19).Visible = True
                     imgMenu(18).Visible = False
            
            Case 23: imgMenu(22).Visible = True
                     imgMenu(21).Visible = False
                     
            Case 26: imgMenu(25).Visible = True
                     imgMenu(24).Visible = False
            
            Case 29: imgMenu(27).Visible = False
                     imgMenu(28).Visible = True
            
            Case 30: imgMenu(31).Visible = True
                     imgMenu(32).Visible = False
            
            Case 33: imgMenu(34).Visible = True
                     imgMenu(35).Visible = False
        End Select
    End If
    
End Sub

Private Sub mn_Agenda_Click()
     FRM_Agenda.Show
End Sub

Private Sub mn_Apartados_Click()
    
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


Private Sub mn_Caja_Click()
    
    Call checarPermisos("FRM_Caja", FRM_Menu.menuBarra2.Panels(8).Text)
    
    If permAcceso = "SI" Then
        
        FRM_Caja.Show vbModal
    Else
        MsgBox "Opción no disponible.", vbInformation
    End If
    
    
End Sub

Private Sub mn_CatMembresias_Click()
    CAT_Membresias.Show vbModal
End Sub

Private Sub mn_Clientes_Click()
    FRM_Clientes.Show 'vbModal
End Sub

Private Sub mn_CSI_Click()
    FRM_ConsumoInterno.Show vbModal
    
End Sub

Private Sub mn_DatosSuc_Click()
    FRM_DatosSuc.Show vbModal
End Sub

Private Sub mn_Etiquetas_Click()
    CAT_Etiquetas.Show vbModal
End Sub

Private Sub mn_Gastos_Click()
    FRM_Gastos.Show vbModal
End Sub

Private Sub mn_HistoClieActi_Click()

End Sub

Private Sub mn_HistoClieResumen_Click()
    FRM_HistoClie.Show vbModal
End Sub

Private Sub mn_HistoriVent_Click()
    FRM_HistoProd.Show vbModal
End Sub

Private Sub mn_Indicadores_Click()
    RPT_RrtMain.Show
End Sub

Private Sub mn_Invent_Click()
    FRM_Inventario.Show
End Sub

Private Sub mn_membresiasClts_Click()
    FRM_Membresias.Show vbModal
End Sub

Private Sub mn_Monedero_Click()
    FRM_Monederos.Show vbModal
End Sub

Private Sub mn_Msjs_Click()
    FRM_Mensajes.Show vbModal
End Sub

Private Sub mn_Operaciones_Click()
     
    Call checarPermisos("MDI_OPERACIONES", FRM_Menu.menuBarra2.Panels(8).Text)
    
    If permAcceso = "SI" Then
        MDI_Operaciones.WindowState = vbMaximized
        MDI_Operaciones.Show
    Else
        MsgBox "Opción no disponible.", vbInformation
    End If
     
     
End Sub

Private Sub mn_PagosCom_Click()
    CAT_Pagos.Show vbModal
End Sub

Private Sub mn_PagUser_Click()
    FRM_PagosUsuarios.Show vbModal
End Sub

Private Sub mn_Pedidos_Click()
    CAT_Pedidos.Show
    
End Sub

Private Sub mn_Periodos_Click()
    CAT_Periodos.Show vbModal
End Sub

Private Sub mn_Perm_Click()
        
    Call checarPermisos("FRM_Permisos", FRM_Menu.menuBarra2.Panels(8).Text)
    
    If permAcceso = "SI" Then
        FRM_Permisos.Show vbModal
    Else
        MsgBox "Opción no disponible.", vbInformation
    End If
    
    
End Sub

Private Sub mn_Productos_Click()
        
        
    Call checarPermisos("FRM_Productos", FRM_Menu.menuBarra2.Panels(8).Text)
    
    If permAcceso = "SI" Then
        FRM_Productos.Show 'vbModal
    Else
        MsgBox "Opción no disponible.", vbInformation
    End If
    

End Sub

Private Sub mn_PuntosMone_Click()
    FRM_PuntosMone.Show vbModal
End Sub

Private Sub mn_SegServ_Click()
    FRM_Seguimiento.Show vbModal
    
End Sub

Private Sub mn_Servicios_Click()
    
    Call checarPermisos("FRM_Servicios", FRM_Menu.menuBarra2.Panels(8).Text)
    
    If permAcceso = "SI" Then
        FRM_Servicios.Show vbModal
    Else
        MsgBox "Opción no disponible.", vbInformation
    End If
    
    
    
End Sub

Private Sub mn_Traslado_Click()
    FRM_TrasLados.Show vbModal
End Sub

Private Sub mn_UserPagos_Click()
    FRM_Usuarios_Pagos.Show vbModal
End Sub

Private Sub mn_Usuario_Click()
    
    Call checarPermisos("FRM_Usuarios", FRM_Menu.menuBarra2.Panels(8).Text)
    
    If permAcceso = "SI" Then
        FRM_Usuarios.Show vbModal
    Else
        MsgBox "Opción no disponible.", vbInformation
    End If
    
    
End Sub

Private Sub mn_ventResumen_Click()
'    Call checarPermisos("FRM_HistoVentas", FRM_Menu.menuBarra2.Panels(8).Text)
    
'    If permAcceso = "SI" Then
        FRM_HistoVentas.Show vbModal
'    Else
'        MsgBox "Opción no disponible. Verifique", vbInformation
'    End If
    
End Sub

Private Sub mn_VentTouch_Click()
'    FRM_OperTouch.Show vbModal
    tipo_AccesoTouch = "Indentificador de usuario - Menu"
    FRM_Identificador.Caption = "Indentificador de usuario - Menu"
    FRM_Identificador.Show vbModal
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    checarCarpetaTemp
    'MsgBox direccionSistema
    imgFondo.Picture = LoadPicture(direccionSistema & "\Com\Menu2.jpg")
    'imgMenu.Picture = LoadPicture(direccionSistema & "\Com\Productos_A.jpg")
    'imgMenu.Visible = True
End Sub
Private Sub muestraInfo_Asistencia(perId As Long)

    sql1 = "SELECT T2.PER_ID, T1.PERTP_CODIGO_MEMBRESIA, T2.PER_NOMBRE, T2.PER_PATERNO, T2.PER_MATERNO, T2.PER_FOTO, " & _
    "T1.PERTP_TIPO_ID, T1.PERTP_PER_ID, T1.PERTP_PER_TIPO, T4.mbr_ctmbId, T4.mbr_VentaFolio, T4.MBR_FIN, DATEDIFF(MBR_fIN, CURDATE()) DIAS  " & _
    "FROM PER_TIPO T1, PERSONA T2, CAT_TIPO T3, MEMBRESIAS T4 " & _
    "WHERE T1.PERTP_PER_ID = T2.PER_ID AND T1.PERTP_TIPO_ID = T3.CTPT_ID AND T1.PERTP_PER_TIPO = T3.CTPT_SUBTIPO " & _
    "AND MBR_PERTP_PER_ID = T1.PERTP_PER_ID AND MBR_PERTP_TIPO_ID = T1.PERTP_TIPO_ID " & _
    "AND CURDATE() BETWEEN T4.MBR_INICIO AND T4.MBR_FIN AND MBR_STATUS = 'A'  " & _
    "AND T1.PERTP_PER_ID = '" & perId & "'"
    Set ResAst = con.Execute(sql1)
    If Not ResAst.EOF Then

        Unload TopForm_Asistencia
        If IsNull(ResAst.Fields("PER_fOTO")) = False Then
            checarCarpetaTemp
            Imagen1.Open
            Imagen1.Write ResAst.Fields("PER_FOTO")
            Imagen1.SaveToFile direccionSistema & "\Temp\TempUser.dat", adSaveCreateOverWrite
            Imagen1.Close
            TopForm_Asistencia.iFoto.Picture = LoadPicture(direccionSistema & "\Temp\TempUser.dat")
        Else
            TopForm_Asistencia.iFoto.Picture = LoadPicture("")
        End If

        TopForm_Asistencia.lDatos.Caption = ResAst.Fields("PER_NOMBRE") & " " & ResAst.Fields("PER_PATERNO") & " " & ResAst.Fields("PER_MATERNO") & _
        vbCrLf & ResAst.Fields("PERTP_CODIGO_MEMBRESIA") & vbCrLf & "Tu membresia vence: " & ResAst.Fields("MBR_FIN") & vbCrLf & _
        "Días disponibles: " & ResAst.Fields("DIAS")
        TopForm_Asistencia.Show
'        txtUsuario(0).Text = ""
'        txtUsuario(0).SetFocus
'        Image1(0).Visible = True
'        Image1(1).Visible = False
'        lInfo.Caption = "Bienvenido " & vbCrLf & ResAst.Fields("PER_NOMBRE") & " " & ResAst.Fields("PER_PATERNO") & " " & ResAst.Fields("PER_MATERNO") & _
'        vbCrLf & "Tu membresia vence: " & ResAst.Fields("MBR_FIN") & vbCrLf & _
'        "Días disponibles: " & ResAst.Fields("DIAS")
    End If

End Sub
Private Sub tmr_Asistencias_Timer()
    If Me.Visible = True Then
    
        sql1 = "SELECT COUNT(CLAVE_CLTE) NUM FROM VIEW_ASISTENCIAS WHERE DATE_FORMAT(NOW(), '%d/%m/%y') = DATE_FORMAT(FECHA_HORA, '%d/%m/%y') "
        Set ResAst = con.Execute(sql1)
        
        If Not ResAst.EOF Then
            If numAsistencias = 0 Then
                numAsistencias = ResAst.Fields("NUM")
            End If
            
            If numAsistencias > 0 Then
                If ResAst.Fields("NUM") > numAsistencias Then
                    SQL2 = "SELECT CLAVE_CLTE FROM VIEW_ASISTENCIAS ORDER BY CLAVE_ASTS DESC LIMIT 1"
                    Set ResAst2 = con.Execute(SQL2)
                    If Not ResAst2.EOF Then
                        numAsistencias = ResAst.Fields("NUM")
                        muestraInfo_Asistencia (ResAst2.Fields("CLAVE_CLTE"))
                    End If
                End If
            End If
    
    '        numAsistencias = ResAst.Fields("CLAVE_CLTE")
        End If
    Else
        tmr_Asistencias.Enabled = False
    End If
    
End Sub
