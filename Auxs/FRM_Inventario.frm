VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRM_Inventario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventario"
   ClientHeight    =   9480
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   16455
   Icon            =   "FRM_Inventario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9480
   ScaleWidth      =   16455
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   9495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   16748
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   873
      TabCaption(0)   =   "  Inventarios realizados"
      TabPicture(0)   =   "FRM_Inventario.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MSFlexGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ListaInvent"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "  Nuevo inventario"
      TabPicture(1)   =   "FRM_Inventario.frx":11A4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lBus(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Borde(15)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lProd(16)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Shape1(6)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Shape1(0)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lProd(0)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "listaInventAdd"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "listaResultados"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "time1"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "textBus(0)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cmdAdd"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "cmdFin"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "cmdCancelar"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).ControlCount=   13
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   735
         Left            =   -63360
         Picture         =   "FRM_Inventario.frx":1A7E
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton cmdFin 
         Caption         =   "Finalizar"
         Height          =   735
         Left            =   -65880
         Picture         =   "FRM_Inventario.frx":2348
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton cmdAdd 
         Height          =   495
         Left            =   -71160
         Picture         =   "FRM_Inventario.frx":2C12
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox textBus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   -74880
         TabIndex        =   3
         Top             =   840
         Width           =   3615
      End
      Begin VB.Timer time1 
         Interval        =   500
         Left            =   -66360
         Top             =   360
      End
      Begin MSFlexGridLib.MSFlexGrid ListaInvent 
         Height          =   5895
         Left            =   120
         TabIndex        =   1
         Top             =   3120
         Width           =   17175
         _ExtentX        =   30295
         _ExtentY        =   10398
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         AllowUserResizing=   1
         FormatString    =   $"FRM_Inventario.frx":34DC
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
      Begin MSFlexGridLib.MSFlexGrid listaResultados 
         Height          =   7695
         Left            =   -69480
         TabIndex        =   2
         Top             =   1800
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   13573
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         AllowUserResizing=   1
         FormatString    =   $"FRM_Inventario.frx":3602
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
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2175
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   17175
         _ExtentX        =   30295
         _ExtentY        =   3836
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColor       =   16777215
         AllowUserResizing=   1
         FormatString    =   "Fecha / Hora                                     | Usuario                                    | Clave    "
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
      Begin MSFlexGridLib.MSFlexGrid listaInventAdd 
         Height          =   7695
         Left            =   -74880
         TabIndex        =   8
         Top             =   1800
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   13573
         _Version        =   393216
         FixedCols       =   0
         AllowUserResizing=   1
         FormatString    =   "Codigo escaneado               | Fecha/Hora escaneo                    "
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
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Resultado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   -69480
         TabIndex        =   10
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   60
         Index           =   0
         Left            =   -69480
         Top             =   1680
         Width           =   10815
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   60
         Index           =   6
         Left            =   -74880
         Top             =   1680
         Width           =   5295
      End
      Begin VB.Label lProd 
         BackStyle       =   0  'Transparent
         Caption         =   "Codigos escaneados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   16
         Left            =   -74880
         TabIndex        =   9
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   15
         Left            =   -74880
         Top             =   840
         Width           =   3645
      End
      Begin VB.Label lBus 
         BackStyle       =   0  'Transparent
         Caption         =   "Código producto"
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
         Left            =   -74880
         TabIndex        =   4
         Top             =   600
         Width           =   2655
      End
   End
   Begin VB.Menu mn_menu 
      Caption         =   "Menu"
      Begin VB.Menu mn_nuevo 
         Caption         =   "Nuevo inventario"
      End
      Begin VB.Menu mn_Salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mn_Options 
      Caption         =   "Opciones"
      Begin VB.Menu mn_Export 
         Caption         =   "Exportar lista a excel"
      End
   End
End
Attribute VB_Name = "FRM_Inventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql1 As String
Dim RES1 As Recordset
Dim invtId As Integer


Private Sub addInventario()



        sql1 = "INSERT INTO INVENTARIO_DETALLE (INVDT_INVID, INVDT_CODIGO, INVDT_FECHAHORA) VALUES ('" & invtId & "', '" & textBus(0).Text & "', now()) "
        con.Execute (sql1)

        listaInventAdd.AddItem ""
        listaInventAdd.TextMatrix(listaInventAdd.Rows - 1, 0) = textBus(0).Text
        listaInventAdd.TextMatrix(listaInventAdd.Rows - 1, 1) = Now

        textBus(0).SetFocus
        textBus(0).SelStart = 0
        textBus(0).SelLength = Len(textBus(0).Text)
End Sub

Private Sub cargaIncial()
    ListaInvent.Rows = 1
    listaInventAdd.Rows = 1
    SSTab1.Tab = 0
    SSTab1.TabEnabled(1) = False
End Sub

Private Sub cmdAdd_Click()
    If textBus(0) <> "" Then
        addInventario
    End If
    
End Sub

Private Sub cmdFin_Click()
Dim ques As String

ques = MsgBox("Va a finalizar el inventario. Al terminar no podrá anexar mas productos a la lista. " & vbCrLf & vbCrLf & "¿Continuar? ", vbYesNo + vbQuestion)
If ques = vbYes Then
    textBus(0).Enabled = False
    cmdAdd.Enabled = False
    cmdFin.Enabled = False
    
    listaResultados.Rows = 1
    
    sql1 = "SELECT * FROM VIEW_INVENTARIO_DETALLE WHERE CLAVE_INVENTARIO = '" & invtId & "'"
    Set RES1 = con.Execute(sql1)
    
    listaResultados.Redraw = False
    Do While Not RES1.EOF
        listaResultados.AddItem ""
        listaResultados.TextMatrix(listaResultados.Rows - 1, 0) = RES1.Fields("CODIGO")
        listaResultados.TextMatrix(listaResultados.Rows - 1, 2) = RES1.Fields("CANTIDAD_REGISTRADA")
        If IsNull(RES1.Fields("PRODUCTO")) Then
            listaResultados.TextMatrix(listaResultados.Rows - 1, 1) = "NO ENCONTRADO"
            listaResultados.TextMatrix(listaResultados.Rows - 1, 4) = ""
            listaResultados.TextMatrix(listaResultados.Rows - 1, 5) = ""
            listaResultados.TextMatrix(listaResultados.Rows - 1, 3) = "0"
        Else
            listaResultados.TextMatrix(listaResultados.Rows - 1, 1) = RES1.Fields("PRODUCTO")
            listaResultados.TextMatrix(listaResultados.Rows - 1, 4) = (RES1.Fields("CANTIDAD_SISTEMA") - RES1.Fields("CANTIDAD_REGISTRADA"))
            listaResultados.TextMatrix(listaResultados.Rows - 1, 5) = RES1.Fields("ESTATUS")
            listaResultados.TextMatrix(listaResultados.Rows - 1, 3) = RES1.Fields("CANTIDAD_SISTEMA")
        End If
        listaResultados.TextMatrix(listaResultados.Rows - 1, 6) = RES1.Fields("TIPO") & ""
        listaResultados.TextMatrix(listaResultados.Rows - 1, 7) = RES1.Fields("MARCA") & ""

        If listaResultados.TextMatrix(listaResultados.Rows - 1, 4) = "" Then
            listaResultados.Row = listaResultados.Rows - 1
            listaResultados.Col = 1
            listaResultados.CellForeColor = vbRed
        Else
            If Val(listaResultados.TextMatrix(listaResultados.Rows - 1, 4)) > 0 Then
                listaResultados.Row = listaResultados.Rows - 1
                listaResultados.Col = 1
                listaResultados.CellForeColor = vbMagenta
            Else
                If Val(listaResultados.TextMatrix(listaResultados.Rows - 1, 4)) = 0 Then
                    listaResultados.Row = listaResultados.Rows - 1
                    listaResultados.Col = 1
                    listaResultados.CellForeColor = vbBlue
                End If
            End If
        End If
        
        RES1.MoveNext
    Loop
    listaResultados.Redraw = True

End If
End Sub

Private Sub Form_Load()
    cargaIncial
End Sub

Private Sub mn_Export_Click()
    Dim ques As String
    ques = MsgBox("¿Exportar la lista a excel? ", vbYesNo + vbQuestion)
    If ques = vbYes Then
        If SSTab1.Tab = 1 Then
            Call exportExcel(listaResultados)
        Else
            Call exportExcel(ListaInvent)
        End If
    End If
End Sub

Private Sub mn_Nuevo_Click()
    nuevo_Inventario
End Sub
Private Sub nuevo_Inventario()
    sql1 = "INSERT INTO INVENTARIO (INV_FECHAHORA, INV_TIPOID, INV_PERID, INV_PERTIPO) VALUES " & _
    "(NOW(),  '" & FRM_Menu.menuBarra2.Panels(8).Text & "', '" & FRM_Menu.menuBarra2.Panels(7).Text & "', 'U')"
    con.Execute (sql1)

    sql1 = "select last_insert_id() invtId"
    Set RES1 = con.Execute(sql1)
    If Not RES1.EOF Then
        invtId = RES1.Fields("invtId")
    End If

    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(0) = False
    SSTab1.Tab = 1
    listaInventAdd.Rows = 1

End Sub
Private Sub textBus_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
           textBus(0).Text = Replace(textBus(0).Text, "'", "-")
           If Left(textBus(0).Text, 1) = " " Then
                textBus(0).Text = Right(textBus(0).Text, (Len(textBus(0).Text) - 1))
           End If
        'addProducto
        addInventario
    End If
End Sub
Private Sub addProducto()
    Dim cantidad As Long
    Dim encontrado As Boolean
    
    sql1 = "SELECT * FROM VIEW_PRODUCTOS_INVENTARIO WHERE CODIGO = '" & textBus(0).Text & "'"
    Set RES1 = con.Execute(sql1)
    encontrado = False
    listaInventAdd.Redraw = False
    If Not RES1.EOF Then
        
        For b1 = 1 To listaInventAdd.Rows - 1
            If listaInventAdd.TextMatrix(b1, 0) = RES1.Fields("CODIGO") Then
                listaInventAdd.TextMatrix(listaInventAdd.Rows - 1, 3) = Val(listaInventAdd.TextMatrix(listaInventAdd.Rows - 1, 3)) + 1
                encontrado = True
                Exit For
            End If
        Next b1
        
        If encontrado = False Then
            listaInventAdd.AddItem ""
            listaInventAdd.TextMatrix(listaInventAdd.Rows - 1, 0) = RES1.Fields("CODIGO")
            listaInventAdd.TextMatrix(listaInventAdd.Rows - 1, 1) = RES1.Fields("NOMBRE")
            listaInventAdd.TextMatrix(listaInventAdd.Rows - 1, 2) = RES1.Fields("MARCA")
            listaInventAdd.TextMatrix(listaInventAdd.Rows - 1, 3) = "1"
            listaInventAdd.TextMatrix(listaInventAdd.Rows - 1, 4) = ""
            listaInventAdd.TextMatrix(listaInventAdd.Rows - 1, 5) = Now()
        End If
    End If
    listaInventAdd.Redraw = True
End Sub
Private Sub time1_Timer()
    time1.Enabled = False
    SSTab1.width = Me.width - 50
    SSTab1.height = Me.height
    ListaInvent.width = Me.width - 500
    listaResultados.width = Me.width - 6000
    listaInventAdd.height = Me.height - 3000
    listaResultados.height = Me.height - 3000
    
End Sub
