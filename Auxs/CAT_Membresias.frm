VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form CAT_Membresias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Catálogo de Membresias"
   ClientHeight    =   7425
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   13260
   Icon            =   "CAT_Membresias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   13260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   12938
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Lista de membresías"
      TabPicture(0)   =   "CAT_Membresias.frx":058A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Check1"
      Tab(0).Control(1)=   "lista"
      Tab(0).Control(2)=   "lBus(5)"
      Tab(0).Control(3)=   "Label1"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Datos generales"
      TabPicture(1)   =   "CAT_Membresias.frx":05A6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lbStatus"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lUsuario(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lUsuario(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lUsuario(2)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lUsuario(3)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lUsuario(4)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lUsuario(5)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lUsuario(6)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lUsuario(7)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmBoton(1)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cmBoton(0)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtDato(0)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtDato(1)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Command1"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "cmbDato(0)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "cmbDato(1)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtDato(2)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "cmbDato(2)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "txtDato(3)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txtDato(4)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).ControlCount=   20
      Begin VB.CheckBox Check1 
         Height          =   255
         Left            =   -63960
         TabIndex        =   23
         Top             =   480
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.TextBox txtDato 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   4
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   21
         Top             =   3360
         Width           =   1935
      End
      Begin VB.TextBox txtDato 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   3
         Left            =   240
         MaxLength       =   50
         TabIndex        =   7
         Top             =   4800
         Width           =   1935
      End
      Begin VB.ComboBox cmbDato 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   2
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   4080
         Width           =   3375
      End
      Begin VB.TextBox txtDato 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Index           =   2
         Left            =   4920
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   2040
         Width           =   4935
      End
      Begin VB.ComboBox cmbDato 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1200
         Width           =   3375
      End
      Begin VB.ComboBox cmbDato 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2640
         Width           =   3375
      End
      Begin VB.CommandButton Command1 
         Caption         =   ">"
         Height          =   375
         Left            =   3720
         TabIndex        =   5
         Top             =   2640
         Width           =   255
      End
      Begin VB.TextBox txtDato 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   240
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox txtDato 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   240
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1200
         Width           =   4095
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   0
         Left            =   480
         Picture         =   "CAT_Membresias.frx":05C2
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   5400
         Width           =   1695
      End
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   1
         Left            =   2280
         Picture         =   "CAT_Membresias.frx":0E8C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   5400
         Width           =   1695
      End
      Begin MSFlexGridLib.MSFlexGrid lista 
         Height          =   6015
         Left            =   -74880
         TabIndex        =   1
         Top             =   840
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   10610
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         WordWrap        =   -1  'True
         AllowUserResizing=   1
         FormatString    =   $"CAT_Membresias.frx":1756
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
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ver solo activos"
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
         Index           =   5
         Left            =   -63840
         TabIndex        =   24
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Para ordenar la lista de doble clic sobre el título de la columna por la que desea ser ordenada"
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
         Left            =   -74880
         TabIndex        =   22
         Top             =   6960
         Width           =   8535
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Días del periodo"
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
         Index           =   7
         Left            =   240
         TabIndex        =   20
         Top             =   3120
         Width           =   3015
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Dias"
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
         Index           =   6
         Left            =   240
         TabIndex        =   19
         Top             =   4560
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo *"
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
         Index           =   5
         Left            =   240
         TabIndex        =   18
         Top             =   3840
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción *"
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
         Index           =   4
         Left            =   4920
         TabIndex        =   17
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Estatus *"
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
         Index           =   3
         Left            =   4920
         TabIndex        =   16
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Periodo *"
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
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Precio *"
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
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre descriptivo *"
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
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label lbStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Estatus:"
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
         Left            =   4320
         TabIndex        =   12
         Top             =   5880
         Width           =   4695
      End
   End
   Begin VB.Menu mn_Menu 
      Caption         =   "Menu"
      Begin VB.Menu mn_Add 
         Caption         =   "Agregar"
      End
      Begin VB.Menu mn_Edit 
         Caption         =   "Editar"
      End
   End
   Begin VB.Menu mn_Catalogo 
      Caption         =   "Catálogo"
      Begin VB.Menu mn_CatPeriodo 
         Caption         =   "Periodos"
      End
   End
End
Attribute VB_Name = "CAT_Membresias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL As String
Dim RES1 As Recordset
Dim RES2 As Recordset
Dim idMembresia As Long
Dim validar As Boolean

Private Sub Check1_Click()
    cargaLista
End Sub

Private Sub cmbDato_Click(Index As Integer)
    If Index = 2 Then
        If Left(cmbDato(2).Text, 1) = "P" Then
            txtDato(3).Visible = True
            lUsuario(6).Visible = True
        Else
            txtDato(3).Visible = False
            lUsuario(6).Visible = False
            txtDato(3).Text = txtDato(4).Text
        End If
    Else
        If Index = 0 Then
            sql1 = "SELECT CTPR_DIAS FROM CAT_PERIODO WHERE CTID_PERIODO = '" & cmbDato(0).ItemData(cmbDato(0).ListIndex) & "'"
            Set RES1 = con.Execute(sql1)
            'MsgBox SQL1
            'cmbDato(0).Clear
            If Not RES1.EOF Then
                txtDato(4).Text = RES1.Fields("CTPR_DIAS")
            End If
            
        End If
    End If
End Sub

Private Sub cmBoton_Click(Index As Integer)
    
    If Index = 0 Then
        validar = False
        checkCampos
        If validar = False Then
            crearRegistro
        End If
    Else
        If Index = 1 Then
            Dim ques As String
            ques = MsgBox("¿Cancelar?", vbYesNo + vbQuestion)
            If ques = vbYes Then
                cancelar
            End If
        End If
    End If

End Sub
Private Sub cancelar()
    For b1 = 0 To 4
        txtDato(b1).Text = ""
        cargaDatos
    Next b1
End Sub

Private Sub Command1_Click()
    periodoValor = "Membresia"
    CAT_Periodos.Show vbModal

End Sub

Private Sub Form_Load()
    cargaDatos
End Sub
Public Sub cargaPeriodo()
    validar = False
    
    sql1 = "SELECT CTID_PERIODO, CTPR_PERIODO, CTPR_DIAS FROM CAT_PERIODO"
    Set RES1 = con.Execute(sql1)
    cmbDato(0).Clear
    Do While Not RES1.EOF
        cmbDato(0).AddItem RES1.Fields("CTPR_PERIODO")
        cmbDato(0).ItemData(cmbDato(0).ListCount - 1) = RES1.Fields("CTID_PERIODO")
        RES1.MoveNext
    Loop
End Sub

Private Sub cargaDatos()
    cmbDato(1).Clear
    cmbDato(1).AddItem "ACTIVO"
    cmbDato(1).AddItem "INACTIVO"
    cmbDato(1).ListIndex = 0
    cargaPeriodo
    SSTab1.Tab = 0
    SSTab1.TabEnabled(1) = False
    lbStatus.Caption = "Estatus: Agregando membresía"
    validar = False
    cmbDato(2).Clear
    cmbDato(2).AddItem "CONSECUTIVO"
    cmbDato(2).AddItem "PERIODICO"
    cmbDato(2).ListIndex = 0
    
    cargaLista
    
End Sub

Private Sub Lista_Click()
''''''asdsadas
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

Private Sub Lista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lista.Rows > 1 Then
        If Button = vbRightButton Then
            mn_Add.Enabled = True
            mn_Edit.Enabled = True
            PopupMenu mn_menu, vbPopupMenuLeftAlign
        End If
    Else
            mn_Add.Enabled = True
            mn_Edit.Enabled = False
        If Button = vbRightButton Then
            PopupMenu mn_menu, vbPopupMenuLeftAlign
        End If
    End If

End Sub

Private Sub mn_Add_Click()
    Dim ques As String
    
    ques = MsgBox("¿Desea agregar una membresía?", vbYesNo + vbQuestion)
        If ques = vbYes Then
            lbStatus.Caption = "Estatus: Agregando membresía"
            SSTab1.TabEnabled(1) = True
            SSTab1.Tab = 1
            SSTab1.TabEnabled(0) = False
'            txtUsuario(0).SetFocus
'            save = False
        End If
    
End Sub

Private Sub mn_Edit_Click()
    Dim ques As String
    
    ques = MsgBox("¿Desea editar la membresía " & lista.TextMatrix(lista.Row, 1) & "?", vbYesNo + vbQuestion)
        If ques = vbYes Then
            lbStatus.Caption = "Estatus: Editando membresía"
            SSTab1.TabEnabled(1) = True
            SSTab1.Tab = 1
            SSTab1.TabEnabled(0) = False
            idMembresia = lista.TextMatrix(lista.Row, 0)
            cargaEdit (idMembresia)
        End If

End Sub
Private Sub cargaEdit(valor As Long)
    sql1 = "SELECT T1.ctmb_Id, T1.ctmb_Status, IF(ctmb_Status='A', 'ACTIVO', 'INACTIVO') STATUS, " & _
    "T1.ctmb_Precio, T1.ctmb_Nombre, T1.ctmb_Descripcion, T2.ctpr_Periodo, T2.ctpr_DIAS, IF(CTMB_TIPO='P', 'PERIODICO', 'CONSECUTIVO') TIPO, CTMB_DIAS " & _
    "FROM CAT_MEMBRESIAS T1, CAT_PERIODO T2 " & _
    "WHERE T1.CTMB_PERIODOID = T2.CTID_PERIODO AND T1.CTMB_ID = '" & valor & "'"
    Set RES2 = con.Execute(sql1)
    
    If Not RES2.EOF Then
        txtDato(0).Text = RES2.Fields("ctmb_nombre")
        txtDato(1).Text = RES2.Fields("ctmb_precio")
        txtDato(2).Text = RES2.Fields("ctmb_descripcion")
        txtDato(0).Text = RES2.Fields("ctmb_nombre")
        txtDato(3).Text = RES2.Fields("ctmb_dias")
        cmbDato(0).Text = RES2.Fields("CTPR_PERIODO")
        cmbDato(1).Text = RES2.Fields("STATUS")
        cmbDato(2).Text = RES2.Fields("TIPO")
    End If
    
    
End Sub
Private Sub cargaLista()
    Dim texto1 As String
    texto1 = " "
    If Check1.value = Checked Then
        texto1 = texto1 & " AND ctmb_Status = 'A' "
    End If

    sql1 = "SELECT T1.ctmb_Id, T1.ctmb_Status, IF(ctmb_Status='A', 'ACTIVO', 'INACTIVO') STATUS, " & _
    "T1.ctmb_Precio, T1.ctmb_Nombre, T1.ctmb_Descripcion, T2.ctpr_Periodo, T2.ctpr_DIAS, IF(CTMB_TIPO='P', 'PERIODICO', 'CONSECUTIVO') TIPO, CTMB_DIAS " & _
    "FROM CAT_MEMBRESIAS T1, CAT_PERIODO T2 " & _
    "WHERE T1.CTMB_PERIODOID = T2.CTID_PERIODO " & texto1 & " order by T1.ctmb_Status ASC"
    'MsgBox sql1
    Set RES1 = con.Execute(sql1)
    lista.Rows = 1
    Do While Not RES1.EOF
        lista.AddItem ""
        lista.TextMatrix(lista.Rows - 1, 0) = RES1.Fields("CTMB_ID")
        lista.TextMatrix(lista.Rows - 1, 1) = RES1.Fields("CTMB_NOMBRE")
        lista.TextMatrix(lista.Rows - 1, 2) = RES1.Fields("CTMB_PRECIO")
        lista.TextMatrix(lista.Rows - 1, 3) = RES1.Fields("CTPR_PERIODO")
        lista.TextMatrix(lista.Rows - 1, 5) = RES1.Fields("STATUS")
        lista.TextMatrix(lista.Rows - 1, 6) = RES1.Fields("TIPO")
        lista.TextMatrix(lista.Rows - 1, 7) = RES1.Fields("CTMB_DIAS")
        lista.TextMatrix(lista.Rows - 1, 8) = RES1.Fields("CTMB_DESCRIPCION")
        lista.TextMatrix(lista.Rows - 1, 4) = RES1.Fields("CTPR_DIAS")
        RES1.MoveNext
    Loop
End Sub
Private Sub crearRegistro()

    If lbStatus.Caption = "Estatus: Agregando membresía" Then
        If vaiidar = False Then
            sql1 = "INSERT INTO CAT_MEMBRESIAS  (CTMB_STATUS, CTMB_PRECIO, CTMB_PERIODOID, CTMB_NOMBRE, CTMB_DESCRIPCION, CTMB_TIPO, CTMB_DIAS) VALUES " & _
            "('" & Left(cmbDato(1).Text, 1) & "', '" & txtDato(1).Text & "', '" & cmbDato(0).ItemData(cmbDato(0).ListIndex) & "', " & _
            "'" & txtDato(0).Text & "', '" & txtDato(2).Text & "', '" & Left(cmbDato(2).Text, 1) & "', '" & txtDato(3).Text & "')"
            'MsgBox SQL1
            con.Execute (sql1)
            
            MsgBox "Registro de membresía realizado.", vbInformation
            cargaDatos
            cancelar
        End If
    Else
        If lbStatus.Caption = "Estatus: Editando membresía" Then
            sql1 = "UPDATE CAT_MEMBRESIAS SET CTMB_NOMBRE = '" & txtDato(0).Text & "', " & _
            "CTMB_STATUS='" & Left(cmbDato(1).Text, 1) & "', " & _
            "CTMB_PERIODOID='" & cmbDato(0).ItemData(cmbDato(0).ListIndex) & "', " & _
            "CTMB_PRECIO='" & txtDato(1).Text & "', " & _
            "CTMB_TIPO='" & Left(cmbDato(2).Text, 1) & "', " & _
            "CTMB_DIAS='" & txtDato(3).Text & "', " & _
            "CTMB_DESCRIPCION='" & txtDato(2).Text & "' " & _
            "WHERE CTMB_ID = '" & idMembresia & "'"
            con.Execute (sql1)
            MsgBox "Registro de membresía realizado.", vbInformation
            cargaDatos
            cancelar
            
        End If
    End If
End Sub
Private Sub checkCampos()

    If txtDato(3).Visible = False And txtDato(3).Text = "" Then
        txtDato(3).Text = txtDato(4).Text
    End If
    
    For b1 = 0 To 2
        If txtDato(b1).Text = "" Then
            validar = True
            MsgBox "Se requiere un valor. Verifique.", vbInformation
            Exit Sub
        End If
    Next b1
    
    For b1 = 0 To 2
        If cmbDato(b1).Text = "" Then
            validar = True
            MsgBox "Se requiere un valor. Verifique.", vbInformation
            Exit Sub
        End If
    Next b1
    
    If Left(cmbDato(2).Text, 1) = "P" Then
        If txtDato(3).Text = "" Then
            validar = True
            MsgBox "Se requiere un valor. Verifique.", vbInformation
            Exit Sub
        Else
             If Val(txtDato(3).Text) > Val(txtDato(4).Text) Then
                validar = True
                MsgBox "Los días no pueden ser mayor que los días del periodo.", vbInformation
                Exit Sub
             End If
        End If
    End If
End Sub


Private Sub txtDato_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 1: Call NumerosPunto(KeyAscii)
        Case 3: Call Numeros(KeyAscii)
        Case 4: Call Numeros(KeyAscii)
    End Select
End Sub
