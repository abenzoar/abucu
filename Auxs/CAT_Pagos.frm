VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form CAT_Pagos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogo de pagos y comisiones"
   ClientHeight    =   7185
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   12255
   Icon            =   "CAT_Pagos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   12255
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Lista de pagos"
      TabPicture(0)   =   "CAT_Pagos.frx":058A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Lista"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Datos generales"
      TabPicture(1)   =   "CAT_Pagos.frx":05A6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lUsuario(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lUsuario(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lUsuario(2)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lUsuario(3)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lUsuario(4)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lUsuario(5)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lbStatus"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lUsuario(6)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lUsuario(7)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lUsuario(8)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtDato(0)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "cmbDato(0)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "cmbDato(1)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtDato(1)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "cmbDato(2)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "cmBoton(0)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "cmBoton(1)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Command1"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Check1"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "cmbDato(3)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "cmbDato(4)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "txtDato(2)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).ControlCount=   22
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
         Height          =   1095
         Index           =   2
         Left            =   240
         MaxLength       =   1500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   4920
         Width           =   11535
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
         Index           =   4
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   3840
         Width           =   2895
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
         Index           =   3
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   3840
         Width           =   2895
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Aplica sobre productos o servicios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   3000
         Width           =   4095
      End
      Begin VB.CommandButton Command1 
         Caption         =   ">"
         Height          =   375
         Left            =   3720
         TabIndex        =   16
         Top             =   2040
         Width           =   255
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
         Left            =   2040
         Picture         =   "CAT_Pagos.frx":05C2
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   6240
         Width           =   1695
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
         Left            =   240
         Picture         =   "CAT_Pagos.frx":0E8C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   6240
         Width           =   1695
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
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2040
         Width           =   2895
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
         Left            =   7920
         MaxLength       =   50
         TabIndex        =   8
         Top             =   2040
         Width           =   2055
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2040
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
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1080
         Width           =   3855
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
         Top             =   1080
         Width           =   4095
      End
      Begin MSFlexGridLib.MSFlexGrid Lista 
         Height          =   5535
         Left            =   -74880
         TabIndex        =   1
         Top             =   600
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   9763
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         AllowUserResizing=   1
         FormatString    =   $"CAT_Pagos.frx":1756
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
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Cálculo"
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
         Index           =   8
         Left            =   240
         TabIndex        =   23
         Top             =   4680
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Subtipo "
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
         Left            =   3600
         TabIndex        =   21
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Aplica a:"
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
         Top             =   3600
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
         Left            =   4080
         TabIndex        =   15
         Top             =   6720
         Width           =   4695
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Periodo de pago *"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   12
         Top             =   2520
         Width           =   6735
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo pago *"
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
         Left            =   4560
         TabIndex        =   10
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad de pago *"
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
         Left            =   7920
         TabIndex        =   9
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Periodo de pago *"
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
         TabIndex        =   6
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label lUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Concepto de pago *"
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
         Left            =   4560
         TabIndex        =   4
         Top             =   840
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
         TabIndex        =   3
         Top             =   840
         Width           =   2415
      End
   End
   Begin VB.Menu mn_Pagos 
      Caption         =   "Pagos"
      Begin VB.Menu mn_Add 
         Caption         =   "Agregar"
      End
      Begin VB.Menu mn_Edit 
         Caption         =   "Editar"
      End
   End
   Begin VB.Menu mn_Cat 
      Caption         =   "Catálogo"
      Begin VB.Menu mn_Periodo 
         Caption         =   "Periodo"
      End
   End
End
Attribute VB_Name = "CAT_Pagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql1 As String
Dim RES1 As Recordset
Dim RES2 As Recordset
Dim validar As Boolean
Dim idPago As Long

Private Sub Check1_Click()
    If Check1.value = Checked Then
        cmbDato(3).Enabled = True
        cmbDato(4).Enabled = True
    Else
        cargaSubTipos
    End If
End Sub

Private Sub cmbDato_Click(Index As Integer)
    If Index = 0 Then
        cmbDato(2).Clear
        If cmbDato(0).ListIndex = 1 Then
            cmbDato(2).AddItem "EFECTIVO"
        Else
            cmbDato(2).AddItem "PORCENTAJE"
            cmbDato(2).AddItem "CALCULADO"
        End If
        cmbDato(2).ListIndex = 0
    Else
        If Index = 1 Then
                sql1 = "select * from cat_periodo " & _
                "where ctid_Periodo = '" & cmbDato(1).ItemData(cmbDato(1).ListIndex) & "'"
                Set RES1 = con.Execute(sql1)
                
                If Not RES1.EOF Then
                    lUsuario(5).Caption = "DIAS: " & RES1.Fields("CTPR_dIAS")
                Else
                    lUsuario(5).Caption = ""
                End If
        Else
            If Index = 2 Then
                
                If cmbDato(2).Text = "PORCENTAJE" Then
                    txtDato(2).Enabled = False
                Else
                    txtDato(2).Enabled = True
                End If
            End If
        End If
    End If

End Sub
Public Sub cargaPeriodo()
    sql1 = "SELECT CTID_PERIODO, CTPR_PERIODO, CTPR_DIAS FROM CAT_PERIODO"
    Set RES1 = con.Execute(sql1)
    cmbDato(1).Clear
    Do While Not RES1.EOF
        cmbDato(1).AddItem RES1.Fields("CTPR_PERIODO")
        cmbDato(1).ItemData(cmbDato(1).ListCount - 1) = RES1.Fields("CTID_PERIODO")
        RES1.MoveNext
    Loop
End Sub
Private Sub cmBoton_Click(Index As Integer)
    If Index = 0 Then
        crearRegistro
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
    limpiarDatos
    cargaDatos
End Sub
Private Sub crearRegistro()
    checkCampos
    txtDato(2).Text = Replace(txtDato(2).Text, "'", "\'")
    If lbStatus.Caption = "Estatus: Agregando pago" Then
        If vaiidar = False Then
            sql1 = "INSERT INTO CAT_PAGOS (CTPG_NOMBRE, CTPG_TIPOPAGO, CTPG_IDPERIODO, CTPG_VALOR, CTPG_TIPOVALOR, " & _
            "CTPG_APLICAVALORES, CTPG_APLICATIPO, CTPG_APLICASUBTIPO, CTPG_REGLA) VALUES " & _
            "('" & txtDato(0).Text & "', '" & Left(cmbDato(0).Text, 1) & "', '" & cmbDato(1).ItemData(cmbDato(1).ListIndex) & "', " & _
            "'" & txtDato(1).Text & "', '" & Left(cmbDato(2).Text, 1) & "', '" & Check1.value & "', " & _
            "'" & Left(cmbDato(3).Text, 1) & "', '" & Left(cmbDato(4).Text, 1) & "', '" & txtDato(2).Text & "')"
            con.Execute (sql1)
            
            MsgBox "Registro de tipo de tipo pago realizado.", vbInformation
            cargaDatos
            limpiarDatos
        End If
    Else
        If lbStatus.Caption = "Estatus: Editando pago" Then
            sql1 = "UPDATE CAT_PAGOS SET CTPG_NOMBRE = '" & txtDato(0).Text & "', " & _
            "CTPG_TIPOPAGO='" & Left(cmbDato(0).Text, 1) & "', " & _
            "CTPG_IDPERIODO='" & cmbDato(1).ItemData(cmbDato(1).ListIndex) & "', " & _
            "CTPG_VALOR='" & txtDato(1).Text & "', " & _
            "CTPG_TIPOVALOR='" & Left(cmbDato(2).Text, 1) & "', " & _
            "CTPG_APLICAVALORES='" & Check1.value & "', " & _
            "CTPG_APLICATIPO='" & Left(cmbDato(3).Text, 1) & "', " & _
            "CTPG_APLICASUBTIPO='" & Left(cmbDato(4).Text, 1) & "', " & _
            "CTPG_REGLA='" & txtDato(2).Text & "' " & _
            "WHERE CTPG_ID = '" & idPago & "'"
            'MsgBox sql1
            con.Execute (sql1)
            MsgBox "Registro de tipo de tipo pago realizado.", vbInformation
            cargaDatos
            limpiarDatos
            
        End If
    End If
End Sub
Private Sub limpiarDatos()
    For b1 = 0 To 1
        txtDato(b1).Text = ""
    Next b1
    
End Sub
Private Sub checkCampos()
    For b1 = 0 To 1
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
    
End Sub

Private Sub cmdPeriodo_Click()
    cargaPeriodo
End Sub

Private Sub Command1_Click()
    mn_Periodo_Click
End Sub

Private Sub Form_Load()
    cargaDatos
End Sub
Private Sub cargaDatos()
    cmbDato(0).Clear
    cmbDato(0).AddItem "COMISIONES"
    cmbDato(0).AddItem "HONORARIOS/PAGOS FIJOS"
    cargaPeriodo
    lUsuario(5).Caption = ""
    SSTab1.Tab = 0
    SSTab1.TabEnabled(1) = False
    lbStatus.Caption = "Estatus: Agregando pago"
    validar = False

    Check1.value = Unchecked
    
    cargaSubTipos
    
    cargaLista
End Sub
Private Sub cargaSubTipos()
    cmbDato(3).Clear
    cmbDato(3).AddItem "PRODUCTOS"
    cmbDato(3).AddItem "SERVICIOS"
    
    cmbDato(4).Clear
    cmbDato(4).AddItem "GENERAL"
    cmbDato(4).AddItem "SUBTIPO ESPECIFICO"
    cmbDato(3).Enabled = False
    cmbDato(4).Enabled = False
    
End Sub
Private Sub cargaLista()
    sql1 = "SELECT CTPG_ID, CTPG_NOMBRE, IF(CTPG_TIPOPAGO='C', 'COMISION', 'HONORARIOS/PAGOS FIJOS') TIPO_PAGO, CTPG_IDPERIODO, " & _
    "CTPG_VALOR, IF(CTPG_TIPOVALOR='E', 'EFECTIVO', 'PORCENTAJE') TIPO_VALOR, CONCAT(CTPR_PERIODO, ' DIAS: ', CTPR_DIAS) PERIODO, " & _
    "IF(CTPG_APLICAVALORES='1', 'SI', 'NO') APLICA, IF(CTPG_APLICATIPO=NULL, 'NO', IF(CTPG_APLICATIPO='P', 'PRODUCTOS', 'SERVICIOS')) APLICA_TIPO, " & _
    "IF(CTPG_APLICASUBTIPO IS NULL, 'NO', (IF(CTPG_APLICASUBTIPO='G', 'GENERAL', 'SUBTIPOS'))) APLICA_SUBTIPO  " & _
    "FROM CAT_PAGOS, CAT_PERIODO WHERE CTPG_IDPERIODO = CTID_PERIODO"
    Set RES1 = con.Execute(sql1)
    lista.Rows = 1
    Do While Not RES1.EOF
        lista.AddItem ""
        lista.TextMatrix(lista.Rows - 1, 0) = RES1.Fields("CTPG_ID")
        lista.TextMatrix(lista.Rows - 1, 1) = RES1.Fields("CTPG_NOMBRE")
        lista.TextMatrix(lista.Rows - 1, 2) = RES1.Fields("TIPO_PAGO")
        lista.TextMatrix(lista.Rows - 1, 3) = RES1.Fields("PERIODO")
        lista.TextMatrix(lista.Rows - 1, 4) = RES1.Fields("CTPG_VALOR")
        lista.TextMatrix(lista.Rows - 1, 5) = RES1.Fields("TIPO_VALOR")
        lista.TextMatrix(lista.Rows - 1, 6) = RES1.Fields("APLICA")
        lista.TextMatrix(lista.Rows - 1, 7) = RES1.Fields("APLICA_TIPO")
        lista.TextMatrix(lista.Rows - 1, 8) = RES1.Fields("APLICA_SUBTIPO")
        RES1.MoveNext
    Loop
    
End Sub

Private Sub Lista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lista.Rows > 1 Then
        If Button = vbRightButton Then
            mn_Add.Enabled = True
            mn_Edit.Enabled = True
            PopupMenu mn_Pagos, vbPopupMenuLeftAlign
        End If
    Else
            mn_Add.Enabled = True
            mn_Edit.Enabled = False
        If Button = vbRightButton Then
            PopupMenu mn_Pagos, vbPopupMenuLeftAlign
        End If
    End If

End Sub

Private Sub mn_Add_Click()
    Dim ques As String
    
    ques = MsgBox("¿Desea agregar un tipo de pago?", vbYesNo + vbQuestion)
        If ques = vbYes Then
            lbStatus.Caption = "Estatus: Agregando pago"
            SSTab1.TabEnabled(1) = True
            SSTab1.Tab = 1
            SSTab1.TabEnabled(0) = False
'            txtUsuario(0).SetFocus
'            save = False
        End If

    
End Sub

Private Sub mn_Edit_Click()
    Dim ques As String
    
    ques = MsgBox("¿Desea editar el tipo de pago " & lista.TextMatrix(lista.Row, 1) & "?", vbYesNo + vbQuestion)
        If ques = vbYes Then
            lbStatus.Caption = "Estatus: Editando pago"
            SSTab1.TabEnabled(1) = True
            SSTab1.Tab = 1
            SSTab1.TabEnabled(0) = False
            idPago = lista.TextMatrix(lista.Row, 0)
            cargaEdit (idPago)
        End If
    
End Sub
Private Sub cargaEdit(b1 As Long)
    sql1 = "SELECT CTPG_ID, CTPG_NOMBRE, IF(CTPG_TIPOPAGO='C', 'COMISIONES', 'HONORARIOS/PAGOS FIJOS') TIPO_PAGO, CTPG_IDPERIODO, " & _
    "CTPG_VALOR, IF(CTPG_TIPOVALOR='E', 'EFECTIVO', 'PORCENTAJE') TIPO_VALOR, CTPR_PERIODO, CONCAT(CTPR_PERIODO, ' DIAS: ', CTPR_DIAS) PERIODO, " & _
    "CTPG_APLICAVALORES, IF(CTPG_APLICATIPO=NULL, 'NO', IF(CTPG_APLICATIPO='P', 'PRODUCTOS', 'SERVICIOS')) APLICA_TIPO, " & _
    "IF(CTPG_APLICASUBTIPO IS NULL, 'NO', (IF(CTPG_APLICASUBTIPO='G', 'GENERAL', 'SUBTIPO ESPECIFICO'))) APLICA_SUBTIPO, CTPG_REGLA  " & _
    "FROM CAT_PAGOS, CAT_PERIODO WHERE CTPG_IDPERIODO = CTID_PERIODO AND CTPG_ID = '" & b1 & "'"
    Set RES2 = con.Execute(sql1)
    
    If Not RES2.EOF Then
        txtDato(0).Text = RES2.Fields("CTPG_NOMBRE")
        txtDato(1).Text = RES2.Fields("CTPG_VALOR")
        cmbDato(0).Text = RES2.Fields("TIPO_PAGO")
        cmbDato(1).Text = RES2.Fields("CTPR_PERIODO")
        cmbDato(2).Text = RES2.Fields("TIPO_VALOR")
        cmbDato(3).Text = RES2.Fields("APLICA_TIPO")
        cmbDato(4).Text = RES2.Fields("APLICA_SUBTIPO")
        Check1.value = RES2.Fields("CTPG_APLICAVALORES")
        txtDato(2).Text = RES2.Fields("CTPG_REGLA") & ""
    End If
End Sub
Private Sub mn_Periodo_Click()
    periodoValor = "Pagos"
    CAT_Periodos.Show vbModal
End Sub
