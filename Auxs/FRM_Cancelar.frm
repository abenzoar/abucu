VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRM_Cancelar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancelacion"
   ClientHeight    =   9165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7110
   Icon            =   "FRM_Cancelar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   7110
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtMotivo 
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
      Left            =   120
      MaxLength       =   2500
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1560
      Width           =   6855
   End
   Begin VB.TextBox txtPass 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1680
      Picture         =   "FRM_Cancelar.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8040
      Width           =   1695
   End
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
      Height          =   975
      Left            =   3480
      Picture         =   "FRM_Cancelar.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8040
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid lista1 
      Height          =   3975
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3960
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   7011
      _Version        =   393216
      FixedRows       =   0
      FixedCols       =   0
      HighLight       =   0
      GridLinesFixed  =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblAutoriza 
      Caption         =   "lblAutoriza"
      Height          =   255
      Index           =   2
      Left            =   5640
      TabIndex        =   9
      Top             =   -5000
      Width           =   1095
   End
   Begin VB.Label lblAutoriza 
      Caption         =   "lblAutoriza"
      Height          =   255
      Index           =   1
      Left            =   5640
      TabIndex        =   8
      Top             =   -5000
      Width           =   1095
   End
   Begin VB.Label lblAutoriza 
      Caption         =   "lblAutoriza"
      Height          =   255
      Index           =   0
      Left            =   5640
      TabIndex        =   7
      Top             =   -5000
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Motivo"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Conrtraseña de persona que autoriza"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "FRM_Cancelar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim autorizado As Boolean
Dim resUsuario As Recordset
Dim sql1 As String
Private Sub Command1_Click()
    verificarUsuario
    If autorizado = True Then
        
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()

If cancelarMotivo = "TICKET" Then
    Me.Caption = "Autorización Re-Impresion de Ticket"
End If

autorizado = False
Me.height = 5925
lista1.Visible = False
Command1.Top = 3960
Command2.Top = 3960
carga_numeros
        lblAutoriza(0).Caption = ""
        lblAutoriza(1).Caption = ""
        lblAutoriza(2).Caption = ""


End Sub
Private Sub carga_numeros()
    lista1.Rows = 4
    lista1.Cols = 3
    lista1.ColWidth(0) = 2200
    lista1.ColWidth(1) = 2200
    lista1.ColWidth(2) = 2200
    fila = 0
    Col = 0
    For b1 = 1 To 9
        lista1.TextMatrix(fila, Col) = b1
        lista1.Col = Col
        lista1.Row = fila
        lista1.CellAlignment = 4
        
        Col = Col + 1
        If Col = 3 Then
            fila = fila + 1
            Col = 0
        End If
        If b1 = 9 Then
            lista1.TextMatrix(fila, coi + 1) = "0"
            lista1.Col = Col + 1
            lista1.Row = fila
            lista1.CellAlignment = 4
        End If
    Next b1
End Sub



Private Sub txtPass_DblClick()
    '    Shell "osk.exe"
    If Me.height = 9600 Then
    Else
        If Me.height = 5925 Then
            Me.Top = Me.Top - 2500
            Me.height = 9600
            lista1.Visible = True
            Command1.Top = 8040
            Command2.Top = 8040
            
        Else
            Me.height = 5925
        End If
    End If
End Sub


Private Sub txtPass_KeyPress(KeyAscii As Integer)
    txtPass.PasswordChar = "*"
    txtPass.FontBold = True
    txtPass.FontSize = 36
    If KeyAscii = 13 Then
        verificarUsuario
    Else
        If KeyAscii = 27 Then
            Unload Me
        End If
    End If

End Sub


Private Sub verificarUsuario()
    If txtPass.Text = "" Then
        MsgBox "Verifique los datos.", vbInformation
        Exit Sub
    Else
        If txtMotivo.Text = "" Then
            MsgBox "Verifique los datos.", vbInformation
            Exit Sub
        End If
    End If
    
    Dim fila As Integer
    Dim columna As Integer
    sql1 = "SELECT * FROM VIEW_PERSONA WHERE TIPO = 'USUARIO' AND PASS = MD5('" & txtPass.Text & "') AND PERTP_STATUS = 'A' "
    Set resUsuario = con.Execute(sql1)
    salida = True
    If Not resUsuario.EOF Then
            
        Call checarPermisos("MDI_OPERACIONES2", resUsuario.Fields("TIPOID"))
        
        If permAcceso = "SI" Then
            lblAutoriza(0).Caption = resUsuario.Fields("perid")
            lblAutoriza(1).Caption = resUsuario.Fields("tipoid")
            lblAutoriza(2).Caption = resUsuario.Fields("tipo_tipo")
            
            If cancelarMotivo = "OPERACION" Then
                MDI_Operaciones.cancelFila (FrmFocus.lista.Row)
            Else
                If cancelarMotivo = "TICKET" Then
                    sql1 = "UPDATE VENTA_DETALLE SET VENDET_NOTAMESA = NULL, vendet_MotivoReimpresion = '" & txtMotivo.Text & "', " & _
                    "vendet_AutorizaPerIdPrint = '" & FRM_Cancelar.lblAutoriza(0).Caption & "', vendet_AutorizaTipoIdPrint = '" & FRM_Cancelar.lblAutoriza(1).Caption & "', vendet_AutorizaTipoPrint = '" & FRM_Cancelar.lblAutoriza(2).Caption & "'  " & _
                    "WHERE VENDET_FOLIO = '" & FRM_OperTouch.lista_detalle.TextMatrix(FRM_OperTouch.lista_detalle.Row, 1) & "' AND VENDET_NOTAMESA = 'A' AND VENDET_PRODUCTOID = '" & FRM_OperTouch.lista_Producto.TextMatrix(FRM_OperTouch.lista_Producto.Row, 8) & "' AND VENDET_ID = '" & FRM_OperTouch.lista_Producto.TextMatrix(FRM_OperTouch.lista_Producto.Row, 10) & "'"
                    con.Execute (sql1)
                    FRM_OperTouch.lista_Producto.TextMatrix(FRM_OperTouch.lista_Producto.Row, 5) = "NO"
                Else
                    If cancelarMotivo = "OPERACION_ALL" Then
                        MDI_Operaciones.canelarOperacion (FrmFocus.lista.TextMatrix(FrmFocus.lista.Row, 0))
                        MDIC_OperTickets.cargaTickets
            '            canelarOperacion (FrmFocus.lista.TextMatrix(FrmFocus.lista.Row, 0))
            '            MDIC_OperTickets.cargaTickets
                        
                    End If
                End If
            End If
            Unload Me
        Else
            MsgBox "Usuario no permitido. Verifique.", vbInformation
            lblAutoriza(0).Caption = ""
            lblAutoriza(1).Caption = ""
            lblAutoriza(2).Caption = ""
            txtPass.Text = ""
        End If
    Else
        MsgBox "Contraseña incorrecta. Verifique", vbExclamation
         txtPass.Text = ""
        lblAutoriza(0).Caption = ""
        lblAutoriza(1).Caption = ""
        lblAutoriza(2).Caption = ""
         
    End If
    
    
End Sub


