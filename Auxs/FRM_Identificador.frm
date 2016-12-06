VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRM_Identificador 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Identificador de usuario"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7080
   Icon            =   "FRM_Identificador.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   7080
   StartUpPosition =   1  'CenterOwner
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
      Picture         =   "FRM_Identificador.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5400
      Width           =   1695
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
      Picture         =   "FRM_Identificador.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid lista1 
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   1320
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
   Begin VB.TextBox txtPass 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "FRM_Identificador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL1 As String
Dim resUsuario As Recordset
Dim resUsuMesa As Recordset
Dim resSuc As Recordset
Dim resAsistencia As Recordset
Dim salida As Boolean
Dim altura As Long

Private Sub Command1_Click()
        verificarUsuario
End Sub

Private Sub Command2_Click()
Unload Me
'            If Me.Caption = "Indentificador de usuario - Menu" Then
'                Unload Me
'            End If
End Sub

Private Sub Form_Load()
salida = True
Me.height = 1515
carga_numeros
'If Me.Caption = "Indentificador de usuario - Menu" Then
If tipo_AccesoTouch = "Indentificador de usuario - Menu" Then

Else
'    Call ConexionDB("localhost", "auXs_Db", "root", "9807288")
    loadDb = True
    Call buscarConexiones("actual")

End If

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

Private Sub Form_Unload(Cancel As Integer)
If Me.Caption = "Indentificador de usuario - Menu" Then

Else
    If salida = True Then
        Cancel = 0
        End
    End If
End If

End Sub

Private Sub lista1_Click()
    txtPass.PasswordChar = "*"
    txtPass.FontBold = True
    txtPass.FontSize = 36
    txtPass = txtPass & lista1.TextMatrix(lista1.Row, lista1.Col)
End Sub

Private Sub txtPass_DblClick()
    '    Shell "osk.exe"
    If Me.height = 7020 Then
    Else
        If Me.height = 1515 Then
            Me.Top = Me.Top - 2500
            Me.height = 7020
        Else
            Me.height = 1515
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
            If Me.Caption = "Indentificador de usuario - Menu" Then
                Unload Me
            End If
        End If
    End If

End Sub


Private Sub verificarUsuario()
    Dim fila As Integer
    Dim columna As Integer
    SQL1 = "SELECT * FROM VIEW_PERSONA WHERE TIPO = 'USUARIO' AND PASS = MD5('" & txtPass.Text & "') AND PERTP_STATUS = 'A' "
    Set resUsuario = con.Execute(SQL1)
    salida = True
    If Not resUsuario.EOF Then
        '' Para verificar si se permite accesar si no ha checado asistencia
        SQL1 = "SELECT SUC_VENTA_ASISTENCIA FROM SUCURSAL "
        Set resSuc = con.Execute(SQL1)
        
        If Not resSuc.EOF Then
            If resSuc.Fields("SUC_VENTA_ASISTENCIA") = "S" Then
                'resAsistencia
            End If
        End If
        
        salida = False
        
        ''--Para las cancelaciones
        If tipoIdentificador = "PRODUCTO-OPERACION" Then
            'tipoIdentificador = "N"
            tipoIdentificador = "PRODUCTO-OPERACION"
            FRM_NotaProducto.Show vbModal
            'cancelFila (FrmFocus.lista.Row)
        Else
        
        
            ''--Para la comandera
            If Me.Caption = "Indentificador de usuario - Menu" Then
    
            Else
                FRM_OperTouch.carga_mesas
            End If
            FRM_OperTouch.lblUserId(0).Caption = resUsuario.Fields("PERID")
            FRM_OperTouch.lblUserId(1).Caption = resUsuario.Fields("TIPOID")
            FRM_OperTouch.lblUserId(2).Caption = resUsuario.Fields("TIPO_TIPO")
            FRM_OperTouch.Caption = "Operaciones   - " & resUsuario.Fields("USUARIO")
            SQL1 = "SELECT MESA FROM VIEW_VENTAS WHERE STATUS = 'ABIERTO' AND MESA IS NOT NULL AND " & _
            "CONCAT(USU_PERID, USU_TIPOID, USU_TIPO) = '" & resUsuario.Fields("PERID") & resUsuario.Fields("TIPOID") & resUsuario.Fields("TIPO_TIPO") & "' "
            Set resUsuMesa = con.Execute(SQL1)
            Do While Not resUsuMesa.EOF
                For b1 = 1 To FRM_OperTouch.lista_Mesa.Rows - 1
                    For c1 = 0 To 1
                        If FRM_OperTouch.lista_Mesa.TextMatrix(b1, c1) = resUsuMesa.Fields("MESA") Then
                            FRM_OperTouch.lista_Mesa.Row = b1
                            FRM_OperTouch.lista_Mesa.Col = c1
                            FRM_OperTouch.lista_Mesa.CellFontBold = True
                            FRM_OperTouch.lista_Mesa.CellFontSize = 16
                            FRM_OperTouch.lista_Mesa.CellForeColor = vbBlue
                            
                        End If
                    Next c1
                Next b1
                resUsuMesa.MoveNext
            Loop
            Unload Me
            
    '        FRM_OperTouch.protector = True
    '        FRM_OperTouch.Timer_Protect.Enabled = True
            FRM_OperTouch.Timer_tiempo.Enabled = True
            FRM_OperTouch.tiempo = 0
            
            If Me.Caption = "Indentificador de usuario - Menu" Then
                FRM_OperTouch.Show
            Else
                FRM_OperTouch.Show
            End If
        End If
    Else
        MsgBox "Contraseña incorrecta. Verifique", vbExclamation
         txtPass.Text = ""
    End If
    
    
End Sub
