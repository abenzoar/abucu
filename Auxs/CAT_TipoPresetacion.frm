VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form CAT_TipoPresetacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogo de tipos de presentación"
   ClientHeight    =   7725
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   6165
   Icon            =   "CAT_TipoPresetacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   6165
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Atlanta"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   240
      MaskColor       =   &H0000FFFF&
      TabIndex        =   8
      Top             =   720
      UseMaskColor    =   -1  'True
      Width           =   210
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Atlanta"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   2235
      MaskColor       =   &H0000FFFF&
      TabIndex        =   7
      Top             =   720
      UseMaskColor    =   -1  'True
      Width           =   210
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Atlanta"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   3675
      MaskColor       =   &H0000FFFF&
      TabIndex        =   6
      Top             =   720
      UseMaskColor    =   -1  'True
      Width           =   210
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
      Height          =   735
      Index           =   0
      Left            =   3840
      Picture         =   "CAT_TipoPresetacion.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtProd 
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
      MaxLength       =   65
      TabIndex        =   0
      Top             =   1680
      Width           =   3135
   End
   Begin MSFlexGridLib.MSFlexGrid listCatalogo 
      Height          =   5055
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   8916
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
      FormatString    =   "Clave     | Descripción                                                  "
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
   Begin VB.Label lProd 
      BackStyle       =   0  'Transparent
      Caption         =   "Unidad medida"
      BeginProperty Font 
         Name            =   "Atlanta"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   555
      TabIndex        =   11
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lProd 
      BackStyle       =   0  'Transparent
      Caption         =   "Talla gral"
      BeginProperty Font 
         Name            =   "Atlanta"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   18
      Left            =   2475
      TabIndex        =   10
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lProd 
      BackStyle       =   0  'Transparent
      Caption         =   "Talla especifica"
      BeginProperty Font 
         Name            =   "Atlanta"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   19
      Left            =   3915
      TabIndex        =   9
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Agregar"
      BeginProperty Font 
         Name            =   "Atlanta"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   5
      Top             =   7440
      Width           =   1575
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos en lista:"
      BeginProperty Font 
         Name            =   "Atlanta"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   4
      Top             =   7440
      Width           =   3375
   End
   Begin VB.Label lProd 
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción de la presentación"
      BeginProperty Font 
         Name            =   "Atlanta"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H00800080&
      BorderWidth     =   4
      Height          =   435
      Index           =   16
      Left            =   240
      Top             =   1680
      Width           =   3165
   End
   Begin VB.Image Image2 
      Height          =   8415
      Index           =   1
      Left            =   -360
      Picture         =   "CAT_TipoPresetacion.frx":0E54
      Stretch         =   -1  'True
      Top             =   -360
      Width           =   15255
   End
   Begin VB.Menu mn_Ayuda 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "CAT_TipoPresetacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim SQL1 As String
    Dim RES1 As Recordset
    Dim tipoId As Long
    Dim Tipo As String

Private Sub cmBoton_Click(Index As Integer)
    If txtProd(0).Text <> "" Then
        If lInfo(0).Caption = "Agregar" Then
            guardarTipo
        Else
            editarTipo
        End If
    Else
        MsgBox "Se ha detectado un error. Por favor verifique.", vbExclamation
    End If

End Sub

Private Sub guardarTipo()
    

    SQL1 = "INSERT INTO CAT_PRESENTACION (CTPS_NOMBRE, CTPS_TIPO) VALUES " & _
    "('" & txtProd(0).Text & "', '" & Tipo & "')"
    con.Execute (SQL1)
    
    MsgBox "Información guardada.", vbInformation
    
    lInfo(0).Caption = "Agregar"
    cargaPresentacion (Tipo)
    FRM_Productos.cmdPresentacion_Click
    
End Sub
Private Sub editarTipo()
    SQL1 = "UPDATE CAT_PRESENTACION SET CTPS_NOMBRE = '" & txtProd(0).Text & "' " & _
    "WHERE CTPs_ID = '" & tipoId & "'"
    con.Execute (SQL1)
    
    MsgBox "Información guardada.", vbInformation
    lInfo(0).Caption = "Agregar"
    cargaPresentacion (Tipo)
    FRM_Productos.cmdPresentacion_Click
End Sub
Private Sub Form_Load()
    Option1_Click (0)
    Option1(0).value = True
    lInfo(0).Caption = "Agregar"
End Sub

Private Sub listCatalogo_DblClick()
    txtProd(0).Text = listCatalogo.TextMatrix(listCatalogo.Row, 1)
    tipoId = listCatalogo.TextMatrix(listCatalogo.Row, 0)
    lInfo(0).Caption = "Editar"
End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0:
            Tipo = "U"
            cargaPresentacion ("U")
        Case 1:
            Tipo = "T"
            cargaPresentacion ("T")
        Case 2:
            Tipo = "M"
            cargaPresentacion ("M")
    End Select

End Sub

Private Sub cargaPresentacion(tipoPresen As String)

    SQL1 = ("SELECT CTPS_ID, CTPS_NOMBRE, CTPS_DESCRIPCION FROM CAT_PRESENTACION WHERE CTPS_TIPO = '" & tipoPresen & "' ORDER BY CTPS_NOMBRE")
    Set RES1 = con.Execute(SQL1)
    
    listCatalogo.Rows = 1
    Do While Not RES1.EOF
        listCatalogo.AddItem ""
        listCatalogo.TextMatrix(listCatalogo.Rows - 1, 0) = RES1.Fields("CTPS_ID")
        listCatalogo.TextMatrix(listCatalogo.Rows - 1, 1) = RES1.Fields("CTPS_NOMBRE")
        'listCatalogo.TextMatrix(listCatalogo.Rows - 1, 2) = RES1.Fields("CTPS_Descripcion") & ""
        RES1.MoveNext
    Loop

End Sub

