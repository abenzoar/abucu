VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form CAT_Periodos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Periodos de tiempo"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10275
   Icon            =   "CAT_Periodos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   10275
   StartUpPosition =   1  'CenterOwner
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
      Index           =   2
      Left            =   2880
      MaxLength       =   7
      TabIndex        =   8
      Top             =   600
      Width           =   2655
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
      Left            =   6000
      Picture         =   "CAT_Periodos.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
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
      Index           =   1
      Left            =   120
      MaxLength       =   100
      TabIndex        =   1
      Top             =   1320
      Width           =   5655
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
      Left            =   120
      MaxLength       =   15
      TabIndex        =   0
      Top             =   600
      Width           =   2655
   End
   Begin MSFlexGridLib.MSFlexGrid listCatalogo 
      Height          =   5055
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   8916
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      AllowUserResizing=   1
      FormatString    =   "Clave     | Periodo                  | Días del periodo         | Descripción                                         "
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
      Caption         =   "Catidad días del periodo"
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
      Left            =   2880
      TabIndex        =   9
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label lInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Agregar"
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
      Left            =   8040
      TabIndex        =   7
      Top             =   7080
      Width           =   2055
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos en lista:"
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
      Index           =   10
      Left            =   120
      TabIndex        =   6
      Top             =   7080
      Width           =   5775
   End
   Begin VB.Label lProd 
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción"
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
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label lProd 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo (Nombre) *"
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
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "CAT_Periodos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim SQL1 As String
    Dim RES1 As Recordset
    Dim periodoId As Long


Private Sub cmBoton_Click(Index As Integer)
    If txtProd(0).Text <> "" Then
        If lInfo(0).Caption = "Agregar" Then
            guardarPeriodo
        Else
            editaPeriodo
        End If
        
        If periodoValor = "Pagos" Then
            CAT_Pagos.cargaPeriodo
        Else
            If periodoValor = "Membresia" Then
                CAT_Membresias.cargaPeriodo
            Else
                If periodoValor = "Apartado" Then
                    FRM_Apartados.Aprt_cargaPeriodo
                End If
            End If
        End If
    Else
        MsgBox "Se ha detectado un error. Por favor verifique.", vbExclamation
    End If
End Sub
Private Sub editaPeriodo()
    SQL1 = "UPDATE CAT_PERIODO SET CTPR_PERIODO = '" & txtProd(0).Text & "', " & _
    "CTPR_DESCRIPCION = '" & txtProd(1).Text & "', " & _
    "CTPR_DIAS = '" & txtProd(2).Text & "' " & _
    "WHERE CTID_PERIODO = '" & periodoId & "'"
    con.Execute (SQL1)
    
    MsgBox "Información guardada.", vbInformation
    txtProd(1).Text = ""
    lInfo(0).Caption = "Agregar"
    cargaLista
    'checkProducto
End Sub
Private Sub guardarPeriodo()
    SQL1 = "INSERT INTO CAT_PERIODO (CTPR_PERIODO, CTPR_DESCRIPCION, CTPR_DIAS) VALUES " & _
    "('" & txtProd(0).Text & "', '" & txtProd(1).Text & "', '" & txtProd(2).Text & "')"
    con.Execute (SQL1)
    
    MsgBox "Información guardada.", vbInformation
    
    txtProd(1).Text = ""
    txtProd(2).Text = ""
    lInfo(0).Caption = "Agregar"
    cargaLista
    'checkProducto
End Sub

Private Sub Form_Load()
    cargaLista
    lInfo(0).Caption = "Agregar"
End Sub
Private Sub cargaLista()
    SQL1 = "SELECT CTID_PERIODO, CTPR_PERIODO, CTPR_DIAS,  CTPR_dESCRIPCION FROM CAT_PERIODO " & _
    "WHERE CTPR_PERIODO LIKE '%" & txtProd(0).Text & "%' "
    Set RES1 = con.Execute(SQL1)
    
    listCatalogo.Rows = 1
    
    Do While Not RES1.EOF
        listCatalogo.AddItem ""
        listCatalogo.TextMatrix(listCatalogo.Rows - 1, 0) = RES1.Fields("CTID_PERIODO")
        listCatalogo.TextMatrix(listCatalogo.Rows - 1, 1) = RES1.Fields("CTPR_PERIODO")
        listCatalogo.TextMatrix(listCatalogo.Rows - 1, 2) = RES1.Fields("CTPR_DIAS")
        If IsNull(RES1.Fields("CTPR_DESCRIPCION")) Then
            listCatalogo.TextMatrix(listCatalogo.Rows - 1, 3) = ""
        Else
            listCatalogo.TextMatrix(listCatalogo.Rows - 1, 3) = RES1.Fields("CTPR_DESCRIPCION")
        End If
        RES1.MoveNext
    Loop

    lInfo(10).Caption = "Tipos en lista: " & listCatalogo.Rows - 1
End Sub


Private Sub listCatalogo_DblClick()
    txtProd(0).Text = listCatalogo.TextMatrix(listCatalogo.Row, 1)
    txtProd(1).Text = listCatalogo.TextMatrix(listCatalogo.Row, 3)
    txtProd(2).Text = listCatalogo.TextMatrix(listCatalogo.Row, 2)
    periodoId = listCatalogo.TextMatrix(listCatalogo.Row, 0)
    lInfo(0).Caption = "Editar"

End Sub

Private Sub txtProd_Change(Index As Integer)
    If Index = 0 And lInfo(0).Caption = "Agregar" Then
        cargaLista
    End If
End Sub

Private Sub txtProd_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = 13 And Index = 0 And lInfo(0).Caption = "Agregar" Then
        cargaLista
    Else
        If KeyAscii = 27 Then
            Unload Me
        End If
    End If
End Sub

