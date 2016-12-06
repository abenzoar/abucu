VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form CAT_Marca 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogo de marcas"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   10305
   Icon            =   "CAT_Marca.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   10305
   StartUpPosition =   1  'CenterOwner
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
      Left            =   8640
      Picture         =   "CAT_Marca.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
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
      Left            =   2880
      MaxLength       =   65
      TabIndex        =   1
      Top             =   1080
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
      MaxLength       =   65
      TabIndex        =   0
      Top             =   1080
      Width           =   2655
   End
   Begin MSFlexGridLib.MSFlexGrid listCatalogo 
      Height          =   5055
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   8916
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      AllowUserResizing=   1
      FormatString    =   "Clave     | Marca                                   | Descripción                                         "
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
   Begin VB.Shape Borde 
      BorderColor     =   &H00800080&
      BorderWidth     =   4
      Height          =   435
      Index           =   16
      Left            =   120
      Top             =   1080
      Width           =   2685
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H00800080&
      BorderWidth     =   4
      Height          =   435
      Index           =   0
      Left            =   2880
      Top             =   1080
      Width           =   5685
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
      Top             =   6840
      Width           =   2055
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Marcas en lista:"
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
      Top             =   6840
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
      Left            =   2880
      TabIndex        =   5
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label lProd 
      BackStyle       =   0  'Transparent
      Caption         =   "Marca *"
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
      Top             =   840
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   11760
      Index           =   1
      Left            =   -120
      Picture         =   "CAT_Marca.frx":0E54
      Stretch         =   -1  'True
      Top             =   -4320
      Width           =   15255
   End
   Begin VB.Menu mn_Ayuda 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "CAT_Marca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim sql1 As String
    Dim res1 As Recordset
    Dim marcaId As Long

Private Sub cmBoton_Click(Index As Integer)
    
    Call checarPermisos("CAT_MARCA", FRM_Menu.menuBarra2.Panels(8).Text)
    
    If permAdd = "SI" Then
        If txtProd(0).Text <> "" Then
            If lInfo(0).Caption = "Agregar" Then
                guardarMarca
            Else
                editaMarca
            End If
        Else
            MsgBox "Se ha detectado un error. Por favor verifique.", vbExclamation
        End If
    Else
        MsgBox "Opción no disponible.", vbInformation
    End If
        
End Sub
Private Sub editaMarca()
    sql1 = "UPDATE CAT_MARCA SET CTMR_MARCA = '" & txtProd(0).Text & "', " & _
    "CTMR_DESCRIPCION = '" & txtProd(1).Text & "' " & _
    "WHERE CTMR_ID = '" & marcaId & "'"
    con.Execute (sql1)
    
    MsgBox "Información guardada.", vbInformation
    txtProd(1).Text = ""
    lInfo(0).Caption = "Agregar"
    cargaLista
    checkProducto
End Sub
Private Sub guardarMarca()
    On Error Resume Next
    Err.Clear
    sql1 = "INSERT INTO CAT_MARCA (CTMR_MARCA, CTMR_DESCRIPCION) VALUES " & _
    "('" & txtProd(0).Text & "', '" & txtProd(1).Text & "')"
    con.Execute (sql1)
    
    If Err.Number <> 0 Then
        MsgBox "Se detecto un error, verifique los valores de la marca.", vbInformation
    Else
    
    
        MsgBox "Información guardada.", vbInformation
        
        txtProd(1).Text = ""
        lInfo(0).Caption = "Agregar"
        cargaLista
        checkProducto
    End If
End Sub
Private Sub checkProducto()
    If FRM_Productos.Visible = True Then
        FRM_Productos.cmd_Marca_Click
    End If

End Sub
Private Sub Form_Load()
    cargaLista
    lInfo(0).Caption = "Agregar"
End Sub
Private Sub cargaLista()
    On Error Resume Next
    sql1 = "SELECT CTMR_ID, CTMR_MARCA, CTMR_DESCRIPCION FROM CAT_MARCA " & _
    "WHERE CTMR_MARCA LIKE '%" & txtProd(0).Text & "%' "
    Set res1 = con.Execute(sql1)
    
    listCatalogo.Rows = 1
    
    Do While Not res1.EOF
        listCatalogo.AddItem ""
        listCatalogo.TextMatrix(listCatalogo.Rows - 1, 0) = res1.Fields("CTMR_ID")
        listCatalogo.TextMatrix(listCatalogo.Rows - 1, 1) = res1.Fields("CTMR_MARCA")
        If IsNull(res1.Fields("CTMR_DESCRIPCION")) Then
            listCatalogo.TextMatrix(listCatalogo.Rows - 1, 2) = ""
        Else
            listCatalogo.TextMatrix(listCatalogo.Rows - 1, 2) = res1.Fields("CTMR_DESCRIPCION")
        End If
        res1.MoveNext
    Loop

    lInfo(10).Caption = "Marcas en lista: " & listCatalogo.Rows - 1
End Sub

Private Sub listCatalogo_DblClick()
    
    Call checarPermisos("CAT_MARCA", FRM_Menu.menuBarra2.Panels(8).Text)
    
    If permEdit = "SI" Then
        txtProd(0).Text = listCatalogo.TextMatrix(listCatalogo.Row, 1)
        txtProd(1).Text = listCatalogo.TextMatrix(listCatalogo.Row, 2)
        marcaId = listCatalogo.TextMatrix(listCatalogo.Row, 0)
        lInfo(0).Caption = "Editar"
    Else
        MsgBox "Opción no disponible.", vbInformation
    End If
    
End Sub

Private Sub mn_Ayuda_Click()
    MsgBox "Para agregar una marca escriba los valores en los cuadros de texto y de clic en aceptar. " & vbCrLf & vbCrLf & _
    "Para editar una marca de doble clic sobre la marca que desea editar en la lista, los valores se mostrarán en los cuadros de texto y podrá cambiarlos, al concluir de clic en aceptar." & vbCrLf & vbCrLf & _
    "En la parte inferior derecha se muestra una leyenda para verificar si esta agregando o editando.", vbInformation
End Sub

Private Sub txtProd_Change(Index As Integer)
    If Index = 0 And lInfo(0).Caption = "Agregar" Then
        cargaLista
    End If
End Sub

Private Sub txtProd_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = 13 And Index = 0 And lInfo(0).Caption = "Agregar" Then
        cargaLista
    End If
End Sub
