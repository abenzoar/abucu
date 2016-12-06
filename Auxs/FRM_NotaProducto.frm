VERSION 5.00
Begin VB.Form FRM_NotaProducto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nota de información para producto"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11250
   Icon            =   "FRM_NotaProducto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   11250
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9960
      Picture         =   "FRM_NotaProducto.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9960
      Picture         =   "FRM_NotaProducto.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtDescripcion 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1200
      Width           =   9375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   9375
   End
End
Attribute VB_Name = "FRM_NotaProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    If tipoIdentificador = "PRODUCTO-OPERACION" Then
        tipoIdentificador = "N"
        MDI_Operaciones.cancelFila (FrmFocus.lista.Row)
    Else
        SQL1 = "UPDATE VENTA_DETALLE SET VENDET_DESCRIPCION = '" & txtDescripcion.Text & "' " & _
                "WHERE VENDET_FOLIO = '" & FRM_OperTouch.lista_detalle.TextMatrix(1, 1) & "' " & _
                "AND VENDET_PRODUCTOID = '" & FRM_OperTouch.lista_Producto.TextMatrix(FRM_OperTouch.lista_Producto.Row, 8) & "' AND VENDET_ID = '" & FRM_OperTouch.lista_Producto.TextMatrix(FRM_OperTouch.lista_Producto.Row, 10) & "'"
        con.Execute (SQL1)
        FRM_OperTouch.lista_Producto.TextMatrix(FRM_OperTouch.lista_Producto.Row, 4) = txtDescripcion.Text
    End If
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If tipoIdentificador = "PRODUCTO-OPERACION" Then
        
    Else
        Label1.Caption = "Mesa: " & FRM_OperTouch.lista_detalle.TextMatrix(1, 0) & vbCrLf & _
        "Atiende: " & FRM_OperTouch.lista_detalle.TextMatrix(1, 8) & vbCrLf & _
        "Producto: " & FRM_OperTouch.lista_Producto.TextMatrix(FRM_OperTouch.lista_Producto.Row, 0)
    End If
End Sub

Private Sub txtDescripcion_DblClick()
'    On Error Resume Next
'    Shell "osk.exe"
    Set formDescripcion = FRM_NotaProducto
    teclado = "obser_touch"
    FRM_Teclado.Show

End Sub
