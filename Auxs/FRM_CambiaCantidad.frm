VERSION 5.00
Begin VB.Form FRM_CambiaCantidad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar cantidad"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10410
   Icon            =   "FRM_CambiaCantidad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   10410
   StartUpPosition =   1  'CenterOwner
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
      Left            =   9000
      Picture         =   "FRM_CambiaCantidad.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
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
      Left            =   7560
      Picture         =   "FRM_CambiaCantidad.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmd_Mas 
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
      Left            =   5160
      Picture         =   "FRM_CambiaCantidad.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmd_Menos 
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
      Left            =   120
      Picture         =   "FRM_CambiaCantidad.frx":2328
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txt_Cantidad 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1680
      Width           =   3735
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
      Height          =   1215
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   9615
   End
End
Attribute VB_Name = "FRM_CambiaCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Mas_Click()
txt_Cantidad.Text = Val(txt_Cantidad) + 1


End Sub

Private Sub cmd_Menos_Click()
If Val(txt_Cantidad.Text) >= 1 Then

txt_Cantidad.Text = Val(txt_Cantidad) - 1
End If

End Sub

Private Sub Command1_Click()
    If Me.Caption = "Agregar cantidad" Then
        sql1 = "UPDATE VENTA_DETALLE SET VENDET_CANTIDAD = '" & Val(txt_Cantidad.Text) & "' " & _
                "WHERE VENDET_FOLIO = '" & FRM_OperTouch.lista_detalle.TextMatrix(1, 1) & "' " & _
                "AND VENDET_PRODUCTOID = '" & FRM_OperTouch.lista_Producto.TextMatrix(FRM_OperTouch.lista_Producto.Row, 8) & "' AND VENDET_ID = '" & FRM_OperTouch.lista_Producto.TextMatrix(FRM_OperTouch.lista_Producto.Row, 10) & "'"
        con.Execute (sql1)
        FRM_OperTouch.lista_Producto.TextMatrix(FRM_OperTouch.lista_Producto.Row, 1) = Val(txt_Cantidad.Text)
    Else
        If Me.Caption = "Agregar tiempo" Then
            sql1 = "UPDATE VENTA_DETALLE SET VENDET_TIEMPO = '" & Val(txt_Cantidad.Text) & "' " & _
                    "WHERE VENDET_FOLIO = '" & FRM_OperTouch.lista_detalle.TextMatrix(1, 1) & "' " & _
                    "AND VENDET_PRODUCTOID = '" & FRM_OperTouch.lista_Producto.TextMatrix(FRM_OperTouch.lista_Producto.Row, 8) & "' AND VENDET_ID = '" & FRM_OperTouch.lista_Producto.TextMatrix(FRM_OperTouch.lista_Producto.Row, 10) & "'"
            con.Execute (sql1)
            FRM_OperTouch.lista_Producto.TextMatrix(FRM_OperTouch.lista_Producto.Row, 2) = Val(txt_Cantidad.Text)
        Else
            If Me.Caption = "Agregar personas" Then
                sql1 = "UPDATE VENTAS SET VENT_PERSONAS = '" & Val(txt_Cantidad.Text) & "' " & _
                "WHERE VENT_IDFOLIO = '" & FRM_OperTouch.lista_detalle.TextMatrix(1, 1) & "'"
                con.Execute (sql1)
                FRM_OperTouch.lista_detalle.TextMatrix(FRM_OperTouch.lista_detalle.Row, 3) = Val(txt_Cantidad.Text)
            Else
                If Me.Caption = "Agregar asiento" Then
                    sql1 = "UPDATE VENTA_DETALLE SET VENDET_ASIENTO = '" & Val(txt_Cantidad.Text) & "' " & _
                            "WHERE VENDET_FOLIO = '" & FRM_OperTouch.lista_detalle.TextMatrix(1, 1) & "' " & _
                            "AND VENDET_PRODUCTOID = '" & FRM_OperTouch.lista_Producto.TextMatrix(FRM_OperTouch.lista_Producto.Row, 8) & "' AND VENDET_ID = '" & FRM_OperTouch.lista_Producto.TextMatrix(FRM_OperTouch.lista_Producto.Row, 10) & "'"
                    con.Execute (sql1)
                    FRM_OperTouch.lista_Producto.TextMatrix(FRM_OperTouch.lista_Producto.Row, 3) = Val(txt_Cantidad.Text)
                End If
            End If
        End If
    End If
FRM_OperTouch.lista_Mesa_Click
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
'    If FRM_CambiaCantidad.Caption = "Agregar personas" Then
        Label1.Caption = "Mesa: " & FRM_OperTouch.lista_detalle.TextMatrix(1, 0) & vbCrLf & _
        "Atiende: " & FRM_OperTouch.lista_detalle.TextMatrix(1, 8)
'    Else
'        Label1.Caption = "Mesa: " & FRM_OperTouch.lista_detalle.TextMatrix(1, 0) & vbCrLf & _
'        "Atiende: " & FRM_OperTouch.lista_detalle.TextMatrix(1, 8) & vbCrLf & _
'        "Producto: " & FRM_OperTouch.lista_Producto.TextMatrix(FRM_OperTouch.lista_Producto.Row, 0)
'    End If
End Sub
