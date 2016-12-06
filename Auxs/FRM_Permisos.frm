VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FRM_Permisos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Permisos y accesos"
   ClientHeight    =   8865
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   15510
   Icon            =   "FRM_Permisos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   15510
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   8895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15495
      _ExtentX        =   27331
      _ExtentY        =   15690
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Lista general de permisos y accesos"
      TabPicture(0)   =   "FRM_Permisos.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Lista"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Asignar permisos"
      TabPicture(1)   =   "FRM_Permisos.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Line1(5)"
      Tab(1).Control(1)=   "Label1(5)"
      Tab(1).Control(2)=   "Line1(2)"
      Tab(1).Control(3)=   "Label1(2)"
      Tab(1).Control(4)=   "lblDatos(1)"
      Tab(1).Control(5)=   "imgFoto(1)"
      Tab(1).Control(6)=   "txtClave(1)"
      Tab(1).Control(7)=   "MSFlexGrid1"
      Tab(1).Control(8)=   "cmBoton(3)"
      Tab(1).Control(9)=   "cmBoton(4)"
      Tab(1).ControlCount=   10
      Begin VB.CommandButton cmBoton 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aceptar permisos"
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
         Index           =   4
         Left            =   -71400
         Picture         =   "FRM_Permisos.frx":05C2
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1680
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
         Index           =   3
         Left            =   -69480
         Picture         =   "FRM_Permisos.frx":0E8C
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1680
         Width           =   1695
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   5655
         Left            =   -74760
         TabIndex        =   6
         Top             =   2760
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   9975
         _Version        =   393216
         Cols            =   6
         FormatString    =   "Módulo                                           | Acceso     | Adición    | Edición     | Modificación    | Eliminación      "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtClave 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   -73440
         TabIndex        =   2
         Top             =   2160
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid Lista 
         Height          =   7935
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   13996
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         FormatString    =   $"FRM_Permisos.frx":1756
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
      Begin VB.Image imgFoto 
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Index           =   1
         Left            =   -74760
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblDatos 
         BackStyle       =   0  'Transparent
         Caption         =   "Ninguno"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   -73440
         TabIndex        =   5
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario seleccionado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Index           =   2
         Left            =   -74760
         TabIndex        =   4
         Top             =   720
         Width           =   2175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   2
         X1              =   -74760
         X2              =   -71760
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Clave/Código   F3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Index           =   5
         Left            =   -73440
         TabIndex        =   3
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004080&
         Index           =   5
         X1              =   -73440
         X2              =   -71760
         Y1              =   2040
         Y2              =   2040
      End
   End
   Begin VB.Menu mn_Perm 
      Caption         =   "Permisos"
      Begin VB.Menu mn_Edit 
         Caption         =   "Editar/Asignar"
      End
   End
   Begin VB.Menu mn_Busqueda 
      Caption         =   "Busqueda"
      Begin VB.Menu mn_Usuarios 
         Caption         =   "Usuarios"
      End
   End
End
Attribute VB_Name = "FRM_Permisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL1 As String
Dim RES1 As Recordset
Private Sub Form_Load()
    SSTab1.Tab = 0
    SSTab1.TabEnabled(1) = False
    Lista.ColWidth(6) = 0
    Lista.ColWidth(7) = 0
    cargaLista
End Sub

Private Sub cargaLista()
    Lista.Rows = 1
    SQL1 = "SELECT * FROM VIEW_PERMISOS ORDER BY TIPO, MODULO ASC"
    Set RES1 = con.Execute(SQL1)

    Do While Not RES1.EOF
        Lista.AddItem ""
        Lista.TextMatrix(Lista.Rows - 1, 0) = RES1.Fields("Tipo")
        Lista.TextMatrix(Lista.Rows - 1, 1) = RES1.Fields("Modulo")
        Lista.TextMatrix(Lista.Rows - 1, 2) = RES1.Fields("ACCESO")
        Lista.TextMatrix(Lista.Rows - 1, 3) = RES1.Fields("CREACION")
        Lista.TextMatrix(Lista.Rows - 1, 4) = RES1.Fields("MODIFICAR")
        Lista.TextMatrix(Lista.Rows - 1, 5) = RES1.Fields("ELIMINAR")
        Lista.TextMatrix(Lista.Rows - 1, 6) = RES1.Fields("CLAVE_MODULO")
        Lista.TextMatrix(Lista.Rows - 1, 7) = RES1.Fields("CLAVE_TIPOUSUARIO")

        Lista.Row = Lista.Rows - 1
        Lista.Col = 2
        Lista.CellFontName = "Wingdings"
        Lista.CellFontBold = True
        Lista.CellFontSize = 16

        If RES1.Fields("ACCESO") = "SI" Then
            Lista.TextMatrix(Lista.Rows - 1, 2) = Chr(254)
        Else
            Lista.TextMatrix(Lista.Rows - 1, 2) = Chr(168)
        End If
        Lista.Col = 3
        Lista.CellFontName = "Wingdings"
        Lista.CellFontBold = True
        Lista.CellFontSize = 16
        If RES1.Fields("CREACION") = "SI" Then
            Lista.TextMatrix(Lista.Rows - 1, 3) = Chr(254)
        Else
            Lista.TextMatrix(Lista.Rows - 1, 3) = Chr(168)
        End If
        Lista.Col = 4
        Lista.CellFontName = "Wingdings"
        Lista.CellFontBold = True
        Lista.CellFontSize = 16
        If RES1.Fields("MODIFICAR") = "SI" Then
            Lista.TextMatrix(Lista.Rows - 1, 4) = Chr(254)
        Else
            Lista.TextMatrix(Lista.Rows - 1, 4) = Chr(168)
        End If
        Lista.Col = 5
        Lista.CellFontName = "Wingdings"
        Lista.CellFontBold = True
        Lista.CellFontSize = 16
        If RES1.Fields("ELIMINAR") = "SI" Then
            Lista.TextMatrix(Lista.Rows - 1, 5) = Chr(254)
        Else
            Lista.TextMatrix(Lista.Rows - 1, 5) = Chr(168)
        End If


        RES1.MoveNext
    Loop
End Sub

Private Sub Lista_DblClick()
    checar_Valores
End Sub
Private Sub checar_Valores()
    Dim b1 As Long

    Select Case Lista.Col
        Case 2:
            b1 = Lista.Row
            Lista.Row = b1
            Lista.Col = 2
            If Lista.TextMatrix(b1, 2) = Chr(168) Then
                Call actualizaDatos(Lista.TextMatrix(b1, 6), Lista.TextMatrix(b1, 7), "1", "ACCESO")
                Lista.TextMatrix(b1, 2) = Chr(254)
            Else
                Call actualizaDatos(Lista.TextMatrix(b1, 6), Lista.TextMatrix(b1, 7), "0", "ACCESO")
                Lista.TextMatrix(b1, 2) = Chr(168)
            End If
        Case 3:
            b1 = Lista.Row
            Lista.Row = b1
            Lista.Col = 3
            If Lista.TextMatrix(b1, 3) = Chr(168) Then
                Call actualizaDatos(Lista.TextMatrix(b1, 6), Lista.TextMatrix(b1, 7), "1", "CREACION")
                Lista.TextMatrix(b1, 3) = Chr(254)
            Else
                Call actualizaDatos(Lista.TextMatrix(b1, 6), Lista.TextMatrix(b1, 7), "0", "CREACION")
                Lista.TextMatrix(b1, 3) = Chr(168)
            End If
        Case 4:
            b1 = Lista.Row
            Lista.Row = b1
            Lista.Col = 4
            If Lista.TextMatrix(b1, 4) = Chr(168) Then
                Call actualizaDatos(Lista.TextMatrix(b1, 6), Lista.TextMatrix(b1, 7), "1", "MODIFICAR")
                Lista.TextMatrix(b1, 4) = Chr(254)
            Else
                Call actualizaDatos(Lista.TextMatrix(b1, 6), Lista.TextMatrix(b1, 7), "0", "MODIFICAR")
                Lista.TextMatrix(b1, 4) = Chr(168)
            End If
        Case 5:
            b1 = Lista.Row
            Lista.Row = b1
            Lista.Col = 5
            If Lista.TextMatrix(b1, 5) = Chr(168) Then
                Call actualizaDatos(Lista.TextMatrix(b1, 6), Lista.TextMatrix(b1, 7), "1", "ELIMINAR")
                Lista.TextMatrix(b1, 5) = Chr(254)
            Else
                Call actualizaDatos(Lista.TextMatrix(b1, 6), Lista.TextMatrix(b1, 7), "0", "ELIMINAR")
                Lista.TextMatrix(b1, 5) = Chr(168)
            End If



    End Select


End Sub
Private Sub actualizaDatos(idModulo As Long, idTipo As Long, valorPerm As String, tipo As String)
    Select Case tipo
        Case "ACCESO":
            SQL1 = "UPDATE PERMISOS SET PERM_ACCESO = '" & valorPerm & "' WHERE PERM_MODULO = '" & idModulo & "' AND PERM_TIPO = '" & idTipo & "'"
            'MsgBox SQL1
            con.Execute (SQL1)
        Case "CREACION":
            SQL1 = "UPDATE PERMISOS SET PERM_CREACION = '" & valorPerm & "' WHERE PERM_MODULO = '" & idModulo & "' AND PERM_TIPO = '" & idTipo & "'"
            con.Execute (SQL1)
        Case "MODIFICAR":
            SQL1 = "UPDATE PERMISOS SET PERM_MODIFICACION = '" & valorPerm & "' WHERE PERM_MODULO = '" & idModulo & "' AND PERM_TIPO = '" & idTipo & "'"
            con.Execute (SQL1)
        Case "ELIMINAR":
            SQL1 = "UPDATE PERMISOS SET PERM_ELIMINAR = '" & valorPerm & "' WHERE PERM_MODULO = '" & idModulo & "' AND PERM_TIPO = '" & idTipo & "'"
            con.Execute (SQL1)
    End Select
End Sub

Private Sub mn_Usuarios_Click()
        tipoBusqueda = "U"
        modBusqueda = "Permisos"
        BUSQ_Usuarios.Caption = "Búsqueda de usuarios."
        BUSQ_Usuarios.Show vbModal
End Sub
