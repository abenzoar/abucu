VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form BUSQ_Usuarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Búsqueda de personas"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14550
   Icon            =   "BUSQ_Usuarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   14550
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   10
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Editar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9000
      Picture         =   "BUSQ_Usuarios.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
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
      Height          =   735
      Left            =   7560
      Picture         =   "BUSQ_Usuarios.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10440
      TabIndex        =   3
      Text            =   "30"
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   3855
   End
   Begin MSFlexGridLib.MSFlexGrid lista 
      Height          =   4935
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   8705
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      AllowUserResizing=   1
      FormatString    =   $"BUSQ_Usuarios.frx":109E
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Telefono(s)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   11
      Top             =   360
      Width           =   3615
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H0000C000&
      BorderWidth     =   4
      Height          =   435
      Index           =   1
      Left            =   4320
      Top             =   720
      Width           =   2685
   End
   Begin VB.Image imgFoto 
      BorderStyle     =   1  'Fixed Single
      Height          =   2295
      Index           =   0
      Left            =   11640
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Imagen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   11640
      TabIndex        =   8
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Observaciones: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   3
      Left            =   11280
      TabIndex        =   7
      Top             =   4320
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Registros en la lista: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   6360
      Width           =   4455
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H0000C000&
      BorderWidth     =   4
      Height          =   435
      Index           =   0
      Left            =   240
      Top             =   720
      Width           =   3885
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Núm registros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10440
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Texto a buscar (Nombre(s) o apellidos)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin VB.Image Image2 
      Height          =   9855
      Index           =   1
      Left            =   -240
      Picture         =   "BUSQ_Usuarios.frx":1181
      Stretch         =   -1  'True
      Top             =   0
      Width           =   17655
   End
End
Attribute VB_Name = "BUSQ_Usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    tipoPersona = "CLIENTE"
    ADD_Cliente.txtUsuario(0).Text = UCase(Text1.Text)
    ADD_Cliente.Show vbModal
End Sub

Private Sub Command2_Click()
    If lista.TextMatrix(lista.Row, 4) <> "" And lista.Rows > 1 Then
        tipoPersona = "CLIENTE_EDIT"
        'ADD_Cliente.txtUsuario(0).Text = UCase(Text1.Text)
        ADD_Cliente.Show vbModal
    Else
        MsgBox "No se puede realizar la acción. Verifique.", vbInformation
    End If
    
End Sub

Private Sub Form_Load()
    lista.Rows = 1
    lista.ColWidth(4) = 0
    lista.ColWidth(5) = 0
    
    If tipoBusqueda = "C" Then
        Command1.Visible = True
    Else
        Command1.Visible = False
    End If
    buscarUsuario
End Sub

Private Sub Lista_Click()
    Label1(3).Caption = "Descripción: " & lista.TextMatrix(lista.Row, 6)
    sql1 = "SELECT PER_FOTO FROM PERSONA WHERE PER_ID = '" & lista.TextMatrix(lista.Row, 4) & "'"
    Set res1 = con.Execute(sql1)
    
    If IsNull(res1.Fields("PER_fOTO")) = False Then
        Dim Imagen1 As Stream
        Set Imagen1 = New Stream
        Imagen1.Type = adTypeBinary
        checarCarpetaTemp
        Imagen1.Open
        Imagen1.Write res1.Fields("PER_FOTO")
        Imagen1.SaveToFile direccionSistema & "\Temp\TempUser.dat", adSaveCreateOverWrite
        Imagen1.Close
        imgFoto(0).Picture = LoadPicture(direccionSistema & "\Temp\TempUser.dat")
    Else
        imgFoto(0).Picture = LoadPicture("")
    End If

End Sub
Private Sub lista_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Lista_DblClick
    End If
End Sub
Private Sub Lista_DblClick()
    If lista.TextMatrix(lista.Row, 1) <> "" Then
        ''''Para asistencias
        If modBusqueda = "Asistencia" Then
            FRM_Asistencias.txtUsuario(0).Text = lista.TextMatrix(lista.Row, 1)
            Unload Me
            FRM_Asistencias.cmdCheck_Click
        Else
        '''Para operaciones
            If modBusqueda = "Operaciones" And tipoBusqueda = "C" Then
                FrmFocus.txtClave(2).Text = lista.TextMatrix(lista.Row, 1)
                Unload Me
                FrmFocus.cmdOperCheck_Click (2)
            Else
                If modBusqueda = "Operaciones" And tipoBusqueda = "U" Then
                    FrmFocus.txtClave(1).Text = lista.TextMatrix(lista.Row, 1)
                    Unload Me
                    FrmFocus.cmdOperCheck_Click (1)
                Else
                    If modBusqueda = "ConsumoInterno" And tipoBusqueda = "U" Then
                        FRM_ConsumoInterno.lblUserId(3).Caption = lista.TextMatrix(lista.Row, 4)
                        FRM_ConsumoInterno.lblUserId(4).Caption = lista.TextMatrix(lista.Row, 5)
                        FRM_ConsumoInterno.lblUserId(5).Caption = "U"
                        FRM_ConsumoInterno.lblDatos(0).Caption = lista.TextMatrix(lista.Row, 2)
                        Call FRM_ConsumoInterno.cargaFotoMostrador("U", 1)
                        Unload Me
                    Else
                        If modBusqueda = "Apartado" Then
                            FRM_Apartados.txtClave(2).Text = lista.TextMatrix(lista.Row, 1)
                            FRM_Apartados.aprt_checkCliente
                            Unload Me
                        Else
                            If modBusqueda = "Permisos" Then
                                FRM_Permisos.txtClave(2).Text = lista.TextMatrix(lista.Row, 1)
                                Unload Me
                                'FrmFocus.cmdOperCheck_Click (2)
                                
                            End If

                        End If
                    End If
                End If
            End If
        End If
        
        
    Else
        MsgBox "La persona no tiene código de membresía. Verifique. ", vbInformation
    End If
End Sub

Private Sub lista_GotFocus()
    ConScroll lista
End Sub


Private Sub lista_LostFocus()
    SinScroll lista
End Sub

Private Sub lista_SelChange()
    Lista_Click
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        buscarUsuario
    Else
        If KeyAscii = 27 Then
            Unload Me
        End If
    End If
End Sub
Public Sub buscarUsuario()
    Dim sql1 As String
    Dim res1 As Recordset
    Dim texto As String
    
    texto = ""
    
    If modBusqueda = "Asistencia" Then
        texto = " AND T4.PERTP_PER_TIPO IN ('C', 'U') "
    Else
        texto = " AND T4.PERTP_PER_TIPO = '" & tipoBusqueda & "'  "
    End If
    
    If Text3.Text <> "" Then
       texto = texto & " AND T2.PER_TEL1 LIKE '%" & Text3.Text & "%' "
    End If

    sql1 = "SELECT T4.PERTP_CODIGO_MEMBRESIA, if(PERTP_PER_TIPO= 'C', 'Cliente', 'USUARIO') TIPO, " & _
    "T2.PER_NOMBRE, T2.PER_PATERNO, T2.PER_MATERNO, T2.PER_ID, T2.PER_TEL1 TEL,  T4.PERTP_TIPO_ID, T2.PER_DESCRIPCION, IF(T4.PERTP_MEMBRESIA = 'S', 'SI', 'NO') MEMBRESIA " & _
    "FROM PERSONA T2, CAT_TIPO T3, PER_tIPO T4 " & _
    "WHERE T4.PERTP_TIPO_ID = T3.CTPT_ID AND T4.PERTP_PER_TIPO = T3.CTPT_SUBTIPO AND T2.PER_ID = T4.PERTP_PER_ID " & _
    "AND concat(T2.PER_NOMBRE, ' ', T2.PER_PATERNO, ' ', T2.PER_MATERNO) LIKE '%" & Text1.Text & "%' " & _
    "AND T4.PERTP_STATUS = 'A' " & texto & _
    "ORDER BY T4.PERTP_PERALTA_FECHA DESC"

    '"Limit 0, " & Val(Text2.Text) & " ORDER BY T4.PERTP_ALTA DESC"
    'MsgBox sql1
    lista.Redraw = False
    Set res1 = con.Execute(sql1)
        lista.Rows = 1
    Do While Not res1.EOF
        lista.AddItem ""
        lista.TextMatrix(lista.Rows - 1, 0) = lista.Rows - 1
        lista.TextMatrix(lista.Rows - 1, 1) = "" & res1.Fields("pertp_codigo_membresia")
        lista.TextMatrix(lista.Rows - 1, 2) = res1.Fields("PER_NOMBRE") & " " & res1.Fields("PER_PATERNO") & " " & res1.Fields("PER_MATERNO")
        lista.TextMatrix(lista.Rows - 1, 3) = res1.Fields("TIPO")
        lista.TextMatrix(lista.Rows - 1, 4) = res1.Fields("PER_ID")
        lista.TextMatrix(lista.Rows - 1, 5) = res1.Fields("PERTP_TIPO_ID")
        lista.TextMatrix(lista.Rows - 1, 6) = res1.Fields("MEMBRESIA")
        lista.TextMatrix(lista.Rows - 1, 7) = res1.Fields("TEL") & ""
        lista.TextMatrix(lista.Rows - 1, 8) = res1.Fields("PER_DESCRIPCION") & ""
        res1.MoveNext
    Loop
    lista.Redraw = True

    Label3.Caption = "Registros en la lista: " & lista.Rows - 1
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    buscarUsuario
Else
    Call Numeros(KeyAscii)
    
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        buscarUsuario
    Else
        If KeyAscii = 27 Then
            Unload Me
        End If
    End If

End Sub
