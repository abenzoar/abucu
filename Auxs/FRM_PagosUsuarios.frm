VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FRM_PagosUsuarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pagos a usuarios"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   14925
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid lista1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5318
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      AllowUserResizing=   1
      FormatString    =   " Usuario      | Nombre     | Apellido Paterno | Apellido Materno  | Tipo usuario      "
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
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3015
      Left            =   8520
      TabIndex        =   2
      Top             =   840
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5318
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      AllowUserResizing=   1
      FormatString    =   "Periodo inicio | Periodo fin  | Monto       | Estatus       "
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
   Begin VB.Label lTitulo 
      BackStyle       =   0  'Transparent
      Caption         =   "Lista de usuarios con asignación de pagos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   7455
   End
End
Attribute VB_Name = "FRM_PagosUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL1 As String
Dim RES1 As Recordset
Dim RES2 As Recordset
Private Sub Form_Load()
    cargaListaUsers
End Sub
Private Sub cargaListaUsers()
    
    SQL1 = "SELECT T2.PERTP_PER_ID, T1.PER_NOMBRE, T1.PER_PATERNO, T1.PER_MATERNO, T2.PERTP_USUARIO, T5.CTPT_TIPO " & _
    "FROM PERSONA T1, PER_TIPO T2, CAT_PAGOS T3, COMISIONES T4, CAT_TIPO T5 " & _
    "WHERE T1.PER_ID = T2.PERTP_PER_ID AND T4.PG_PERTP_PER_ID = T2.PERTP_PER_ID AND T2.PERTP_TIPO_ID = T5.CTPT_ID AND T4.PG_STATUS = 'A' AND " & _
    "T4.PG_PERTP_TIPO_ID = T2.PERTP_TIPO_ID AND T4.PG_PERTP_PER_ID = T2.PERTP_PER_ID AND T4.PG_PERTP_PER_TIPO = T2.PERTP_PER_TIPO " & _
    "GROUP BY T2.PERTP_PER_ID, T1.PER_NOMBRE, T1.PER_PATERNO, T1.PER_MATERNO, T2.PERTP_USUARIO, T5.CTPT_TIPO  "
    Set RES1 = con.Execute(SQL1)
    lista1.Rows = 1
    
    Do While Not RES1.EOF
        lista1.AddItem ""
        lista1.TextMatrix(lista1.Rows - 1, 0) = RES1.Fields("PERTP_PER_ID")
        lista1.TextMatrix(lista1.Rows - 1, 1) = RES1.Fields("PERTP_USUARIO")
        lista1.TextMatrix(lista1.Rows - 1, 2) = RES1.Fields("PER_NOMBRE")
        lista1.TextMatrix(lista1.Rows - 1, 3) = RES1.Fields("PER_PATERNO")
        lista1.TextMatrix(lista1.Rows - 1, 4) = RES1.Fields("PER_MATERNO")
        lista1.TextMatrix(lista1.Rows - 1, 5) = RES1.Fields("CTPTP_TIPO")
        RES1.MoveNext
    Loop

End Sub

Private Sub lista1_Click()
    cargaPagos (lista1.Row)
End Sub
Private Sub cargaPagos(fila As Long)
    SQL1 = ""
End Sub
