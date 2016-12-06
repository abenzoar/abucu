VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRM_Membresias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Membresias"
   ClientHeight    =   10065
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10065
   ScaleWidth      =   16905
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   9975
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   16935
      _ExtentX        =   29871
      _ExtentY        =   17595
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "FRM_Membresias.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Borde(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lBus(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Borde(17)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Borde(16)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lBus(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lBus(4)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Borde(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lBus(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lBus(5)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Borde(3)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Lista"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "textBus(1)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmbStatus(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "textBus(0)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdAccion(7)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "time_size"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "textBus(2)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "textBus(3)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      Begin VB.TextBox textBus 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   10680
         TabIndex        =   12
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox textBus 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   8760
         TabIndex        =   10
         Top             =   480
         Width           =   1695
      End
      Begin VB.Timer time_size 
         Interval        =   500
         Left            =   11880
         Top             =   840
      End
      Begin VB.CommandButton cmdAccion 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exportar detalle"
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
         Index           =   7
         Left            =   14760
         Picture         =   "FRM_Membresias.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox textBus 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   3015
      End
      Begin VB.ComboBox cmbStatus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Selecciona el tipo de clasificación a la que pertenece el producto, o agrega o edita los existentes"
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox textBus 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   6600
         TabIndex        =   2
         Top             =   480
         Width           =   1935
      End
      Begin MSFlexGridLib.MSFlexGrid Lista 
         Height          =   8655
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   16575
         _ExtentX        =   29236
         _ExtentY        =   15266
         _Version        =   393216
         Cols            =   14
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"FRM_Membresias.frx":05A6
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
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   3
         Left            =   10680
         Top             =   480
         Width           =   1725
      End
      Begin VB.Label lBus 
         BackStyle       =   0  'Transparent
         Caption         =   "Folio venta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   10680
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lBus 
         BackStyle       =   0  'Transparent
         Caption         =   "Folio venta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   8760
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   2
         Left            =   8760
         Top             =   480
         Width           =   1725
      End
      Begin VB.Label lBus 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   11400
         TabIndex        =   9
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label lBus 
         BackStyle       =   0  'Transparent
         Caption         =   "Días para vencimiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   6600
         TabIndex        =   8
         Top             =   240
         Width           =   1935
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   16
         Left            =   240
         Top             =   480
         Width           =   3045
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   17
         Left            =   3480
         Top             =   480
         Width           =   2925
      End
      Begin VB.Label lBus 
         BackStyle       =   0  'Transparent
         Caption         =   "Estatus"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3480
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
      Begin VB.Shape Borde 
         BorderColor     =   &H000080FF&
         BorderWidth     =   4
         Height          =   435
         Index           =   0
         Left            =   6600
         Top             =   480
         Width           =   1965
      End
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H000080FF&
      BorderWidth     =   4
      Height          =   435
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   1965
   End
   Begin VB.Label lBus 
      BackStyle       =   0  'Transparent
      Caption         =   "Marca"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "FRM_Membresias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RES1 As Recordset
Dim sql1 As String

Private Sub cmbStatus_Click(Index As Integer)
    cargaLista
End Sub

Private Sub cmdAccion_Click(Index As Integer)
    ques = MsgBox("¿Exportar la lista a excel? ", vbYesNo + vbQuestion)
    If ques = vbYes Then
        Call exportExcel(lista)
    End If

End Sub

Private Sub Form_Load()
        
    cmbStatus(0).Clear
    cmbStatus(0).AddItem "TODOS"
    cmbStatus(0).AddItem "VIGENTE"
    cmbStatus(0).AddItem "VENCIDO"
    cmbStatus(0).AddItem "ANTICIPADA"
    cmbStatus(0).AddItem "VENCIDO SIN RENOVACION"
    SSTab1.TabCaption(0) = ""
    
    cargaLista
End Sub

Private Sub lista_DblClick()

    Call ordenarLista(lista)

End Sub

Private Sub lista_GotFocus()
    ConScroll lista

End Sub

Private Sub lista_LostFocus()
    SinScroll lista
End Sub

Private Sub textBus_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        cargaLista
    End If
End Sub

Private Sub time_size_Timer()
    time_size.Enabled = False
    SSTab1.width = Me.width - 50
    SSTab1.height = Me.height
    lista.width = Me.width - 500

End Sub


Private Sub cargaLista()
    Dim texto1 As String
    texto1 = ""
    

    
    If textBus(1).Text <> "" Then
        texto1 = texto1 & " AND DIAS_RESTANTES <= " & Val(textBus(1).Text) & " AND DIAS_RESTANTES > 0 "
    End If
    
    If textBus(2).Text <> "" Then
        texto1 = texto1 & " AND FOLIO LIKE upper('%" & textBus(2).Text & "%') "
    End If
        
        
    If cmbStatus(0).Text = "VENCIDO SIN RENOVACION" Then
        texto1 = texto1 & "AND upper(STATUS) LIKE upper('%VENCIDO%')  AND PER_ID NOT IN (SELECT PER_ID FROM VIEW_MEMBRESIAS_ASIGNADAS WHERE STATUS = 'VIGENTE')"
    Else
        If cmbStatus(0).Text <> "TODOS" Then
            texto1 = texto1 & "AND upper(STATUS) LIKE upper('%" & cmbStatus(0).Text & "%') "
        End If
    End If
        
    texto1 = texto1 & " order by ADQUIRIO DESC "
        
    sql1 = "SELECT * fROM VIEW_MEMBRESIAS_asignadas WHERE " & _
    "UPPER(CLIENTE) LIKE UPPER('%" & textBus(0).Text & "%') " & _
    texto1
    
    Set RES1 = con.Execute(sql1)
    
    lista.Redraw = False
    lista.Rows = 1
    Do While Not RES1.EOF
        lista.AddItem ""
        lista.TextMatrix(lista.Rows - 1, 0) = RES1.Fields("PER_ID")
        lista.TextMatrix(lista.Rows - 1, 1) = RES1.Fields("CLIENTE")
        lista.TextMatrix(lista.Rows - 1, 2) = RES1.Fields("MEMBRESIA")
        lista.TextMatrix(lista.Rows - 1, 3) = RES1.Fields("DIAS_MEM")
        lista.TextMatrix(lista.Rows - 1, 4) = RES1.Fields("INICIO")
        lista.TextMatrix(lista.Rows - 1, 5) = RES1.Fields("FIN")
        lista.TextMatrix(lista.Rows - 1, 6) = RES1.Fields("ADQUIRIO")
        lista.TextMatrix(lista.Rows - 1, 7) = RES1.Fields("FOLIO")
        lista.TextMatrix(lista.Rows - 1, 8) = RES1.Fields("DIAS_RESTANTES")
        lista.TextMatrix(lista.Rows - 1, 9) = RES1.Fields("STATUS")
        lista.TextMatrix(lista.Rows - 1, 10) = RES1.Fields("PER_EMAIL") & ""
        lista.TextMatrix(lista.Rows - 1, 11) = RES1.Fields("PER_TEL1") & ""
        lista.TextMatrix(lista.Rows - 1, 12) = RES1.Fields("PER_TEL2") & ""
        lista.TextMatrix(lista.Rows - 1, 13) = RES1.Fields("recomendado_por") & ""
        RES1.MoveNext
    Loop
    lista.Redraw = True
    lBus(4).Caption = "Resultados: " & lista.Rows - 1

End Sub
