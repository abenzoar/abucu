VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form MDIC_OperTickets 
   Caption         =   "Tickets del día"
   ClientHeight    =   8865
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17835
   Icon            =   "MDIC_OperTickets.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8865
   ScaleWidth      =   17835
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbMesa 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   11760
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   960
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox textBus 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   360
      Index           =   1
      Left            =   2160
      TabIndex        =   17
      Text            =   "CLIENTE"
      Top             =   1005
      Width           =   4215
   End
   Begin VB.TextBox textBus 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   330
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Text            =   "FOLIO"
      Top             =   1005
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker dtFecha1 
      Height          =   375
      Left            =   12120
      TabIndex        =   14
      Top             =   480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   192937985
      CurrentDate     =   41358
   End
   Begin VB.Timer tCargaTickets 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   8520
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Horario sucursal"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Index           =   2
      Left            =   14760
      TabIndex        =   13
      Top             =   600
      Value           =   1  'Checked
      Width           =   1600
   End
   Begin VB.Timer timerCobro 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1440
      Top             =   8520
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Ayer"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Index           =   1
      Left            =   16440
      TabIndex        =   12
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Hoy"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Index           =   0
      Left            =   14040
      TabIndex        =   11
      Top             =   600
      Width           =   1335
   End
   Begin VB.OptionButton Opt 
      Caption         =   "Ver cerradas"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   195
      Index           =   2
      Left            =   9480
      TabIndex        =   10
      Top             =   600
      Width           =   1695
   End
   Begin VB.OptionButton Opt 
      Caption         =   "Ver abiertas"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   195
      Index           =   1
      Left            =   7800
      TabIndex        =   9
      Top             =   600
      Width           =   1575
   End
   Begin VB.OptionButton Opt 
      Caption         =   "Ver todas"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   195
      Index           =   0
      Left            =   6480
      TabIndex        =   8
      Top             =   600
      Width           =   1215
   End
   Begin VB.Timer tTime 
      Interval        =   250
      Left            =   2040
      Top             =   8520
   End
   Begin MSFlexGridLib.MSFlexGrid lista 
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   17415
      _ExtentX        =   30718
      _ExtentY        =   12515
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      BackColorFixed  =   9520683
      ForeColorFixed  =   16777215
      BackColorBkg    =   15329769
      GridColor       =   16711680
      AllowUserResizing=   1
      FormatString    =   $"MDIC_OperTickets.frx":058A
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Desde:"
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
      Index           =   4
      Left            =   11400
      TabIndex        =   15
      Top             =   600
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      Index           =   3
      X1              =   4200
      X2              =   6240
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Canceladas: "
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
      Index           =   3
      Left            =   4200
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   2
      Left            =   5400
      TabIndex        =   6
      Top             =   480
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      Index           =   1
      X1              =   2160
      X2              =   3960
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cerradas: "
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
      Index           =   1
      Left            =   2160
      TabIndex        =   5
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Abiertas: "
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
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      Index           =   2
      X1              =   120
      X2              =   1920
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lista de operaciones en el día"
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
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      Index           =   0
      X1              =   120
      X2              =   16200
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "MDIC_OperTickets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql1 As String
Dim res1 As Recordset
Dim tickAb As Long
Dim tickCe As Long
Dim tickCa As Long

Private Sub Check1_Click(Index As Integer)
    cargaTickets
End Sub

Private Sub Command1_Click()
    timerCobro.Enabled = True
End Sub

Private Sub cmbMesa_Click()
    updateOperMesa (cmbMesa.ItemData(cmbMesa.ListIndex))
    
End Sub

Private Sub updateOperMesa(idMesa As Long)
    If idMesa <> 0 Then
        sql1 = "UPDATE VENTAS SET VENT_MESA = '" & idMesa & "' WHERE VENT_IDFOLIO = '" & lista.TextMatrix(lista.Row, 0) & "'"
        'MsgBox idMesa
        con.Execute (sql1)
        
        lista.TextMatrix(lista.Row, 10) = idMesa
        cmbMesa.Visible = False
    End If
End Sub

Private Sub cmbMesa_LostFocus()
    cmbMesa.Visible = False
End Sub

Private Sub dtFecha1_Change()
    'Opt(0).value = True
    cargaTickets
End Sub

Private Sub dtFecha1_Click()
    'Opt(0).value = True
    cargaTickets
'    cargaTickets
End Sub

Private Sub dtFecha1_KeyDown(KeyCode As Integer, Shift As Integer)
    'Opt(0).value = True
    cargaTickets
End Sub

Private Sub Form_Load()

'    If Val(MDI_Operaciones.StatusBar1.Panels(4).Text) >= 1 Then
'        Unload Me
'        Exit Sub
'    End If
    dtFecha1 = Date
    cmbMesa.Visible = False
    If mesas = False Then
        lista.ColWidth(10) = 0
    End If
    Set FrmFocus = Me
    numFrmTicket = numFrmTicket + 1
    Opt(0).value = True
    Check1(2).value = Checked
     
    lista.ColWidth(5) = 0
    lista.ColWidth(6) = 0
    lista.ColWidth(7) = 0
        
    If FRM_Menu.menuBarra2.Panels(14).Text = "A" Then
        lista.ColWidth(9) = 2500
    Else
        lista.ColWidth(9) = 0
    End If
        
    
    MDI_Operaciones.StatusBar1.Panels(4).Text = Val(MDI_Operaciones.StatusBar1.Panels(4).Text) + 1
    
     Me.Caption = "Tickets del día                 " & MDI_Operaciones.StatusBar1.Panels(4).Text
    
'    Opt(2).value = True
    'cargaTickets
End Sub
Public Sub cargaTickets()
Dim tipo As String
    TTime.Enabled = True
    tipo = ""
    If Opt(0).value = True Then
        tipo = " ID_STATUS IN ('G', 'P', 'C', 'A') "
    Else
        If Opt(1).value = True Then
             tipo = " ID_STATUS = 'G' "
        Else
            If Opt(2).value = True Then
                 tipo = " ID_STATUS = 'P' "
            End If
        End If
    End If
    
    'tipo = tipo & " and (date_format(FechaHora, '%x-%m-%d') = date_format(now(),'%x-%m-%d')"
    If Check1(1).value = Checked Then
         tipo = tipo & " and date_format(FechaHora_dOS, '%Y-%m-%d') between  date_format(DATE_SUB(NOW(),  INTERVAL 1 DAY) ,'%Y-%m-%d') AND date_format(now() ,'%Y-%m-%d') "
    Else
        If Check1(0).value = Checked Then
''" & Format(dtFecha1, "yyyy-MM-dd") & "'
            tipo = tipo & " and (date_format(FechaHora_dOS, '%Y-%m-%d') >= date_format(now(),'%Y-%m-%d'))"
'            tipo = tipo & " and (date_format(FechaHora, '%Y-%m-%d') >= '" & Format(dtFecha1, "yyyy-MM-dd") & "'"
        Else
            If Check1(2).value = Checked Then
                If FRM_Menu.menuBarra2.Panels(13).Text = "M" Then
'                    tipo = tipo & " and (date_format(FechaHora, '%Y-%m-%d') >= date_format(now(),'%Y-%m-%d'))"
                    tipo = tipo & " and (date_format(FechaHora_DOS, '%Y-%m-%d') >= '" & Format(dtFecha1, "yyyy-MM-dd") & "') "
                Else
                    If FRM_Menu.menuBarra2.Panels(13).Text = "D" Then
                         'tipo = tipo & " and date_format(FechaHora, '%x-%m-%d %T') between CONCAT((DATE_FORMAT(NOW(), '%x-%m-%d')), ' ', '" & FRM_Menu.menuBarra2.Panels(11).Text & "' )  AND  CONCAT((DATE_FORMAT(DATE_add(NOW(), INTERVAL 1 DAY), '%x-%m-%d')), ' ', '" & FRM_Menu.menuBarra2.Panels(12).Text & "')"
                         'tipo = tipo & " and FechaHora between CONCAT((DATE_FORMAT(NOW(), '%x-%m-%d')), ' ', '" & FRM_Menu.menuBarra2.Panels(11).Text & "' )  AND  CONCAT((DATE_FORMAT(DATE_add(NOW(), INTERVAL 1 DAY), '%x-%m-%d')), ' ', '" & FRM_Menu.menuBarra2.Panels(12).Text & "')"
                        'tipo = tipo & "AND T2.FECHAHORA BETWEEN CONCAT((DATE_FORMAT(NOW(), '%x-%m-%d')), ' ', T1.SUC_HORAENTRADA) AND  CONCAT((DATE_FORMAT(DATE_add(NOW(), INTERVAL 1 DAY), '%x-%m-%d')), ' ', T1.SUC_HORASALIDA)"
                        
                        'MsgBox Format(Time, "Long Time") & " > " & FRM_Menu.menuBarra2.Panels(11).Text
                        'MsgBox Format(Time, "Short Time") & " > " & Format(FRM_Menu.menuBarra2.Panels(11).Text, "Short Time")
                        If Format(Time, "Short Time") > Format(FRM_Menu.menuBarra2.Panels(11).Text, "Short Time") Then
'                            tipo = tipo & " AND T2.FECHAHORA_dos BETWEEN CONCAT((DATE_FORMAT(NOW(), '%Y-%m-%d')), ' ', T1.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(DATE_ADD(NOW(), INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T1.SUC_HORASALIDA) "
                            tipo = tipo & " AND T2.FECHAHORA_dos BETWEEN CONCAT(('" & Format(dtFecha1, "yyyy-MM-dd") & "'), ' ', T1.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(DATE_ADD(NOW(), INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T1.SUC_HORASALIDA) "
                        Else
'                            tipo = tipo & " AND T2.FECHAHORA_dos BETWEEN CONCAT((DATE_FORMAT(DATE_SUB(NOW(), INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T1.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(NOW(), '%Y-%m-%d')), ' ', T1.SUC_HORASALIDA)"
                            tipo = tipo & " AND T2.FECHAHORA_dos BETWEEN CONCAT((DATE_FORMAT(DATE_SUB(('" & Format(dtFecha1, "yyyy-MM-dd") & "'), INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T1.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(NOW(), '%Y-%m-%d')), ' ', T1.SUC_HORASALIDA)"
                        End If
                                
'                        If Format(Time, "Short Time") > Format(FRM_Menu.menuBarra2.Panels(11).Text, "Short Time") Then
'                            tipo = tipo & " AND vent_fechaHora_cobro BETWEEN CONCAT(('" & Format(dtFecha1(0), "yyyy-MM-dd") & "'), ' ', T5.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT(DATE_ADD('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T5.SUC_HORASALIDA) "
'                        Else
'                            tipo = tipo & " AND vent_fechaHora_cobro BETWEEN CONCAT((DATE_FORMAT(DATE_SUB('" & Format(dtFecha1(0), "yyyy-MM-dd") & "', INTERVAL 1 DAY), '%Y-%m-%d')), ' ', T5.SUC_HORAENTRADA) AND CONCAT((DATE_FORMAT('" & Format(dtFecha1(1), "yyyy-MM-dd") & "', '%Y-%m-%d')), ' ', T5.SUC_HORASALIDA)"
'                        End If
                        
                        
                    End If
                End If
                
            End If
        End If
    End If
    'MsgBox tipo
    Dim busFolio, busCliente As String
    
    If textBus(0).Text = "FOLIO" Then
        busFolio = ""
    Else
        busFolio = textBus(0).Text
    End If
    If textBus(1).Text = "CLIENTE" Then
        busCliente = ""
    Else
        busCliente = textBus(1).Text
    End If
    
    tipo = tipo & " AND FOLIO LIKE '%" & busFolio & "%' " & _
    "AND upper(CLIENTE) LIKE upper('%" & busCliente & "%') "
    
    tipo = tipo & " ORDER BY FECHAHORA DESC"
    
    sql1 = "SELECT * FROM SUCURSAL T1, VIEW_VENTAS T2 WHERE " & tipo & " "
    'MsgBox SQL1
'    Text2.Text = SQL1
    Set res1 = con.Execute(sql1)
  
    'MsgBox SQL1
    
    tickAb = 0
    tickCe = 0
    tickCa = 0
    
    lista.Rows = 1
    lista.Redraw = False
    Do While Not res1.EOF
      lista.AddItem ""
      lista.TextMatrix(lista.Rows - 1, 0) = res1.Fields("FOLIO")
      lista.TextMatrix(lista.Rows - 1, 1) = res1.Fields("FECHAHORA_DOS")
      lista.TextMatrix(lista.Rows - 1, 2) = res1.Fields("USUARIO")
      lista.TextMatrix(lista.Rows - 1, 3) = res1.Fields("CLIENTE")
      lista.TextMatrix(lista.Rows - 1, 4) = res1.Fields("STATUS")
      lista.TextMatrix(lista.Rows - 1, 5) = FormatCurrency(res1.Fields("SUB_TOTAL"))
      lista.TextMatrix(lista.Rows - 1, 6) = FormatCurrency(res1.Fields("DESCUENTO"))
      lista.TextMatrix(lista.Rows - 1, 7) = FormatCurrency(res1.Fields("TOTAL"))
      lista.TextMatrix(lista.Rows - 1, 8) = res1.Fields("OPERACIONES")
      lista.TextMatrix(lista.Rows - 1, 9) = res1.Fields("SUB_STATUS")
      lista.TextMatrix(lista.Rows - 1, 10) = res1.Fields("MESA") & ""
      
      
      
      If res1.Fields("ID_STATUS") = "G" Then
          tickAb = tickAb + 1
          lista.Row = lista.Rows - 1
          lista.Col = 4
          lista.CellForeColor = &H40C0&
          lista.Col = 0
          lista.CellForeColor = &H40C0&
      Else
        If res1.Fields("ID_STATUS") = "C" Then
            tickCa = tickCa + 1
            lista.Row = lista.Rows - 1
            lista.Col = 4
            lista.CellForeColor = vbRed
            lista.Col = 0
            lista.CellForeColor = vbRed
        Else
            If res1.Fields("ID_STATUS") = "P" Then
                tickCe = tickCe + 1
            End If
        End If
      End If
      
      res1.MoveNext
    Loop
    lista.Redraw = True

    lInfo(0).Caption = tickAb
    lInfo(1).Caption = tickCe
    lInfo(2).Caption = tickCa
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    numFrmTicket = numFrmTicket - 1
    MDI_Operaciones.StatusBar1.Panels(4).Text = Val(MDI_Operaciones.StatusBar1.Panels(4).Text) - 1
End Sub

Private Sub Lista_Click()
''''asdasdasd
    If lista.Rows > 1 Then
        lista.Row = lista.MouseRow
        Set FrmFocus = Me
    End If
End Sub

Private Sub Lista_DblClick()
    If lista.Row > 0 Then
        
        If lista.Col = 10 Then
            If mesas = True Then
                cargaMesa
                cmbMesa.Top = lista.CellTop + lista.Top
                cmbMesa.Left = lista.CellLeft + lista.Left
                cmbMesa.width = lista.CellWidth
                If lista.TextMatrix(lista.Row, lista.Col) <> "" Then
                    cmbMesa.AddItem lista.TextMatrix(lista.Row, lista.Col)
                    cmbMesa.Text = lista.TextMatrix(lista.Row, lista.Col)
                End If
                cmbMesa.Visible = True
                cmbMesa.SetFocus
            End If
        Else
            If lista.TextMatrix(lista.Row, 4) <> "APARTADOS" Then
                operacionTicket
            Else
                MsgBox "Un apartado no puede verse en operaciones. Verifique en el módulo de apartados.", vbExclamation
            End If
        End If
    End If
End Sub

Private Sub cargaMesa()
    
    sql1 = "SELECT * FROM VIEW_MESAS_ESTADO WHERE ESTADO = 'DISPONIBLE' ORDER BY MESA_ID"
    Set res1 = con.Execute(sql1)
    
    cmbMesa.Clear
    Do While Not res1.EOF
        cmbMesa.AddItem res1.Fields("MESA_ID")
        cmbMesa.ItemData(cmbMesa.ListCount - 1) = res1.Fields("mesa_id")
        res1.MoveNext
    Loop

End Sub


Public Sub operacionTicket()
        Set FrmOper = New MDIC_Operaciones
        If nForms <= 0 Then
            nForms = 1
        Else
            nForms = nForms + 1
        End If
        tikcet = True
        folioTicket = lista.TextMatrix(lista.Row, 0)
        FrmOper.Caption = sCaption & nForms
        FrmOper.lInfo(2).Caption = lista.TextMatrix(lista.Row, 4)
        MDI_Operaciones.WindowState = vbMaximized
        MDI_Operaciones.Show
    
    If lista.TextMatrix(lista.Row, 4) = "CERRADO" Then
        FrmOper.txtClave(0).Enabled = False
        FrmOper.txtClave(1).Enabled = False
        FrmOper.txtClave(2).Enabled = False
        FrmOper.txtDesc(0).Enabled = False
        FrmOper.txtDesc(1).Enabled = False
        
        MsgBox "La operación esta cerrada, solo podrá verificar información y actualizar datos de los usuarios.", vbInformation
    Else
        If lista.TextMatrix(lista.Row, 4) = "CANCELADO" Then
            MsgBox "La operación " & lista.TextMatrix(lista.Row, 0) & " se ha cancelado y no puede utilizarse. Verfique.", vbInformation
        End If
    End If

End Sub
Private Sub lista_GotFocus()
    Set FrmFocus = Me
    ConScroll lista
End Sub

Private Sub lista_LostFocus()
    SinScroll lista
End Sub

Private Sub Lista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lista.Rows > 1 Then
        Lista_Click
        If Button = vbRightButton Then
            MDI_Operaciones.mn_PrintTicket.Caption = "Imprimir ticket folio: " & lista.TextMatrix(lista.Row, 0) & " Cliente: " & lista.TextMatrix(lista.Row, 3)
            MDI_Operaciones.mn_PrintPreTicket.Caption = "Imprimir pre-ticket folio: " & lista.TextMatrix(lista.Row, 0) & " Cliente: " & lista.TextMatrix(lista.Row, 3)
            MDI_Operaciones.mn_CancelOperTicket.Caption = "Cancelar operación folio: " & lista.TextMatrix(lista.Row, 0) & " Cliente: " & lista.TextMatrix(lista.Row, 3)
            
            If lista.TextMatrix(lista.Row, 4) = "CERRADO" Then
                MDI_Operaciones.mn_PrintTicket.Enabled = True
                MDI_Operaciones.mn_CancelOperTicket.Enabled = True
                MDI_Operaciones.mn_PrintPreTicket.Visible = False
            Else
                
                MDI_Operaciones.mn_PrintTicket.Enabled = False
                MDI_Operaciones.mn_CancelOperTicket.Enabled = True
                MDI_Operaciones.mn_PrintPreTicket.Visible = True
            End If
            PopupMenu MDI_Operaciones.mn_TicketsPrint, vbPopupMenuLeftAlign

        End If
    End If

End Sub


Private Sub Opt_Click(Index As Integer)
    cargaTickets
End Sub

Private Sub tCargaTickets_Timer()
     tCargaTickets.Enabled = False
    cargaTickets
End Sub

Private Sub textBus_GotFocus(Index As Integer)
    If Index = 0 Then
        If textBus(Index).Text = "FOLIO" Then
            textBus(Index).Text = ""
        End If
    Else
        If Index = 1 Then
            If textBus(Index).Text = "CLIENTE" Then
                textBus(Index).Text = ""
            End If
        End If
    End If

End Sub

Private Sub textBus_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        cargaTickets
    End If
End Sub

Private Sub textBus_LostFocus(Index As Integer)
    If Index = 0 Then
        If textBus(Index).Text = "" Then
            textBus(Index).Text = "FOLIO"
        End If
    Else
        If Index = 1 Then
            If textBus(Index).Text = "" Then
                textBus(Index).Text = "CLIENTE"
            End If
        End If
    End If
End Sub

Private Sub timerCobro_Timer()
    timerCobro.Enabled = False
    operacionTicket
    
End Sub

Private Sub TTime_Timer()
    TTime.Enabled = False
    lista.width = Me.width - 500
    lista.height = Me.height - 2100
End Sub
