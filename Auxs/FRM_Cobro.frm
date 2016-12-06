VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FRM_Cobro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cobro"
   ClientHeight    =   9525
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6690
   ClipControls    =   0   'False
   Icon            =   "FRM_Cobro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9525
   ScaleWidth      =   6690
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer timeSalir 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2160
      Top             =   120
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Salir (Esc)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   4
      Left            =   4080
      Picture         =   "FRM_Cobro.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "imprimir ticket"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   3
      Left            =   120
      Picture         =   "FRM_Cobro.frx":0E54
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2280
      Width           =   3615
   End
   Begin VB.TextBox txtCambio2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1605
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Text            =   "$0.0"
      Top             =   360
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.TextBox txtPrintCopias 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      TabIndex        =   22
      Text            =   "2"
      Top             =   8760
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8760
      Width           =   375
   End
   Begin VB.TextBox txtPrint 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   8760
      Width           =   4215
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   2
      Left            =   4560
      Picture         =   "FRM_Cobro.frx":171E
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   1
      Left            =   2400
      Picture         =   "FRM_Cobro.frx":1FE8
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Aceptar e imprimir ticket"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   0
      Left            =   240
      Picture         =   "FRM_Cobro.frx":28B2
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6960
      Width           =   1695
   End
   Begin VB.TextBox txtPago 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   840
      Index           =   4
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4080
      Width           =   5775
   End
   Begin VB.TextBox txtCambio 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1125
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "$0.0"
      Top             =   5520
      Width           =   6015
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1815
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   3201
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   706
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   " Efectivo"
      TabPicture(0)   =   "FRM_Cobro.frx":317C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "label0(5)"
      Tab(0).Control(1)=   "txtPago(0)"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   " Tarjeta"
      TabPicture(1)   =   "FRM_Cobro.frx":3716
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtPago(1)"
      Tab(1).Control(1)=   "label0(6)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Monedero"
      TabPicture(2)   =   "FRM_Cobro.frx":3CB0
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "label0(7)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "txtPago(2)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Otro"
      TabPicture(3)   =   "FRM_Cobro.frx":424A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtPago(3)"
      Tab(3).Control(1)=   "label0(8)"
      Tab(3).ControlCount=   2
      Begin VB.TextBox txtPago 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Index           =   3
         Left            =   -74880
         TabIndex        =   8
         Top             =   480
         Width           =   5775
      End
      Begin VB.TextBox txtPago 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   5775
      End
      Begin VB.TextBox txtPago 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Index           =   1
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Width           =   5775
      End
      Begin VB.TextBox txtPago 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Index           =   0
         Left            =   -74880
         TabIndex        =   0
         Top             =   480
         Width           =   5775
      End
      Begin VB.Label label0 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Otro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Index           =   8
         Left            =   -72480
         TabIndex        =   21
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Label label0 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Monedero"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Index           =   7
         Left            =   2520
         TabIndex        =   20
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Label label0 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tarjeta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Index           =   6
         Left            =   -72480
         TabIndex        =   19
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Label label0 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Efectivo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Index           =   5
         Left            =   -72480
         TabIndex        =   18
         Top             =   1440
         Width           =   3375
      End
   End
   Begin VB.TextBox txtTot 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "$0.0"
      Top             =   480
      Width           =   6015
   End
   Begin VB.Label label0 
      BackStyle       =   0  'Transparent
      Caption         =   "Copias"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   5760
      TabIndex        =   23
      Top             =   8520
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   4680
      Picture         =   "FRM_Cobro.frx":4266
      Top             =   8880
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   4680
      Picture         =   "FRM_Cobro.frx":4B30
      Top             =   8880
      Width           =   480
   End
   Begin VB.Label label0 
      BackStyle       =   0  'Transparent
      Caption         =   "Estatus"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   4560
      TabIndex        =   16
      Top             =   8520
      Width           =   855
   End
   Begin VB.Label label0 
      BackStyle       =   0  'Transparent
      Caption         =   "Impresora actual"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   15
      Top             =   8520
      Width           =   2175
   End
   Begin VB.Label label0 
      BackStyle       =   0  'Transparent
      Caption         =   "Pago"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   10
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label label0 
      BackStyle       =   0  'Transparent
      Caption         =   "Cambio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   5160
      Width           =   2175
   End
   Begin VB.Label label0 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "FRM_Cobro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim sql1 As String
Dim RES1 As Recordset
Dim RES2 As Recordset
Dim resEnvioManil As Recordset
Dim sqlEnvio As String
Dim tiempo As Integer
Dim ResDependiente As Recordset

Private Sub checkMonedero()
'    For b1 = 1 To FrmFocus.lista.Rows - 1
'        FrmFocus.lista.TextMatrix(FrmFocus.lista.Rows - 1, 15) = "MND"
'    Next b1
'
    Call FrmFocus.addMonedero(0, Val(Format(txtPago(2).Text, "General Number")))
End Sub
Private Sub cmdOpcion_Click(Index As Integer)
    If Index = 2 Then
        Unload Me
    Else
        If Index = 3 Then
            If tipoCobro = "OPERACIONES_TOUCH" Then
                nota (FRM_OperTouch.lista_detalle.TextMatrix(1, 1))
                MsgBox "Ticket impreso." & vbCrLf & vbCrLf & "Verifique.", vbInformation
            Else
                If tipoCobro = "OPERACIONES" Then
                    nota (FrmFocus.lInfo(1).Caption)
                    MsgBox "Ticket impreso." & vbCrLf & vbCrLf & "Verifique.", vbInformation
                Else
                    MsgBox "Opción no disponible para este módulo. Verifique con el Administrador del Sistema.", vbInformation
                End If
            End If
        Else
            If Index = 4 Then
                    Unload Me
                If tipoCobro = "OPERACIONES" Then
                    Unload FrmFocus
                Else
                    If tipoCobro = "OPERACIONES_TOUCH" Then
                        FRM_OperTouch.carga_mesas
                        FRM_OperTouch.carga_Detalle (0)
                        FRM_OperTouch.cmd_cerrar.Enabled = False
                    End If
                End If
            Else
                If Val(Format(txtPago(4).Text, "General Number")) >= Val(Format(txtTot.Text, "General Number")) Then
                    If tipoCobro = "OPERACIONES" Then
                        Select Case Index
                            Case 2: Unload Me
                            Case 0:
                            'MsgBox Val(Format(txtPago(2).Text, "General Number"))
                            If Val(Format(txtPago(2).Text, "General Number")) > 0 Then
                                checkMonedero
                            End If
                            
                            cobrar (Index)
                            Case 1:
                            'MsgBox Val(Format(txtPago(2).Text, "General Number"))
                            If Val(Format(txtPago(2).Text, "General Number")) > 0 Then
                                checkMonedero
                            End If
                            cobrar (Index)
                        End Select
                    Else
                        If tipoCobro = "APARTADOS1" Then
                            Select Case Index
                                Case 2: Unload Me
                                Case 0: cobrarApartado (Index)
                                Case 1: cobrarApartado (Index)
                            End Select
                        Else
                            If tipoCobro = "APARTADOS2" Then
                                Select Case Index
                                    Case 2: Unload Me
                                    Case 0: cobrarApartado2 (Index)
                                    Case 1: cobrarApartado2 (Index)
                                End Select
                            Else
                                If tipoCobro = "CAMBIOS" Then
                                    Select Case Index
                                        Case 2: Unload Me
                                        Case 0: cobrarCambio (Index)
                                        Case 1: cobrarCambio (Index)
                                    End Select
                                Else
                                    If tipoCobro = "OPERACIONES_TOUCH" Then
                                        Select Case Index
                                            Case 2: Unload Me
                                            Case 0:
                                            'MsgBox Val(Format(txtPago(2).Text, "General Number"))
                                            If Val(Format(txtPago(2).Text, "General Number")) > 0 Then
                                                checkMonedero
                                            End If
                                            cobrar_touch (Index)
                                            Case 1:
                                            'MsgBox Val(Format(txtPago(2).Text, "General Number"))
                                            If Val(Format(txtPago(2).Text, "General Number")) > 0 Then
                                                checkMonedero
                                            End If
                                            cobrar_touch (Index)
                                            FRM_OperTouch.tiempo = 0
                                            FRM_OperTouch.Timer_tiempo.Enabled = True
                                        End Select
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else
                    MsgBox "Verifique la cantidad de pago. ", vbExclamation
                End If
            End If
        End If
    End If
End Sub
Private Sub cobrarCambio(ticket As Integer)
    FRM_Cambios.realizarCambios
    If ticket = 0 Then
        For b1 = 1 To Val(txtPrintCopias.Text)
            notaCambio (folioCambio)
            MsgBox "Operación realizada. " & vbCrLf & vbCrLf & "Impresión ticket " & b1 & " de " & txtPrintCopias.Text, vbInformation
        Next b1
    End If
    Unload Me
End Sub
Private Sub cobrarApartado(ticket As Integer)
    
    FRM_Apartados.apartado_crearApartado

    If ticket = 0 Then
        For b1 = 1 To Val(txtPrintCopias.Text)
            If tipoAprt = "APRT" Then
                notaApartado (folioTicket)
            Else
                If tipoAprt = "CRED" Then
                    notaCredito (folioTicket)
                End If
            End If
            MsgBox "Operación realizada. " & vbCrLf & vbCrLf & "Impresión ticket " & b1 & " de " & txtPrintCopias.Text, vbInformation
        Next b1
    End If
    Unload Me
End Sub
Private Sub cobrarApartado2(ticket As Integer)
    
    FRM_Apartados.apartado_crearApartado2
    If ticket = 0 Then
        For b1 = 1 To Val(txtPrintCopias.Text)
            If tipoAprt = "APRT" Then
                notaApartado (folioTicket)
            Else
                If tipoAprt = "CRED" Then
                    notaCredito (folioTicket)
                End If
            End If
            MsgBox "Operación realizada. " & vbCrLf & vbCrLf & "Impresión ticket " & b1 & " de " & txtPrintCopias.Text, vbInformation
        Next b1
    End If
    Unload Me
End Sub
Private Sub cobrar_touch(ticket As Integer)
'    On Error Resume Next
    Dim textoCita As String
    Dim Imagen1 As Stream
    Set Imagen1 = New Stream
    Imagen1.Type = adTypeBinary
    
'    SQL1 = "UPDATE VENTAS SET VENT_CLIEPERID = '" & FrmFocus.lblClieId(0).Caption & "', " & _
'    "VENT_CLIETIPOID = '" & FrmFocus.lblClieId(1).Caption & "', " & _
'    "VENT_CLIETIPO = '" & FrmFocus.lblClieId(2).Caption & "', " & _
'    "VENT_SUBTOTAL = '" & Val(Format(FrmFocus.txtSub.Text, "General Number")) & "', " & _
'    "VENT_DESCUENTO = '" & Val(Format(FrmFocus.txtDesc(0).Text, "General Number")) & "', " & _
'    "VENT_TOTAL = '" & Val(Format(txtTot.Text, "General Number")) & "', " & _
'    "VENT_PAGADO = '" & Val(Format(txtPago(4).Text, "General Number")) & "', " & _
'    "VENT_CAMBIO = '" & Val(Format(txtCambio.Text, "General Number")) & "', " & _
'    "VENT_PAGOEFECTIVO = '" & Val(Format(txtPago(0).Text, "General Number")) & "', " & _
'    "VENT_PAGOTARJETA = '" & Val(Format(txtPago(1).Text, "General Number")) & "', " & _
'    "VENT_PAGOCHEQUE = '" & Val(Format(txtPago(2).Text, "General Number")) & "', " & _
'    "VENT_STATUS = 'P', " & _
'    "VENT_OBSERVACIONES = '" & FrmFocus.txtObservacion.Text & "', " & _
'    "VENT_FECHAHORA_COBRO = '" & Format(Date, "yyyy-MM-dd") & " " & Format(Time, "HH:MM:SS") & "' " & _
'    "WHERE VENT_IDFOLIO = '" & FrmFocus.lInfo(1).Caption & "' "
    
    sql1 = "UPDATE VENTAS SET  " & _
    "VENT_SUBTOTAL = '" & Val(Format(FRM_OperTouch.lista_detalle.TextMatrix(FRM_OperTouch.lista_detalle.Row, 3), "General Number")) & "', " & _
    "VENT_DESCUENTO = '" & Val(0) & "', " & _
    "VENT_TOTAL = '" & Val(Format(txtTot.Text, "General Number")) & "', " & _
    "VENT_PAGADO = '" & Val(Format(txtPago(4).Text, "General Number")) & "', " & _
    "VENT_CAMBIO = '" & Val(Format(txtCambio.Text, "General Number")) & "', " & _
    "VENT_PAGOEFECTIVO = '" & Val(Format(txtPago(0).Text, "General Number")) & "', " & _
    "VENT_PAGOTARJETA = '" & Val(Format(txtPago(1).Text, "General Number")) & "', " & _
    "VENT_PAGOCHEQUE = '" & Val(Format(txtPago(2).Text, "General Number")) & "', " & _
    "VENT_STATUS = 'P', " & _
    "VENT_FECHAHORA_COBRO = '" & Format(Date, "yyyy-MM-dd") & " " & Format(Time, "HH:MM:SS") & "' " & _
    "WHERE VENT_IDFOLIO = '" & FRM_OperTouch.lista_detalle.TextMatrix(FRM_OperTouch.lista_detalle.Row, 1) & "' "
    con.Execute (sql1)
    
'    For b1 = 1 To FRM_OperTouch.lista_Producto.Rows - 1
'        If FrmFocus.lista.TextMatrix(b1, 1) = "MND" Then
'            SQL1 = "INSERT INTO MONEDERO (MND_TIPOGENERA, MND_CLIEPERID, MND_CLIETIPOID, MND_CLIETIPO, MND_VENTFOLIO, MND_USERPERID, MND_USERTIPOID, MND_USERTIPO, MND_PUNTOS, MND_TIPO, MND_FECHAHORA) " & _
'            "VALUES ('V', '" & FrmFocus.lblClieId(0).Caption & "', '" & FrmFocus.lblClieId(1).Caption & "', '" & FrmFocus.lblClieId(2).Caption & "', '" & FrmFocus.lInfo(1).Caption & "', " & _
'            "'" & FrmFocus.lblUserId(0).Caption & "', '" & FrmFocus.lblUserId(1).Caption & "', '" & FrmFocus.lblUserId(2).Caption & "',  '" & Val(Format(FrmFocus.lista.TextMatrix(b1, 5), "General Number")) & "', 'E', NOW() ) "
'            con.Execute (SQL1)
'        End If
'        If FrmFocus.lista.TextMatrix(b1, 6) = "M" Then
'            Dim fechaFin As Date
'            Dim fechaIni As Date
'
'            SQL1 = "SELECT MAX(MBR_FIN) as fecha FROM MEMBRESIAS WHERE MBR_PERTP_PER_ID = '" & FrmFocus.lblClieId(0).Caption & "' "
'            Set RES1 = con.Execute(SQL1)
'
'            If IsNull(RES1.Fields("Fecha")) Then
'
'                fechaIni = Date
'                fechaFin = fechaIni + (Val(FrmFocus.lista.TextMatrix(b1, 3)) * Val(FrmFocus.lista.TextMatrix(b1, 13)))
'
'
'            Else
'
'                If RES1.Fields("Fecha") < Date Then
'                    fechaIni = Date
'                Else
'                    fechaIni = RES1.Fields("Fecha") + 1
'                End If
'                'fechaFin = fechaIni
'                fechaFin = fechaIni + (Val(FrmFocus.lista.TextMatrix(b1, 3)) * Val(FrmFocus.lista.TextMatrix(b1, 13)))
'                'MsgBox fechaFin & "  " & (Val(FrmFocus.Lista.TextMatrix(b1, 3)) * Val(FrmFocus.Lista.TextMatrix(b1, 13)))
'
'            End If
'          '  MsgBox fechaIni & "  " & fechaFin
'            SQL1 = "INSERT INTO MEMBRESIAS " & _
'            "(MBR_CTMBID, MBR_PERTP_TIPO_ID, MBR_PERTP_PER_ID, MBR_PERTP_PER_TIPO, MBR_INICIO, MBR_FIN, MBR_STATUS, MBR_FECHA, MBR_VENTAFOLIO) " & _
'            "VALUES ('" & FrmFocus.lista.TextMatrix(b1, 1) & "', '" & FrmFocus.lblClieId(1).Caption & "', '" & FrmFocus.lblClieId(0).Caption & "', " & _
'            "'" & FrmFocus.lblClieId(2).Caption & "', '" & Format(fechaIni, "yyyy-MM-dd") & "',  " & _
'            "'" & Format(fechaFin, "yyyy-MM-dd") & "', 'A', NOW(), '" & FrmFocus.lInfo(1).Caption & "')"
'            con.Execute (SQL1)
'
'
''            SQL1 = "UPDATE MEMBRESIAS SET MBR_STATUS = 'A' " & _
''            "WHERE MBR_VENTAFOLIO = '" & FrmFocus.lInfo(1).Caption & "' AND MBR_STATUS = 'I'"
''            con.Execute (SQL1)
'        End If
'    Next b1
    
    
'    FrmFocus.lInfo(2).Caption = "PAGADO"
    Me.height = 4155
    txtCambio2.Visible = True
    cmdOpcion(3).Visible = True
    cmdOpcion(4).Visible = True
    SSTab1.Visible = False
'    tipoCobro = "-"
    label0(0).Caption = "Cambio"
    'timeSalir.Enabled = True
    txtCambio2.SetFocus
    
    If ticket = 0 Then
        For b1 = 1 To Val(txtPrintCopias.Text)
'            nota (FrmFocus.lInfo(1).Caption)
'            If mesas = True Then
'                notaPago (FRM_OperTouch.lista_detalle.TextMatrix(FRM_OperTouch.lista_detalle.Row, 0))
'            Else
'                MsgBox FRM_OperTouch.lista_detalle.TextMatrix(FRM_OperTouch.lista_detalle.Row, 1)
                nota (FRM_OperTouch.lista_detalle.TextMatrix(1, 1))
'
 '           End If
'    Unload Me
'    Unload FrmFocus
            
            If Image1(1).Visible = False Then
                MsgBox "Operación realizada. " & vbCrLf & vbCrLf & "Impresión ticket " & b1 & " de " & txtPrintCopias.Text, vbInformation
            End If
        Next b1
    Else
        MsgBox "Operación realizada." & vbCrLf & vbclrf & "Cambio: " & txtCambio.Text & txtPrintCopias.Text, vbInformation
    End If
    
timeSalir.Enabled = True
    'StartUpPosition = 1
    
''''Enviar el email
''''Enviar el email
''''Enviar el email
    Dim mensaje As String
    Dim tipoPago As String
    Dim totalCompra As Double
    
    mensaje = ""
    sqlEnvio = "SELECT msj_anexo, msj_anexo_nombre, msj_nombre, msj_descripcion, msj_copia FROM MENSAJES_EMAIL WHERE MSJ_TIPO = 'V' "
    Set resEnvioManil = con.Execute(sqlEnvio)
    If Not resEnvioManil.EOF Then
    
        tipoPago = ""
        totalCompra = txtTot.Text
        If IsNull(resEnvioManil.Fields("msj_copia")) = False Then

                If Val(Format(txtPago(0).Text, "General Number")) > 0 Then
                    tipoPago = tipoPago & " Efectivo "
                Else
                    If Val(Format(txtPago(1).Text, "General Number")) > 0 Then
                        tipoPago = tipoPago & " Tarjeta "
                    Else
                        If Val(Format(txtPago(2).Text, "General Number")) > 0 Then
                            tipoPago = tipoPago & " Puntos "
                        End If
                    End If
                End If
        End If
    
        mensaje = resEnvioManil.Fields("MSJ_DESCRIPCION")
        If IsNull(resEnvioManil.Fields("msj_anexo")) = False Then
            checarCarpetaTemp
            Imagen1.Open
            Imagen1.Write resEnvioManil.Fields("msj_anexo")
            Imagen1.SaveToFile direccionSistema & "\Temp\" & resEnvioManil.Fields("msj_anexo_nombre"), adSaveCreateOverWrite
            Imagen1.Close
            adjuntoDir = direccionSistema & "\Temp\" & resEnvioManil.Fields("msj_anexo_nombre")
        Else
            adjuntoDir = ""
        End If
        If FrmFocus.lblDatos(3).Caption <> "" Then
            Call enviar_Mail("MENSAJES", resEnvioManil.Fields("MSJ_NOMBRE"), FrmFocus.lblDatos(3).Caption, mensaje)
        End If
        If IsNull(resEnvioManil.Fields("msj_copia")) = False Then

            mensaje = mensaje & vbCrLf & vbCrLf & "Información adicional de la compra: " & vbCrLf & vbCrLf & _
            "Cliente: " & FrmFocus.lblDatos(2).Caption & vbCrLf & _
            FRM_Menu.menuBarra2.Panels(3).Text & vbCrLf & _
            "Folio: " & FrmFocus.lInfo(1).Caption & vbCrLf & _
            "Total compra: " & totalCompra & vbCrLf & _
            "Tipo de Pago: " & tipoPago & "-"
            
            'MsgBox mensaje

            Call enviar_Mail("MENSAJES", resEnvioManil.Fields("MSJ_NOMBRE") & " Mensaje copia", resEnvioManil.Fields("msj_copia"), mensaje)
        End If
    End If

    
    
    
End Sub

Private Sub cobrar(ticket As Integer)
'    On Error Resume Next
    Dim textoCita As String
    Dim Imagen1 As Stream
    Set Imagen1 = New Stream
    
    Imagen1.Type = adTypeBinary
    
    sql1 = "UPDATE VENTAS SET VENT_CLIEPERID = '" & FrmFocus.lblClieId(0).Caption & "', " & _
    "VENT_CLIETIPOID = '" & FrmFocus.lblClieId(1).Caption & "', " & _
    "VENT_CLIETIPO = '" & FrmFocus.lblClieId(2).Caption & "', " & _
    "VENT_SUBTOTAL = '" & Val(Format(FrmFocus.txtSub.Text, "General Number")) & "', " & _
    "VENT_DESCUENTO = '" & Val(Format(FrmFocus.txtDesc(0).Text, "General Number")) & "', " & _
    "VENT_TOTAL = '" & Val(Format(txtTot.Text, "General Number")) & "', " & _
    "VENT_PAGADO = '" & Val(Format(txtPago(4).Text, "General Number")) & "', " & _
    "VENT_CAMBIO = '" & Val(Format(txtCambio.Text, "General Number")) & "', " & _
    "VENT_PAGOEFECTIVO = '" & Val(Format(txtPago(0).Text, "General Number")) & "', " & _
    "VENT_PAGOTARJETA = '" & Val(Format(txtPago(1).Text, "General Number")) & "', " & _
    "VENT_PAGOCHEQUE = '" & Val(Format(txtPago(2).Text, "General Number")) & "', " & _
    "VENT_STATUS = 'P', " & _
    "VENT_OBSERVACIONES = '" & FrmFocus.txtObservacion.Text & "', " & _
    "VENT_FECHAHORA_COBRO = '" & Format(Date, "yyyy-MM-dd") & " " & Format(Time, "HH:MM:SS") & "' " & _
    "WHERE VENT_IDFOLIO = '" & FrmFocus.lInfo(1).Caption & "' "
    con.Execute (sql1)
    
    ''''Para los puntos''''
    If FrmFocus.lblDatos(5).Caption = "SI" Then
        
        Dim dia As String
        Dim DIA2 As Integer
        Dim puntosLista As Boolean
        Dim totPuntosProd As Double
        Dim montoMone As Double
        Dim tipoMone As String
        
        dia = Format(Date, "dddd")
        Select Case dia
            Case "domingo": DIA2 = 1
            Case "lunes": DIA2 = 2
            Case "martes": DIA2 = 3
            Case "miercoles": DIA2 = 4
            Case "jueves": DIA2 = 5
            Case "viernes": DIA2 = 6
            Case "sabado": DIA2 = 7
        End Select
        
        puntosLista = False
        totPuntosProd = 0
        
        sql1 = "SELECT * FROM CAT_PUNTOS_DIAS T1, CAT_PUNTOS T2 WHERE T1.PNTDS_DIA = '" & DIA2 & "' AND T2.PNT_ID = PNTDS_PNTID AND PNT_STATUS = 'A' "
        Set RES1 = con.Execute(sql1)
            
        If Not RES1.EOF Then
            ''''Va solo por los productos
            If RES1.Fields("PNT_TIPO") = "P" Then
                puntosLista = True
                tipoMone = RES1.Fields("PNT_TIPOVALOR")
                montoMone = RES1.Fields("PNT_VALOR")
            End If
            ''''Va por el total de la vanta
            If RES1.Fields("PNT_TIPO") = "T" Then
                
            End If
            '''Va por solo las membresias
            If RES1.Fields("PNT_TIPO") = "M" Then
                
            End If
            
        End If
        
    End If
    
    ''''Para los puntos''''
    
    For b1 = 1 To FrmFocus.lista.Rows - 1
        ''''Para meter a monedero en caso de que utilice
        If FrmFocus.lista.TextMatrix(b1, 1) = "MND" Then
            sql1 = "INSERT INTO MONEDERO (MND_TIPOGENERA, MND_CLIEPERID, MND_CLIETIPOID, MND_CLIETIPO, MND_VENTFOLIO, MND_USERPERID, MND_USERTIPOID, MND_USERTIPO, MND_PUNTOS, MND_TIPO, MND_FECHAHORA) " & _
            "VALUES ('V', '" & FrmFocus.lblClieId(0).Caption & "', '" & FrmFocus.lblClieId(1).Caption & "', '" & FrmFocus.lblClieId(2).Caption & "', '" & FrmFocus.lInfo(1).Caption & "', " & _
            "'" & FrmFocus.lblUserId(0).Caption & "', '" & FrmFocus.lblUserId(1).Caption & "', '" & FrmFocus.lblUserId(2).Caption & "',  '" & Val(Format(FrmFocus.lista.TextMatrix(b1, 5), "General Number")) & "', 'E', NOW() ) "
            con.Execute (sql1)
        End If
        ''''Si es una membresia la que se vende
        ''''Si es una membresia la que se vende
        If FrmFocus.lista.TextMatrix(b1, 6) = "M" Then
            Dim fechaFin As Date
            Dim fechaIni As Date
            sql1 = "SELECT MAX(MBR_FIN) as fecha FROM MEMBRESIAS WHERE MBR_PERTP_PER_ID = '" & FrmFocus.lblClieId(0).Caption & "' "
            Set RES1 = con.Execute(sql1)
            
            If IsNull(RES1.Fields("Fecha")) Then
            
                fechaIni = Date
                fechaFin = fechaIni + (Val(FrmFocus.lista.TextMatrix(b1, 3)) * Val(FrmFocus.lista.TextMatrix(b1, 13)))
            Else
                If RES1.Fields("Fecha") < Date Then
                    fechaIni = Date
                Else
                    fechaIni = RES1.Fields("Fecha") + 1
                End If
                fechaFin = fechaIni + (Val(FrmFocus.lista.TextMatrix(b1, 3)) * Val(FrmFocus.lista.TextMatrix(b1, 13)))
            End If
            sql1 = "INSERT INTO MEMBRESIAS " & _
            "(MBR_CTMBID, MBR_PERTP_TIPO_ID, MBR_PERTP_PER_ID, MBR_PERTP_PER_TIPO, MBR_INICIO, MBR_FIN, MBR_STATUS, MBR_FECHA, MBR_VENTAFOLIO) " & _
            "VALUES ('" & FrmFocus.lista.TextMatrix(b1, 1) & "', '" & FrmFocus.lblClieId(1).Caption & "', '" & FrmFocus.lblClieId(0).Caption & "', " & _
            "'" & FrmFocus.lblClieId(2).Caption & "', '" & Format(fechaIni, "yyyy-MM-dd") & "',  " & _
            "'" & Format(fechaFin, "yyyy-MM-dd") & "', 'A', NOW(), '" & FrmFocus.lInfo(1).Caption & "')"
            con.Execute (sql1)
            
            encuentra_Promocion ("M") 'Busca un promocion relacionada con la venta de membresias
            If PromoEncontrada = True Then
                'Si es por recomendacion
                sql1 = "SELECT PERTP_RECOMENDADO_ID, PERTP_RECOMENDADO_TIPO_ID, PERTP_RECOMENDADO_TIPO FROM PER_TIPO " & _
                "WHERE PERTP_TIPO_ID = '" & FrmFocus.lblClieId(1).Caption & "' AND  PERTP_PER_ID = '" & FrmFocus.lblClieId(0).Caption & "' "
                Set RES1 = con.Execute(sql1)
                
                If Not RES1.EOF Then
                    sql1 = "SELECT MAX(MBR_FIN) FECHA FROM MEMBRESIAS WHERE MBR_PERTP_PER_ID = '" & RES1.Fields("PERTP_RECOMENDADO_ID") & "'"
                    Set RES2 = con.Execute(sql1)
                    ''vALIDA SI LA MEMBRESIA ACTUAL DEL CLIENTE ESTA VIGENTE
                    If RES2.Fields("FECHA") >= Date Then
                        sql1 = "INSERT INTO MEMBRESIAS " & _
                        "(MBR_CTMBID, MBR_PERTP_TIPO_ID, MBR_PERTP_PER_ID, MBR_PERTP_PER_TIPO, MBR_INICIO, MBR_FIN, MBR_STATUS, MBR_FECHA, MBR_VENTAFOLIO) " & _
                        "VALUES ('" & promoMembId & "', '" & RES1.Fields("PERTP_RECOMENDADO_TIPO_ID") & "', '" & RES1.Fields("PERTP_RECOMENDADO_ID") & "', " & _
                        "'" & RES1.Fields("PERTP_RECOMENDADO_TIPO") & "', '" & Format((RES2.Fields("FECHA") + 1), "yyyy-MM-dd") & "',  " & _
                        "'" & Format((RES2.Fields("FECHA") + 1 + promoMemDias), "yyyy-MM-dd") & "', 'A', NOW(), '" & FrmFocus.lInfo(1).Caption & "')"
                        con.Execute (sql1)
                    End If
                End If
                
                
            End If
            PromoEncontrada = False
            
        End If
        ''''Si es una membresia la que se vende
        
        ''''Para los puntos''''
        ''''Para los puntos''''
        
        
        '''Actualiza inventario
        '''Actualiza inventario y cuenta puntos por productos
        If FrmFocus.lista.TextMatrix(b1, 6) = "P" Then
            sql1 = "UPDATE PRODUCTOS SET PROD_CANT = PROD_CANT - " & Val(FrmFocus.lista.TextMatrix(b1, 3)) & " " & _
            "WHERE PROD_ID = '" & FrmFocus.lista.TextMatrix(b1, 7) & "' AND PROD_SERV = 'P' AND PROD_INVENTARIO = 'S' "
            'MsgBox SQL1
            con.Execute (sql1)
        
                If puntosLista = True Then
                    totPuntosProd = FrmFocus.lista.TextMatrix(b1, 14)
                End If
            
            If FrmFocus.lista.TextMatrix(b1, 18) = "D" Then
                sql1 = "SELECT * from view_PRODUCTO_DEPENDIENTEs " & _
                "WHERE ID = '" & FrmFocus.lista.TextMatrix(b1, 7) & "' AND CODIGO = '" & FrmFocus.lista.TextMatrix(b1, 1) & "' "
                Set RES1 = con.Execute(sql1)
                
                Do While Not RES1.EOF
                    sql1 = "UPDATE PRODUCTOS SET PROD_CANT = PROD_CANT - '" & RES1.Fields("CANTIDAD_EQUI") & "'  " & _
                    "WHERE PROD_ID  = '" & RES1.Fields("ID_DEPEN") & "'  AND PROD_CODIGO = '" & RES1.Fields("CODIGO_DEPEN") & "'"
                    con.Execute (sql1)
                    'MsgBox sql1
                    RES1.MoveNext
                    
                Loop
            End If
        
        
        End If
        '''Actualiza inventario

    Next b1
    
    '''Inserta Registro de puntos generados
    If puntosLista = True Then
        If Val(Format(FrmFocus.txtDesc(1).Text, "General Number")) > 0 Then
            totPuntosProd = totPuntosProd - (totPuntosProd * (Val(Format(FrmFocus.txtDesc(1).Text, "General Number")) / 100))
        End If
        
        If tipoMone = "P" Then
            totPuntosProd = ((montoMone / 100) * totPuntosProd)
        Else
            If tipoMone = "T" Then
                totPuntosProd = totPuntosProd
            End If
        End If
        
        sql1 = "INSERT INTO MONEDERO (MND_TIPOGENERA, MND_CLIEPERID, MND_CLIETIPOID, MND_CLIETIPO, MND_VENTFOLIO, MND_USERPERID, MND_USERTIPOID, MND_USERTIPO, MND_PUNTOS, MND_TIPO, MND_FECHAHORA) VALUES " & _
        "('V', '" & FrmFocus.lblClieId(0).Caption & "', '" & FrmFocus.lblClieId(1).Caption & "', '" & FrmFocus.lblClieId(2).Caption & "', '" & FrmFocus.lInfo(1).Caption & "' , " & _
        "'" & FrmFocus.lblUserId(0).Caption & "', '" & FrmFocus.lblUserId(1).Caption & "', '" & FrmFocus.lblUserId(2).Caption & "', '" & totPuntosProd & "', 'R', NOW())"
        con.Execute (sql1)
        
    End If
    
    '''Inserta Registro de puntos generados
    
    
    
    
    FrmFocus.lInfo(2).Caption = "PAGADO"
    Me.height = 4155
    txtCambio2.Visible = True
    cmdOpcion(3).Visible = True
    cmdOpcion(4).Visible = True
    SSTab1.Visible = False
    txtCambio2.SetFocus
    label0(0).Caption = "Cambio"
    txtCambio2.SetFocus
    
    If ticket = 0 Then
        For b1 = 1 To Val(txtPrintCopias.Text)
            If mesas = True Then
                If tipoTicket = "CORTO" Then
                    notaPago (FrmFocus.lInfo(1).Caption)
                Else
                    nota (FrmFocus.lInfo(1).Caption)
                End If
            Else
                nota (FrmFocus.lInfo(1).Caption)
            End If
            If Image1(1).Visible = False Then
                MsgBox "Operación realizada. " & vbCrLf & vbCrLf & "Impresión ticket " & b1 & " de " & txtPrintCopias.Text, vbInformation
            End If
        Next b1
    Else
        MsgBox "Operación realizada." & vbCrLf & vbclrf & "Cambio: " & txtCambio.Text & txtPrintCopias.Text, vbInformation
    End If
    
timeSalir.Enabled = True
    'StartUpPosition = 1
    
''''Enviar el email
''''Enviar el email
''''Enviar el email
    Dim mensaje As String
    Dim tipoPago As String
    Dim totalCompra As Double
    
    mensaje = ""
    sqlEnvio = "SELECT msj_anexo, msj_anexo_nombre, msj_nombre, msj_descripcion, msj_copia FROM MENSAJES_EMAIL WHERE MSJ_TIPO = 'V' "
    Set resEnvioManil = con.Execute(sqlEnvio)
    If Not resEnvioManil.EOF Then
    
        tipoPago = ""
        totalCompra = txtTot.Text
        If IsNull(resEnvioManil.Fields("msj_copia")) = False Then

                If Val(Format(txtPago(0).Text, "General Number")) > 0 Then
                    tipoPago = tipoPago & " Efectivo "
                Else
                    If Val(Format(txtPago(1).Text, "General Number")) > 0 Then
                        tipoPago = tipoPago & " Tarjeta "
                    Else
                        If Val(Format(txtPago(2).Text, "General Number")) > 0 Then
                            tipoPago = tipoPago & " Puntos "
                        End If
                    End If
                End If
        End If
    
        mensaje = resEnvioManil.Fields("MSJ_DESCRIPCION")
        If IsNull(resEnvioManil.Fields("msj_anexo")) = False Then
            checarCarpetaTemp
            Imagen1.Open
            Imagen1.Write resEnvioManil.Fields("msj_anexo")
            Imagen1.SaveToFile direccionSistema & "\Temp\" & resEnvioManil.Fields("msj_anexo_nombre"), adSaveCreateOverWrite
            Imagen1.Close
            adjuntoDir = direccionSistema & "\Temp\" & resEnvioManil.Fields("msj_anexo_nombre")
        Else
            adjuntoDir = ""
        End If
        If FrmFocus.lblDatos(3).Caption <> "" Then
            Call enviar_Mail("MENSAJES", resEnvioManil.Fields("MSJ_NOMBRE"), FrmFocus.lblDatos(3).Caption, mensaje)
        End If
        If IsNull(resEnvioManil.Fields("msj_copia")) = False Then

            mensaje = mensaje & vbCrLf & vbCrLf & "Información adicional de la compra: " & vbCrLf & vbCrLf & _
            "Cliente: " & FrmFocus.lblDatos(2).Caption & vbCrLf & _
            FRM_Menu.menuBarra2.Panels(3).Text & vbCrLf & _
            "Folio: " & FrmFocus.lInfo(1).Caption & vbCrLf & _
            "Total compra: " & totalCompra & vbCrLf & _
            "Tipo de Pago: " & tipoPago & "-"
            
            'MsgBox mensaje

            Call enviar_Mail("MENSAJES", resEnvioManil.Fields("MSJ_NOMBRE") & " Mensaje copia", resEnvioManil.Fields("msj_copia"), mensaje)
        End If
    End If

    
    
    
End Sub

Private Sub Command1_Click()
    impresoraTicket = "cobro"
    PRINT_Impresora.Show vbModal
End Sub

Private Sub Form_Load()
    tiempo = 20
    SSTab1.Tab = 0
    Me.height = 10020
    
    Call bloquear_cierre(Me, True, True, True)
    
    txtCambio2.Visible = False
    cmdOpcion(3).Visible = False
    cmdOpcion(4).Visible = False

'    MsgBox Val(Format(txtTot.Text, "General Number"))
'    If Val(Format(txtTot.Text, "General Number")) <= 0 Then
'        SSTab1.TabEnabled(1) = False
'        SSTab1.TabEnabled(2) = False
'        SSTab1.TabEnabled(3) = False
'    Else
'        SSTab1.TabEnabled(1) = True
'        SSTab1.TabEnabled(2) = True
'        SSTab1.TabEnabled(3) = True
'    End If
    
    If Val(Format(txtPago(4).Text, "General Number")) >= Val(Format(txtTot.Text, "General Number")) Then
        cmdOpcion(0).Enabled = True
        cmdOpcion(1).Enabled = True
    Else
        cmdOpcion(0).Enabled = False
        cmdOpcion(1).Enabled = False
    End If
    checkIMpresora
    'txtPago(0).SetFocus
End Sub
Private Sub checkIMpresora()
    txtPrint.Text = Printer.DeviceName
    
    sql1 = "SELECT SUC_ESTATUSTICKET, SUC_TICKET_COPIA FROM SUCURSAL"
    Set RES1 = con.Execute(sql1)
    
    If Not RES1.EOF Then
        If RES1.Fields("SUC_ESTATUSTICKET") = 1 Then
            Image1(0).Visible = True
            Image1(1).Visible = False
        Else
            Image1(1).Visible = True
            Image1(0).Visible = False
        
        End If
        txtPrintCopias.Text = RES1.Fields("SUC_TICKET_COPIA")
    Else
            Image1(1).Visible = True
            Image1(0).Visible = False
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    If FrmFocus.lInfo(2).Caption = "PAGADO" Then
'        Unload FrmFocus
'    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
    Dim monedero As Double
    Select Case SSTab1.Tab
        Case 0:
            txtPago(0).SelStart = 0
            txtPago(0).SelLength = Len(txtPago(0).Text)
            txtPago(0).SetFocus
        Case 1:
            If Val(Format(txtPago(0).Text, "General Number")) = 0 Then
                txtPago(1).Text = txtTot.Text
            ''Else
                '''
            End If
        
            txtPago(1).SelStart = 0
            txtPago(1).SelLength = Len(txtPago(1).Text)
            txtPago(1).SetFocus
        Case 2:
            If modBusqueda = "Apartado" Then
                monedero = Val(Format(FRM_Apartados.lblDatos(6).Caption, "General Number"))
            Else
                monedero = Val(Format(FrmFocus.lblDatos(6).Caption, "General Number"))
            End If
            
            If Val(monedero) > 0 Then
                txtPago(2).Enabled = True
                    
                If Val(monedero) >= Val(Format(txtTot.Text, "General Number")) Then
                    txtPago(2).Text = txtTot.Text
                Else
                    txtPago(2).Text = Val(monedero)
                End If
                    
                    'monedero = True
                    
'                    txtClave(0).Text = Right(txtClave(0).Text, (Len(txtClave(0).Text) - 2))
            Else
                txtPago(2).Enabled = False
            End If
            
            
            txtPago(2).SelStart = 0
            txtPago(2).SelLength = Len(txtPago(2).Text)
            txtPago(2).SetFocus

    End Select
End Sub

Private Sub timeSalir_Timer()
    tiempo = tiempo - 1
    cmdOpcion(4).Caption = "Salir (Esc) " & tiempo
    If tiempo = 0 Then
        timeSalir.Enabled = False
        cmdOpcion_Click (4)
    End If
End Sub

Private Sub txtCambio_Change()
    txtCambio2.Text = txtCambio.Text
End Sub

Private Sub txtCambio2_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Or 13 Then
'    nota (FrmFocus.lInfo(1).Caption)
    Unload Me
    Unload FrmFocus
End If
End Sub

Private Sub txtPago_Change(Index As Integer)
    Dim monedero As Double
    If Index < 4 Then
        txtPago(4).Text = FormatCurrency(Val(Format(txtPago(0).Text)) + Val(Format(txtPago(1).Text)) + Val(Format(txtPago(2).Text)) + Val(Format(txtPago(3).Text)))
        
        If Index = 2 Then
            If modBusqueda = "Apartado" Then
                monedero = Val(Format(FRM_Apartados.lblDatos(6).Caption, "General Number"))
            Else
                monedero = Val(Format(FrmFocus.lblDatos(6).Caption, "General Number"))
            End If
            If Val(monedero) < Val(Format(txtPago(2).Text, "General Number")) Then
                txtPago(2).Text = "0"
            Else
                If Val(Format(txtPago(2).Text, "General Number")) > Val(Format(txtTot.Text, "General Number")) Then
                    txtPago(2).Text = Val(Format(txtTot.Text, "General Number"))
                End If
            End If
        End If
        
        If Val(Format(txtPago(4).Text, "General Number")) >= Val(Format(txtTot.Text, "General Number")) Then
            txtPago(4).ForeColor = vbBlack
            cmdOpcion(0).Enabled = True
            cmdOpcion(1).Enabled = True
        Else
            txtPago(4).ForeColor = &HFF&
            txtCambio.Text = "0"
            cmdOpcion(0).Enabled = False
            cmdOpcion(1).Enabled = False
        End If
        
    txtCambio.Text = FormatCurrency((Val(Format(txtPago(4).Text, "General number")) - Val(Format(txtTot.Text, "General Number"))))
    
            If txtPago(Index).Text = "" Then
                'txtCambio.Text = "0"
                cmdOpcion(0).Enabled = False
                cmdOpcion(1).Enabled = False
            End If
    
        
    
    End If
End Sub

Private Sub txtPago_DblClick(Index As Integer)
    Shell "osk.exe"
End Sub

Private Sub txtPago_GotFocus(Index As Integer)
    If Val(Format(txtPago(Index).Text, "General Number")) = 0 Then
        txtPago(Index).SelStart = 0
        txtPago(Index).SelLength = Len(txtPago(Index).Text)
        'txtPago(Index).SetFocus
    End If
End Sub

Private Sub txtPago_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
If keysacii = 27 Then
    cmdOpcion_Click (2)
Else
    If KeyAscii = 13 Then
        If cmdOpcion(0).Enabled = True Then
    
            If Val(Format(txtCambio.Text, "General Number")) >= 0 And Val(Format(txtPago(4).Text, "General Number")) >= Val(Format(txtTot.Text, "General Number")) Then
                
                If Index = 2 Then
                    FrmFocus.addMonedero
                End If
                
                
                cmdOpcion_Click (0)
            Else
                MsgBox "Opción no disponible por falta de información. Verifique.", vbInformation
            End If
        Else
            MsgBox "Opción no permitida. Verifique.", vbInformation
        End If
    Else
        Call NumerosPunto(KeyAscii)
    End If
End If
End Sub

Private Sub txtPrintCopias_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        sql1 = "UPDATE SUCURSAL SET SUC_TICKET_COPIA = '" & Val(txtPrintCopias.Text) & "'"
        con.Execute (sql1)
        
        MsgBox "Se ha actualizado el valor para el número de copias del ticket.", vbInformation
    End If
End Sub
