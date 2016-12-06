Attribute VB_Name = "FCN_Variables"
Public PromoEncontrada As Boolean
Public PromoTipo As String
Public promoId As Long
Public promoMembId As Long
Public promoMemDias As Long

Public cancelarMotivo As String
Public tipoIdentificador As String
Public formDescripcion As Form
Public teclado As String
Public tipo_AccesoTouch As String
Public direccionSistema As String
Public tipoCatTipo As String
Public tipoBusqueda As String
Public usuarioInicial As Boolean
Public modBusqueda As String
Public impresoraTicket As String
Public periodoValor As String
Public idUserHuella As Long
Public tipoCita As String
Public clavesCitas(100, 100)
Public tipoHuellas As String
Public tipoCobro As String
Public idAgenda As Long
Public dbActual As String
Public loadDb As Boolean
Public tipoAprt As String
Public mesas As Boolean
Public tipoTicket As String
'Public tTip As clss_ToolTip
Public tipoPersona As String
'''''''''''''''''''''
'''''ToolTip
Public TT1 As New clss_ToolTipText
Public TT2 As New clss_ToolTipText
Public TT3 As New clss_ToolTipText
Public TT4 As New clss_ToolTipText
Public TT5 As New clss_ToolTipText
Public TT6 As New clss_ToolTipText
Public TT7 As New clss_ToolTipText
Public TT8 As New clss_ToolTipText
Public TT9 As New clss_ToolTipText
Public TT10 As New clss_ToolTipText
Public TT11 As New clss_ToolTipText
Public TT12 As New clss_ToolTipText
Public TT13 As New clss_ToolTipText
Public TT14 As New clss_ToolTipText
Public TT15 As New clss_ToolTipText
Public TT16 As New clss_ToolTipText
Public TT17 As New clss_ToolTipText
Public TT18 As New clss_ToolTipText
Public TT19 As New clss_ToolTipText
Public TT20 As New clss_ToolTipText
'''''''''''''

Public FrmOper As MDIC_Operaciones
Public FrmOper2 As MDIC_Operaciones2
Public FrmTickets As MDIC_OperTickets
Public numFrmTicket As Integer
Public numFrmOper As Integer
Public numFrmOper2 As Integer
Public tikcet As Boolean
Public folioTicket As Long
Public folioCambio As Long


Public nForms As Integer
Public FrmFocus As Form
Public Const sCaption = "Operacion "

Public vendetId As Long

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Function CapsLockOn() As Boolean
    Dim iKeyState As Integer
    iKeyState = GetKeyState(vbKeyCapital)
    CapsLockOn = (iKeyState = 1 Or iKeyState = -127)
End Function

Public Sub checarCarpetaTemp()
On Error Resume Next
    Dim valor As String
    
    valor = Dir(direccionSistema & "\Temp", vbDirectory)
    'MsgBox valor
    If valor <> "" Then
    Else
        MkDir (direccionSistema & "\Temp")
    End If

End Sub

Public Sub ordenarLista(lista As MSFlexGrid)

        Static Modo  As Boolean
        If Modo Then
            lista.Col = lista.MouseCol
            lista.Sort = 2
            Modo = False
        Else
            lista.Col = lista.MouseCol
            lista.Sort = 1
            Modo = True
        End If

End Sub

Public Sub encuentra_Promocion(tipo As String)
Dim SQL As String
Dim RES1 As Recordset
Dim RES2 As Recordset
Dim RES3 As Recordset
Select Case tipo
    'Case m: Membresias
    Case "M": sql1 = "SELECT * FROM PROMOCIONES WHERE PROMO_RELACION = 'M' "
    Set RES1 = con.Execute(sql1)
    If RES1.EOF = False Then
        'Si obtiene por recomendacion
        If RES1.Fields("PROMO_OBTIENEPOR") = "R" Then
            PromoEncontrada = True
            sql1 = "SELECT * FROM PROMOCIONES_RELACION WHERE PROMR_PROMOID = '" & RES1.Fields("Prom_id") & "'"
            Set RES2 = con.Execute(sql1)
            If RES2.EOF = False Then
                sql1 = "SELECT * FROM CAT_MEMBRESIAS WHERE CTMB_ID = '" & RES2.Fields("PROMR_CTMB_ID") & "'"
                Set RES3 = con.Execute(sql1)
                If RES3.EOF = False Then
                    promoMembId = RES3.Fields("CTMB_ID")
                    promoMemDias = RES3.Fields("CTMB_DIAS")
                End If
            End If
        End If
    
    
    End If

End Select

End Sub

    Public Function ImprimirLogo(PathImagen As String, Alignment As AlignmentConstants, pY As Long, tAltura As Long)
        Dim tAncho As Long
        Dim xFoto As IPictureDisp
        Set xFoto = LoadPicture(PathImagen)
        tAncho = Round(Printer.ScaleX(xFoto.width, vbHimetric, vbMillimeters))
        tAncho = 40
        Select Case Alignment
            Case vbCenter
                pX = (Printer.ScaleWidth - tAncho) \ 2
            Case vbLeftJustify
                pX = 2
            Case vbRightJustify
                pX = Printer.ScaleWidth - tAncho - 2
        End Select
        Printer.PaintPicture LoadPicture(PathImagen), pX, pY, tAncho, tAltura
        Set xFoto = Nothing
    End Function
