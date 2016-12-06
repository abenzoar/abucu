Attribute VB_Name = "FCN_tICKET"
'Option Explicit
Public Sub abrirCajon()
    Printer.KillDoc
    Printer.Font = "Courier New"
    Printer.FontSize = 10
    Printer.FontBold = True
    
    Printer.Print " "
    Printer.EndDoc
    
End Sub
Public Sub impresionRenglones(texto As String)

Dim texto2 As String
Dim renglones As Integer
Dim largo As Long
Dim b1 As Long

'texto1 = texto
largo = 25

If Len(texto) > largo Then
    renglones = Len(texto) / largo
    If renglones - Int(renglonres) >= (0.5) Then
        renglones = Int(renglones) + 1
    Else
        renglones = Round(renglones, 0)
    End If
    
    
    For b1 = 1 To renglones
        texto2 = Left(texto, largo)
        If Len(texto) > largo Then
            texto = Right(texto, ((Len(texto)) - largo))
            Printer.Print texto2
        Else
            Printer.Print texto
            Exit For
        End If
    Next b1
    

    
Else
    Printer.Print texto
End If

End Sub



Public Sub nota(folio As String)
   ' On Error Resume Next
    
    Dim sql1 As String
    Dim RES3 As Recordset
    Dim RES4 As Recordset
    Dim RES5 As Recordset
    Dim RES6 As Recordset

    Dim SUBTOTAL
    Dim DESCUENTO
    Dim total
    Dim PAGOEFECTIVO
    Dim PAGOTARJETA
    Dim PAGOCHEQUE
    Dim PAGADO
    Dim CAMBIO
    Dim OBSERVACIONES As String
    Dim IVA As String
    Dim Lineas() As String
    Dim ASIENTO As String
    
    sql1 = "select * from SUCURSAL"
    Set RES3 = con.Execute(sql1)
    IVA = "N"
    If Not RES3.EOF Then
        IVA = RES3.Fields("SUC_IVA")
        If RES3.Fields("SUC_ESTATUSTICKET") = 1 Then
            
        Else
            MsgBox "El status del ticket está desactivado.", vbInformation
            Exit Sub
        End If
    Else
        MsgBox "No se puede imprimir el ticket por que no tiene información referente a la sucursal del negocio. Verifique.", vbInformation
        Exit Sub
    End If
    
    
    sql1 = "SELECT T1.vent_fechahora_cobro, T1.VENT_MESA MESA, CONCAT(T2.PER_NOMBRE, ' ', T2.PER_PATERNO, ' ', T2.PER_MATERNO) CLIENTE, " & _
    "CONCAT(T3.PER_NOMBRE, ' ', T3.PER_PATERNO, ' ', T3.PER_MATERNO) USUARIO, CONCAT(T4.PER_NOMBRE, ' ', T4.PER_PATERNO, ' ', T4.PER_MATERNO) CAJA, " & _
    "T5.VENDET_PRODCODIGO, T5.VENDET_PRODSERV, T5.VENDET_PRODUCTOID, T5.VENDET_PRODUCTONOMBRE, T5.VENDET_ASIENTO ASIENTO, T5.VENDET_PRECIO, SUM(T5.VENDET_CANTIDAD) VENDET_CANTIDAD, T5.venDet_Descuento, " & _
    "T1.VENT_PAGOEFECTIVO, VENT_PERSONAS PERSONAS, VENT_PAGOTARJETA, VENT_PAGOCHEQUE, VENT_PAGADO, VENT_CAMBIO, (SELECT SUM(VENDET_PRECIO * VENDET_CANTIDAD) FROM VENTA_DETALLE WHERE VENDET_FOLIO =  T1.VENT_IDFOLIO and VENDET_STATUS = 'A') VENT_SUBTOTAL, " & _
    "(IF (T1.VENT_DESCUENTO = 0, (select sum(t4A.venDet_Descuento) from venta_detalle T4A where (t4A.venDet_Folio = t1.vent_IdFolio AND T4A.VENDET_STATUS = 'A')) ,  IF(T1.VENT_DESCUENTO IS NULL, 0, T1.VENT_DESCUENTO)    )) VENT_DESCUENTO, VENT_OBSERVACIONES, (((SELECT SUM(VENDET_PRECIO * VENDET_CANTIDAD) FROM VENTA_DETALLE WHERE VENDET_FOLIO =  T1.VENT_IDFOLIO and VENDET_STATUS = 'A')) - ((IF (T1.VENT_DESCUENTO = 0, (select sum(t4A.venDet_Descuento) from venta_detalle T4A where (t4A.venDet_Folio = t1.vent_IdFolio AND T4A.VENDET_STATUS = 'A')) ,  IF(T1.VENT_DESCUENTO IS NULL, 0, T1.VENT_DESCUENTO)    )) )) VENT_TOTAL, " & _
    "(SELECT T41.TOTAL FROM VIEW_MONEDERO_CLIENTES T41 WHERE T2.PER_ID = T41.PER_ID) MONEDERO, (SELECT SUM(MONEDERO) FROM VIEW_PUNTOS_ADMIN WHERE FOLIO = '" & folio & "' AND TIPO = 'RECIBE') MONE_RECIBE, (SELECT SUM(MONEDERO) FROM VIEW_PUNTOS_ADMIN WHERE FOLIO = '" & folio & "' AND TIPO = 'ENTREGA') MONE_ENTREGA " & _
    "FROM VENTAS T1, PERSONA T2, PERSONA T3, VENTA_DETALLE T5, PERSONA T4 " & _
    "Where T1.VENT_IDFOLIO = '" & folio & "' And T1.VENT_CLIEPERID = T2.PER_ID And T5.VENDET_VENDPERID = T3.PER_ID And T1.VENT_IDFOLIO = T5.VENDET_FOLIO And T1.VENT_VENDPERID = T4.PER_ID and T5.vendet_Status = 'A' " & _
    "GROUP BY T1.vent_fechahora_cobro, T1.VENT_MESA, CONCAT(T2.PER_NOMBRE, ' ', T2.PER_PATERNO, ' ', T2.PER_MATERNO), CONCAT(T3.PER_NOMBRE, ' ', T3.PER_PATERNO, ' ', T3.PER_MATERNO) , CONCAT(T4.PER_NOMBRE, ' ', T4.PER_PATERNO, ' ', T4.PER_MATERNO) , " & _
    "T5.VENDET_PRODCODIGO, T5.VENDET_PRODUCTONOMBRE, T5.VENDET_PRECIO, T5.venDet_Descuento, T1.VENT_PAGOEFECTIVO , VENT_PAGOTARJETA, VENT_PAGOCHEQUE, VENT_PAGADO, VENT_CAMBIO, VENT_SUBTOTAL, VENT_dESCUENTO, VENT_OBSERVACIONES, VENt_TOTAL, T5.VENDET_ASIENTO ORDER BY T5.VENDET_ASIENTO ASC "


    
    Set RES4 = con.Execute(sql1)
    If Not RES4.EOF Then
        SUBTOTAL = RES4.Fields("VENT_SUBTOTAL")
        DESCUENTO = RES4.Fields("VENT_dESCUENTO")
        total = RES4.Fields("VENT_TOTAL")
        PAGOEFECTIVO = RES4.Fields("VENT_PAGOEFECTIVO")
        PAGOTARJETA = RES4.Fields("VENT_PAGOTARJETA")
        PAGOCHEQUE = RES4.Fields("VENT_PAGOCHEQUE")
        PAGADO = RES4.Fields("VENT_PAGADO")
        CAMBIO = RES4.Fields("VENT_CAMBIO")
        If IsNull(RES4.Fields("MONEDERO")) Then
            monedero = 0
        Else
            monedero = Val(RES4.Fields("MONEDERO"))
        End If
        If IsNull(RES4.Fields("MONE_RECIBE")) Then
            mone_recibe = 0
        Else
            mone_recibe = Val(RES4.Fields("MONE_RECIBE"))
        End If
        If IsNull(RES4.Fields("MONE_ENTREGA")) Then
            mone_entrega = 0
        Else
            mone_entrega = Val(RES4.Fields("MONE_ENTREGA"))
        End If
        OBSERVACIONES = RES4.Fields("VENT_OBSERVACIONES")
    Else
        MsgBox "No se puede imprimir el ticket por que no tiene información referente a la venta referida. " & vbCrLf & vbCrLf & "Posible causa: Falta de información en la lista. " & vbCrLf & vbCrLf & "Verifique.", vbInformation
        Exit Sub
    End If
    
    Printer.KillDoc
    
'    Printer.PaintPicture FRM_Menu.imgInfo(1).Picture
    Printer.ScaleMode = vbMillimeters
    ImprimirLogo (direccionSistema & "\Temp\TempSucur.dat"), vbLeftJustify, 5, 10
    
    
    Printer.Font = "Courier New"
    
    Printer.FontSize = RES3.Fields("SUC_TICKET_SIZE_1")
    Printer.FontBold = True
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print UCase(RES3.Fields("SUC_NOMBRE")) & vbCrLf
    Printer.Print UCase(RES3.Fields("SUC_RAZON_SOCIAL"))
    Printer.FontBold = False
    Printer.Print UCase(RES3.Fields("SUC_DIR_CALLE"))
    Printer.Print UCase(RES3.Fields("SUC_DIR_NUM_EXT") & " " & RES3.Fields("SUC_DIR_NUM_INT"))
    Printer.Print UCase(RES3.Fields("SUC_dIR_COLONIA"))
    Printer.Print "CP: "; UCase(RES3.Fields("SUC_DIR_Cp"))
    Printer.Print UCase(RES3.Fields("SUC_DIR_CIUDAD")) '& " " & RES1.Fields("Municipio")
    Printer.Print UCase(RES3.Fields("SUC_RFC"))
    Printer.Print "TELS: " & RES3.Fields("SUC_TEL1") & " " & RES3.Fields("SUC_TEL2") & vbCrLf
    Printer.FontSize = RES3.Fields("SUC_TICKET_SIZE_2")
    Printer.FontBold = True
    Printer.Print "FECHA DE OPERACIÓN: "
    Printer.Print Format(RES4.Fields("vent_fechahora_cobro"), "dddd dd-mm-yyyy") & " " & Format(RES4.Fields("vent_fechahora_cobro"), "Short Time") & vbCrLf
    Printer.Print "FOLIO:   " & Format(folio, "0000000")
    If IsNull(RES4.Fields("MESA")) = False Then
        Printer.Print "MESA:      " & RES4.Fields("MESA") & ""
        Printer.Print "PERSONAS:  " & RES4.Fields("PERSONAS") & "" & vbCrLf
    End If
'    Printer.FontSize = 12
    
    Printer.FontBold = True
    Printer.Print "CLIENTE: "
    Printer.Print UCase(RES4.Fields("CLIENTE")) '& vbCrLf
    Printer.FontBold = False
    Printer.Print "MOSTRADOR: "
    Printer.Print UCase(RES4.Fields("CAJA"))
    'Printer.FontBold = False
    Printer.Print "- - - - - - - - - - - - - - -"
    Printer.Print "DETALLE DE OPERACIÓN:" & vbCrLf
    ASIENTO = ""
    
    'For b1 = 1 To FrmFocus.ListaOper.Rows - 1
    Do While Not RES4.EOF
        clave = RES4.Fields("VENDET_PRODCODIGO")
        If Len(clave) > 17 Then
            clave = Left(clave, 17)
        Else
            clave = clave & String(17 - Len(clave), " ")
        End If
        Nombre2 = RES4.Fields("VENDET_PRODUCTONOMBRE")
        If Len(Nombre2) > 28 Then
            Nombre2 = Left(Nombre2, 28)
        Else
            Nombre2 = Nombre2 & String(28 - Len(Nombre2), " ")
        End If
        cantidad = RES4.Fields("VENDET_CANTIDAD")
        Precio = RES4.Fields("VENDET_PRECIO")
        desc = RES4.Fields("venDet_Descuento")
        tot = (RES4.Fields("VENDET_CANTIDAD") * RES4.Fields("VENDET_PRECIO") - RES4.Fields("VENDET_DESCUENTO"))
        If Len(Precio) > 9 Then
            Precio = Left(Precio, 9)
        End If
        If Len(RES4.Fields("USUARIO")) > 28 Then
            atendio = Left(RES4.Fields("USUARIO"), 28)
        Else
            atendio = RES4.Fields("USUARIO")
        End If
        
        cantidad = cantidad & String(10 - Len(cantidad), " ")
        Precio = Precio & String(12 - Len(Precio), " ")
        desc = FormatCurrency(desc) & String(11 - Len(FormatCurrency(desc)), " ")
'        Tot = Tot & String(9 - Len(Tot), " ")
        titulo1 = "Cant" & String(10 - Len("Cant"), " ")
        titulo2 = "Precio" & String(10, " ")
        titulo3 = "Desc" & String(6, " ")
        titulo4 = "Total" & String(4, " ")
        titulo5 = "Atendio" & String(12, " ")
                
'        Printer.Print UCase(clave & vbCrLf & nombre2 & vbCrLf & Titulo1 & " " & Titulo2 & " " & Titulo3 & " " & Titulo4 & _
'        vbCrLf & cantidad & " " & PRECIO & " " & desc & " " & Tot & vbCrLf & Titulo5 & vbCrLf & Atendio & vbCrLf)
        Printer.FontBold = True
        If ASIENTO <> RES4.Fields("ASIENTO") Then
            Printer.Print "----------------------------------------"
            Printer.FontSize = 11
            Printer.FontBold = True
            Printer.Print "Asiento: " & RES4.Fields("ASIENTO")
            ASIENTO = RES4.Fields("ASIENTO")
        
        Else
        End If
        
        Printer.FontSize = 9
        Printer.FontBold = True
        
        Printer.Print UCase(clave)
        Printer.Print UCase(Nombre2)
        Printer.FontBold = False
        Printer.Print UCase(titulo1) & " " & UCase(titulo2)
        Printer.FontBold = True
        Printer.Print cantidad & " " & FormatCurrency(Precio)
        Printer.FontBold = False
        Printer.Print titulo3 & " " & titulo4
        Printer.FontBold = True
        Printer.Print desc & FormatCurrency(tot)
        Printer.FontBold = False
'        Printer.Print Titulo5 & vbCrLf & UCase(Atendio) & vbCrLf

        If RES4.Fields("VENDET_PRODSERV") = "M" Then
            sql1 = "SELECT * FROM MEMBRESIAS WHERE MBR_VENTAFOLIO = '" & folio & "' AND MBR_CTMBID = '" & RES4.Fields("VENDET_PRODUCTOID") & "' "
            Set RES6 = con.Execute(sql1)
                        
            If Not RES6.EOF Then
                Printer.Print RES6.Fields("MBR_INICIO") & " - " & RES6.Fields("MBR_FIN")
            End If
        End If

                       
    RES4.MoveNext
    Loop
    
    Printer.Print "- - - - - - - - - - - - - - - "
    If Len(Horario) > 0 Then
'        Printer.FontSize = 12
        Printer.FontSize = RES3.Fields("SUC_TICKET_SIZE_2")
        Printer.Print Horario
        Printer.Print "- - - - - - - - - - - - - - - "
    End If
    Printer.FontBold = True
    If IVA = "S" Then
        'Printer.Print "  SUB TOTAL: " & FormatCurrency((total / (1.16)))
        Printer.Print "  SUB TOTAL: " & FormatCurrency((SUBTOTAL / (1.16)))
        Printer.Print "  DESCUENTO: " & FormatCurrency(DESCUENTO)
        Printer.Print "        IVA: " & FormatCurrency((total) - (total / (1.16)))
    Else
        'Printer.Print "  SUB TOTAL: " & FormatCurrency((total) - (DESCUENTO))
        Printer.Print "  SUB TOTAL: " & FormatCurrency((SUBTOTAL))
        Printer.Print "  DESCUENTO: " & FormatCurrency(DESCUENTO)
        'Printer.Print "        IVA: " & FormatCurrency(0)
    End If
    Printer.Print "      TOTAL: " & FormatCurrency(total)
    Printer.Print "- - - - - - - - - - - - - - - "
    Printer.Print "FORMA DE PAGO:"
    If Val(PAGOEFECTIVO) > 0 Then
    Printer.Print "   EFECTIVO: " & FormatCurrency(PAGOEFECTIVO)
    End If
    If Val(PAGOTARJETA) > 0 Then
    Printer.Print "    TARJETA: " & FormatCurrency(PAGOTARJETA)
    End If
    If Val(PAGOCHEQUE) > 0 Then
    Printer.Print "   MONEDERO: " & FormatCurrency(PAGOCHEQUE)
    End If
    Printer.Print "     PAGADO: " & FormatCurrency(PAGADO)
    Printer.Print "     CAMBIO: " & FormatCurrency(CAMBIO)
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    If Val(monedero) > 0 Then
    Printer.Print "  --- MONEDERO ---"
    Printer.Print "   RECIBIDO: " & FormatCurrency(mone_recibe)
    Printer.Print "   APLICADO: " & FormatCurrency(mone_entrega)
    Printer.Print "      TOTAL: " & FormatCurrency(monedero)
    End If
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    Printer.FontBold = True
    
    Printer.Print "           OBSERVACIONES    " & vbCrLf
    Printer.FontBold = False
'    Printer.FontSize = 10
    Printer.FontSize = RES3.Fields("SUC_TICKET_SIZE_1")
    If Len(OBSERVACIONES) > 0 Then
    
        Lineas = Split(OBSERVACIONES, vbNewLine)
    
        For b1 = 0 To UBound(Lineas)
            'Printer.Print UCase(Lineas(b1))
            
            For c1 = 0 To ((Round((Len(Lineas(b1)) / 35))) + 1)
                If Len(Lineas(b1)) >= 35 Then
                    'largo = True
                    Printer.Print Left(Lineas(b1), 35)
                    Lineas(b1) = Right(Lineas(b1), (Len(Lineas(b1)) - 35))
                Else
                    Printer.Print Lineas(b1)
                    Exit For
                End If
            Next c1
        Next b1
        
        Printer.Print "- - - - - - - - - - - - - - - - - - "
    End If
    'Call Centrar(Eslogan, 15)
    Printer.Print vbCrLf & UCase(RES3.Fields("SUC_SLOGAN")) & vbCrLf
    'Call Centrar(Web, 15)
    Printer.Print UCase(RES3.Fields("SUC_PAGINA_WEB"))
    Printer.Print UCase(RES3.Fields("SUC_EMAIL"))

    impresionRenglones (RES3.Fields("SUC_INFORMACION"))


'    Lineas = Split(RES3.Fields("SUC_INFORMACION"), vbNewLine)
'    For b1 = 0 To UBound(Lineas)
'        For c1 = 0 To ((Round((Len(Lineas(b1)) / 35))) + 1)
'            If Len(Lineas(b1)) >= 35 Then
'                'largo = True
'                Printer.Print Left(Lineas(b1), 35)
'                Lineas(b1) = Right(Lineas(b1), (Len(Lineas(b1)) - 35))
'            Else
'                Printer.Print Lineas(b1)
'                Exit For
'            End If
'        Next c1
'    Next b1
    'Printer.Print RES3.Fields("SUC_INFORMACION") & vbCrLf & vbCrLf
'    Printer.FontSize = 12
'    Printer.FontName = "Control"
'    Printer.Print "P"  'Cut
    Printer.EndDoc
    

End Sub
Public Sub notaPago(folio As String)
    'On Error Resume Next
    
    Dim sql1 As String
    Dim RES3 As Recordset
    Dim RES4 As Recordset
    Dim RES5 As Recordset

    Dim SUBTOTAL
    Dim DESCUENTO
    Dim total
    Dim PAGOEFECTIVO
    Dim PAGOTARJETA
    Dim PAGOCHEQUE
    Dim PAGADO
    Dim CAMBIO
    Dim IVA As String
    Dim OBSERVACIONES As String

    Dim Lineas() As String
    
    sql1 = "select * from SUCURSAL"
    Set RES3 = con.Execute(sql1)
    IVA = "N"
    If Not RES3.EOF Then
        IVA = RES3.Fields("SUC_IVA")
        If RES3.Fields("SUC_ESTATUSTICKET") = 1 Then
            
        Else
            MsgBox "El status del ticket está desactivado.", vbInformation
            Exit Sub
        End If
    Else
        MsgBox "No se puede imprimir el ticket por que no tiene información referente a la sucursal del negocio. Verifique.", vbInformation
        Exit Sub
    End If
    
'    SQL1 = "SELECT T1.vent_fechahora_cobro, T1.VENT_MESA MESA, CONCAT(T2.PER_NOMBRE, ' ', T2.PER_PATERNO, ' ', T2.PER_MATERNO) CLIENTE, " & _
'    "CONCAT(T3.PER_NOMBRE, ' ', T3.PER_PATERNO, ' ', T3.PER_MATERNO) USUARIO, " & _
'    "CONCAT(T4.PER_NOMBRE, ' ', T4.PER_PATERNO, ' ', T4.PER_MATERNO) CAJA, " & _
'    "T5.VENDET_PRODCODIGO, T5.VENDET_PRODUCTONOMBRE, T5.VENDET_PRECIO, T5.VENDET_CANTIDAD, T5.venDet_Descuento, " & _
'    "T1.VENT_PAGOEFECTIVO, VENT_PAGOTARJETA, VENT_PAGOCHEQUE, VENT_PAGADO, VENT_CAMBIO, VENT_SUBTOTAL, VENT_dESCUENTO, VENT_OBSERVACIONES, VENt_TOTAL, (SELECT T41.TOTAL FROM VIEW_MONEDERO_CLIENTES T41 WHERE T2.PER_ID = T41.PER_ID) MONEDERO, (SELECT SUM(MONEDERO) FROM VIEW_PUNTOS_ADMIN WHERE FOLIO = '" & folio & "' AND TIPO = 'RECIBE') MONE_RECIBE, (SELECT SUM(MONEDERO) FROM VIEW_PUNTOS_ADMIN WHERE FOLIO = '" & folio & "' AND TIPO = 'ENTREGA') MONE_ENTREGA " & _
'    "FROM VENTAS T1, PERSONA T2, PERSONA T3, VENTA_DETALLE T5, PERSONA T4 " & _
'    "WHERE T1.VENT_IDFOLIO = '" & folio & "' AND T1.VENT_CLIEPERID = T2.PER_ID AND T5.VENDET_VENDPERID = T3.PER_ID AND " & _
'    "T1.VENT_IDFOLIO = T5.VENDET_FOLIO AND T1.VENT_VENDPERID = T4.PER_ID "

'    SQL1 = "SELECT T1.vent_fechahora_cobro, T1.vent_fechahora, T1.VENT_MESA MESA, CONCAT(T2.PER_NOMBRE, ' ', T2.PER_PATERNO, ' ', T2.PER_MATERNO) CLIENTE, " & _
'    "CONCAT(T3.PER_NOMBRE, ' ', T3.PER_PATERNO, ' ', T3.PER_MATERNO) USUARIO, CONCAT(T4.PER_NOMBRE, ' ', T4.PER_PATERNO, ' ', T4.PER_MATERNO) CAJA, " & _
'    "T5.VENDET_PRODCODIGO, T5.VENDET_PRODUCTONOMBRE, T5.VENDET_PRECIO, SUM(T5.VENDET_CANTIDAD) VENDET_CANTIDAD, T5.venDet_Descuento, T1.VENT_SUBTOTAL, T1.VENT_DESCUENTO, T1.VENT_TOTAL, " & _
'    "T1.VENT_PAGOEFECTIVO, VENT_PAGOTARJETA, VENT_PAGOCHEQUE, VENT_PAGADO, VENT_CAMBIO, (SELECT SUM(VENDET_PRECIO * VENDET_CANTIDAD) FROM VENTA_DETALLE WHERE VENDET_FOLIO =  T1.VENT_IDFOLIO) VENT_SUBTOTAL1, " & _
'    "(IF (T1.VENT_DESCUENTO = 0, (select sum(t4A.venDet_Descuento) from venta_detalle T4A where (t4A.venDet_Folio = t1.vent_IdFolio)) ,  (T1.VENT_DESCUENTO)   )) VENT_DESCUENTO1, VENT_OBSERVACIONES, (((SELECT SUM(VENDET_PRECIO * VENDET_CANTIDAD) FROM VENTA_DETALLE WHERE VENDET_FOLIO =  T1.VENT_IDFOLIO)) - ((IF (T1.VENT_DESCUENTO = 0, (select sum(t4A.venDet_Descuento) from venta_detalle T4A where (t4A.venDet_Folio = t1.vent_IdFolio)) ,  (T1.VENT_DESCUENTO)   )) )) VENT_TOTAL1, " & _
'    "(SELECT T41.TOTAL FROM VIEW_MONEDERO_CLIENTES T41 WHERE T2.PER_ID = T41.PER_ID) MONEDERO, (SELECT SUM(MONEDERO) FROM VIEW_PUNTOS_ADMIN WHERE FOLIO = '" & folio & "' AND TIPO = 'RECIBE') MONE_RECIBE, (SELECT SUM(MONEDERO) FROM VIEW_PUNTOS_ADMIN WHERE FOLIO = '" & folio & "' AND TIPO = 'ENTREGA') MONE_ENTREGA " & _
'    "FROM VENTAS T1, PERSONA T2, PERSONA T3, VENTA_DETALLE T5, PERSONA T4 " & _
'    "Where T1.VENT_IDFOLIO = '" & folio & "' And T1.VENT_CLIEPERID = T2.PER_ID And T5.VENDET_VENDPERID = T3.PER_ID And T1.VENT_IDFOLIO = T5.VENDET_FOLIO And T1.VENT_VENDPERID = T4.PER_ID " & _
'    "GROUP BY T1.vent_fechahora_cobro, T1.VENT_MESA, CONCAT(T2.PER_NOMBRE, ' ', T2.PER_PATERNO, ' ', T2.PER_MATERNO), CONCAT(T3.PER_NOMBRE, ' ', T3.PER_PATERNO, ' ', T3.PER_MATERNO) , CONCAT(T4.PER_NOMBRE, ' ', T4.PER_PATERNO, ' ', T4.PER_MATERNO) , " & _
'    "T5.VENDET_PRODCODIGO, T5.VENDET_PRODUCTONOMBRE, T5.VENDET_PRECIO, T5.venDet_Descuento, T1.VENT_PAGOEFECTIVO , VENT_PAGOTARJETA, VENT_PAGOCHEQUE, VENT_PAGADO, VENT_CAMBIO, VENT_SUBTOTAL, VENT_dESCUENTO, VENT_OBSERVACIONES, VENt_TOTAL "


    sql1 = "SELECT T1.vent_fechahora_cobro, T1.VENT_MESA MESA, T1.VENT_PERSONAS PERSONAS, CONCAT(T2.PER_NOMBRE, ' ', T2.PER_PATERNO, ' ', T2.PER_MATERNO) CLIENTE, " & _
    "CONCAT(T3.PER_NOMBRE, ' ', T3.PER_PATERNO, ' ', T3.PER_MATERNO) USUARIO, CONCAT(T4.PER_NOMBRE, ' ', T4.PER_PATERNO, ' ', T4.PER_MATERNO) CAJA, " & _
    "T5.VENDET_PRODCODIGO, T5.VENDET_PRODUCTONOMBRE, T5.VENDET_PRECIO, SUM(T5.VENDET_CANTIDAD) VENDET_CANTIDAD, T5.venDet_Descuento, " & _
    "T1.VENT_PAGOEFECTIVO, VENT_PAGOTARJETA, VENT_PAGOCHEQUE, VENT_PAGADO, VENT_CAMBIO, (SELECT SUM(VENDET_PRECIO * VENDET_CANTIDAD) FROM VENTA_DETALLE WHERE VENDET_FOLIO =  T1.VENT_IDFOLIO and VENDET_STATUS = 'A') VENT_SUBTOTAL, " & _
    "(IF (T1.VENT_DESCUENTO = 0, (select sum(t4A.venDet_Descuento) from venta_detalle T4A where (t4A.venDet_Folio = t1.vent_IdFolio and T4A.VENDET_STATUS = 'A')) ,  IF(T1.VENT_DESCUENTO IS NULL, 0, T1.VENT_DESCUENTO)    )) VENT_DESCUENTO, VENT_OBSERVACIONES, (((SELECT SUM(VENDET_PRECIO * VENDET_CANTIDAD) FROM VENTA_DETALLE WHERE VENDET_FOLIO =  T1.VENT_IDFOLIO AND VENDET_STATUS = 'A')) - ((IF (T1.VENT_DESCUENTO = 0, (select sum(t4A.venDet_Descuento) from venta_detalle T4A where (t4A.venDet_Folio = t1.vent_IdFolio and T4A.VENDET_STATUS = 'A')) ,  IF(T1.VENT_DESCUENTO IS NULL, 0, T1.VENT_DESCUENTO)    )) )) VENT_TOTAL, " & _
    "(SELECT T41.TOTAL FROM VIEW_MONEDERO_CLIENTES T41 WHERE T2.PER_ID = T41.PER_ID) MONEDERO, (SELECT SUM(MONEDERO) FROM VIEW_PUNTOS_ADMIN WHERE FOLIO = '" & folio & "' AND TIPO = 'RECIBE') MONE_RECIBE, (SELECT SUM(MONEDERO) FROM VIEW_PUNTOS_ADMIN WHERE FOLIO = '" & folio & "' AND TIPO = 'ENTREGA') MONE_ENTREGA " & _
    "FROM VENTAS T1, PERSONA T2, PERSONA T3, VENTA_DETALLE T5, PERSONA T4 " & _
    "Where T1.VENT_IDFOLIO = '" & folio & "' And T1.VENT_CLIEPERID = T2.PER_ID And T5.VENDET_VENDPERID = T3.PER_ID And T1.VENT_IDFOLIO = T5.VENDET_FOLIO And T1.VENT_VENDPERID = T4.PER_ID and T5.vendet_Status = 'A'" & _
    "GROUP BY T1.vent_fechahora_cobro, T1.VENT_MESA, CONCAT(T2.PER_NOMBRE, ' ', T2.PER_PATERNO, ' ', T2.PER_MATERNO), CONCAT(T3.PER_NOMBRE, ' ', T3.PER_PATERNO, ' ', T3.PER_MATERNO) , CONCAT(T4.PER_NOMBRE, ' ', T4.PER_PATERNO, ' ', T4.PER_MATERNO) , " & _
    "T5.VENDET_PRODCODIGO, T5.VENDET_PRODUCTONOMBRE, T5.VENDET_PRECIO, T5.venDet_Descuento, T1.VENT_PAGOEFECTIVO , VENT_PAGOTARJETA, VENT_PAGOCHEQUE, VENT_PAGADO, VENT_CAMBIO, VENT_SUBTOTAL, VENT_dESCUENTO, VENT_OBSERVACIONES, VENt_TOTAL "

    Set RES4 = con.Execute(sql1)
    If Not RES4.EOF Then
        SUBTOTAL = RES4.Fields("VENT_SUBTOTAL")
        DESCUENTO = RES4.Fields("VENT_dESCUENTO")
        total = RES4.Fields("VENT_TOTAL")
        PAGOEFECTIVO = RES4.Fields("VENT_PAGOEFECTIVO")
        PAGOTARJETA = RES4.Fields("VENT_PAGOTARJETA")
        PAGOCHEQUE = RES4.Fields("VENT_PAGOCHEQUE")
        PAGADO = RES4.Fields("VENT_PAGADO")
        CAMBIO = RES4.Fields("VENT_CAMBIO")
        If IsNull(RES4.Fields("MONEDERO")) Then
            monedero = 0
        Else
            monedero = Val(RES4.Fields("MONEDERO"))
        End If
        If IsNull(RES4.Fields("MONE_RECIBE")) Then
            mone_recibe = 0
        Else
            mone_recibe = Val(RES4.Fields("MONE_RECIBE"))
        End If
        If IsNull(RES4.Fields("MONE_ENTREGA")) Then
            mone_entrega = 0
        Else
            mone_entrega = Val(RES4.Fields("MONE_ENTREGA"))
        End If
        OBSERVACIONES = RES4.Fields("VENT_OBSERVACIONES")
    Else
        MsgBox "No se puede imprimir el ticket por que no tiene información referente a la venta referida. " & vbCrLf & vbCrLf & "Posible causa: Falta de información en la lista. " & vbCrLf & vbCrLf & "Verifique.", vbInformation
        Exit Sub
    End If
    
    Printer.KillDoc
    Printer.Font = "Courier New"
    Printer.FontSize = RES3.Fields("SUC_TICKET_SIZE_1")
    Printer.FontBold = True
    
    Printer.PaintPicture FRM_Menu.imgInfo(1).Picture
    
    Printer.Print UCase(RES3.Fields("SUC_NOMBRE")) & vbCrLf
    Printer.Print UCase(RES3.Fields("SUC_RAZON_SOCIAL"))
    Printer.FontBold = False
    Printer.Print UCase(RES3.Fields("SUC_DIR_CALLE"))
    Printer.Print UCase(RES3.Fields("SUC_DIR_NUM_EXT") & " " & RES3.Fields("SUC_DIR_NUM_INT"))
    Printer.Print UCase(RES3.Fields("SUC_dIR_COLONIA"))
    Printer.Print "CP: "; UCase(RES3.Fields("SUC_DIR_Cp"))
    Printer.Print UCase(RES3.Fields("SUC_DIR_CIUDAD")) '& " " & RES1.Fields("Municipio")
    Printer.Print UCase(RES3.Fields("SUC_RFC"))
    'Printer.Print "Villahermosa" & " " & "Centro" & " " & "Tabasco"
    Printer.Print "TELS: " & RES3.Fields("SUC_TEL1") & " " & RES3.Fields("SUC_TEL2") & vbCrLf
'    Printer.FontSize = 12
    Printer.FontSize = RES3.Fields("SUC_TICKET_SIZE_2")
    Printer.FontBold = True
    Printer.Print "FECHA DE OPERACIÓN: "
    Printer.Print Format(RES4.Fields("vent_fechahora_cobro"), "dddd dd-mm-yyyy") & " " & Format(RES4.Fields("vent_fechahora_cobro"), "Short Time") & vbCrLf
    Printer.Print "FOLIO:   " & Format(folio, "0000000")
    If IsNull(RES4.Fields("MESA")) = False Then
        Printer.Print "MESA:      " & RES4.Fields("MESA") & ""
        Printer.Print "PERSONAS:  " & RES4.Fields("PERSONAS") & "" & vbCrLf
    End If
    Printer.FontBold = True
    Printer.Print "CLIENTE: "
    Printer.Print UCase(RES4.Fields("CLIENTE")) '& vbCrLf
    Printer.FontBold = False
    Printer.Print "MOSTRADOR: "
    Printer.Print UCase(RES4.Fields("CAJA"))
    Printer.Print "- - - - - - - - - - - - - - -"
    
    Printer.Print "- - - - - - - - - - - - - - - "
    If Len(Horario) > 0 Then
        Printer.FontSize = RES3.Fields("SUC_TICKET_SIZE_2")
        Printer.Print Horario
        Printer.Print "- - - - - - - - - - - - - - - "
    End If
    Printer.FontBold = True
    If IVA = "S" Then
        Printer.Print "  SUB TOTAL: " & FormatCurrency((total / (1.16)))
        Printer.Print "  DESCUENTO: " & FormatCurrency(DESCUENTO)
        Printer.Print "        IVA: " & FormatCurrency((total) - (total / (1.16)))
    Else
        Printer.Print "  SUB TOTAL: " & FormatCurrency((total) - (DESCUENTO))
        Printer.Print "  DESCUENTO: " & FormatCurrency(DESCUENTO)
    End If
    Printer.Print "      TOTAL: " & FormatCurrency(total)
    Printer.Print "- - - - - - - - - - - - - - - "
    Printer.Print "FORMA DE PAGO:"
    If Val(PAGOEFECTIVO) > 0 Then
    Printer.Print "   EFECTIVO: " & FormatCurrency(PAGOEFECTIVO)
    End If
    If Val(PAGOTARJETA) > 0 Then
    Printer.Print "    TARJETA: " & FormatCurrency(PAGOTARJETA)
    End If
    If Val(PAGOCHEQUE) > 0 Then
    Printer.Print "   MONEDERO: " & FormatCurrency(PAGOCHEQUE)
    End If
    Printer.Print "     PAGADO: " & FormatCurrency(PAGADO)
    Printer.Print "     CAMBIO: " & FormatCurrency(CAMBIO)
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    If Val(monedero) > 0 Then
    Printer.Print "  --- MONEDERO ---"
    Printer.Print "   RECIBIDO: " & FormatCurrency(mone_recibe)
    Printer.Print "   APLICADO: " & FormatCurrency(mone_entrega)
    Printer.Print "      TOTAL: " & FormatCurrency(monedero)
    End If
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    Printer.FontBold = True
    
    Printer.Print "           OBSERVACIONES    " & vbCrLf
    Printer.FontBold = False
    Printer.FontSize = RES3.Fields("SUC_TICKET_SIZE_1")
    If Len(OBSERVACIONES) > 0 Then
    
        Lineas = Split(OBSERVACIONES, vbNewLine)
    
        For b1 = 0 To UBound(Lineas)
            'Printer.Print UCase(Lineas(b1))
            
            For c1 = 0 To ((Round((Len(Lineas(b1)) / 35))) + 1)
                If Len(Lineas(b1)) >= 35 Then
                    'largo = True
                    Printer.Print Left(Lineas(b1), 35)
                    Lineas(b1) = Right(Lineas(b1), (Len(Lineas(b1)) - 35))
                Else
                    Printer.Print Lineas(b1)
                    Exit For
                End If
            Next c1
        Next b1
        
        Printer.Print "- - - - - - - - - - - - - - - - - - "
    End If
    'Call Centrar(Eslogan, 15)
    Printer.Print vbCrLf & UCase(RES3.Fields("SUC_SLOGAN")) & vbCrLf
    'Call Centrar(Web, 15)
    Printer.Print UCase(RES3.Fields("SUC_PAGINA_WEB"))
    Printer.Print UCase(RES3.Fields("SUC_EMAIL"))


    impresionRenglones (RES3.Fields("SUC_INFORMACION"))


'    Lineas = Split(RES3.Fields("SUC_INFORMACION"), vbNewLine)
'    For b1 = 0 To UBound(Lineas)
'        For c1 = 0 To ((Round((Len(Lineas(b1)) / 35))) + 1)
'            If Len(Lineas(b1)) >= 35 Then
'                'largo = True
'                Printer.Print Left(Lineas(b1), 35)
'                Lineas(b1) = Right(Lineas(b1), (Len(Lineas(b1)) - 35))
'            Else
'                Printer.Print Lineas(b1)
'                Exit For
'            End If
'        Next c1
'    Next b1
    'Printer.Print RES3.Fields("SUC_INFORMACION") & vbCrLf & vbCrLf
'    Printer.FontSize = 12
'    Printer.FontName = "Control"
'    Printer.Print "P"  'Cut
    Printer.EndDoc
    

End Sub

Public Sub nota_Mesa(folio As String)
    On Error Resume Next
    
    Dim sql1 As String
    Dim RES3 As Recordset
    Dim RES4 As Recordset
    Dim RES5 As Recordset

    Dim SUBTOTAL
    Dim DESCUENTO
    Dim total
    Dim MESA
    Dim PAGOEFECTIVO
    Dim PAGOTARJETA
    Dim PAGOCHEQUE
    Dim PAGADO
    Dim CAMBIO
   Dim OBSERVACIONES As String
    Dim Lineas() As String
    
    sql1 = "select * from SUCURSAL"
    Set RES3 = con.Execute(sql1)
    If Not RES3.EOF Then
        If RES3.Fields("SUC_ESTATUSTICKET") = 1 Then
            
        Else
            MsgBox "El status del ticket está desactivado.", vbInformation
            Exit Sub
        End If
    Else
        MsgBox "No se puede imprimir el ticket por que no tiene información referente a la sucursal del negocio. Verifique.", vbInformation
        Exit Sub
    End If
    
    sql1 = "SELECT T1.vent_fechahora, T1.VENT_MESA MESA, VENT_PERSONAS PERSONAS, T5.VENDET_NOTAMESA, T1.vent_fechahora_cobro, CONCAT(T2.PER_NOMBRE, ' ', T2.PER_PATERNO, ' ', T2.PER_MATERNO) CLIENTE, " & _
    "CONCAT(T3.PER_NOMBRE, ' ', T3.PER_PATERNO, ' ', T3.PER_MATERNO) USUARIO, CONCAT(T4.PER_NOMBRE, ' ', T4.PER_PATERNO, ' ', T4.PER_MATERNO) CAJA, " & _
    "T5.VENDET_PRODCODIGO, T5.VENDET_PRODUCTONOMBRE, T5.VENDET_PRECIO,  SUM(T5.VENDET_CANTIDAD) VENDET_CANTIDAD, T5.venDet_Descuento, " & _
    "T1.VENT_PAGOEFECTIVO, VENT_PAGOTARJETA, VENT_PAGOCHEQUE, VENT_PAGADO, VENT_OBSERVACIONES, VENT_CAMBIO, (SELECT SUM(VENDET_PRECIO * VENDET_CANTIDAD) FROM VENTA_DETALLE WHERE VENDET_FOLIO =  T1.VENT_IDFOLIO) SUBTOTAL, (SELECT SUM(VENDET_DESCUENTO) FROM VENTA_DETALLE WHERE VENDET_FOLIO =  T1.VENT_IDFOLIO) DESCUENTO, (SELECT ((SUM(VENDET_PRECIO * VENDET_CANTIDAD)) - SUM(VENDET_DESCUENTO)) FROM VENTA_DETALLE WHERE VENDET_FOLIO =  T1.VENT_IDFOLIO) TOTAL " & _
    "FROM VENTAS T1, PERSONA T2, PERSONA T3, VENTA_DETALLE T5, PERSONA T4 " & _
    "Where T1.VENT_IDFOLIO = '" & folio & "'  And T1.VENT_CLIEPERID = T2.PER_ID And T5.VENDET_VENDPERID = T3.PER_ID And T1.VENT_IDFOLIO = T5.VENDET_FOLIO And T1.VENT_VENDPERID = T4.PER_ID and T5.vendet_Status = 'A' " & _
    "GROUP BY T1.vent_fechahora, T1.VENT_MESA, T5.VENDET_NOTAMESA, T1.vent_fechahora_cobro, CONCAT(T2.PER_NOMBRE, ' ', T2.PER_PATERNO, ' ', T2.PER_MATERNO), " & _
    "CONCAT(T3.PER_NOMBRE, ' ', T3.PER_PATERNO, ' ', T3.PER_MATERNO) , CONCAT(T4.PER_NOMBRE, ' ', T4.PER_PATERNO, ' ', T4.PER_MATERNO) , " & _
    "T5.VENDET_PRODCODIGO , T5.VENDET_PRODUCTONOMBRE, T5.VENDET_PRECIO, T5.venDet_Descuento, T1.VENT_PAGOEFECTIVO, VENT_PAGOTARJETA, VENT_PAGOCHEQUE, VENT_PAGADO, VENT_OBSERVACIONES, VENT_CAMBIO "
    
    Set RES4 = con.Execute(sql1)
    If Not RES4.EOF Then
        SUBTOTAL = RES4.Fields("SUBTOTAL")
        DESCUENTO = RES4.Fields("DESCUENTO")
        total = RES4.Fields("TOTAL")
        PAGOEFECTIVO = RES4.Fields("VENT_PAGOEFECTIVO")
        PAGOTARJETA = RES4.Fields("VENT_PAGOTARJETA")
        PAGOCHEQUE = RES4.Fields("VENT_PAGOCHEQUE")
        PAGADO = RES4.Fields("VENT_PAGADO")
        CAMBIO = RES4.Fields("VENT_CAMBIO")
        OBSERVACIONES = RES4.Fields("VENT_OBSERVACIONES") & ""
        MESA = RES4.Fields("MESA")
    Else
        MsgBox "No se puede imprimir el pre-ticket por que no tiene información referente a la venta referida. " & vbCrLf & vbCrLf & "Posible causa: Falta de información en la lista. " & vbCrLf & vbCrLf & "Verifique.", vbInformation
        Exit Sub
    End If
    
    Printer.KillDoc
    'Printer.FontSize = 12
    'Printer.FontName = "Control"
    'Printer.Print "C"  'open Drawer 1 at 50ms
    Printer.Font = "Courier New"
    Printer.FontSize = 9
    Printer.FontBold = True
    
    'Call Centrar(Nombre, 12)
    'Printer.Print UCase(RES3.Fields("SUC_RAZON_SOCIAL"))
    
    'Call Centrar(Sucursal, 12)
    Printer.Print UCase(RES3.Fields("SUC_NOMBRE")) & vbCrLf
    Printer.FontSize = 9
    Printer.FontBold = False

'    Printer.Print UCase(RES3.Fields("SUC_DIR_Calle") & " " & RES3.Fields("SUC_DIR_NUM_EXT") & " " & RES3.Fields("SUC_DIR_NUM_INT"))
'    Printer.Print UCase(RES3.Fields("SUC_dIR_COLONIA") & " CP:" & RES3.Fields("SUC_DIR_Cp"))
'    Printer.Print UCase(RES1.Fields("SUC_DIR_CIUDAD")) '& " " & RES1.Fields("Municipio")
'    Printer.Print "Villahermosa" & " " & "Centro" & " " & "Tabasco"
'    Printer.Print "TELS: " & RES3.Fields("SUC_TEL1"); " " & RES3.Fields("SUC_TEL2") & vbCrLf
    Printer.FontSize = 9
    Printer.FontBold = True
    Printer.Print "NOTA COCINA - MESA:  " & MESA & vbCrLf
    Printer.Print "FECHA DE OPERACIÓN: "
    Printer.Print Format(RES4.Fields("vent_fechahora"), "dddd dd-mm-yyyy") & " " & Format(RES4.Fields("vent_fechahora"), "Short Time") & vbCrLf
    Printer.Print "FOLIO: " & Format(folio, "0000000")
    If IsNull(RES4.Fields("MESA")) = False Then
        Printer.Print "MESA:      " & RES4.Fields("MESA") & ""
        Printer.Print "PERSONAS:  " & RES4.Fields("PERSONAS") & "" & vbCrLf
    End If
    Printer.FontSize = 9
    Printer.FontBold = True
    Printer.Print "CLIENTE: "
    Printer.Print RES4.Fields("CLIENTE") & vbCrLf
    Printer.FontBold = False
    Printer.Print "MOSTRADOR: "
    Printer.Print RES4.Fields("CAJA")
    Printer.Print "- - - - - - - - - - - - - - -"
    Printer.Print "DETALLE DE OPERACIÓN:" & vbCrLf
    'For b1 = 1 To FrmFocus.ListaOper.Rows - 1
    Do While Not RES4.EOF
        If RES4.Fields("VENDET_NOTAMESA") <> "A" Or IsNull(RES4.Fields("VENDET_NOTAMESA")) = True Then
            clave = RES4.Fields("VENDET_PRODCODIGO")
            If Len(clave) > 17 Then
                clave = Left(clave, 17)
            Else
                clave = clave & String(17 - Len(clave), " ")
            End If
            Nombre2 = RES4.Fields("VENDET_PRODUCTONOMBRE")
            If Len(Nombre2) > 28 Then
                Nombre2 = Left(Nombre2, 28)
            Else
                Nombre2 = Nombre2 & String(28 - Len(Nombre2), " ")
            End If
            cantidad = RES4.Fields("VENDET_CANTIDAD")
            Precio = RES4.Fields("VENDET_PRECIO")
            desc = RES4.Fields("venDet_Descuento")
            tot = (RES4.Fields("VENDET_CANTIDAD") * RES4.Fields("VENDET_PRECIO") - RES4.Fields("VENDET_DESCUENTO"))
            If Len(Precio) > 9 Then
                Precio = Left(Precio, 9)
            End If
            If Len(RES4.Fields("USUARIO")) > 28 Then
                atendio = Left(RES4.Fields("USUARIO"), 28)
            Else
                atendio = RES4.Fields("USUARIO")
            End If
                    
            cantidad = cantidad & String(10 - Len(cantidad), " ")
            Precio = Precio & String(12 - Len(Precio), " ")
            desc = FormatCurrency(desc) & String(11 - Len(FormatCurrency(desc)), " ")
    '        Tot = Tot & String(9 - Len(Tot), " ")
            titulo1 = "Cant" & String(10 - Len("Cant"), " ")
            titulo2 = "Precio" & String(10, " ")
            titulo3 = "Desc" & String(6, " ")
            titulo4 = "Total" & String(4, " ")
            titulo5 = "Atendio" & String(12, " ")
            
            Printer.FontBold = True
            Printer.Print UCase(clave)
            Printer.Print UCase(titulo1) & " " & cantidad
            Printer.Print UCase(Nombre2)
            'Printer.FontBold = False
    '        Printer.FontBold = True
    '        Printer.Print cantidad & " " & FormatCurrency(Precio)
    '        Printer.FontBold = False
    '        Printer.Print Titulo3 & " " & Titulo4
    '        Printer.FontBold = True
    '        Printer.Print desc & FormatCurrency(Tot)
            Printer.FontBold = False
'            Printer.Print Titulo5 & vbCrLf & UCase(Atendio) & vbCrLf
        End If
    RES4.MoveNext
    Loop
    
    
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    Printer.FontBold = True
    If Len(OBSERVACIONES) > 0 Then
        Printer.Print "           OBSERVACIONES    "
        Printer.FontSize = 10
        Lineas = Split(OBSERVACIONES, vbNewLine)
    
        For b1 = 0 To UBound(Lineas)
            For c1 = 0 To ((Round((Len(Lineas(b1)) / 35))) + 1)
                If Len(Lineas(b1)) >= 35 Then
                    Printer.Print Left(Lineas(b1), 35)
                    Lineas(b1) = Right(Lineas(b1), (Len(Lineas(b1)) - 35))
                Else
                    Printer.Print Lineas(b1)
                    Exit For
                End If
            Next c1
        Next b1
    End If
        
    Printer.FontSize = 9
    Printer.FontBold = False
    
    
    
    Printer.FontBold = True

    Printer.Print "- - - - - - - - - - - - - - - - - - "
    
    Printer.Print vbCrLf & "TICKET COCINA - MESA: " & MESA
    'Printer.Print RES3.Fields("SUC_INFORMACION") & vbCrLf & vbCrLf
'    Printer.FontSize = 12
'    Printer.FontName = "Control"
'    Printer.Print "P"  'Cut
    Printer.EndDoc
    

End Sub
Public Sub nota_Cocina(folio As String, tipo As String)
    On Error Resume Next
    
    Dim sql1 As String
    Dim RES3 As Recordset
    Dim RES4 As Recordset
    Dim RES5 As Recordset
    Dim resPrinter As Recordset
    Dim impresora, impresora2 As String

    Dim SUBTOTAL
    Dim DESCUENTO
    Dim total, cantidad, Precio, desc, tot, b1, c1
    Dim MESA
    Dim PAGOEFECTIVO
    Dim PAGOTARJETA
    Dim PAGOCHEQUE
    Dim PAGADO
    Dim CAMBIO
    Dim OBSERVACIONES As String
    Dim Lineas() As String
    Dim tiempo As String
    Dim clave, Nombre2, atendio, titulo1, titulo2, titulo3, titulo4, titulo5 As String
    Dim prt As Printer
    
    sql1 = "select * from SUCURSAL"
    Set RES3 = con.Execute(sql1)
    If Not RES3.EOF Then
        If RES3.Fields("SUC_ESTATUSTICKET") = 1 Then
            
        Else
            MsgBox "El status del ticket está desactivado.", vbInformation
            Exit Sub
        End If
    Else
        MsgBox "No se puede imprimir el ticket por que no tiene información referente a la sucursal del negocio. Verifique.", vbInformation
        Exit Sub
    End If
    
    sql1 = "SELECT T3.CTPT_IMPRESORA FROM VENTA_DETALLE T1, PRODUCTOS T2, CAT_TIPO T3 " & _
    "Where T3.CTPT_ID = T2.PROD_TIPO And T3.CTPT_SUBTIPO = T2.PROD_SUBTIPO And T2.PROD_ID = T1.VENDET_PRODUCTOID " & _
    "AND T1.VENDET_FOLIO = '" & folio & "'  GROUP BY T3.CTPT_IMPRESORA "
    Set resPrinter = con.Execute(sql1)
    
    Do While Not resPrinter.EOF
        impresora = resPrinter.Fields("CTPT_IMPRESORA")
        impresora2 = resPrinter.Fields("CTPT_IMPRESORA")
        If Left(impresora, 1) = "\" Then
            impresora = Replace(impresora, "\", "\\")
        End If
        'MsgBox impresora
        sql1 = "SELECT T1.vent_fechahora, T1.VENT_MESA MESA, T5.VENDET_NOTAMESA, T1.vent_fechahora_cobro, CONCAT(T2.PER_NOMBRE, ' ', T2.PER_PATERNO, ' ', T2.PER_MATERNO) CLIENTE, " & _
        "CONCAT(T3.PER_NOMBRE, ' ', T3.PER_PATERNO, ' ', T3.PER_MATERNO) USUARIO, CONCAT(T4.PER_NOMBRE, ' ', T4.PER_PATERNO, ' ', T4.PER_MATERNO) CAJA, " & _
        "T5.VENDET_PRODCODIGO, T5.VENDET_STATUS, T5.VENDET_PRODUCTONOMBRE, T5.VENDET_PRECIO,  SUM(T5.VENDET_CANTIDAD) VENDET_CANTIDAD, T5.venDet_Descuento, T5.VENDET_DESCRIPCION, T5.VENDET_TIEMPO, T6.CTPT_IMPRESORA,  " & _
        "T1.VENT_PAGOEFECTIVO, T5.VENDET_ASIENTO ASIENTO, VENT_PAGOTARJETA, VENT_PAGOCHEQUE, VENT_PAGADO, VENT_OBSERVACIONES, VENT_CAMBIO, (SELECT SUM(VENDET_PRECIO * VENDET_CANTIDAD) FROM VENTA_DETALLE WHERE VENDET_FOLIO =  T1.VENT_IDFOLIO) SUBTOTAL, (SELECT SUM(VENDET_DESCUENTO) FROM VENTA_DETALLE WHERE VENDET_FOLIO =  T1.VENT_IDFOLIO) DESCUENTO, (SELECT ((SUM(VENDET_PRECIO * VENDET_CANTIDAD)) - SUM(VENDET_DESCUENTO)) FROM VENTA_DETALLE WHERE VENDET_FOLIO =  T1.VENT_IDFOLIO) TOTAL " & _
        "FROM VENTAS T1, PERSONA T2, PERSONA T3, VENTA_DETALLE T5, PERSONA T4, CAT_TIPO T6, PRODUCTOS T7 " & _
        "Where T1.VENT_IDFOLIO = '" & folio & "'  And T1.VENT_CLIEPERID = T2.PER_ID And T5.VENDET_VENDPERID = T3.PER_ID And T1.VENT_IDFOLIO = T5.VENDET_FOLIO And T1.VENT_VENDPERID = T4.PER_ID AND T6.CTPT_ID = T7.PROD_TIPO AND T6.CTPT_SUBTIPO = T7.PROD_SUBTIPO AND T7.PROD_ID = T5.VENDET_PRODUCTOID AND T6.CTPT_IMPRESORA =  '" & impresora & "'" & _
        "GROUP BY T1.vent_fechahora, T1.VENT_MESA, T5.VENDET_NOTAMESA, T1.vent_fechahora_cobro, CONCAT(T2.PER_NOMBRE, ' ', T2.PER_PATERNO, ' ', T2.PER_MATERNO), " & _
        "CONCAT(T3.PER_NOMBRE, ' ', T3.PER_PATERNO, ' ', T3.PER_MATERNO) , CONCAT(T4.PER_NOMBRE, ' ', T4.PER_PATERNO, ' ', T4.PER_MATERNO) , " & _
        "T5.VENDET_PRODCODIGO , T5.VENDET_PRODUCTONOMBRE, T5.VENDET_PRECIO, T5.venDet_Descuento, T1.VENT_PAGOEFECTIVO, VENT_PAGOTARJETA, VENT_PAGOCHEQUE, VENT_PAGADO, VENT_OBSERVACIONES, VENT_CAMBIO, T6.CTPT_IMPRESORA, T5.VENDET_TIEMPO, VENDET_DESCRIPCION, VENDET_ASIENTO order BY T5.VENDET_TIEMPO, VENDET_ASIENTO  ASC "
        
        Set RES4 = con.Execute(sql1)
        If Not RES4.EOF Then
            SUBTOTAL = RES4.Fields("SUBTOTAL")
            DESCUENTO = RES4.Fields("DESCUENTO")
            total = RES4.Fields("TOTAL")
            PAGOEFECTIVO = RES4.Fields("VENT_PAGOEFECTIVO")
            PAGOTARJETA = RES4.Fields("VENT_PAGOTARJETA")
            PAGOCHEQUE = RES4.Fields("VENT_PAGOCHEQUE")
            PAGADO = RES4.Fields("VENT_PAGADO")
            CAMBIO = RES4.Fields("VENT_CAMBIO")
            OBSERVACIONES = RES4.Fields("VENT_OBSERVACIONES") & ""
            MESA = RES4.Fields("MESA")
        Else
            MsgBox "No se puede imprimir el pre-ticket por que no tiene información referente a la venta referida. " & vbCrLf & vbCrLf & "Posible causa: Falta de información en la lista. " & vbCrLf & vbCrLf & "Verifique.", vbInformation
            Exit Sub
        End If
        
        'If IsNull(RES4.Fields("CTPT_IMPRESORA")) = False Then
            For Each prt In Printers
                If prt.DeviceName = impresora2 Then
                    Set Printer = prt
                End If
            Next
        'End If
        
        Printer.KillDoc
        
        Printer.Font = "Courier New"
        Printer.FontSize = 9
        Printer.FontBold = True
        
        Printer.Print UCase(RES3.Fields("SUC_NOMBRE")) & vbCrLf
        Printer.FontSize = 9
        Printer.FontBold = False
    
        Printer.FontSize = 9
        Printer.FontBold = True
        Printer.Print "NOTA COCINA - MESA:  " & MESA & vbCrLf
        Printer.Print "FECHA DE OPERACIÓN: "
        Printer.Print Format(RES4.Fields("vent_fechahora"), "dddd dd-mm-yyyy") & " " & Format(RES4.Fields("vent_fechahora"), "Short Time") & vbCrLf
        Printer.Print "FOLIO: " & Format(folio, "0000000") & vbCrLf
        Printer.FontSize = 10
        Printer.FontBold = True
        Printer.Print "FECHA/HORA DE IMPRESION: "
        Printer.Print Format(Date, "dddd dd-mm-yyyy")
        Printer.Print Format(Time, "Short Time") & vbCrLf
        Printer.FontSize = 9
        Printer.FontBold = True
        Printer.Print "CLIENTE: "
        Printer.Print RES4.Fields("CLIENTE") & vbCrLf
        Printer.FontBold = False
        Printer.Print "MOSTRADOR: "
        Printer.Print RES4.Fields("CAJA")
        Printer.Print "- - - - - - - - - - - - - - -"
        Printer.Print "DETALLE DE OPERACIÓN:" & vbCrLf
        tiempo = ""
        ASIENTO = ""
        
        Do While Not RES4.EOF
                    
            clave = RES4.Fields("VENDET_PRODCODIGO")
            If Len(clave) > 17 Then
                clave = Left(clave, 17)
            Else
                clave = clave & String(17 - Len(clave), " ")
            End If
            Nombre2 = RES4.Fields("VENDET_PRODUCTONOMBRE")
            cantidad = RES4.Fields("VENDET_CANTIDAD")
            Precio = RES4.Fields("VENDET_PRECIO")
            desc = RES4.Fields("venDet_Descuento")
            tot = (RES4.Fields("VENDET_CANTIDAD") * RES4.Fields("VENDET_PRECIO") - RES4.Fields("VENDET_DESCUENTO"))
            If Len(Precio) > 9 Then
                Precio = Left(Precio, 9)
            End If
            If Len(RES4.Fields("USUARIO")) > 28 Then
                atendio = Left(RES4.Fields("USUARIO"), 28)
            Else
                atendio = RES4.Fields("USUARIO")
            End If
                    
            cantidad = cantidad & String(4 - Len(cantidad), " ")
            Precio = Precio & String(12 - Len(Precio), " ")
            desc = FormatCurrency(desc) & String(11 - Len(FormatCurrency(desc)), " ")
    '        Tot = Tot & String(9 - Len(Tot), " ")
            titulo1 = "Cant" & String(10 - Len("Cant"), " ")
            titulo2 = "Precio" & String(10, " ")
            titulo3 = "Desc" & String(6, " ")
            titulo4 = "Total" & String(4, " ")
            titulo5 = "Atendio" & String(12, " ")
            
            Printer.FontBold = True
            
            If (RES4.Fields("VENDET_NOTAMESA") <> "A" Or IsNull(RES4.Fields("VENDET_NOTAMESA")) = True) And tipo = "GENERAL" Then

                If tiempo <> RES4.Fields("VENDET_TIEMPO") Then

                    Printer.Print "----------------------------------------"
                    Printer.FontSize = 11
                    Printer.FontBold = True
                    Printer.Print "Tiempo: " & RES4.Fields("VENDET_TIEMPO")
                    tiempo = RES4.Fields("VENDET_TIEMPO")
                Else
                End If
                Printer.FontSize = 9
                Printer.FontBold = True
                Printer.Print "Asiento: " & RES4.Fields("ASIENTO")
                impresionRenglones (cantidad & " " & Nombre2)
                
                Printer.FontBold = False
                impresionRenglones (RES4.Fields("VENDET_DESCRIPCION"))
                Printer.FontBold = False
            Else
                If RES4.Fields("VENDET_NOTAMESA") = "A" And RES4.Fields("VENDET_STATUS") = "C" And tipo = "CANCEL" Then
    
                    If tiempo <> RES4.Fields("VENDET_TIEMPO") Then
    
                        Printer.Print "----------------------------------------"
                        Printer.FontSize = 11
                        Printer.FontBold = True
                        Printer.Print "Tiempo: " & RES4.Fields("VENDET_TIEMPO")
                        tiempo = RES4.Fields("VENDET_TIEMPO")
                    Else
                    End If
                    Printer.FontSize = 9
                    Printer.FontBold = True
                    Printer.Print "----CANCELACION----"
                    
                    Printer.Print "Asiento: " & RES4.Fields("ASIENTO")
                    impresionRenglones (cantidad & " " & Nombre2)
                    
                    Printer.FontBold = False
                    impresionRenglones (RES4.Fields("VENDET_DESCRIPCION"))
                    Printer.Print "----CANCELACION----"
                    Printer.FontBold = False
                
                End If
            End If
        RES4.MoveNext
        Loop
        
        
        Printer.Print "- - - - - - - - - - - - - - - - - - "
        Printer.FontBold = True
        If Len(OBSERVACIONES) > 0 Then
            Printer.Print "           OBSERVACIONES    "
            Printer.FontSize = 10
            Lineas = Split(OBSERVACIONES, vbNewLine)
        
            For b1 = 0 To UBound(Lineas)
                For c1 = 0 To ((Round((Len(Lineas(b1)) / 35))) + 1)
                    If Len(Lineas(b1)) >= 35 Then
                        Printer.Print Left(Lineas(b1), 35)
                        Lineas(b1) = Right(Lineas(b1), (Len(Lineas(b1)) - 35))
                    Else
                        Printer.Print Lineas(b1)
                        Exit For
                    End If
                Next c1
            Next b1
        End If
            
        Printer.FontSize = 9
        Printer.FontBold = False
        
        
        
        Printer.FontBold = True
    
        Printer.Print "- - - - - - - - - - - - - - - - - - "
        
        Printer.Print vbCrLf & "TICKET COCINA - MESA: " & MESA
        'Printer.Print RES3.Fields("SUC_INFORMACION") & vbCrLf & vbCrLf
    '    Printer.FontSize = 12
    '    Printer.FontName = "Control"
    '    Printer.Print "P"  'Cut
        Printer.EndDoc
'        RES4.MoveNext
'        Loop
    resPrinter.MoveNext
    Loop

End Sub





Public Sub notaGasto(folio As String)
    'On Error Resume Next
    
    Dim sql1 As String
    Dim RES3 As Recordset
    Dim RES4 As Recordset
    Dim resGasto As Recordset

    Dim SUBTOTAL
    Dim DESCUENTO
    Dim total
    Dim PAGOEFECTIVO
    Dim PAGOTARJETA
    Dim PAGOCHEQUE
    Dim PAGADO
    Dim CAMBIO
    Dim Lineas() As String
    
    sql1 = "select * from SUCURSAL"
    Set RES3 = con.Execute(sql1)
    If Not RES3.EOF Then
        If RES3.Fields("SUC_ESTATUSTICKET") = 1 Then
            
        Else
            MsgBox "El status del ticket está desactivado.", vbInformation
            Exit Sub
        End If
    Else
        MsgBox "No se puede imprimir el ticket por que no tiene información referente a la sucursal del negocio. Verifique.", vbInformation
        Exit Sub
    End If
    
    sql1 = "SELECT * FROM VIEW_GASTOS WHERE ID = '" & folio & "'"
    Set resGasto = con.Execute(sql1)
    
    Printer.KillDoc
    Printer.FontSize = 12
    Printer.Font = "Courier New"
    Printer.FontSize = 10
    Printer.FontBold = True
    
    Printer.Print RES3.Fields("SUC_RAZON_SOCIAL")
    Printer.Print RES3.Fields("SUC_NOMBRE") & vbCrLf
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.Print RES3.Fields("SUC_DIR_Calle") & " " & RES3.Fields("SUC_DIR_NUM_EXT") & " " & RES3.Fields("SUC_DIR_NUM_INT")
    Printer.Print RES3.Fields("SUC_dIR_COLONIA") & " CP:" & RES3.Fields("SUC_DIR_Cp")
'    Printer.Print RES1.Fields("Estado") & " " & RES1.Fields("Municipio")
    Printer.Print "Villahermosa" & " " & "Centro" & " " & "Tabasco"
    Printer.Print RES3.Fields("SUC_TEL1") & "" & " " & RES3.Fields("SUC_TEL2") & ""; vbCrLf
    Printer.Print "Fecha de registro"
    Printer.Print Format(resGasto.Fields("registro"), "dddd dd-mm-yyyy") & " " & Format(resGasto.Fields("registro"), "Short Time") & vbCrLf
    Printer.Print "Fecha de Inicio - Fecha de Fin"
    Printer.Print Format(resGasto.Fields("fecha_hora"), "Short Date") & " " & Format(resGasto.Fields("fecha_fin"), "Short Date") & vbCrLf
    Printer.Print "FOLIO: " & Format(folio, "0000000")
    Printer.Print "Atendió: "
    Printer.Print resGasto.Fields("USUARIO")
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    Printer.Print "Tipo de gasto: "
    Printer.Print resGasto.Fields("TIPO_GASTO")
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    Printer.Print "Descripción: "
    Printer.Print resGasto.Fields("GST_DESCRIPCION")
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    Printer.Print "Total gasto:   "
    Printer.Print FormatCurrency(resGasto.Fields("GASTO"))
    Printer.Print "Comprobante de gasto:  " & resGasto.Fields("COMPROBANTE")
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    'Call Centrar(Eslogan, 15)
    Printer.Print vbCrLf & RES3.Fields("SUC_SLOGAN") & vbCrLf
    'Call Centrar(Web, 15)
    Printer.Print RES3.Fields("SUC_PAGINA_WEB")
    Printer.Print RES3.Fields("SUC_EMAIL")


    impresionRenglones (RES3.Fields("SUC_INFORMACION"))

'    Lineas = Split(RES3.Fields("SUC_INFORMACION"), vbNewLine)
'    For b1 = 1 To UBound(Lineas)
'        Printer.Print Lineas(b1)
'    Next b1

    Printer.EndDoc
    

End Sub


Public Sub notaUsuario(fila As Long)
    'On Error Resume Next
    
    Dim sql1 As String
    Dim RES3 As Recordset
    Dim RES4 As Recordset
    Dim RES5 As Recordset

    Dim SUBTOTAL
    Dim DESCUENTO
    Dim total
    Dim PAGOEFECTIVO
    Dim PAGOTARJETA
    Dim PAGOCHEQUE
    Dim PAGADO
    Dim CAMBIO
    Dim Honorarios As Double
    Dim Comision As Double
    Dim ConsumoI As Double
    Dim totalVenta As Double
    
    sql1 = "select * from SUCURSAL"
    Set RES3 = con.Execute(sql1)
    If Not RES3.EOF Then
    Else
        MsgBox "No se puede imprimir el ticket por que no tiene información referente a la sucursal del negocio. Verifique.", vbInformation
        Exit Sub
    End If
    Honorarios = 0
    Comision = 0
    ConsumoI = 0
    
    totalVenta = Val(Format(FRM_Caja.lista2.TextMatrix(fila, 6), "General Number"))
    
    For b1 = 1 To FRM_Caja.listaPagos.Rows - 1
        If FRM_Caja.listaPagos.TextMatrix(b1, 1) = "DEDUCCION" Then
            ConsumoI = ConsumoI + Val(Format(FRM_Caja.listaPagos.TextMatrix(b1, 4), "General Number"))
        End If
    Next b1
    For b1 = 1 To FRM_Caja.listaPagos.Rows - 1
        If FRM_Caja.listaPagos.TextMatrix(b1, 1) = "HONORARIOS" Then
            Honorarios = Honorarios + Val(Format(FRM_Caja.listaPagos.TextMatrix(b1, 4), "General Number"))
        Else
            If FRM_Caja.listaPagos.TextMatrix(b1, 1) = "COMISION" Then
                Comision = Comision + ((Val(Format(FRM_Caja.listaPagos.TextMatrix(b1, 2), "General Number")) * (0.01)) * (totalVenta - ConsumoI))
            Else
            End If
        End If

    Next b1
    
    Printer.KillDoc
'    Printer.FontSize = 12
'    Printer.FontName = "Control"
'    Printer.Print "C"  'open Drawer 1 at 50ms
    Printer.Font = "Courier New"
    Printer.FontSize = 10
    Printer.FontBold = True
    
    'Call Centrar(Nombre, 12)
    Printer.Print RES3.Fields("SUC_RAZON_SOCIAL")
    
    'Call Centrar(Sucursal, 12)
    Printer.Print RES3.Fields("SUC_NOMBRE") & vbCrLf
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.Print RES3.Fields("SUC_DIR_Calle") & " " & RES3.Fields("SUC_DIR_NUM_EXT") & " " & RES3.Fields("SUC_DIR_NUM_INT")
    Printer.Print RES3.Fields("SUC_dIR_COLONIA") & " CP:" & RES3.Fields("SUC_DIR_Cp")
    Printer.Print "Villahermosa" & " " & "Centro" & " " & "Tabasco"
    Printer.Print RES3.Fields("SUC_TEL1") & "" & " " & RES3.Fields("SUC_TEL2") & ""; vbCrLf
    Printer.Print Format(Date, "dddd dd-mm-yyyy") & " " & Format(Time, "Short Time") & vbCrLf
    Printer.Print "Periodo del " & FRM_Caja.dtFecha1(0) & " al " & FRM_Caja.dtFecha1(1)
    Printer.Print "Corte del día por usuario"
    Printer.Print "Usuario: "
    Printer.Print FRM_Menu.menuBarra2.Panels(5).Text & vbCrLf
    Printer.Print "Generó: "
    Printer.Print FRM_Caja.lista2.TextMatrix(fila, 0)
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    Printer.Print "Productos" & "       " & "Cantidad"
    Printer.Print FRM_Caja.lista2.TextMatrix(fila, 2) & String(16 - (Len(FRM_Caja.lista2.TextMatrix(fila, 2))), " ") & FRM_Caja.lista2.TextMatrix(fila, 3)
    Printer.Print "Servicios" & "       " & "Cantidad"
    Printer.Print FRM_Caja.lista2.TextMatrix(fila, 4) & String(16 - (Len(FRM_Caja.lista2.TextMatrix(fila, 4))), " ") & FRM_Caja.lista2.TextMatrix(fila, 5)
    Printer.Print "Apartados" & "       " & "Cantidad"
    Printer.Print FRM_Caja.lista2.TextMatrix(fila, 6) & String(16 - (Len(FRM_Caja.lista2.TextMatrix(fila, 6))), " ")
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    Printer.FontSize = 10
    Printer.Print "     TOTAL VENTA:  " & FormatCurrency(FRM_Caja.lista2.TextMatrix(fila, 7))
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    Printer.Print "    DEDUCCION CI:  " & FormatCurrency(ConsumoI) & vbCrLf
    Printer.Print "   TOTAL COMSION:  " & FormatCurrency(Comision) & vbCrLf
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    Printer.Print "PAGO HONORARIOS*:  " & FormatCurrency(Honorarios)
    Printer.Print "El pago de honorarios aplica al "
    Printer.Print "periodo que corresponbda. " & vbCrLf & "Se menciona como referencia" & vbCrLf
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    Printer.Print "Recibio: " & vbCrLf & Vbcrklf & vbclrf
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    Printer.Print vbCrLf & RES3.Fields("SUC_SLOGAN") & vbCrLf
'    Printer.FontSize = 12
'    Printer.FontName = "Control"
'    Printer.Print "P"  'Cut
    Printer.EndDoc
    

End Sub

Public Sub notaApartado(folio As String)
'    On Error Resume Next
    
    Dim sql1 As String
    Dim ResAprt As Recordset
    Dim ResAprt2 As Recordset
    Dim ResPago As Recordset
    Dim resSucur As Recordset
    
    Dim SUBTOTAL
    Dim DESCUENTO
    Dim total
    Dim PAGOEFECTIVO
    Dim PAGOTARJETA
    Dim PAGOCHEQUE
    Dim PAGADO
    Dim CAMBIO
    
    Dim APRT_FOLIO
    Dim PAGOS
    Dim FALTANTE
    Dim DIAS_LIQUI
    Dim DIAS_TRANS
    Dim monto
    Dim FechaApartado
    
    sql1 = "select * from SUCURSAL"
    Set resSucur = con.Execute(sql1)
    If Not resSucur.EOF Then
        If resSucur.Fields("SUC_ESTATUSTICKET") = 1 Then
            
        Else
            MsgBox "El status del ticket está desactivado.", vbInformation
            Exit Sub
        End If
    Else
        MsgBox "No se puede imprimir el ticket por que no tiene información referente a la sucursal del negocio. Verifique.", vbInformation
        Exit Sub
    End If


    
sql1 = "SELECT T2.FOLIO_VENTA, T2.FOLIO_APRT, T1.FECHA FECHA_APARTADO, T2.FECHA FECHA_PAGO, T1.CLIENTE, T1.VENDEDOR, T2.MOSTRADOR, T1.PRODUCTO, T1.CODIGO, " & _
"T1.PRECIO, T1.CANTIDAD, T1.DESCUENTO PROD_dESC, T1.TOTAL_PROD, T2.SUBTOTAL, T2.DESCUENTO, T2.TOTAL, " & _
"T2.PAGADO, T2.CAMBIO, T2.EFECTIVO, T2.TARJETA, T2.CHEQUE, T1.TOTAL TOTAL_ADEUDO, T2.PAGO,  " & _
"T1.PAGADO TOTAL_PAGADO, T1.ADEUDO TOTAL_FALTANTE, T1.DIAS, T1.TRANSCURRIDOS, T1.LIQUIDACION, (SELECT T41.TOTAL FROM VIEW_MONEDERO_CLIENTES T41 WHERE T1.aprt_clieperid = T41.PER_ID) MONEDERO, (SELECT SUM(MONEDERO) FROM VIEW_PUNTOS_ADMIN WHERE FOLIO = '" & folio & "' AND TIPO = 'RECIBE') MONE_RECIBE, (SELECT SUM(MONEDERO) FROM VIEW_PUNTOS_ADMIN WHERE FOLIO = '" & folio & "' AND TIPO = 'ENTREGA') MONE_ENTREGA,  " & _
"T1.tel1 , T1.tel2, T1.email, t1.FOLIO FROM VIEW_APARTADOS T1, VIEW_PAGOS_APARTTOTAL T2 " & _
"Where T1.folio = T2.FOLIO_APRT AND T2.FOLIO_VENTA = '" & folio & "'"
    
Set ResAprt = con.Execute(sql1)
Set ResAprt2 = con.Execute(sql1)
    
PAGOS = 0
FALTANTE = 0
monto = 0
monedero = 0

    If Not ResAprt.EOF Then
        monto = monto + Val(ResAprt2.Fields("TOTAL_ADEUDO"))
        PAGOS = PAGOS + Val(ResAprt2.Fields("TOTAL_PAGADO"))
        FALTANTE = FALTANTE + Val(ResAprt2.Fields("TOTAL_FALTANTE"))
        FechaApartado = ResAprt2.Fields("FECHA_APARTADO")
        
        SUBTOTAL = ResAprt.Fields("SUBTOTAL")
        DESCUENTO = ResAprt.Fields("DESCUENTO")
        total = ResAprt.Fields("TOTAL")
        PAGOEFECTIVO = ResAprt.Fields("EFECTIVO")
        PAGOTARJETA = ResAprt.Fields("TARJETA")
        PAGOCHEQUE = ResAprt.Fields("CHEQUE")
        PAGADO = ResAprt.Fields("PAGADO")
        CAMBIO = ResAprt.Fields("CAMBIO")
        APRT_FOLIO = ResAprt.Fields("FOLIO_APRT")
        VENT_FOLIO = ResAprt.Fields("FOLIO_VENTA")
        DIAS_LIQUI = ResAprt.Fields("DIAS")
        DIAS_TRANS = ResAprt.Fields("TRANSCURRIDOS")
        FECHA_LIQUI = Format(ResAprt.Fields("LIQUIDACION"), "Short Date")
        tel1 = ResAprt.Fields("TEL1")
        tel2 = ResAprt.Fields("TEL2")
        email = ResAprt.Fields("email")
        folio = ResAprt.Fields("FOLIO")
        pago = ResAprt.Fields("PAGO")
        'OBSERVACIONES = RES4.Fields("SUC_INFORMACION")
        
        If IsNull(ResAprt.Fields("MONEDERO")) Then
            monedero = 0
        Else
            monedero = Val(ResAprt.Fields("MONEDERO"))
        End If
        If IsNull(ResAprt.Fields("MONE_RECIBE")) Then
            mone_recibe = 0
        Else
            mone_recibe = Val(ResAprt.Fields("MONE_RECIBE"))
        End If
        If IsNull(ResAprt.Fields("MONE_ENTREGA")) Then
            mone_entrega = 0
        Else
            mone_entrega = Val(ResAprt.Fields("MONE_ENTREGA"))
        End If
    Else
        MsgBox "No se puede imprimir el ticket por que no tiene información referente al cobro referido. Verifique.", vbInformation
        Exit Sub
    End If
    
    Printer.KillDoc

    Printer.Font = "Courier New"
    Printer.FontSize = resSucur.Fields("SUC_TICKET_SIZE_1")
    Printer.FontBold = True
    
    Printer.Print resSucur.Fields("SUC_RAZON_SOCIAL")
    
    Printer.Print resSucur.Fields("SUC_NOMBRE") & vbCrLf
    Printer.Print resSucur.Fields("SUC_DIR_Calle") & " " & resSucur.Fields("SUC_DIR_NUM_EXT") & " " & resSucur.Fields("SUC_DIR_NUM_INT")
    Printer.Print resSucur.Fields("SUC_dIR_COLONIA") & " CP:" & resSucur.Fields("SUC_DIR_Cp")
'    Printer.Print RES1.Fields("Estado") & " " & RES1.Fields("Municipio")
    Printer.Print "Villahermosa" & " " & "Centro" & " " & "Tabasco"
    Printer.Print resSucur.Fields("SUC_TEL1") & "" & " " & resSucur.Fields("SUC_TEL2") & ""; vbCrLf
    Printer.Print Format(ResAprt.Fields("FECHA_PAGO"), "dddd dd-mm-yyyy") & " " & Format(ResAprt.Fields("FECHA_PAGO"), "Short Time") & vbCrLf
    Printer.Print "FOLIO: " & Format(folio, "0000000")
'    Printer.FontSize = resSucur.Fields("SUC_TICKET_SIZE_2")
    Printer.Print "CLIENTE: "
    Printer.Print ResAprt.Fields("CLIENTE") & vbCrLf
    Printer.Print "COBRO: "
    Printer.Print ResAprt.Fields("MOSTRADOR")
    Printer.Print "PAGO DE APARTADO"
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    'For b1 = 1 To FrmFocus.ListaOper.Rows - 1
    numpago = 0
    Do While Not ResAprt.EOF
        numpago = numpago + 1
        clave = ResAprt.Fields("CODIGO")
        If Len(clave) > 8 Then
            clave = Left(clave, 8)
        Else
            clave = clave & String(8 - Len(clave), " ")
        End If
        Nombre2 = ResAprt.Fields("PRODUCTO")
        If Len(Nombre2) > 16 Then
            Nombre2 = Left(Nombre2, 16)
        Else
            Nombre2 = Nombre2 & String(16 - Len(Nombre2), " ")
        End If
        cantidad = ResAprt.Fields("CANTIDAD")
        Precio = ResAprt.Fields("PRECIO")
        desc = ResAprt.Fields("PROD_DESC")
        tot = ResAprt.Fields("TOTAL_PROD") '* ResAprt.Fields("APRT_pRODPRECIO") - ResAprt.Fields("APRT_DESC")
        If Len(Precio) > 9 Then
            Precio = Left(Precio, 9)
        End If
        If Len(ResAprt.Fields("VENDEDOR")) > 30 Then
            atendio = Left(ResAprt.Fields("VENDEDOR"), 30)
        Else
            atendio = ResAprt.Fields("VENDEDOR")
        End If
        
        Precio = Precio & String(9 - Len(Precio), " ")
        cantidad = cantidad & String(6 - Len(cantidad), " ")
        desc = desc & String(9 - Len(desc), " ")
'        Tot = Tot & String(9 - Len(Tot), " ")
        titulo1 = "Cant" & String(6 - Len("Cant"), " ")
        titulo2 = "Precio" & String(3, " ")
        titulo3 = "Desc" & String(5, " ")
        titulo4 = "Total" & String(4, " ")
        titulo5 = "Atendio" & String(12, " ")
                
        Printer.Print clave & " " & Nombre2 & vbCrLf & titulo1 & " " & titulo2 & " " & titulo3 & " " & titulo4 & _
        vbCrLf & cantidad & " " & Precio & " " & desc & " " & tot & vbCrLf & titulo5 & vbCrLf & atendio
        
                
    ResAprt.MoveNext
    Loop
    
    
    Printer.Print "- - - - - - - - - - - - - - - - - - "
'    Printer.FontSize = 8
    Printer.Print Horario
    Printer.Print "- - - - - - - - - - - - - - - - - - "
'    Printer.Print "SUB TOTAL APARTADO:  " & FormatCurrency(SUBTOTAL)
'
'    If Val(DESCUENTO) > 0 Then
'    Printer.Print "         DESCUENTO:  " & FormatCurrency(DESCUENTO)
'    End If
    Printer.Print "   PAGO APARTADO:  " & FormatCurrency(pago)
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    If Val(PAGOEFECTIVO) > 0 Then
    Printer.Print "   PAGO EFECTIVO:  " & FormatCurrency(PAGOEFECTIVO)
    End If
    If Val(PAGOTARJETA) > 0 Then
    Printer.Print "    PAGO TARJETA:  " & FormatCurrency(PAGOTARJETA)
    End If
    If Val(PAGOCHEQUE) > 0 Then
    Printer.Print "   PAGO MONEDERO:  " & FormatCurrency(PAGOCHEQUE)
    End If

    Printer.Print "          PAGADO:  " & FormatCurrency(PAGADO)
    Printer.Print "          CAMBIO:  " & FormatCurrency(CAMBIO)
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    Printer.Print "--- INFORMACIÓN DEL APARTADO ---"
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    Printer.Print "FFECHA/HORA APARTADO: "
    Printer.Print FechaApartado
    Printer.Print "FOLIO APARTADO:      " & folio
    Printer.Print "FOLIO PAGO APARTADO: " & VENT_FOLIO
    Printer.Print "TOTAL APARTADO:      " & FormatCurrency(monto)
'    Printer.Print "PAGOS REALIZADOS     " & numpago
    Printer.Print "TOTAL PAGOS          " & FormatCurrency(PAGOS)
    Printer.Print "TOTAL FALTANTE       " & FormatCurrency(FALTANTE)
    Printer.Print "DIAS PARA LIQUIDAR:  " & DIAS_LIQUI
    Printer.Print "DIAS TRANSCURRIDOS:  " & DIAS_TRANS
    Printer.Print "FECHA LIQUIDACION:   " & FECHA_LIQUI
    Printer.Print "- - - - - - - - - - - - - - - - - - - -     "
    If Val(mone_entrega) < 0 Or Val(mone_recibe) > 0 Then
    Printer.Print "  --- MONEDERO ---"
    Printer.Print "   RECIBIDO: " & FormatCurrency(mone_recibe)
    Printer.Print "   APLICADO: " & FormatCurrency(mone_entrega)
    Printer.Print "      TOTAL: " & FormatCurrency(monedero)
    End If
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    Printer.Print "TEL CLIENTE:   "
    Printer.Print tel1 & " " & tel2
    Printer.Print "EMAIL CLIENTE: " & email
    Printer.Print email
    Printer.Print "- - - - - - - - - - - - - - - - - - - -     "
    Printer.Print vbCrLf & resSucur.Fields("SUC_SLOGAN") & vbCrLf
    'Call Centrar(Web, 15)
    Printer.Print resSucur.Fields("SUC_PAGINA_WEB")
    Printer.Print resSucur.Fields("SUC_EMAIL")
    Printer.Print "           OBSERVACIONES    " & vbCrLf
    Printer.FontBold = False
'    Printer.FontSize = 10
    
    'Printer.FontSize = RES3.Fields("SUC_TICKET_SIZE_1")
    
    If Len(resSucur.Fields("SUC_INFORMACION")) > 0 Then
    
        Lineas = Split(resSucur.Fields("SUC_INFORMACION"), vbNewLine)
    
        For b1 = 0 To UBound(Lineas)
            'Printer.Print UCase(Lineas(b1))
            
            For c1 = 0 To ((Round((Len(Lineas(b1)) / 35))) + 1)
                If Len(Lineas(b1)) >= 35 Then
                    'largo = True
                    Printer.Print Left(Lineas(b1), 35)
                    Lineas(b1) = Right(Lineas(b1), (Len(Lineas(b1)) - 35))
                Else
                    Printer.Print Lineas(b1)
                    Exit For
                End If
            Next c1
        Next b1
        
        Printer.Print "- - - - - - - - - - - - - - - - - - "
    End If
    
    
'    Dim Lineas() As String
'    Lineas = Split(resSucur.Fields("SUC_INFORMACION"), vbNewLine)
'    For b1 = 1 To UBound(Lineas)
'        Printer.Print Lineas(b1)
'    Next b1
    'Printer.Print RES3.Fields("SUC_INFORMACION") & vbCrLf & vbCrLf
'    Printer.FontSize = 12
'    Printer.FontName = "Control"
'    Printer.Print "P"  'Cut
    Printer.EndDoc
    
End Sub

Public Sub notaCredito(folio As String)
'    On Error Resume Next
    
    Dim sql1 As String
    Dim ResAprt As Recordset
    Dim ResAprt2 As Recordset
    Dim ResPago As Recordset
    Dim resSucur As Recordset
    
    Dim SUBTOTAL
    Dim DESCUENTO
    Dim total
    Dim PAGOEFECTIVO
    Dim PAGOTARJETA
    Dim PAGOCHEQUE
    Dim PAGADO
    Dim CAMBIO
    
    Dim APRT_FOLIO
    Dim PAGOS
    Dim FALTANTE
    Dim DIAS_LIQUI
    Dim DIAS_TRANS
    Dim monto
    Dim FechaApartado
    
    sql1 = "select * from SUCURSAL"
    Set resSucur = con.Execute(sql1)
    If Not resSucur.EOF Then
        If resSucur.Fields("SUC_ESTATUSTICKET") = 1 Then
            
        Else
            MsgBox "El status del ticket está desactivado.", vbInformation
            Exit Sub
        End If
    Else
        MsgBox "No se puede imprimir el ticket por que no tiene información referente a la sucursal del negocio. Verifique.", vbInformation
        Exit Sub
    End If


    
sql1 = "SELECT T2.FOLIO_VENTA, T2.FOLIO_APRT, T1.FECHA FECHA_APARTADO, T2.FECHA FECHA_PAGO, T1.CLIENTE, T1.VENDEDOR, T2.MOSTRADOR, T1.PRODUCTO, T1.CODIGO, " & _
"T1.PRECIO, T1.CANTIDAD, T1.DESCUENTO PROD_dESC, T1.TOTAL_PROD, T2.SUBTOTAL, T2.DESCUENTO, T2.TOTAL, " & _
"T2.PAGADO, T2.CAMBIO, T2.EFECTIVO, T2.TARJETA, T2.CHEQUE, T1.TOTAL TOTAL_ADEUDO, T2.PAGO,  " & _
"T1.PAGADO TOTAL_PAGADO, T1.ADEUDO TOTAL_FALTANTE, T1.DIAS, T1.TRANSCURRIDOS, T1.LIQUIDACION, (SELECT T41.TOTAL FROM VIEW_MONEDERO_CLIENTES T41 WHERE T1.APRT_CLIEPERID = T41.PER_ID) MONEDERO, " & _
"T1.tel1 , T1.tel2, T1.email, t1.FOLIO FROM VIEW_APARTADOS T1, VIEW_PAGOS_APARTTOTAL T2 " & _
"Where T1.folio = T2.FOLIO_APRT AND T2.FOLIO_VENTA = '" & folio & "'"
    
Set ResAprt = con.Execute(sql1)
Set ResAprt2 = con.Execute(sql1)
    
PAGOS = 0
FALTANTE = 0
monto = 0
monedero = 0

    If Not ResAprt.EOF Then
        monto = monto + Val(ResAprt2.Fields("TOTAL_ADEUDO"))
        PAGOS = PAGOS + Val(ResAprt2.Fields("TOTAL_PAGADO"))
        FALTANTE = FALTANTE + Val(ResAprt2.Fields("TOTAL_FALTANTE"))
        FechaApartado = ResAprt2.Fields("FECHA_APARTADO")
        
        SUBTOTAL = ResAprt.Fields("SUBTOTAL")
        DESCUENTO = ResAprt.Fields("DESCUENTO")
        total = ResAprt.Fields("TOTAL")
        PAGOEFECTIVO = ResAprt.Fields("EFECTIVO")
        PAGOTARJETA = ResAprt.Fields("TARJETA")
        PAGOCHEQUE = ResAprt.Fields("CHEQUE")
        PAGADO = ResAprt.Fields("PAGADO")
        CAMBIO = ResAprt.Fields("CAMBIO")
        APRT_FOLIO = ResAprt.Fields("FOLIO_APRT")
        VENT_FOLIO = ResAprt.Fields("FOLIO_VENTA")
        DIAS_LIQUI = ResAprt.Fields("DIAS")
        DIAS_TRANS = ResAprt.Fields("TRANSCURRIDOS")
        FECHA_LIQUI = Format(ResAprt.Fields("LIQUIDACION"), "Short Date")
        tel1 = ResAprt.Fields("TEL1")
        tel2 = ResAprt.Fields("TEL2")
        email = ResAprt.Fields("email")
        folio = ResAprt.Fields("FOLIO")
        pago = ResAprt.Fields("PAGO")
        If IsNull(ResAprt.Fields("MONEDERO")) Then
            monedero = 0
        Else
            monedero = Val(ResAprt.Fields("MONEDERO"))
        End If
    Else
        MsgBox "No se puede imprimir el ticket por que no tiene información referente al cobro referido. Verifique.", vbInformation
        Exit Sub
    End If
    
    Printer.KillDoc

    Printer.Font = "Courier New"
    Printer.FontSize = 8
    Printer.FontBold = True
    
    Printer.Print resSucur.Fields("SUC_RAZON_SOCIAL")
    
    Printer.Print resSucur.Fields("SUC_NOMBRE") & vbCrLf
    Printer.FontSize = 8
    Printer.FontBold = False
    Printer.Print resSucur.Fields("SUC_DIR_Calle") & " " & resSucur.Fields("SUC_DIR_NUM_EXT") & " " & resSucur.Fields("SUC_DIR_NUM_INT")
    Printer.Print resSucur.Fields("SUC_dIR_COLONIA") & " CP:" & resSucur.Fields("SUC_DIR_Cp")
'    Printer.Print RES1.Fields("Estado") & " " & RES1.Fields("Municipio")
    Printer.Print "Villahermosa" & " " & "Centro" & " " & "Tabasco"
    Printer.Print resSucur.Fields("SUC_TEL1") & "" & " " & resSucur.Fields("SUC_TEL2") & ""; vbCrLf
    Printer.Print Format(ResAprt.Fields("FECHA_PAGO"), "dddd dd-mm-yyyy") & " " & Format(ResAprt.Fields("FECHA_PAGO"), "Short Time") & vbCrLf
    Printer.Print "FOLIO: " & Format(folio, "0000000")
    Printer.Print "CLIENTE: "
    Printer.Print ResAprt.Fields("CLIENTE") & vbCrLf
    Printer.Print "COBRO: "
    Printer.Print ResAprt.Fields("MOSTRADOR")
    Printer.Print "OPERACIÓN DE CRÉDITO"
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    'For b1 = 1 To FrmFocus.ListaOper.Rows - 1
    numpago = 0
    Do While Not ResAprt.EOF
        numpago = numpago + 1
        clave = ResAprt.Fields("CODIGO")
        If Len(clave) > 8 Then
            clave = Left(clave, 8)
        Else
            clave = clave & String(8 - Len(clave), " ")
        End If
        Nombre2 = ResAprt.Fields("PRODUCTO")
        If Len(Nombre2) > 16 Then
            Nombre2 = Left(Nombre2, 16)
        Else
            Nombre2 = Nombre2 & String(16 - Len(Nombre2), " ")
        End If
        cantidad = ResAprt.Fields("CANTIDAD")
        Precio = ResAprt.Fields("PRECIO")
        desc = ResAprt.Fields("PROD_DESC")
        tot = ResAprt.Fields("TOTAL_PROD") '* ResAprt.Fields("APRT_pRODPRECIO") - ResAprt.Fields("APRT_DESC")
        If Len(Precio) > 9 Then
            Precio = Left(Precio, 9)
        End If
        If Len(ResAprt.Fields("VENDEDOR")) > 30 Then
            atendio = Left(ResAprt.Fields("VENDEDOR"), 30)
        Else
            atendio = ResAprt.Fields("VENDEDOR")
        End If
        
        Precio = Precio & String(9 - Len(Precio), " ")
        cantidad = cantidad & String(6 - Len(cantidad), " ")
        desc = desc & String(9 - Len(desc), " ")
'        Tot = Tot & String(9 - Len(Tot), " ")
        titulo1 = "Cant" & String(6 - Len("Cant"), " ")
        titulo2 = "Precio" & String(3, " ")
        titulo3 = "Desc" & String(5, " ")
        titulo4 = "Total" & String(4, " ")
        titulo5 = "Atendio" & String(12, " ")
                
        Printer.Print clave & " " & Nombre2 & vbCrLf & titulo1 & " " & titulo2 & " " & titulo3 & " " & titulo4 & _
        vbCrLf & cantidad & " " & Precio & " " & desc & " " & tot & vbCrLf & titulo5 & vbCrLf & atendio
        
                
    ResAprt.MoveNext
    Loop
    
    
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    Printer.FontSize = 8
    Printer.Print Horario
    Printer.Print "- - - - - - - - - - - - - - - - - - "
'    Printer.Print "SUB TOTAL APARTADO:  " & FormatCurrency(SUBTOTAL)
'
'    If Val(DESCUENTO) > 0 Then
'    Printer.Print "         DESCUENTO:  " & FormatCurrency(DESCUENTO)
'    End If
    Printer.Print "            PAGO:  " & FormatCurrency(pago)
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    If Val(PAGOEFECTIVO) > 0 Then
    Printer.Print "   PAGO EFECTIVO:  " & FormatCurrency(PAGOEFECTIVO)
    End If
    If Val(PAGOTARJETA) > 0 Then
    Printer.Print "    PAGO TARJETA:  " & FormatCurrency(PAGOTARJETA)
    End If
    If Val(PAGOCHEQUE) > 0 Then
    Printer.Print "     PAGO CHEQUE:  " & FormatCurrency(PAGOCHEQUE)
    End If

    Printer.Print "          PAGADO:  " & FormatCurrency(PAGADO)
    Printer.Print "          CAMBIO:  " & FormatCurrency(CAMBIO)
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    Printer.Print "--- INFORMACIÓN DEL CREDITO ---"
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    Printer.Print "FFECHA/HORA OPERACIÓN: "
    Printer.Print FechaApartado
    Printer.Print "FOLIO CREDITO:      " & folio
    Printer.Print "FOLIO PAGO CREDITO: " & VENT_FOLIO
    Printer.Print "TOTAL CREDITO:      " & FormatCurrency(monto)
'    Printer.Print "PAGOS REALIZADOS     " & numpago
    Printer.Print "TOTAL PAGOS          " & FormatCurrency(PAGOS)
    Printer.Print "TOTAL FALTANTE       " & FormatCurrency(FALTANTE)
    'Printer.Print "DIAS PARA LIQUIDAR:  " & DIAS_LIQUI
    Printer.Print "DIAS TRANSCURRIDOS:  " & DIAS_TRANS
'    Printer.Print "FECHA LIQUIDACION:   " & FECHA_LIQUI
    Printer.Print "- - - - - - - - - - - - - - - - - - - -     "
    If Val(monedero) > 0 Then
    Printer.Print "  TOTAL MONEDERO:  " & FormatCurrency(monedero)
    End If
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    Printer.Print "TEL CLIENTE:   "
    Printer.Print tel1 & " " & tel2
    Printer.Print "EMAIL CLIENTE: " & email
    Printer.Print email
    Printer.Print "- - - - - - - - - - - - - - - - - - - -     "
    Printer.Print vbCrLf & resSucur.Fields("SUC_SLOGAN") & vbCrLf
    'Call Centrar(Web, 15)
    Printer.Print resSucur.Fields("SUC_PAGINA_WEB")
    Printer.Print resSucur.Fields("SUC_EMAIL")
    Dim Lineas() As String
    Lineas = Split(resSucur.Fields("SUC_INFORMACION"), vbNewLine)
    For b1 = 1 To UBound(Lineas)
        Printer.Print Lineas(b1)
    Next b1
    'Printer.Print RES3.Fields("SUC_INFORMACION") & vbCrLf & vbCrLf
'    Printer.FontSize = 12
'    Printer.FontName = "Control"
'    Printer.Print "P"  'Cut
    Printer.EndDoc
    
End Sub


Public Sub infoApartado(texto As String)
    On Error Resume Next
    
    Dim sql1 As String
    Dim RES3 As Recordset
    Dim RES4 As Recordset
    Dim RES5 As Recordset

    Dim SUBTOTAL
    Dim DESCUENTO
    Dim total
    Dim PAGOEFECTIVO
    Dim PAGOTARJETA
    Dim PAGOCHEQUE
    Dim PAGADO
    Dim CAMBIO

    
    sql1 = "select * from SUCURSAL"
    Set RES3 = con.Execute(sql1)
    If Not RES3.EOF Then
        If RES3.Fields("SUC_ESTATUSTICKET") = 1 Then
            
        Else
            MsgBox "El status del ticket está desactivado.", vbInformation
            Exit Sub
        End If
    Else
        MsgBox "No se puede imprimir el ticket por que no tiene información referente a la sucursal del negocio. Verifique.", vbInformation
        Exit Sub
    End If
    

    
    Printer.KillDoc
    Printer.FontSize = 12
    'Printer.FontName = "Control"
    'Printer.Print "C"  'open Drawer 1 at 50ms
    Printer.Font = "Courier New"
    Printer.FontSize = 10
    Printer.FontBold = True
    
    'Call Centrar(Nombre, 12)
    Printer.Print RES3.Fields("SUC_RAZON_SOCIAL")
    
    'Call Centrar(Sucursal, 12)
    Printer.Print RES3.Fields("SUC_NOMBRE") & vbCrLf
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.Print RES3.Fields("SUC_DIR_Calle") & " " & RES3.Fields("SUC_DIR_NUM_EXT") & " " & RES3.Fields("SUC_DIR_NUM_INT")
    Printer.Print RES3.Fields("SUC_dIR_COLONIA") & " CP:" & RES3.Fields("SUC_DIR_Cp")
'    Printer.Print RES1.Fields("Estado") & " " & RES1.Fields("Municipio")
    Printer.Print "Villahermosa" & " " & "Centro" & " " & "Tabasco"
    Printer.Print RES3.Fields("SUC_TEL1") & "" & " " & RES3.Fields("SUC_TEL2") & ""; vbCrLf
    Printer.Print Format(RES4.Fields("vent_fechahora_cobro"), "dddd dd-mm-yyyy") & " " & Format(RES4.Fields("vent_fechahora_cobro"), "Short Time") & vbCrLf
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    Printer.Print "Información de apartado"
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    'For b1 = 1 To FrmFocus.ListaOper.Rows - 1
    Printer.Print texto
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    Printer.FontSize = 10
    Printer.Print Horario
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    'Call Centrar(Eslogan, 15)
    Printer.Print vbCrLf & RES3.Fields("SUC_SLOGAN") & vbCrLf
    'Call Centrar(Web, 15)
    Printer.Print RES3.Fields("SUC_PAGINA_WEB")
    Printer.Print RES3.Fields("SUC_EMAIL")
    Dim Lineas() As String
    Lineas = Split(RES3.Fields("SUC_INFORMACION"), vbNewLine)
    For b1 = 1 To UBound(Lineas)
        Printer.Print Lineas(b1)
    Next b1
    Printer.EndDoc
    

End Sub
Public Sub resumenUsuario()
    'On Error Resume Next
    
    Dim sql1 As String
    Dim resSucursal As Recordset
    
    sql1 = "select * from SUCURSAL"
    Set resSucursal = con.Execute(sql1)
    If Not resSucursal.EOF Then
        If resSucursal.Fields("SUC_ESTATUSTICKET") = 1 Then
            
        Else
            MsgBox "El status del ticket está desactivado.", vbInformation
            Exit Sub
        End If
    Else
        MsgBox "No se puede imprimir el ticket por que no tiene información referente a la sucursal del negocio. Verifique.", vbInformation
        Exit Sub
    End If
        
    Printer.KillDoc
    Printer.FontSize = 8
    Printer.Font = "Courier New"
    Printer.FontSize = 8
    Printer.FontBold = True
    Printer.Print resSucursal.Fields("SUC_RAZON_SOCIAL")
    Printer.Print resSucursal.Fields("SUC_NOMBRE")
    Printer.Print ""
    Printer.FontSize = 8
    Printer.FontBold = False
    Printer.Print resSucursal.Fields("SUC_DIR_Calle") & " " & resSucursal.Fields("SUC_DIR_NUM_EXT") & " " & resSucursal.Fields("SUC_DIR_NUM_INT")
    Printer.Print resSucursal.Fields("SUC_dIR_COLONIA") & " CP:" & resSucursal.Fields("SUC_DIR_Cp")
'    Printer.Print RES1.Fields("Estado") & " " & RES1.Fields("Municipio")
    Printer.Print "Villahermosa" & " " & "Centro" & " " & "Tabasco"
    Printer.Print resSucursal.Fields("SUC_TEL1") & "" & " " & resSucursal.Fields("SUC_TEL2") & ""
    'Printer.Print ""
    Printer.Print "Fecha de Corte "
    Printer.Print Format(FRM_Caja.dtFecha1(0), "Long Date")
    If FRM_Caja.dtFecha1(0) <> FRM_Caja.dtFecha1(1) Then
        Printer.Print " a " & Format(FRM_Caja.dtFecha1(1), "Long Date")
    End If
'    Printer.Print ""
    Printer.Print "Fecha/Hora Impresion: "
    Printer.Print Format(Date, "Long Date")
    Printer.Print Format(Time, "Short Time")
 '   Printer.Print ""
    Printer.Print "Usuario al corte: "
    Printer.Print FRM_Menu.menuBarra2.Panels(5).Text
  '  Printer.Print "- - - - - - - - - - - - - - - - - - "
    Printer.Print "Corte de caja por usuario: "
   ' Printer.Print "- - - - - - - - - - - - - - - - - - "
    '''''-----Datos generales del Usuario
    Printer.Print "Usuario:"
    Printer.Print FRM_Caja.lista4.TextMatrix(FRM_Caja.lista4.Row, 0)
    Printer.Print "Fecha/Hora corte:"
    Printer.Print FRM_Caja.lista4.TextMatrix(FRM_Caja.lista4.Row, 1)
    Printer.Print "Corte de sesión:"
    Printer.Print FRM_Caja.lista4.TextMatrix(FRM_Caja.lista4.Row, 6)
    Printer.Print "Corte general:"
    Printer.Print FRM_Caja.lista4.TextMatrix(FRM_Caja.lista4.Row, 11)
    Printer.Print "Detalle del corte del usuario: "
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    'Printer.Print "Descripcion - Precio - Cant - Total    "
    For b1 = 1 To FRM_Caja.Lista6.Rows - 1
        Printer.Print FRM_Caja.Lista6.TextMatrix(b1, 1)
        Printer.Print FRM_Caja.Lista6.TextMatrix(b1, 3) & "  " & FRM_Caja.Lista6.TextMatrix(b1, 4) & " " & FRM_Caja.Lista6.TextMatrix(b1, 5)
    Next b1

    Printer.Print "- - - - - - - - - - - - - - - - - - "
    Printer.FontSize = 8
    Printer.Print resSucursal.Fields("SUC_SLOGAN") & vbCrLf
    Printer.Print resSucursal.Fields("SUC_PAGINA_WEB")
    Printer.Print resSucursal.Fields("SUC_EMAIL")
    Dim Lineas() As String
    Lineas = Split(resSucursal.Fields("SUC_INFORMACION"), vbNewLine)
    For b1 = 1 To UBound(Lineas)
        Printer.Print Lineas(b1)
    Next b1
    Printer.EndDoc
    

End Sub
Public Sub resumenCaja()
    'On Error Resume Next
    
    Dim sql1 As String
    Dim resSucursal As Recordset
    
    sql1 = "select * from SUCURSAL"
    Set resSucursal = con.Execute(sql1)
    If Not resSucursal.EOF Then
        If resSucursal.Fields("SUC_ESTATUSTICKET") = 1 Then
            
        Else
            MsgBox "El status del ticket está desactivado.", vbInformation
            Exit Sub
        End If
    Else
        MsgBox "No se puede imprimir el ticket por que no tiene información referente a la sucursal del negocio. Verifique.", vbInformation
        Exit Sub
    End If
        
    Printer.KillDoc
    Printer.FontSize = 12
    Printer.Font = "Courier New"
    Printer.FontSize = 10
    Printer.FontBold = True
    Printer.Print resSucursal.Fields("SUC_RAZON_SOCIAL")
    Printer.Print resSucursal.Fields("SUC_NOMBRE")
    Printer.Print ""
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.Print resSucursal.Fields("SUC_DIR_Calle") & " " & resSucursal.Fields("SUC_DIR_NUM_EXT") & " " & resSucursal.Fields("SUC_DIR_NUM_INT")
    Printer.Print resSucursal.Fields("SUC_dIR_COLONIA") & " CP:" & resSucursal.Fields("SUC_DIR_Cp")
'    Printer.Print RES1.Fields("Estado") & " " & RES1.Fields("Municipio")
    Printer.Print "Villahermosa" & " " & "Centro" & " " & "Tabasco"
    Printer.Print resSucursal.Fields("SUC_TEL1") & "" & " " & resSucursal.Fields("SUC_TEL2") & ""
    Printer.Print ""
    Printer.Print "Fecha de Corte "
    Printer.Print Format(FRM_Caja.dtFecha1(0), "Long Date")
    If FRM_Caja.dtFecha1(0) <> FRM_Caja.dtFecha1(1) Then
        Printer.Print " a " & Format(FRM_Caja.dtFecha1(1), "Long Date")
    End If
    Printer.Print ""
    Printer.Print "Fecha/Hora Impresion: "
    Printer.Print Format(Date, "Long Date")
    Printer.Print Format(Time, "Short Time")
    Printer.Print ""
    Printer.Print "Usuario al corte: "
    Printer.Print FRM_Menu.menuBarra2.Panels(5).Text
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    Printer.Print "Resumen del corte de caja: "
    For b1 = 0 To FRM_Caja.lista.Rows - 1
        Printer.Print Left(FRM_Caja.lista.TextMatrix(b1, 0), 10) & String(10 - Len((Left(FRM_Caja.lista.TextMatrix(b1, 0), 10))), " ") & _
        FRM_Caja.lista.TextMatrix(b1, 1) & String(10 - Len(FRM_Caja.lista.TextMatrix(b1, 1)), " ") & _
        FRM_Caja.lista.TextMatrix(b1, 2) & String(10 - Len(FRM_Caja.lista.TextMatrix(b1, 2)), " ")
    Next b1

    Printer.Print "- - - - - - - - - - - - - - - - - - "
    Printer.FontSize = 10
    Printer.Print resSucursal.Fields("SUC_SLOGAN") & vbCrLf
    Printer.Print resSucursal.Fields("SUC_PAGINA_WEB")
    Printer.Print resSucursal.Fields("SUC_EMAIL")
    Dim Lineas() As String
    Lineas = Split(resSucursal.Fields("SUC_INFORMACION"), vbNewLine)
    For b1 = 1 To UBound(Lineas)
        Printer.Print Lineas(b1)
    Next b1
    Printer.EndDoc
    

End Sub


Public Sub resumenCaja2()
   
   ' On Error Resume Next
    
    Dim sql1 As String
    Dim resSucursal As Recordset
    Dim monto As Double
    Dim ESP As Double
    
    sql1 = "select * from SUCURSAL"
    Set resSucursal = con.Execute(sql1)
    If Not resSucursal.EOF Then
        ESP = resSucursal.Fields("SUC_ESP")
        If resSucursal.Fields("SUC_ESTATUSTICKET") = 1 Then
            
        Else
            MsgBox "El status del ticket está desactivado.", vbInformation
            Exit Sub
        End If
    Else
        MsgBox "No se puede imprimir el ticket por que no tiene información referente a la sucursal del negocio. Verifique.", vbInformation
        Exit Sub
    End If
        
    Printer.KillDoc
    Printer.FontSize = 12
    Printer.Font = "Courier New"
    Printer.FontSize = 10
    Printer.FontBold = True
    Printer.Print resSucursal.Fields("SUC_RAZON_SOCIAL")
    Printer.Print resSucursal.Fields("SUC_NOMBRE")
    Printer.Print ""
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.Print resSucursal.Fields("SUC_DIR_Calle") & " " & resSucursal.Fields("SUC_DIR_NUM_EXT") & " " & resSucursal.Fields("SUC_DIR_NUM_INT")
    Printer.Print resSucursal.Fields("SUC_dIR_COLONIA") & " CP:" & resSucursal.Fields("SUC_DIR_Cp")
'    Printer.Print RES1.Fields("Estado") & " " & RES1.Fields("Municipio")
    Printer.Print "Villahermosa" & " " & "Centro" & " " & "Tabasco"
    Printer.Print resSucursal.Fields("SUC_TEL1") & "" & " " & resSucursal.Fields("SUC_TEL2") & ""
    Printer.Print ""
    Printer.Print "Fecha de Corte "
    Printer.Print Format(FRM_Caja.dtFecha1(0), "Long Date")
    If FRM_Caja.dtFecha1(0) <> FRM_Caja.dtFecha1(1) Then
        Printer.Print " a " & Format(FRM_Caja.dtFecha1(1), "Long Date")
    End If
    Printer.Print ""
    Printer.Print "Fecha/Hora Impresion: "
    Printer.Print Format(Date, "Long Date")
    Printer.Print Format(Time, "Short Time")
    Printer.Print ""
    Printer.Print "Usuario al corte: "
    Printer.Print FRM_Menu.menuBarra2.Panels(5).Text
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    Printer.Print "Resumen del corte de caja: "
    

    monto = (Val(Format(FRM_Caja.lista.TextMatrix(14, 1), "General Number")) * (ESP))
    SUBTOTAL = monto + (Val(Format(FRM_Caja.lista.TextMatrix(15, 1), "General Number")))
    total = SUBTOTAL
    
    For b1 = 0 To FRM_Caja.lista.Rows - 1
        
        If b1 >= 12 And b1 <= 17 Then
            If b1 = 12 Then
                Printer.Print FRM_Caja.lista.TextMatrix(b1, 0) & String(10 - Len(FRM_Caja.lista.TextMatrix(b1, 0)), " ") & FormatCurrency(SUBTOTAL) & String(10 - Len(FormatCurrency(SUBTOTAL)), " ") & FRM_Caja.lista.TextMatrix(b1, 2) & String(10 - Len(FRM_Caja.lista.TextMatrix(b1, 2)), " ")
            Else
                If b1 = 13 Then
                    Printer.Print FRM_Caja.lista.TextMatrix(b1, 0) & String(10 - Len(FRM_Caja.lista.TextMatrix(b1, 0)), " ") & FormatCurrency(total) & String(10 - Len(FormatCurrency(total)), " ") & FRM_Caja.lista.TextMatrix(b1, 2) & String(10 - Len(FRM_Caja.lista.TextMatrix(b1, 2)), " ")
                Else
                    If b1 = 14 Then
                        Printer.Print FRM_Caja.lista.TextMatrix(b1, 0) & String(10 - Len(FRM_Caja.lista.TextMatrix(b1, 0)), " ") & _
                        FormatCurrency(monto) & String(10 - Len(FormatCurrency(monto)), " ") & _
                        FRM_Caja.lista.TextMatrix(b1, 2) & String(10 - Len(FRM_Caja.lista.TextMatrix(b1, 2)), " ")
                    Else
                        If b1 = 17 Then
                            'Printer.Print FRM_Caja.lista.TextMatrix(b1, 0) & String(10 - Len(FRM_Caja.lista.TextMatrix(b1, 0)), " ") & FormatCurrency(monto) & String(10 - Len(FormatCurrency(monto)), " ") & FRM_Caja.lista.TextMatrix(b1, 2) & String(10 - Len(FRM_Caja.lista.TextMatrix(b1, 2)), " ")
                        Else
                            
                            Printer.Print Left(FRM_Caja.lista.TextMatrix(b1, 0), 10) & String(10 - Len(Left(FRM_Caja.lista.TextMatrix(b1, 0), 10)), " ") & _
                            FRM_Caja.lista.TextMatrix(b1, 1) & String(10 - Len(FRM_Caja.lista.TextMatrix(b1, 1)), " ") & _
                            FRM_Caja.lista.TextMatrix(b1, 2) & String(10 - Len(FRM_Caja.lista.TextMatrix(b1, 2)), " ")
                        End If
                    End If
                End If
            End If
        End If
    Next b1

    Printer.Print "- - - - - - - - - - - - - - - - - - "
    Printer.FontSize = 10
    Printer.Print resSucursal.Fields("SUC_SLOGAN") & vbCrLf
    Printer.Print resSucursal.Fields("SUC_PAGINA_WEB")
    Printer.Print resSucursal.Fields("SUC_EMAIL")
    Dim Lineas() As String
    Lineas = Split(resSucursal.Fields("SUC_INFORMACION"), vbNewLine)
    For b1 = 1 To UBound(Lineas)
        Printer.Print Lineas(b1)
    Next b1
    Printer.EndDoc
    

End Sub

Public Sub resumenDetalleCaja()
    On Error Resume Next
    
    Dim sql1 As String
    Dim resSucursal As Recordset
    
    sql1 = "select * from SUCURSAL"
    Set resSucursal = con.Execute(sql1)
    If Not resSucursal.EOF Then
        If resSucursal.Fields("SUC_ESTATUSTICKET") = 1 Then
            
        Else
            MsgBox "El status del ticket está desactivado.", vbInformation
            Exit Sub
        End If
    Else
        MsgBox "No se puede imprimir el ticket por que no tiene información referente a la sucursal del negocio. Verifique.", vbInformation
        Exit Sub
    End If
        
    Printer.KillDoc
    Printer.FontSize = 12
    Printer.Font = "Courier New"
    Printer.FontSize = 10
    Printer.FontBold = True
    Printer.Print resSucursal.Fields("SUC_RAZON_SOCIAL")
    Printer.Print resSucursal.Fields("SUC_NOMBRE")
    Printer.Print ""
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.Print resSucursal.Fields("SUC_DIR_Calle") & " " & resSucursal.Fields("SUC_DIR_NUM_EXT") & " " & resSucursal.Fields("SUC_DIR_NUM_INT")
    Printer.Print resSucursal.Fields("SUC_dIR_COLONIA") & " CP:" & resSucursal.Fields("SUC_DIR_Cp")
'    Printer.Print RES1.Fields("Estado") & " " & RES1.Fields("Municipio")
    Printer.Print "Villahermosa" & " " & "Centro" & " " & "Tabasco"
    Printer.Print resSucursal.Fields("SUC_TEL1") & "" & " " & resSucursal.Fields("SUC_TEL2") & ""
    Printer.Print ""
    Printer.Print "Fecha de Corte "
    Printer.Print Format(FRM_Caja.dtFecha1(0), "Long Date")
    If FRM_Caja.dtFecha1(0) <> FRM_Caja.dtFecha1(1) Then
        Printer.Print " a " & Format(FRM_Caja.dtFecha1(1), "Long Date")
    End If
    Printer.Print ""
    Printer.Print "Fecha/Hora Impresion: "
    Printer.Print Format(Date, "Long Date")
    Printer.Print Format(Time, "Short Time")
    Printer.Print ""
    Printer.Print "Usuario al corte: "
    Printer.Print FRM_Menu.menuBarra2.Panels(5).Text
    '''''Resumen'''''''''''''
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    Printer.Print "Resumen del corte de caja: "
    For b1 = 0 To FRM_Caja.lista.Rows - 1
        Printer.Print FRM_Caja.lista.TextMatrix(b1, 0) & String(10 - Len(FRM_Caja.lista.TextMatrix(b1, 0)), " ") & _
        FRM_Caja.lista.TextMatrix(b1, 1) & String(10 - Len(FRM_Caja.lista.TextMatrix(b1, 1)), " ") & _
        FRM_Caja.lista.TextMatrix(b1, 2) & String(10 - Len(FRM_Caja.lista.TextMatrix(b1, 2)), " ")
    Next b1

    Printer.Print "- - - - - - - - - - - - - - - - - - "
    
    '''''Resumen'''''''''''''
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    Printer.Print "Detalle del corte de caja "
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    '''''Resumen'''''''''''''
    Printer.Print "Ventas"
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    Printer.FontSize = 7
    With FRM_Caja.Lista3
        For b1 = 0 To .Rows - 1
            Printer.Print .TextMatrix(b1, 0) & .TextMatrix(b1, 12) & .TextMatrix(b1, 13)
        Next b1
    End With
    
    '''''Gastos'''''''''''''
    Printer.Print "Gastos"
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    Printer.FontSize = 7
    With FRM_Caja.Lista3
        For b1 = 0 To .Rows - 1
            Printer.Print .TextMatrix(b1, 0) & .TextMatrix(b1, 12) & .TextMatrix(b1, 13)
        Next b1
    End With
    
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    
    Printer.FontSize = 10
    Printer.Print resSucursal.Fields("SUC_SLOGAN") & vbCrLf
    Printer.Print resSucursal.Fields("SUC_PAGINA_WEB")
    Printer.Print resSucursal.Fields("SUC_EMAIL")
    Dim Lineas() As String
    Lineas = Split(resSucursal.Fields("SUC_INFORMACION"), vbNewLine)
    For b1 = 1 To UBound(Lineas)
        Printer.Print Lineas(b1)
    Next b1
    Printer.EndDoc
    

End Sub


Public Sub notaCambio(folio As String)
    On Error Resume Next
    
    Dim sql1 As String
    Dim ResCambio As Recordset
    Dim ResCambio2 As Recordset
    Dim ResPago As Recordset
    Dim resSucur As Recordset
    
    Dim SUBTOTAL
    Dim DESCUENTO
    Dim total
    Dim PAGOEFECTIVO
    Dim PAGOTARJETA
    Dim PAGOCHEQUE
    Dim PAGADO
    Dim CAMBIO
    
    Dim clave1, clave2, precio1, precio2, Nombre2, nombre3 As String
    
    Dim APRT_FOLIO
    Dim PAGOS
    Dim FALTANTE
    Dim DIAS_LIQUI
    Dim DIAS_TRANS
    Dim monto
    Dim FechaApartado
    
    sql1 = "select * from SUCURSAL"
    Set resSucur = con.Execute(sql1)
    If Not resSucur.EOF Then
        If resSucur.Fields("SUC_ESTATUSTICKET") = 1 Then
            
        Else
            MsgBox "El status del ticket está desactivado.", vbInformation
            Exit Sub
        End If
    Else
        MsgBox "No se puede imprimir el ticket por que no tiene información referente a la sucursal del negocio. Verifique.", vbInformation
        Exit Sub
    End If
    
    
'(SELECT T41.TOTAL FROM VIEW_MONEDERO_CLIENTES T41 WHERE T2.PER_ID = T41.PER_ID) MONEDERO
sql1 = "SELECT * FROM VIEW_CAMBIOS WHERE CMB_ID = '" & folio & "'"
    
    
    
Set ResCambio = con.Execute(sql1)
    If Not ResCambio.EOF Then
        SUBTOTAL = ResCambio.Fields("VENT_SUBTOTAL")
        DESCUENTO = ResCambio.Fields("VENT_dESCUENTO")
        total = ResCambio.Fields("TOT_DIF")
        PAGOEFECTIVO = ResCambio.Fields("VENT_PAGOEFECTIVO")
        PAGOTARJETA = ResCambio.Fields("VENT_PAGOTARJETA")
        PAGOCHEQUE = "0" 'ResCambio.Fields("VENT_PAGOCHEQUE")
        PAGADO = ResCambio.Fields("VENT_PAGADO")
        CAMBIO = ResCambio.Fields("VENT_CAMBIO")
        APRT_FOLIO = ResCambio.Fields("FOLIO_DEVO")
    Else
        MsgBox "No se puede imprimir el ticket por que no tiene información referente al cobro referido. Verifique.", vbInformation
        Exit Sub
    End If
    
    Printer.KillDoc
    Printer.Font = "Courier New"
    Printer.FontSize = 10
    Printer.FontBold = True
    
    Printer.Print resSucur.Fields("SUC_RAZON_SOCIAL")
    
    Printer.Print resSucur.Fields("SUC_NOMBRE") & vbCrLf
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.Print resSucur.Fields("SUC_DIR_Calle") & " " & resSucur.Fields("SUC_DIR_NUM_EXT") & " " & resSucur.Fields("SUC_DIR_NUM_INT")
    Printer.Print resSucur.Fields("SUC_dIR_COLONIA") & " CP:" & resSucur.Fields("SUC_DIR_Cp")
    Printer.Print "Villahermosa" & " " & "Centro" & " " & "Tabasco"
    Printer.Print resSucur.Fields("SUC_TEL1") & "" & " " & resSucur.Fields("SUC_TEL2") & ""; vbCrLf
    Printer.Print Format(ResCambio.Fields("FECHA_DEVO"), "dddd dd-mm-yyyy") & " " & Format(ResCambio.Fields("FECHA_DEVO"), "Short Time") & vbCrLf
    Printer.Print "FOLIO: " & Format(folio, "0000000")
    Printer.Print "CLIENTE: "
    Printer.Print ResCambio.Fields("CLIENTE") & vbCrLf
    Printer.Print "COBRO: "
    Printer.Print ResCambio.Fields("USUARIO_DEVO")
    Printer.Print "CAMBIO POR DEVOLUCIÓN"
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    'For b1 = 1 To FrmFocus.ListaOper.Rows - 1
    Do While Not ResCambio.EOF
        clave1 = ResCambio.Fields("PROD_DEVCODIGO")
        If Len(clave1) > 17 Then
            clave1 = Left(clave1, 17)
        Else
            clave1 = clave1 & String(17 - Len(clave1), " ")
        End If
        Nombre2 = ResCambio.Fields("DEV_PRODUCTO")
        If Len(Nombre2) > 16 Then
            Nombre2 = Left(Nombre2, 16)
        Else
            Nombre2 = Nombre2 & String(16 - Len(Nombre2), " ")
        End If
        precio1 = ResCambio.Fields("DEV_PRECIO")
        If Len(precio1) > 9 Then
            precio1 = Left(precio1, 9)
        End If
        
        
        clave2 = ResCambio.Fields("PROD_CAMCODIGO")
        If Len(clave2) > 17 Then
            clave2 = Left(clave12, 17)
        Else
            clave2 = clave2 & String(17 - Len(clave2), " ")
        End If
        nombre3 = ResCambio.Fields("CAM_PRODUCTO")
        If Len(nombre3) > 16 Then
            nombre3 = Left(nombre3, 16)
        Else
            nombre3 = nombre3 & String(16 - Len(nombre3), " ")
        End If
        precio2 = ResCambio.Fields("CAM_TOTAL")
        If Len(precio2) > 9 Then
            precio2 = Left(precio2, 9)
        End If
        Printer.Print "Producto devuelto: "
        
        Printer.Print "Clave:"
        Printer.Print clave1 & " "
        Printer.Print "Producto:         Precio: "
        Printer.Print Nombre2 & " " & precio1
        Printer.Print "Fecha compra: " & ResCambio.Fields("FECHA_VENTA")
        Printer.Print ""
        Printer.Print "Producto entregado: "
        Printer.Print "Clave:"
        Printer.Print clave2 & " "
        Printer.Print "Producto:         Precio: "
        Printer.Print nombre3 & " " & precio2 & vbCrLf
        
                
    ResCambio.MoveNext
    Loop
    
    
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    Printer.FontSize = 10
    Printer.Print Horario
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    Printer.Print "TOTAL DIFERENCIA CAMBIO: " & FormatCurrency(total)
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    If Val(total) < 0 Then
    Printer.Print "ABONO EN MONEDERO: " & FormatCurrency(total * -1)
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    End If
    If Val(PAGOEFECTIVO) > 0 Then
    Printer.Print "   PAGO EFECTIVO:  " & FormatCurrency(PAGOEFECTIVO)
    End If
    If Val(PAGOTARJETA) > 0 Then
    Printer.Print "    PAGO TARJETA:  " & FormatCurrency(PAGOTARJETA)
    End If
    If Val(PAGOCHEQUE) > 0 Then
    Printer.Print "     PAGO CHEQUE:  " & FormatCurrency(PAGOCHEQUE)
    End If

    Printer.Print "          PAGADO:  " & FormatCurrency(PAGADO)
    Printer.Print "          CAMBIO:  " & FormatCurrency(CAMBIO)
    Printer.Print "- - - - - - - - - - - - - - - - - - "
 
    Printer.Print vbCrLf & resSucur.Fields("SUC_SLOGAN") & vbCrLf
    'Call Centrar(Web, 15)
    Printer.Print resSucur.Fields("SUC_PAGINA_WEB")
    Printer.Print resSucur.Fields("SUC_EMAIL")
    Dim Lineas() As String
    Lineas = Split(resSucur.Fields("SUC_INFORMACION"), vbNewLine)
    For b1 = 1 To UBound(Lineas)
        Printer.Print Lineas(b1)
    Next b1
    'Printer.Print RES3.Fields("SUC_INFORMACION") & vbCrLf & vbCrLf
'    Printer.FontSize = 12
'    Printer.FontName = "Control"
'    Printer.Print "P"  'Cut
    Printer.EndDoc
    
End Sub

Public Sub notaPreTicket(folio As String)
    On Error Resume Next
    
    Dim sql1 As String
    Dim RES3 As Recordset
    Dim RES4 As Recordset
    Dim RES5 As Recordset
    Dim ASIENTO As String

    Dim SUBTOTAL
    Dim DESCUENTO
    Dim total
    Dim PAGOEFECTIVO
    Dim PAGOTARJETA
    Dim PAGOCHEQUE
    Dim PAGADO
    Dim CAMBIO
    Dim IVA As String
   Dim OBSERVACIONES As String
    Dim Lineas() As String
    IVA = "N"
    sql1 = "select * from SUCURSAL"
    Set RES3 = con.Execute(sql1)
    If Not RES3.EOF Then
        IVA = RES3.Fields("SUC_IVA")
        If RES3.Fields("SUC_ESTATUSTICKET") = 1 Then
            
        Else
            MsgBox "El status del ticket está desactivado.", vbInformation
            Exit Sub
        End If
    Else
        MsgBox "No se puede imprimir el ticket por que no tiene información referente a la sucursal del negocio. Verifique.", vbInformation
        Exit Sub
    End If
    
'    SQL1 = "SELECT T1.vent_fechahora, T1.vent_fechahora_cobro, T1.VENT_MESA MESA, CONCAT(T2.PER_NOMBRE, ' ', T2.PER_PATERNO, ' ', T2.PER_MATERNO) CLIENTE, " & _
'    "CONCAT(T3.PER_NOMBRE, ' ', T3.PER_PATERNO, ' ', T3.PER_MATERNO) USUARIO, " & _
'    "CONCAT(T4.PER_NOMBRE, ' ', T4.PER_PATERNO, ' ', T4.PER_MATERNO) CAJA, " & _
'    "T5.VENDET_PRODCODIGO, T5.VENDET_PRODUCTONOMBRE, T5.VENDET_PRECIO, T5.VENDET_CANTIDAD, T5.venDet_Descuento, " & _
'    "T1.VENT_PAGOEFECTIVO, VENT_PAGOTARJETA, VENT_PAGOCHEQUE, VENT_PAGADO, VENT_OBSERVACIONES, VENT_CAMBIO, (SELECT SUM(VENDET_PRECIO * VENDET_CANTIDAD) FROM VENTA_DETALLE WHERE VENDET_FOLIO =  T1.VENT_IDFOLIO) SUBTOTAL, (SELECT SUM(VENDET_DESCUENTO) FROM VENTA_DETALLE WHERE VENDET_FOLIO =  T1.VENT_IDFOLIO) DESCUENTO, (SELECT ((SUM(VENDET_PRECIO * VENDET_CANTIDAD)) - SUM(VENDET_DESCUENTO)) FROM VENTA_DETALLE WHERE VENDET_FOLIO =  T1.VENT_IDFOLIO) TOTAL " & _
'    "FROM VENTAS T1, PERSONA T2, PERSONA T3, VENTA_DETALLE T5, PERSONA T4 " & _
'    "WHERE T1.VENT_IDFOLIO = '" & folio & "' AND T1.VENT_CLIEPERID = T2.PER_ID AND T5.VENDET_VENDPERID = T3.PER_ID AND " & _
'    "T1.VENT_IDFOLIO = T5.VENDET_FOLIO AND T1.VENT_VENDPERID = T4.PER_ID "
    
    sql1 = "SELECT T1.vent_fechahora_cobro, T1.vent_fechahora, T1.VENT_MESA MESA, VENT_PERSONAS PERSONAS, CONCAT(T2.PER_NOMBRE, ' ', T2.PER_PATERNO, ' ', T2.PER_MATERNO) CLIENTE, " & _
    "CONCAT(T3.PER_NOMBRE, ' ', T3.PER_PATERNO, ' ', T3.PER_MATERNO) USUARIO, CONCAT(T4.PER_NOMBRE, ' ', T4.PER_PATERNO, ' ', T4.PER_MATERNO) CAJA, " & _
    "T5.VENDET_PRODCODIGO, T5.VENDET_ASIENTO ASIENTO, T5.VENDET_PRODUCTONOMBRE, T5.VENDET_PRECIO, SUM(T5.VENDET_CANTIDAD) VENDET_CANTIDAD, T5.venDet_Descuento, T1.VENT_SUBTOTAL, T1.VENT_DESCUENTO, T1.VENT_TOTAL, " & _
    "T1.VENT_PAGOEFECTIVO, VENT_PAGOTARJETA, VENT_PAGOCHEQUE, VENT_PAGADO, VENT_CAMBIO, (SELECT SUM(VENDET_PRECIO * VENDET_CANTIDAD) FROM VENTA_DETALLE WHERE VENDET_FOLIO =  T1.VENT_IDFOLIO and VENDET_STATUS = 'A') VENT_SUBTOTAL1, " & _
    "(IF (T1.VENT_DESCUENTO = 0, (select sum(t4A.venDet_Descuento) from venta_detalle T4A where (t4A.venDet_Folio = t1.vent_IdFolio AND T4A.VENDET_STATUS = 'A')) ,  IF(T1.VENT_DESCUENTO IS NULL, 0, T1.VENT_DESCUENTO)    )) VENT_DESCUENTO1, VENT_OBSERVACIONES, (((SELECT SUM(VENDET_PRECIO * VENDET_CANTIDAD) FROM VENTA_DETALLE WHERE VENDET_FOLIO =  T1.VENT_IDFOLIO AND VENDET_STATUS = 'A')) - ((IF (T1.VENT_DESCUENTO = 0, (select sum(t4A.venDet_Descuento) from venta_detalle T4A where (t4A.venDet_Folio = t1.vent_IdFolio AND T4A.VENDET_STATUS = 'A')) ,  IF(T1.VENT_DESCUENTO IS NULL, 0, T1.VENT_DESCUENTO)    )) )) VENT_TOTAL1, " & _
    "(SELECT T41.TOTAL FROM VIEW_MONEDERO_CLIENTES T41 WHERE T2.PER_ID = T41.PER_ID) MONEDERO, (SELECT SUM(MONEDERO) FROM VIEW_PUNTOS_ADMIN WHERE FOLIO = '" & folio & "' AND TIPO = 'RECIBE') MONE_RECIBE, (SELECT SUM(MONEDERO) FROM VIEW_PUNTOS_ADMIN WHERE FOLIO = '" & folio & "' AND TIPO = 'ENTREGA') MONE_ENTREGA " & _
    "FROM VENTAS T1, PERSONA T2, PERSONA T3, VENTA_DETALLE T5, PERSONA T4 " & _
    "Where T1.VENT_IDFOLIO = '" & folio & "' And T1.VENT_CLIEPERID = T2.PER_ID And T5.VENDET_VENDPERID = T3.PER_ID And T1.VENT_IDFOLIO = T5.VENDET_FOLIO And T1.VENT_VENDPERID = T4.PER_ID and T5.vendet_Status = 'A' " & _
    "GROUP BY T1.vent_fechahora_cobro, T1.VENT_MESA, CONCAT(T2.PER_NOMBRE, ' ', T2.PER_PATERNO, ' ', T2.PER_MATERNO), CONCAT(T3.PER_NOMBRE, ' ', T3.PER_PATERNO, ' ', T3.PER_MATERNO) , CONCAT(T4.PER_NOMBRE, ' ', T4.PER_PATERNO, ' ', T4.PER_MATERNO) , " & _
    "T5.VENDET_PRODCODIGO, T5.VENDET_PRODUCTONOMBRE, T5.VENDET_PRECIO, T5.venDet_Descuento, T1.VENT_PAGOEFECTIVO , VENT_PAGOTARJETA, VENT_PAGOCHEQUE, VENT_PAGADO, VENT_CAMBIO, VENT_SUBTOTAL, VENT_dESCUENTO, VENT_OBSERVACIONES, VENt_TOTAL, T5.VENDET_ASIENTO ORDER BY T5.VENDET_ASIENTO ASC "
    
    Set RES4 = con.Execute(sql1)
    If Not RES4.EOF Then
        SUBTOTAL = RES4.Fields("VENT_SUBTOTAL1")
        DESCUENTO = RES4.Fields("VENT_DESCUENTO1")
        total = RES4.Fields("VENT_TOTAL1")
        PAGOEFECTIVO = RES4.Fields("VENT_PAGOEFECTIVO")
        PAGOTARJETA = RES4.Fields("VENT_PAGOTARJETA")
        PAGOCHEQUE = RES4.Fields("VENT_PAGOCHEQUE")
        PAGADO = RES4.Fields("VENT_PAGADO")
        CAMBIO = RES4.Fields("VENT_CAMBIO")
        OBSERVACIONES = RES4.Fields("VENT_OBSERVACIONES") & ""
    Else
        MsgBox "No se puede imprimir el pre-ticket por que no tiene información referente a la venta referida. " & vbCrLf & vbCrLf & "Posible causa: Falta de información en la lista. " & vbCrLf & vbCrLf & "Verifique.", vbInformation
        Exit Sub
    End If
    
    Printer.KillDoc
    'Printer.FontSize = 12
    'Printer.FontName = "Control"
    'Printer.Print "C"  'open Drawer 1 at 50ms
    Printer.Font = "Courier New"
    Printer.FontSize = 9
    Printer.FontBold = True
    
    'Call Centrar(Nombre, 12)
    Printer.Print UCase(RES3.Fields("SUC_RAZON_SOCIAL"))
    
    'Call Centrar(Sucursal, 12)
    Printer.Print UCase(RES3.Fields("SUC_NOMBRE")) & vbCrLf
    Printer.FontSize = 9
    Printer.FontBold = False

    Printer.Print UCase(RES3.Fields("SUC_DIR_Calle") & " " & RES3.Fields("SUC_DIR_NUM_EXT") & " " & RES3.Fields("SUC_DIR_NUM_INT"))
    Printer.Print UCase(RES3.Fields("SUC_dIR_COLONIA") & " CP:" & RES3.Fields("SUC_DIR_Cp"))
    Printer.Print UCase(RES3.Fields("SUC_DIR_CIUDAD")) '& " " & RES3.Fields("Municipio")
'    Printer.Print "Villahermosa" & " " & "Centro" & " " & "Tabasco"
    Printer.Print "TELS: " & RES3.Fields("SUC_TEL1"); " " & RES3.Fields("SUC_TEL2") & vbCrLf
    Printer.FontSize = 9
    Printer.FontBold = True
    Printer.Print "PRE TICKET - NO PAGADO" & vbCrLf
    Printer.Print "FECHA DE OPERACIÓN: "
    Printer.Print Format(RES4.Fields("vent_fechahorA"), "dddd dd-mm-yyyy") & " " & Format(RES4.Fields("vent_fechahora"), "Short Time") & vbCrLf
    Printer.Print "FOLIO: " & Format(folio, "0000000")
    If IsNull(RES4.Fields("MESA")) = False Then
        Printer.Print "MESA:      " & RES4.Fields("MESA") & ""
        Printer.Print "PERSONAS:  " & RES4.Fields("PERSONAS") & "" & vbCrLf
    End If
    Printer.FontSize = 9
    Printer.FontBold = True
    Printer.Print "CLIENTE: "
    Printer.Print RES4.Fields("CLIENTE") & vbCrLf
    Printer.FontBold = False
    Printer.Print "MOSTRADOR: "
    Printer.Print RES4.Fields("CAJA")
    Printer.Print "- - - - - - - - - - - - - - -"
    Printer.Print "DETALLE DE OPERACIÓN:" & vbCrLf
    'For b1 = 1 To FrmFocus.ListaOper.Rows - 1
    ASIENTO = ""
    Do While Not RES4.EOF
        clave = RES4.Fields("VENDET_PRODCODIGO")
        If Len(clave) > 17 Then
            clave = Left(clave, 17)
        Else
            clave = clave & String(17 - Len(clave), " ")
        End If
        Nombre2 = RES4.Fields("VENDET_PRODUCTONOMBRE")
        If Len(Nombre2) > 28 Then
            Nombre2 = Left(Nombre2, 28)
        Else
            Nombre2 = Nombre2 & String(28 - Len(Nombre2), " ")
        End If
        cantidad = RES4.Fields("VENDET_CANTIDAD")
        Precio = RES4.Fields("VENDET_PRECIO")
        desc = RES4.Fields("venDet_Descuento")
        tot = (RES4.Fields("VENDET_CANTIDAD") * RES4.Fields("VENDET_PRECIO") - RES4.Fields("VENDET_DESCUENTO"))
        If Len(Precio) > 9 Then
            Precio = Left(Precio, 9)
        End If
        If Len(RES4.Fields("USUARIO")) > 28 Then
            atendio = Left(RES4.Fields("USUARIO"), 28)
        Else
            atendio = RES4.Fields("USUARIO")
        End If
        
'        cantidad = cantidad & String(6 - Len(cantidad), " ")
'        PRECIO = PRECIO & String(9 - Len(PRECIO), " ")
'        desc = desc & String(9 - Len(desc), " ")
''        Tot = Tot & String(9 - Len(Tot), " ")
'        Titulo1 = "Cant" & String(6 - Len("Cant"), " ")
'        Titulo2 = "Precio" & String(3, " ")
'        Titulo3 = "Desc" & String(5, " ")
'        Titulo4 = "Total" & String(4, " ")
'        Titulo5 = "Atendio" & String(12, " ")
'
'        Printer.Print clave & vbCrLf & nombre2 & vbCrLf & Titulo1 & " " & Titulo2 & " " & Titulo3 & " " & Titulo4 & _
'        vbCrLf & cantidad & " " & PRECIO & " " & desc & " " & Tot & vbCrLf & Titulo5 & vbCrLf & Atendio & vbCrLf
        
        cantidad = cantidad & String(10 - Len(cantidad), " ")
        Precio = Precio & String(12 - Len(Precio), " ")
        desc = FormatCurrency(desc) & String(11 - Len(FormatCurrency(desc)), " ")
'        Tot = Tot & String(9 - Len(Tot), " ")
        titulo1 = "Cant" & String(10 - Len("Cant"), " ")
        titulo2 = "Precio" & String(10, " ")
        titulo3 = "Desc" & String(6, " ")
        titulo4 = "Total" & String(4, " ")
        titulo5 = "Atendio" & String(12, " ")
        
        Printer.FontBold = True
        
        If ASIENTO <> RES4.Fields("ASIENTO") Then
            Printer.Print "----------------------------------------"
            Printer.FontSize = 11
            Printer.FontBold = True
            Printer.Print "Asiento: " & RES4.Fields("ASIENTO")
            ASIENTO = RES4.Fields("ASIENTO")
        
        Else
        End If
        
        Printer.FontSize = 9
        Printer.FontBold = True
        Printer.Print UCase(clave)
        Printer.Print UCase(Nombre2)
        Printer.FontBold = False
        Printer.Print UCase(titulo1) & " " & UCase(titulo2)
        Printer.FontBold = True
        Printer.Print cantidad & " " & FormatCurrency(Precio)
        Printer.FontBold = False
        Printer.Print titulo3 & " " & titulo4
        Printer.FontBold = True
        Printer.Print desc & FormatCurrency(tot)
'        Printer.FontBold = False
'        Printer.Print Titulo5 & vbCrLf & UCase(Atendio) & vbCrLf
                
    RES4.MoveNext
    Loop
    
    
    Printer.Print "- - - - - - - - - - - - - - - - - - "
    If Len(Horario) > 0 Then
        Printer.FontSize = 9
        Printer.Print Horario
        Printer.Print "- - - - - - - - - - - - - - - "
    End If
    Printer.FontBold = True
    
    If IVA = "S" Then
'        Printer.Print "  SUB TOTAL: " & FormatCurrency((total / (1.16)))
        Printer.Print "  SUB TOTAL: " & FormatCurrency((SUBTOTAL / (1.16)))
        Printer.Print "  DESCUENTO: " & FormatCurrency(DESCUENTO)
        Printer.Print "        IVA: " & FormatCurrency((total) - (total / (1.16)))
    Else
'        Printer.Print "  SUB TOTAL: " & FormatCurrency((total) - (DESCUENTO))
        Printer.Print "  SUB TOTAL: " & FormatCurrency((SUBTOTAL))
        Printer.Print "  DESCUENTO: " & FormatCurrency(DESCUENTO)
'        Printer.Print "        IVA: " & FormatCurrency(0)
    End If
        
    Printer.Print "      TOTAL: " & FormatCurrency(total)
    Printer.Print "- - - - - - - - - - - - - - - "
    Printer.FontBold = True
    If Len(OBSERVACIONES) > 0 Then
        Printer.Print "           OBSERVACIONES    "
        Printer.FontSize = 9
        Lineas = Split(OBSERVACIONES, vbNewLine)
    
        For b1 = 0 To UBound(Lineas)
            For c1 = 0 To ((Round((Len(Lineas(b1)) / 35))) + 1)
                If Len(Lineas(b1)) >= 35 Then
                    Printer.Print Left(Lineas(b1), 35)
                    Lineas(b1) = Right(Lineas(b1), (Len(Lineas(b1)) - 35))
                Else
                    Printer.Print Lineas(b1)
                    Exit For
                End If
            Next c1
        Next b1
    End If
        
    Printer.FontSize = 9
    Printer.FontBold = False
    
    Printer.Print vbCrLf & UCase(RES3.Fields("SUC_SLOGAN")) & vbCrLf
    Printer.Print UCase(RES3.Fields("SUC_PAGINA_WEB"))
    Printer.Print UCase(RES3.Fields("SUC_EMAIL"))

    impresionRenglones (RES3.Fields("SUC_INFORMACION"))

'    Lineas = Split(RES3.Fields("SUC_INFORMACION"), vbNewLine)
'    For b1 = 0 To UBound(Lineas)
'        For c1 = 0 To ((Round((Len(Lineas(b1)) / 35))) + 1)
'            If Len(Lineas(b1)) >= 35 Then
'                'largo = True
'                Printer.Print Left(Lineas(b1), 35)
'                Lineas(b1) = Right(Lineas(b1), (Len(Lineas(b1)) - 35))
'            Else
'                Printer.Print Lineas(b1)
'                Exit For
'            End If
'        Next c1
'    Next b1
    
    
    Printer.FontBold = True

    
    Printer.Print vbCrLf & "PRE TICKET - NO PAGADO"
    'Printer.Print RES3.Fields("SUC_INFORMACION") & vbCrLf & vbCrLf
'    Printer.FontSize = 12
'    Printer.FontName = "Control"
'    Printer.Print "P"  'Cut
    Printer.EndDoc
    

End Sub

