Attribute VB_Name = "FCN_EnviandoMail"
'public WithEvents oMail As envioMail



Dim oMail As clss_EnvioMail
Dim msjErroresMail As String
Dim numClie As Long
Dim correoClie As String
Dim infoMail As Boolean
Dim sql1 As String
Dim ResMail As Recordset
Dim mailAsunto As String
Dim mensajeTipo As String
Public adjuntoDir As String
Private Sub checkMailInfo()
    sql1 = "SELECT COUNT(*) NUM FROM  SUCURSAL WHERE SUC_MAIL_CORREO IS NOT NULL"
    Set ResMail = con.Execute(sql1)
    If Not ResMail.EOF Then
        If Val(ResMail.Fields("num")) > 0 Then
            infoMail = True
            sql1 = "SELECT * FROM SUCURSAL WHERE SUC_MAIL_CORREO IS NOT NULL"
            Set ResMail = con.Execute(sql1)
        Else
            infoMail = False
        End If
    End If
End Sub

Public Sub enviar_Mail(mailTipo As String, mailEncabezado As String, mailCorreoClie As String, mailInfo As String)
    'On Error Resume Next
    checkMailInfo
    
    If mailTipo = "CAJA" Then
        'infoMail = False
        
        If infoMail = True Then
            Dim b1 As Long, num As Long
            
            msjErroresMail = ""
            
            
            mensajeTipo = mailInfo
            checkTipoMensaje (mailTipo)
            'correoClie = mailCorreoClie
            mailAsunto = mailEncabezado
            
        Else
            MsgBox "No se puede enviar el correo por falta de información de configuración", vbInformation
            Exit Sub
        End If
    Else
        If mailTipo = "CITA" Then
'            infoMail = False
            If infoMail = True Then
                
                msjErroresMail = ""
                
                
                mensajeTipo = mailInfo
                'checkTipoMensaje (mailTipo)
                correoClie = mailCorreoClie
                mailAsunto = mailEncabezado
                
            Else
                MsgBox "No se puede enviar el correo por falta de información de configuración", vbInformation
                Exit Sub
            End If
        Else
            If mailTipo = "MENSAJES" Then
            
                mensajeTipo = mailInfo
                'checkTipoMensaje (mailTipo)
                correoClie = mailCorreoClie
                mailAsunto = mailEncabezado
            Else
                If mailTipo = "COMPRA" Then
                
                    mensajeTipo = mailInfo
                    'checkTipoMensaje (mailTipo)
                    correoClie = mailCorreoClie
                    mailAsunto = mailEncabezado
                
                End If
            End If
        End If
    End If

    EnviandoMail
    


End Sub
Private Sub checkTipoMensaje(tipo As String)
    Dim infoLista As String
    Dim infoAgrupado As String
    
    If tipo = "CAJA" Then
        FRM_Caja.txtInfo.Text = ""
        FRM_Caja.txtInfo.Text = "Corte de caja " & vbCrLf & vbCrLf & _
        "Fecha: " & Format(Date, "Long Date") & ", Hora: " & Format(Time, "Short Time") & vbCrLf & vbCrLf & _
        "Usuario:  " & FRM_Menu.menuBarra2.Panels(5).Text & vbCrLf & vbCrLf & _
        "Resumen general de la caja: " & vbCrLf
        
        For b1 = 1 To FRM_Caja.Lista.Rows
            FRM_Caja.txtInfo.Text = FRM_Caja.txtInfo.Text & vbCrLf & _
            FRM_Caja.Lista.TextMatrix(b1 - 1, 0) & String(15 - Len(Left(FRM_Caja.Lista.TextMatrix(b1 - 1, 0), 10)), "               ") & _
            FRM_Caja.Lista.TextMatrix(b1 - 1, 1) & String(15 - Len(FRM_Caja.Lista.TextMatrix(b1 - 1, 1)), "               ") & _
            FRM_Caja.Lista.TextMatrix(b1 - 1, 2) & String(8 - Len(FRM_Caja.Lista.TextMatrix(b1 - 1, 2)), "        ")
        Next b1
        
        FRM_Caja.txtInfo.Text = FRM_Caja.txtInfo.Text & vbCrLf & vbCrLf & _
        "Resumen agrupado por insumo: " & vbCrLf
        For b1 = 1 To FRM_Caja.lista5.Rows
            FRM_Caja.txtInfo.Text = FRM_Caja.txtInfo.Text & vbCrLf & _
            FRM_Caja.lista5.TextMatrix(b1 - 1, 0) & String(15 - Len(FRM_Caja.lista5.TextMatrix(b1 - 1, 0)), "               ") & _
            FRM_Caja.lista5.TextMatrix(b1 - 1, 1) & String(35 - Len(Left(FRM_Caja.lista5.TextMatrix(b1 - 1, 1), 35)), "                                   ") & _
            FRM_Caja.lista5.TextMatrix(b1 - 1, 2) & String(20 - Len(FRM_Caja.lista5.TextMatrix(b1 - 1, 2)), "                    ") & _
            FRM_Caja.lista5.TextMatrix(b1 - 1, 3) & String(15 - Len(FRM_Caja.lista5.TextMatrix(b1 - 1, 3)), "               ") & _
            FRM_Caja.lista5.TextMatrix(b1 - 1, 4) & String(8 - Len(FRM_Caja.lista5.TextMatrix(b1 - 1, 4)), "        ") & _
            FRM_Caja.lista5.TextMatrix(b1 - 1, 5) & String(15 - Len(FRM_Caja.lista5.TextMatrix(b1 - 1, 5)), "               ") & _
            FRM_Caja.lista5.TextMatrix(b1 - 1, 6) & String(10 - Len(FRM_Caja.lista5.TextMatrix(b1 - 1, 6) & ""), "          ")
        Next b1
 
        FRM_Caja.txtInfo.Text = FRM_Caja.txtInfo.Text & vbCrLf & vbCrLf & _
        "Membresias asignadas: " & vbCrLf
        For b1 = 1 To FRM_Caja.ListaMbr.Rows
            FRM_Caja.txtInfo.Text = FRM_Caja.txtInfo.Text & vbCrLf & _
            FRM_Caja.ListaMbr.TextMatrix(b1 - 1, 0) & String(7 - Len(FRM_Caja.ListaMbr.TextMatrix(b1 - 1, 0)), "       ") & _
            FRM_Caja.ListaMbr.TextMatrix(b1 - 1, 1) & String(12 - Len(FRM_Caja.ListaMbr.TextMatrix(b1 - 1, 1)), "            ") & _
            FRM_Caja.ListaMbr.TextMatrix(b1 - 1, 2) & String(12 - Len(FRM_Caja.ListaMbr.TextMatrix(b1 - 1, 2)), "            ") & _
            FRM_Caja.ListaMbr.TextMatrix(b1 - 1, 3) & String(30 - Len(FRM_Caja.ListaMbr.TextMatrix(b1 - 1, 3)), "                              ") & _
            Left(FRM_Caja.ListaMbr.TextMatrix(b1 - 1, 4), 30) & String(30 - Len(Left(FRM_Caja.ListaMbr.TextMatrix(b1 - 1, 4), 30)), "                              ") & _
            FRM_Caja.ListaMbr.TextMatrix(b1 - 1, 5) & String(30 - Len(FRM_Caja.ListaMbr.TextMatrix(b1 - 1, 5)), "                              ") & _
            FRM_Caja.ListaMbr.TextMatrix(b1 - 1, 6) & String(5 - Len(FRM_Caja.ListaMbr.TextMatrix(b1 - 1, 6)), "     ") & _
            FRM_Caja.ListaMbr.TextMatrix(b1 - 1, 7) & String(10 - Len(FRM_Caja.ListaMbr.TextMatrix(b1 - 1, 7)), "          ")
        Next b1
                
        FRM_Caja.txtInfo.Text = FRM_Caja.txtInfo.Text & vbCrLf & vbCrLf & _
        "Gastos realizados: " & vbCrLf
        For b1 = 1 To FRM_Caja.listaGST.Rows
            FRM_Caja.txtInfo.Text = FRM_Caja.txtInfo.Text & vbCrLf & _
            FRM_Caja.listaGST.TextMatrix(b1 - 1, 0) & String(7 - Len(FRM_Caja.listaGST.TextMatrix(b1 - 1, 0)), "       ") & _
            FRM_Caja.listaGST.TextMatrix(b1 - 1, 1) & String(27 - Len(FRM_Caja.listaGST.TextMatrix(b1 - 1, 1)), "                           ") & _
            Left(FRM_Caja.listaGST.TextMatrix(b1 - 1, 2), 25) & String(25 - Len(Left(FRM_Caja.listaGST.TextMatrix(b1 - 1, 2), 25)), "                         ") & _
            Left(FRM_Caja.listaGST.TextMatrix(b1 - 1, 3), 21) & String(21 - Len(Left(FRM_Caja.listaGST.TextMatrix(b1 - 1, 3), 21)), "                     ") & _
            FRM_Caja.listaGST.TextMatrix(b1 - 1, 4) & String(15 - Len(FRM_Caja.listaGST.TextMatrix(b1 - 1, 4)), "        ") & _
            FRM_Caja.listaGST.TextMatrix(b1 - 1, 5) & String(15 - Len(FRM_Caja.listaGST.TextMatrix(b1 - 1, 5)), "               ") & _
            Left(FRM_Caja.listaGST.TextMatrix(b1 - 1, 6), 45) & String(45 - Len(Left(FRM_Caja.listaGST.TextMatrix(b1 - 1, 6), 45)), "                                             ")
        Next b1
                
        correoClie = ResMail.Fields("SUC_MAIL_CORREO")
        mensajeTipo = FRM_Caja.txtInfo.Text
    End If

    If tipo = "Cita" Then
        'correoClie = ResMail.Fields("SUC_MAIL_CORREO")
        mensajeTipo = "Gracias por su preferencia. " & vbCrLf & vbCrLf & "Su cita se ha generado satisfactoriamente." & _
        vbCrLf & vbCrLf & "Detalle de su cita: " & vbCrLf & vbclrf & _
        a
    End If
End Sub
Private Sub EnviandoMail()
    'Set oMail = New envioMail
    Set oMail = New clss_EnvioMail
    
    With oMail
         'datos para enviar
        .servidor = ResMail.Fields("SUC_MAIL_SMTP")
        .puerto = ResMail.Fields("SUC_MAIL_PUERTO")
        
        .UseAuntentificacion = "" & ResMail.Fields("SUC_MAIL_AUTEN") & ""
        .ssl = "" & ResMail.Fields("SUC_MAIL_SSL") & ""
        
        .Usuario = ResMail.Fields("SUC_MAIL_USUARIO")
        .PassWord = ResMail.Fields("SUC_MAIL_PASS")
        
        .Asunto = mailAsunto
        .de = ResMail.Fields("SUC_MAIL_CORREO")
'        .de = ResMail.Fields("SUC_NOMBRE")
        If adjuntoDir <> "" Then
            .Adjunto = adjuntoDir
        End If
        If correoClie <> "" Then
            .para = correoClie
        Else
            .para = ResMail.Fields("SUC_MAIL_CORREO")
        End If
        .mensaje = mensajeTipo
        .Enviar_Backup ' manda el mail
    
    
'        Open App.Path & "\LogErrMail.txt" For Output As #1
'        Print #1, Date & "  " & Time & " Envio de Mensajes por correo: " & vbCrLf & vbCrLf & Error
'        Close #1
        Set oMail = Nothing
    End With

End Sub

' envio completo
Private Sub oMail_EnvioCompleto()
    MsgBox "Mensaje enviado", vbInformation
End Sub
' error al enviar
Private Sub oMail_Error(Descripcion As String, Numero As Variant)
    MsgBox Descripcion, vbCritical, Numero
End Sub
