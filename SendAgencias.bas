Attribute VB_Name = "SendaGENCIAS"
Private Function GetTables(xRange As String) As String
    Dim xlSheet As Worksheet
    Dim tableRange As Range
    Dim imgTable As String
    Dim imgObject As ChartObject
    
    Set xlSheet = ThisWorkbook.Sheets("DETALLE")
    xlSheet.Activate
    Set tableRange = xlSheet.Range(xRange)
    tableRange.CopyPicture
    Set imgObject = xlSheet.ChartObjects.Add(tableRange.Left, tableRange.Top, _
                    tableRange.width, tableRange.height)
    imgObject.Activate
    With imgObject.Chart
        .Paste
        .Export Environ("temp") & "\tabla" & ".bmp"
    End With
    imgObject.Delete
    imgTable = Environ("temp") & "\tabla" & ".bmp"
    
    GetTables = imgTable
End Function

Private Function GetGraphs(Optional nGraphs As Integer = 6) As Variant
    Dim xlSheet As Worksheet
    Dim graphPath As String
    Dim graphObjet As ChartObject
    Dim imgGraphsArr() As Variant
    ReDim imgGraphsArr(nGraphs + 1)
    
    Set xlSheet = ThisWorkbook.Sheets("DETALLE")
    xlSheet.Activate
    For I = 0 To nGraphs
        Set graphObject = xlSheet.ChartObjects("grafico" & I)
        graphPath = Environ("temp") & "\" & "grafica" & I & ".bmp"
        graphObject.Chart.Export graphPath
        imgGraphsArr(I) = Environ("temp") & "\" & "grafica" & I & ".bmp"
    Next I
    
    GetGraphs = imgGraphsArr
End Function

Private Function GetEmails(x As String) As String
    Dim xlSheet As Worksheet
    Dim emails As String
    Dim nRange As Integer
    
    Set xlSheet = ThisWorkbook.Sheets("CORREOS")
    xlSheet.Activate
    
    nRange = Range(x, Range(x).End(xlDown)).Count
    emails = Range(x).Value
    For I = 1 To nRange - 1
        emails = emails & "; "
        emails = emails & Range(x).Offset(I, 0).Value
    Next I
    
    GetEmails = emails
End Function

Private Function GetSubject() As String
    Dim xlSheet As Worksheet
    Dim xDia, xMes, xSubject As String
    
    xDia = Format(Date, "dd")
    xMes = Format(Date, "mmmm")
    xMes = Application.WorksheetFunction.Proper(xMes)
    
    Set xlSheet = ThisWorkbook.Sheets("DETALLE")
    xlSheet.Activate
    xHora = Format(xlSheet.Range("I5"), "hh:mm")
    
    xSubject = "[PRIVADO]Seguimiento Intervalos AGENCIAS LATAM OS | " & _
    xDia & " de " & xMes & Space(1) & xHora & " Hrs"
    
    GetSubject = xSubject
End Function

Private Function GetHTMLBody(Optional nGraphs = 5, Optional xWidth = 320, Optional xHeight = 200) As String
    Dim textoHTML As String
    
    textoHTML = "<Body><b> Cordial saludo,</b>" & "<br><br>" & _
                "Seguimiento Intervalos AGENCIAS LATAM" & "<br><br>" & _
                "<b>DETALLE POR SERVICIO</b>" & "<br><br>" & _
                "<img src=""cid:tabla.bmp"" width=650 height=350>" & "<br><br>" & _
                "<b>GRAFICO AGENCIAS</b>" & "<br><br>"
                
    For I = 0 To 2
        textoHTML = textoHTML & _
        "<img src=""cid:grafica" & I & ".bmp"" width=500 height=280>" & _
        "&nbsp;&nbsp;&nbsp;"
    Next I
    
    textoHTML = textoHTML & "<br><br>" & "<b>DISTRIBUCION DE MERCADOS</b>" & "<br><br>"
    For J = 3 To 4
        textoHTML = textoHTML & _
        "<img src=""cid:grafica" & J & ".bmp"" width=xWidth height=xHeight>" & _
        "&nbsp;&nbsp;"
    Next J
    
    textoHTML = textoHTML & "<br>"
    For K = 5 To 6
        textoHTML = textoHTML & _
        "<img src=""cid:grafica" & K & ".bmp"" width=xWidth height=xHeight>" & _
        "&nbsp;&nbsp;"
    Next K
    
    textoHTML = textoHTML & "<br><br><br>" & GetSignature & "</Body>"
    
    GetHTMLBody = textoHTML
End Function

Private Function GetSignature() As String
    Dim xFSO, xTextStream As Object
    Dim sigDir, sigPath, xSignature, xFiles As String
    If Environ("username") = "" Then
        sigDir = Environ("appdata") & "\Microsoft\Signatures"
        sigPath = sigDir & "\" & "firma.htm"
        Set xFSO = CreateObject("Scripting.FileSystemObject")
        Set xTextStream = xFSO.OpenTextFile(sigPath)
        xSignature = xTextStream.ReadAll
        xFiles = "firma_archivos/"
        xSignature = Replace(xSignature, xFiles, sigDir & "\" & xFiles)
    Else
        sigDir = Environ("appdata") & "\Microsoft\Signatures"
        sigPath = sigDir & "\" & Environ("username") & ".htm"
        Set xFSO = CreateObject("Scripting.FileSystemObject")
        Set xTextStream = xFSO.OpenTextFile(sigPath)
        xSignature = xTextStream.ReadAll
        xFiles = Replace(Environ("username"), ".htm", "") & "_archivos/"
        xSignature = Replace(xSignature, xFiles, sigDir & "\" & xFiles)
    End If
    
    GetSignature = xSignature
End Function

Public Sub SendEmails()
Attribute SendEmails.VB_ProcData.VB_Invoke_Func = "W\n14"
    Dim outApp As New Outlook.Application
    Dim outMail As Object
    Set outMail = outApp.CreateItem(0)
    
    Dim varTables As String
    Dim varGraphs As Variant
    Dim varTo, varCC As String
    Dim varHTMLBody, varSubject As String
    
    varTables = GetTables("B3:G20")
    varGraphs = GetGraphs()
    varTo = GetEmails("B3")
    varCC = GetEmails("E3")
    varHTMLBody = GetHTMLBody()
    varSubject = GetSubject()
    
    adjExcel = ThisWorkbook.Path & "\" & ThisWorkbook.Name
    varErr = Err.Description
    
    ThisWorkbook.Sheets("DETALLE").Activate
    If varErr = "" Then
        Dim outATS As Outlook.Attachments
        Set outATS = outMail.Attachments
        Set outAttach = outATS.Add(varTables)
        
        For I = 0 To 6
            Set outAttach = outATS.Add(varGraphs(I))
        Next
        
        With outMail
            .To = varTo
            .CC = varCC
            .Subject = varSubject
            .HTMLBody = varHTMLBody
            .Attachments.Add (adjExcel)
            .SendUsingAccount outApp.Session.Accounts("reportes@almacontactcol.co")
        End With
        
        Response = MsgBox("Desea ver el cuerpo del mensaje antes de enviar ?", vbQuestion + vbYesNo, "CONTROL GTR")
        If Response = vbYes Then
            outMail.Display
        Else
            outMail.Send
        End If
        
        Set outlookApp = Nothing
        Set outMail = Nothing
        Set outATS = Nothing
        Set outAttach = Nothing
        
        mensaje = "Mensaje enviado correctamente"
    Else
        mensaje = "Ha ocurrido un error, por favor vuelve a intentarlo"
    End If
    
    MsgBox mensaje
End Sub
