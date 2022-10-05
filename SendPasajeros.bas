Attribute VB_Name = "SendPasajeros"
Private Function GetTables(xRange As String, x As Integer) As String
    Dim xlSheet As Worksheet
    Dim tableRange As Range
    Dim imgTable As String
    Dim imgObject As ChartObject
    
    Set xlSheet = ThisWorkbook.Sheets("DETALLE")
    xlSheet.Activate
    Set tableRange = xlSheet.Range(xRange)
    tableRange.CopyPicture
    Set imgObject = xlSheet.ChartObjects.Add(tableRange.Left, tableRange.Top, _
                    tableRange.Width, tableRange.Height)
    imgObject.Activate
    With imgObject.Chart
        .Paste
        .Export Environ("temp") & "\tabla" & x & ".bmp"
    End With
    imgObject.Delete
    imgTable = Environ("temp") & "\tabla" & x & ".bmp"
    
    GetTables = imgTable
End Function

Private Function GetGraphs(Optional nGraphs As Integer = 2) As Variant
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
    xHora = Format(xlSheet.Range("M5"), "hh:mm")
    
    xSubject = "[PRIVADO]Seguimiento Intervalos LATAM_PASAJEROS | " & _
    xDia & " de " & xMes & Space(1) & xHora & " Hrs"
    
    GetSubject = xSubject
End Function

Private Function GetHTMLBody(Optional nGraphs = 2) As String
    Dim textoHTML As String
    
    textoHTML = "<Body> Cordial saludo," & "<br><br>" & _
                "Seguimiento Intervalos LATAM_PASAJEROS" & "<br><br>" & _
                "<img src=""cid:tabla1.bmp"" width=1280 height=423>" & "<br><br>" & _
                "GRAFICO PASAJEROS" & "<br><br>"
                
    For I = 0 To nGraphs
        textoHTML = textoHTML & _
        "<img src=""cid:grafica" & I & ".bmp"" width=885 height=402>" & _
        "&nbsp;&nbsp;&nbsp;"
    Next I
    
    textoHTML = textoHTML & "<br><br>" & "RESUMEN CHAT LATAM" & "<br><br>" & _
    "<img src=""cid:tabla2.bmp"" width=700 height=423>" & "<br><br>" & _
    "<br><br><br>" & GetSignature & "</Body>"
    
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
Attribute SendEmails.VB_ProcData.VB_Invoke_Func = "S\n14"
    Dim outApp As New Outlook.Application
    Dim outMail As Object
    Set outMail = outApp.CreateItem(0)
    
    Dim varTables As String
    Dim varGraphs As Variant
    Dim varTo, varCC As String
    Dim varHTMLBody, varSubject As String
    
    varTable1 = GetTables("B3:K23", 1)
    varTable2 = GetTables("B3:K23", 2)
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
        Set outAttach = outATS.Add(varTable1)
        Set outAttach = outATS.Add(varTable2)
        
        For I = 0 To 2
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
