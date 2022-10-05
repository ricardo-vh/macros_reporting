Attribute VB_Name = "SendMensual"
Private Function GetTables(xSheets As Variant, xCicle) As Variant
    Dim xlSheet As Worksheet
    Dim tableRange As Range
    Dim imgTablesArr() As Variant
    Dim imgObject As ChartObject
    
    ReDim imgTablesArr(UBound(xSheets))
    For I = LBound(xSheets) To UBound(xSheets)
        Set xlSheet = ThisWorkbook.Sheets(xSheets(I))
        xlSheet.Activate
        
        Set tableRange = xlSheet.Range("B3:B17", Range("B3:B17").Offset(0, xCicle + 1))
        tableRange.CopyPicture
        Set imgObject = xlSheet.ChartObjects.Add(tableRange.Left, tableRange.Top, _
                        tableRange.Width, tableRange.Height)
        
        imgObject.Activate
        With imgObject.Chart
            .Paste
            .Export Environ("temp") & "\tabla" & I & ".bmp"
        End With
        imgObject.Delete
        
        imgTablesArr(I) = Environ("temp") & "\tabla" & I & ".bmp"
    Next I
    
    GetTables = imgTablesArr
End Function

Private Function GetGraphs(xSheets As Variant) As Variant
    Dim xlSheet As Worksheet
    Dim graphPath As String
    Dim graphObjet As ChartObject
    Dim imgGraphsArr() As Variant
    Dim nDim, counter As Integer
    
    nDim = (UBound(xSheets) * 2) + 1
    ReDim imgGraphsArr(nDim)
    counter = 0
    For I = 0 To UBound(xSheets)
        Set xlSheet = ThisWorkbook.Sheets(xSheets(I))
        xlSheet.Activate
        
        For J = 0 To 1
            Set graphObject = xlSheet.ChartObjects("grafico" & I & J)
            graphPath = Environ("temp") & "\" & "grafico" & I & J & ".bmp"
            graphObject.Chart.Export graphPath
            imgGraphsArr(counter) = Environ("temp") & "\" & "grafico" & I & J & ".bmp"
            counter = counter + 1
        Next J
    Next I
    
    GetGraphs = imgGraphsArr
End Function

Function GetEmails(x As String) As String
    Dim emails As String
    Dim nRange As Integer
    
    ThisWorkbook.Sheets("CORREOS").Activate
    nRange = Range(x, Range(x).End(xlDown)).Count
    emails = Range(x).Value
    For I = 1 To nRange - 1
        emails = emails & "; "
        emails = emails & Range(x).Offset(I, 0).Value
    Next I
    
    GetEmails = emails
End Function

Function GetSubject() As String
    Dim xDia, xMes, xSubject As String
    
    xDia = Format(Date - 1, "dd")
    xMes = Format(Date - 1, "mmmm")
    xMes = Application.WorksheetFunction.Proper(xMes)
    
    xSubject = "Control Mensual LATAM " & xDia & " de " & xMes & " de 2022"
    GetSubject = xSubject
End Function

Function GetSignature() As String
    Dim xFSO, xTextStream As Object
    Dim sigDir, sigPath, xSignature, xFiles As String
    
    sigDir = Environ("appdata") & "\Microsoft\Signatures"
    sigPath = sigDir & "\" & Environ("username") & ".htm"
    Set xFSO = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    Set xTextStream = xFSO.OpenTextFile(sigPath)
    xSignature = xTextStream.ReadAll
    xFiles = Replace(Environ("username"), ".htm", "") & "_archivos/"
    xSignature = Replace(xSignature, xFiles, sigDir & "\" & xFiles)
    
    GetSignature = xSignature
End Function

Function GetHTMLBody(xSheets, xCicle, xSize) As String
    Dim textoHTML As String
    Dim xMes, xDia As String
    
    xDia = Format(Date - 1, "dd")
    xMes = Format(Date - 1, "mmmm")
    xMes = Application.WorksheetFunction.Proper(xMes)
    
    textoHTML = "<Body>" & "Cordial Saludo" & "<br><br>" & _
    "Control mensual de indicadores, actualizado al " & xDia & Space(1) & _
    xMes & " de 2022 - Dentro del consolidado ya se encuentra LUA." & "<br><br>" & _
    "Consolidado" & "<br><br>" & _
    "<img src=""cid:tabla" & ".bmp"" width=1280 height=255>" & "<br><br>"

    For I = LBound(xSheets) To UBound(xSheets)
        textoHTML = textoHTML & _
        "RESUMEN " & xSheets(I) & "<br><br>" & _
        "<img src=""cid:tabla" & I & ".bmp"" width=" & xSize(xCicle) & "height=361>" & "<br><br>" & _
        "GRAFICO " & xSheets(I) & "<br><br>"
        
        For J = 0 To 1
            textoHTML = textoHTML & _
            "<img src=""cid:grafico" & I & J & ".bmp"" width=800 height=350>" & "&nbsp;&nbsp;&nbsp;"
        Next J
        textoHTML = textoHTML & "<br><br>"
    Next I
    
    textoHTML = textoHTML & "<br><br>" & GetSignature & "</Body>"
    
    GetHTMLBody = textoHTML
End Function

Public Sub SendEmails()
Attribute SendEmails.VB_ProcData.VB_Invoke_Func = "F\n14"
    Dim outApp As New Outlook.Application
    Dim outMail As Object
    Dim outPA As Outlook.PropertyAccessor
    Dim outATS As Outlook.Attachments
    Dim outAT As Outlook.Attachment
    
    Dim xSheets As Variant
    Dim xSize As Variant
    Dim varTablesArr, varGraphsArr As Variant
    Dim xCicle As Integer
    Dim varTO, varCC, varSubject, varHTMLBody As String
    Dim varAccount As String
    Dim adjExcel, varErr, mensaje As String
    
    Set outMail = outApp.CreateItem(0)
    Set outATS = outMail.Attachments
    
    Set xlSheet = ThisWorkbook.Sheets("CONSOLIDADO")
    xlSheet.Activate
    Set tableRange = xlSheet.Range("B3:T18")
    tableRange.CopyPicture
    Set imgObject = xlSheet.ChartObjects.Add(tableRange.Left, tableRange.Top, _
                    tableRange.Width, tableRange.Height)
    imgObject.Activate
    With imgObject.Chart
        .Paste
        .Export Environ("temp") & "\tabla" & ".bmp"
    End With
    imgObject.Delete
    imgTableImg = Environ("temp") & "\tabla" & ".bmp"
    
    xSheets = Array("PASAJEROS", "AGENCIAS", "LUA", "SAG5", "LUA ENG", "SAG15", _
    "SAG16", "VENTAS", "TRAVEL", "TARGET ESP", "TARGET ENG", "AGENCIAS PORTUGUES", "EMPRESAS")
    xCicle = InputBox("Ingrese el numero de ciclo a enviar", "Sistema de ingreso", Default)
    xSize = Array(0, 330, 440, 550, 660)
        
    imgTablesArr = GetTables(xSheets, xCicle)
    imgGraphsArr = GetGraphs(xSheets)
    varTO = GetEmails("B3")
    varCC = GetEmails("E3")
    varSubject = GetSubject()
    varHTMLBody = GetHTMLBody(xSheets, xCilcle, xSize)
    varAccount = "reportes@almacontactcol.co"
    
    adjExcel = ThisWorkbook.Path & "\" & ThisWorkbook.Name
    varErr = Err.Description
    If varErr = "" Then
        Set outAttach = outATS.Add(imgTable)
        For I = LBound(imgTablesArr) To UBound(imgTablesArr)
            Set outAttach = outATS.Add(imgTablesArr(I))
        Next I
        
        For J = LBound(imgGraphsArr) To UBound(imgGraphsArr)
            Set outAttach = outATS.Add(imgGraphsArr(J))
        Next J
        
        With outMail
            .To = varTO
            .CC = varCC
            .Subject = varSubject
            .HTMLBody = varHTMLBody
            .Attachments.Add (adjExcel)
            .SendUsingAccount = varAccount
            .Display
        End With
        
        Set outlookApp = Nothing
        Set outMail = Nothing
        mensaje = "Mensaje enviado correctamente"
    Else
        mensaje = "ERROR" & vbNewLine & "Por favor intentalo nuevamente"
    End If
    
    MsgBox mensaje
End Sub
