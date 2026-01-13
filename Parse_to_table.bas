Attribute VB_Name = "Parse_to_table"
Sub ParseXMLToExcelTable()
    Dim xmlDoc As Object
    Dim xmlNode As Object
    Dim xmlNodes As Object
    Dim stepNode As Object
    Dim filePath As String
    Dim ws As Worksheet
    Dim i As Integer
    
    
    folderPath = "J:\ShareENG\Basanel\DASHBOARD - Yellow\Dashboard - SAT\"
    
    ' Get the first file in the folder
    fileName = Dir(folderPath & "*.xml*")

    
    ' Set the file path
    filePath = folderPath & fileName
    
    ' Create the XML document object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    xmlDoc.async = False
    xmlDoc.Load (filePath)
    
    ' Check if the XML file is loaded successfully
    If xmlDoc.parseError.ErrorCode <> 0 Then
        'MsgBox "Error loading XML file: " & xmlDoc.parseError.reason
        Exit Sub
    End If
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    ws.Cells.Clear ' Clear existing data
    
    ' Add headers
    ws.Cells(1, 1).Value = "SeqNo"
    ws.Cells(1, 2).Value = "StepDescription"
    ws.Cells(1, 3).Value = "StepTime"
    ws.Cells(1, 4).Value = "SpinSpeed"
    ws.Cells(1, 5).Value = "SprayManifold1"
    ws.Cells(1, 6).Value = "SprayManifold3"
    ws.Cells(1, 7).Value = "DrainManifold"
    ws.Cells(1, 8).Value = "Description"

    
    ' Find all Step elements
    Set xmlNodes = xmlDoc.getElementsByTagName("Step")
    i = 2 ' Start from the second row
    For Each stepNode In xmlNodes
        ws.Cells(i, 1).Value = stepNode.SelectSingleNode("SeqNo").Text
        ws.Cells(i, 2).Value = stepNode.SelectSingleNode("StepDescription").Text
        ws.Cells(i, 3).Value = stepNode.SelectSingleNode("StepTime").Text
        ws.Cells(i, 4).Value = stepNode.SelectSingleNode("SpinSpeed").Text
        ws.Cells(i, 5).Value = stepNode.SelectSingleNode("SprayManifold1").Text
        ws.Cells(i, 6).Value = stepNode.SelectSingleNode("SprayManifold3").Text
        ws.Cells(i, 7).Value = stepNode.SelectSingleNode("DrainManifold").Text
        i = i + 1
    Next stepNode
    
Set xmlNodes = xmlDoc.SelectSingleNode("//Description")
    ws.Cells(2, 8).Value = xmlNodes.Text
    
    'MsgBox "XML data parsed to Excel table successfully!"
End Sub

