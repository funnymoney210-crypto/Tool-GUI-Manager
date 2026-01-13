Attribute VB_Name = "Module1"

Sub UpdateSpecificStepTimes()
    Dim xmlDoc As Object
    Dim xmlNode As Object
    Dim xmlNodes As Object
    Dim filePath As String
    Dim newFilePath As String
    Dim rng As Range
    Dim WO As String
    

    ETCH1T = UserForm1.lastEtch_TextBox.Value
    Num_of_pl = UserForm1.TextBox1.Value
    WO = UserForm1.Lot_TextBox.Text
    Cu_Thick = UserForm1.TextBox2.Text
   

    If Cu_Thick = "" Then
        If UserForm1.Label24.Caption = "0" Then
        MsgBox "Please fill the required Cu thickness required to etch"
        Exit Sub
        Else
        MsgBox "יש למלא עובי נחושת הנדרש לאיכול במכונה"
        Exit Sub
        End If
    
    End If
   
    If ETCH1T = "" Or Num_of_pl = "" Or WO = "" Then
    
        If UserForm1.Label24.Caption = "0" Then
        MsgBox "Please fill the required Lot#, etch time and # of wafers"
        Exit Sub
        Else
        MsgBox "יש להקליד מספר מנה, זמן איכול הנדרש ומספר פרוסות ואז לנסות שנית"
        Exit Sub
        End If
    End If
    
    If Num_of_pl > 24 Then
        If UserForm1.Label24.Caption = "0" Then
        MsgBox "Run more than 24 wafers is not acceptable. "
        Exit Sub
        Else
        MsgBox "הרצה של יותר מ-24 פרוסות אינה אפשרית, אנא נסה שנית"
        Exit Sub
        End If
    End If
    
    
    If ETCH1T <= 15 Then
    
    DUMPT1C = 0
    
    Else
   
    Set rng = ThisWorkbook.Worksheets("Sheet3").Range("A2:A30").Find(Num_of_pl)
    
    'DUMPT1C = rng.Offset(, 1).Value
    DUMPT1C = Cu_Thick * Num_of_pl * 0.8
    DUMPT1C = Round(DUMPT1C)
    
    If DUMPT1C > 40 Then
    
    DUMPT1C = 40
    
    End If
    
    ETCH1T = ETCH1T - DUMPT1C
    
        If ETCH1T < 0 Then
        
        ETCH1T = UserForm1.lastEtch_TextBox.Value
        DUMPT1C = 0
        
        End If
    
    End If
    
    
    
    
    
    
    
    Call UserForm1.Search_Button_Click
    
    
     
     
    
    


' Description =  Cu Etch = 15; Refresh = 0; Lot = 854865$M_4;Inductors;L0201R82AHS00Y;201;0.82;608
    
Description = "Cu Etch = " & ETCH1T & ";" & "Refresh = " & DUMPT1C & ";" & "Lot = " & WO & "_" & Num_of_pl & ";" & "Cu_Thick = " & Cu_Thick & ";" & UserForm1.Product_TextBox.Value & ";" & UserForm1.ESN_TextBox.Value & ";" & UserForm1.Size_TextBox.Value & ";" & UserForm1.ElValue_TextBox.Value & ";" & UserForm1.Proc_TextBox.Text
    
    
        ' Set the file paths
    filePath = "J:\ShareENG\Basanel\SAT tool DATA\DataFiles 3\Recipes\Cu-150sec.Ch2.xml"
    newFilePath = "J:\ShareENG\Basanel\DASHBOARD - Yellow\Dashboard - SAT\" & "Cu-" & DUMPT1C + ETCH1T & "sec.Ch2." & WO & "_" & Num_of_pl & ".xml"
    
      
    
    ' Create the XML document object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    xmlDoc.async = False
    xmlDoc.Load (filePath)
    
    ' Check if the XML file is loaded successfully
    If xmlDoc.parseError.ErrorCode <> 0 Then
        MsgBox "Error loading XML file: " & xmlDoc.parseError.reason
        Exit Sub
    End If
    
    
    ' Update the Description value
    Set xmlNode = xmlDoc.SelectSingleNode("//Description")
    If Not xmlNode Is Nothing Then
        xmlNode.Text = Description
    End If
    
    
    
    
    
    
    ' Find all Step elements
    Set xmlNodes = xmlDoc.getElementsByTagName("Step")
    For Each stepNode In xmlNodes
        ' Check if the StepDescription is DUMPT1C
        If stepNode.SelectSingleNode("StepDescription").Text = "DUMPT1C" Then
            ' Update the StepTime value
            stepNode.SelectSingleNode("StepTime").Text = DUMPT1C
        End If
        
        If stepNode.SelectSingleNode("StepDescription").Text = "ETCH1T" Then
            ' Update the StepTime value
            stepNode.SelectSingleNode("StepTime").Text = ETCH1T
        End If
       
        
        
        
    Next stepNode
    
    ' Save the updated XML to a new file
    xmlDoc.Save newFilePath
    
    'MsgBox "StepTime value updated successfully!"
End Sub















