VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Upload recipe to SAT"
   ClientHeight    =   14085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21360
   OleObjectBlob   =   "Upload.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub CommandButton1_Click() 'Calculate required etch time [s]: BUTTON

Worksheets("SAT.calc").Range("C13").Value = Me.ComboBox1.Value 'Product/Etch type:

Worksheets("SAT.calc").Range("C14").Value = Me.TextBox2.Value 'Copper thickness [um]:

Worksheets("SAT.calc").Range("C15").Value = Me.TextBox3.Value 'Etch rate [um/min]:

Worksheets("SAT.calc").Range("C16").Value = Me.TextBox4.Value 'PR width [um]:

Worksheets("SAT.calc").Range("C17").Value = Me.TextBox5.Value 'Target element width [um]:

On Error Resume Next
Me.TextBox6.Value = Round(Worksheets("SAT.calc").Range("C18").Value, 0)




End Sub








Private Sub CommandButton5_Click() 'Run Test button


    Dim fso As Object
    Dim sourceFile As String
    Dim destinationFile As String
    

 Application.ScreenUpdating = False

ThisWorkbook.Worksheets("Sheet1").Range("A1:K100").ClearContents

'COPY 03-test.30s.Ch2.xml FILE from folder1 to folder2
'***************************************************************************************************************
    ' Initialize FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Define source and destination file paths
    sourceFile = "J:\ShareENG\Basanel\DASHBOARD - Yellow\Dashboard - SAT\Test_Recipe\03-test.30s.Ch2.xml"
    destinationFile = "J:\ShareENG\Basanel\DASHBOARD - Yellow\Dashboard - SAT\03-test.30s.Ch2.xml"
    
    ' Check if the destination file exists and delete it if it does
    If fso.FileExists(destinationFile) Then
        fso.DeleteFile destinationFile
    End If
    
    ' Copy the file
    fso.CopyFile sourceFile, destinationFile
    
    ' Clean up
    Set fso = Nothing
'***************************************************************************************************************

 
 
Call ParseXMLToExcelTable


Last_Row = ThisWorkbook.Worksheets("Sheet1").Cells(Rows.Count, "A").End(xlUp).Row
With UserForm1.ListBox2
.ColumnCount = 7
.ColumnHeads = True
.ColumnWidths = "50,100,50,50,100,100,100"
.RowSource = "Sheet1!A2:H" & Last_Row
End With




 Application.ScreenUpdating = True




End Sub






Private Sub CommandButton2_Click() 'Create Recipe button

    Application.ScreenUpdating = False


ThisWorkbook.Worksheets("Sheet1").Range("A1:K100").ClearContents


    folderPath = "J:\ShareENG\Basanel\DASHBOARD - Yellow\Dashboard - SAT\"
    
    ' Get the first file in the folder
    fileName = Dir(folderPath & "*.xml*")

    
    ' Set the file path
    filePath = folderPath & fileName
    

    ' Check if the file exists before attempting to delete it "J:\ShareENG\Basanel\DASHBOARD - Yellow\Dashboard - SAT"
    If fileName <> "" Then
        ' Delete .xml file only in
        Kill filePath

    End If



                'Check ER value is acceptable
                    Last_Row = Worksheets("SAT.calc").Cells(Rows.Count, "O").End(xlUp).Row
                    Last_ER_value = Worksheets("SAT.calc").Cells(Last_Row, 15).Value
                    If Int(Worksheets("SAT.calc").Cells(Last_Row, 14).Value) = Date Then
                    
                            If Last_ER_value > 1.2 Or Last_ER_value < 1 Then
                            
                                If UserForm1.Label24.Caption = "0" Then
                                MsgBox "ER value is not in limits, 1 < ER < 1.2. Call Eng"
                                Exit Sub
                                Else
                                MsgBox "קצב איכול אינו עומד בגבולות 1 ל- 1.2 מיקרון/דקה, יש לקרוא למהנדס"
                                Exit Sub
                                End If
                                   
                            End If
                    Else
                    
                        If UserForm1.Label24.Caption = "0" Then
                            MsgBox "Please run for " & Date & " ER evaluation test, and try again"
                            Exit Sub
                        Else
                            MsgBox "יש להריץ בדיקת טסט להיום ולנסות שנית"
                            Exit Sub
                        End If
                    End If







Call UpdateSpecificStepTimes

Call ParseXMLToExcelTable

' Fill ListBox2 with New Created Recipe
Last_Row = ThisWorkbook.Worksheets("Sheet1").Cells(Rows.Count, "A").End(xlUp).Row
With ListBox2
.ColumnCount = 7
.ColumnHeads = True
.ColumnWidths = "50,100,50,50,100,100,100"
.RowSource = "Sheet1!A2:H" & Last_Row
End With




 Application.ScreenUpdating = True


ThisWorkbook.Save

End Sub


Private Sub CommandButton3_Click() 'Upload button


    Dim response As VbMsgBoxResult
    Dim fso As Object
    Dim sourcePath As String
    Dim destinationPath As String
    Dim fileName As String


    
    ' Display the message box
    If UserForm1.Label24.Caption = "0" Then
        response = MsgBox("Do you really want upload the recipe to the SAT?", vbYesNo + vbQuestion, "Last confirmation")
    Else
        response = MsgBox("האם אתה בטוח לשלוח את הרסיפי למכונה?", vbYesNo + vbQuestion, "אישור סופי")
    End If
    
' Check the user's response
If response = vbYes Then
        

 

' MoveXMLFile

    folderPath = "J:\ShareENG\Basanel\DASHBOARD - Yellow\Dashboard - SAT\"
    
    ' Get the first xml file in the folder
    fileName = Dir(folderPath & "*.xml*")
    
    OldfolderPath = folderPath & fileName

    fileName = Replace(fileName, ".xml", "_" & Now & ".xml")
    
    fileName = Replace(fileName, "/", "-")
    
    fileName = Replace(fileName, ":", "-")
    
    NewfolderPath = "J:\ShareENG\Basanel\DASHBOARD - Yellow\Dashboard - SAT\" & fileName
 

    Name OldfolderPath As NewfolderPath
    
    ' Set the source and destination paths
    sourcePath = folderPath & fileName
    'destinationPath = "\\CLASSONE-PC\Users\Public\Documents\Copley Motion\CME 2\New created Recipe for Ch2\"
     destinationPath = "W:\Sat-recipe\"


    DeleteAllFilesInFolder (destinationPath) 'Delete all files in "W:\Sat-recipe"
    
    
    

    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Check if the file exists in the source folder
    If fso.FileExists(sourcePath) Then
        ' Move the file to the destination folder
        fso.MoveFile sourcePath, destinationPath & fileName
            If UserForm1.Label24.Caption = "0" Then
            MsgBox "Recipe " & fileName & " uploaded successfully to SAT"
            Else
            MsgBox " רסיפי הועלה בהצלחה למכונה. יש לכניס סירה עם פרוסות ולהתחיל תהליך האיכול"
            End If
            Call UserForm1.Reset_Button_Click
            
    Else
            If UserForm1.Label24.Caption = "0" Then
            MsgBox "Recipe was not created"
            Exit Sub
            Else
            MsgBox "רסיפי לא נוצר, יש לקרוא למהנדס במקרה שהשגיאה חוזרת"
            End If
    End If





Call Update_Log_File






Me.ListBox2.RowSource = ""
ThisWorkbook.Worksheets("Sheet1").Range("A1:K100").ClearContents

   End If
End Sub



' FilterAndShowEtchData and put in the UserForm2


Private Sub CommandButton6_Click()

Application.ScreenUpdating = False

    Dim ws As Worksheet
    Dim ws_RowSource As Worksheet
    Dim mode_value As Long ' מספר השכיח ביותר בעמודה E ב Sheets(RowSource)
    
    Set ws = ThisWorkbook.Sheets("Log_file")
    Set ws_RowSource = ThisWorkbook.Sheets("RowSource")
    
    Dim lastRow As Long
    Dim last_RowSource_row As Long
    
    ws_RowSource.Range("A2:O1000").ClearContents
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    last_RowSource_row = ws_RowSource.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1

    Dim criteriaProduct As String, criteriaSize As String
    Dim criteriaElValue As String, criteriaStep As String
    Dim i As Long, matchRow As Boolean

    criteriaProduct = UserForm1.Product_TextBox.Value
    criteriaSize = UserForm1.Size_TextBox.Value
    criteriaElValue = UserForm1.ElValue_TextBox.Value
    criteriaStep = UserForm1.Proc_TextBox.Value

    With UserForm2.ListBox1
        .Clear
        .ColumnCount = 15
    End With
    
    
    

    For i = 2 To lastRow
        matchRow = True
        
        If criteriaProduct <> "" And ws.Cells(i, 10).Value <> criteriaProduct Then matchRow = False
        If criteriaSize <> "" And ws.Cells(i, 12).Text <> criteriaSize Then matchRow = False
        If criteriaElValue <> "" And CStr(ws.Cells(i, 13).Value) <> criteriaElValue Then matchRow = False
        
        
        If criteriaStep <> "" And ws.Cells(i, 14).Text <> criteriaStep Then matchRow = False

        If matchRow Then
            Dim rowData(1 To 15) As String
            Dim j As Integer
            For j = 1 To 15
                rowData(j) = ws.Cells(i, j).Text
                ws_RowSource.Cells(last_RowSource_row, j) = rowData(j)
                'rowData(j) = ws.Cells(i, j).Text
            Next j
            last_RowSource_row = last_RowSource_row + 1
            'UserForm2.ListBox1.AddItem Join(rowData, vbTab)
            
        End If
    Next i
    
    
If last_RowSource_row > 2 Then
    With UserForm2.ListBox1
    .ColumnCount = 15
    .ColumnHeads = True
    .ColumnWidths = "50,50,50,5,60,60,100,50,60,60,100,40,50,50,200"
    .RowSource = "RowSource!A2:O" & (last_RowSource_row - 1)
    End With

End If

'------------------------------------------------------------------------------------------------

'Calculate Mode of the Worksheets("RowSource").Range("E1:E1000")  Column("Etch Time [sec]")
Dim modeValue As Variant
Dim meanValue As Double
Dim rng As Range
Dim thickness_total As Double
Dim thickness_avg As Double
Dim Denominator As Long

thickness_total = 0
thickness_avg = 0
Denominator = 0

If last_RowSource_row > 2 Then

    
    Set rng = ws_RowSource.Range(ws_RowSource.Cells(2, 5), ws_RowSource.Cells(last_RowSource_row - 1, 5))
    
    On Error Resume Next
    modeValue = Application.WorksheetFunction.Mode(rng)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        meanValue = Application.WorksheetFunction.Median(rng)
        
        If IsError(meanValue) Or IsEmpty(meanValue) Then
            modeValue = "N/A"
        Else
            modeValue = Round(meanValue, 1)
        End If
    Else
        On Error GoTo 0
        modeValue = Round(modeValue, 1)
    End If
    
    
    ' Thickness of Cu layer vs. etch time dependecy
    
    For i = 2 To last_RowSource_row
    
        If ws_RowSource.Cells(i, 5).Value = modeValue Then
        
        Dim rawVal As Variant
        Dim cleaned As String

        rawVal = ws_RowSource.Cells(i, 9).Value
        cleaned = Trim(Replace(rawVal, "micron", "", , , vbTextCompare))

        thickness_total = thickness_total + CDbl(cleaned)

        
        'thickness_total = thickness_total + CDbl(Trim(Replace(ws_RowSource.(Cells(i, 9).Value), "micron", "", , , vbTextCompare)))
        Denominator = Denominator + 1
        
        End If
    
    
    Next i
    
    thickness_avg = thickness_total / Denominator
    
        If UserForm1.TextBox2.Value = "" Then
        
        UserForm1.lastEtch_TextBox.Value = modeValue
        UserForm2.EtchTime_TextBox.Value = modeValue
        
        Else
    
        UserForm1.lastEtch_TextBox.Value = Int((modeValue / thickness_avg) * UserForm1.TextBox2.Value)
        UserForm2.EtchTime_TextBox.Value = Int((modeValue / thickness_avg) * UserForm1.TextBox2.Value)
        
        End If
    'UserForm1.lastEtch_TextBox.Value = modeValue
    'UserForm2.EtchTime_TextBox.Value = modeValue

Else
    
    UserForm1.lastEtch_TextBox.Value = 0
    UserForm2.EtchTime_TextBox.Value = 0

End If

'------------------------------------------------------------------------------------------------





    UserForm2.Show vbModeless

End Sub

        




Private Sub imgChart_Click()


Dim CurrentChart As Chart
Dim FName As String


ChartName = "ER_Chart"

FName = ThisWorkbook.Path & "\temp.gif"

Set CurrentChart = ThisWorkbook.Sheets("SAT.calc").ChartObjects(ChartName).Chart
CurrentChart.Export fileName:=FName, FilterName:="GIF"




    ' Set Image1 size to match UserForm3
    With UserForm3.Image1
        .Width = UserForm3.InsideWidth
        .Height = UserForm3.InsideHeight
        .PictureSizeMode = fmPictureSizeModeStretch ' Stretch the image to fit the control
    End With




UserForm3.Image1.Picture = LoadPicture(FName)



 UserForm3.Show vbModeless


End Sub





Private Sub CommandButton4_Click() 'Calculate ER [um/min] Button



Dim lastRow As Integer


Init_thick = TextBox7.Value
Final_thick = TextBox8.Value


If Init_thick = "" Or Final_thick = "" Then

    If UserForm1.Label24.Caption = "0" Then
    MsgBox "Please fill Initinal and Final thickness and try again"
    Exit Sub
    Else
    MsgBox "יש להקליד עובי נחושת התחלתי וסופי ואז לנסות שוב"
    Exit Sub
    End If

End If




Me.TextBox9.Value = (Init_thick - Final_thick) * 2 'ER



    If TextBox9.Value > 1.2 Or TextBox9.Value < 1 Then
    TextBox9.ForeColor = RGB(255, 0, 0)
        If UserForm1.Label24.Caption = "0" Then
        MsgBox "ER value is not in limits, 1 < ER < 1.2. Call Eng"
        Else
        MsgBox "קצב איכול אינו עומד בגבולות 1 ל- 1.2, יש לקרוא למהנדס"
        End If
    Else
    TextBox9.ForeColor = RGB(0, 255, 0) 'Green Color
    
    End If


lastRow = ThisWorkbook.Worksheets("SAT.calc").Cells(Worksheets("SAT.calc").Rows.Count, "O").End(xlUp).Row + 1

Worksheets("SAT.calc").Cells(lastRow, "O").Value = TextBox9.Value 'ER
Worksheets("SAT.calc").Cells(lastRow, "N") = Now ' DATE

Worksheets("SAT.calc").Cells(lastRow, "P").Value = TextBox7.Value 'Initial copper thickness [um]:
Worksheets("SAT.calc").Cells(lastRow, "Q").Value = TextBox8.Value 'final copper thickness [um]:


Me.TextBox10.Value = Round(Worksheets("SAT.calc").Range("R2").Value, 2)
Me.TextBox11.Value = Round(Worksheets("SAT.calc").Range("S2").Value, 2)




Call CheckAndDeleteERChart

Call CreateERTrendChart

Call image_toChart

UserForm1.Repaint

ThisWorkbook.Save

End Sub





Sub Reset_Button_Click()

Frame1.Visible = False
Lot_TextBox.Value = ""
TextBox1.Value = ""


Product_TextBox = ""
ESN_TextBox.Value = ""
Size_TextBox.Value = ""
ElValue_TextBox.Value = ""
Proc_TextBox.Value = ""

Me.ComboBox1.Text = "Select Product"

TextBox2.Value = ""
TextBox3.Value = ""
TextBox4.Value = ""
TextBox5.Value = ""
TextBox6.Value = ""
lastEtch_TextBox.Value = ""

End Sub


Sub Search_Button_Click()



Dim lot As String
Dim rng As Range
Dim row_i As Integer

Application.ScreenUpdating = False


lot = Lot_TextBox.Value

    If lot = "" Then
    
    Exit Sub
    End If


lot = Replace(Lot_TextBox.Value, "$M", "")

'ThisWorkbook.Queries("BATCHLST").Refresh

'ActiveWorkbook.RefreshAll
'Worksheets("BATCHLST").ListObjects("BATCHLST").Refresh


Set rng = ThisWorkbook.Worksheets("BATCHLST").Range("A:A").Find(What:=lot, LookAt:=xlWhole)  'Check if the lot number exsists in BATCHLST

If rng Is Nothing Then
    If UserForm1.Label24.Caption = "0" Then
    MsgBox "Lot # is incorrect or finish all processes, please try again", vbExclamation, "Error"
    Exit Sub
    Else
    MsgBox "מספר מנה לא תקין או סיים את כל שלבי הייצור, בבקשה הזן מחדש"
    Exit Sub
    End If
End If



Frame1.Visible = True

row_i = rng.Row



Product_TextBox = ThisWorkbook.Worksheets("BATCHLST").Cells(row_i, 12)
ESN_TextBox.Value = ThisWorkbook.Worksheets("BATCHLST").Cells(row_i, 2)
Size_TextBox.Value = Format(ThisWorkbook.Worksheets("BATCHLST").Cells(row_i, 4).Value, "0000")
ElValue_TextBox.Value = ThisWorkbook.Worksheets("BATCHLST").Cells(row_i, 5)
Proc_TextBox.Value = ThisWorkbook.Worksheets("BATCHLST").Cells(row_i, 8).Text
'Product_TextBox
'ESN_TextBox
'Size_TextBox
'ElValue_TextBox
'Proc_TextBox



 'Unload UserForm
 'Load UserForm
 'UserForm.Show



Application.ScreenUpdating = True

End Sub

Private Sub UserForm_Initialize()

Frame1.Visible = False


ListBox3.ColumnWidths = "100;150"

Me.ComboBox1.RowSource = "SAT.calc!F2:F10"
ComboBox1.Text = "Select Product"


Me.TextBox10.Value = Round(Worksheets("SAT.calc").Range("R2").Value, 2)
Me.TextBox11.Value = Round(Worksheets("SAT.calc").Range("S2").Value, 2)


Call image_toChart


        Image2.Picture = LoadPicture("J:\ShareENG\Basanel\DASHBOARD - Yellow\Dashboard - SAT\ISRAEL.jpg") 'Israel flag
        Image2.PictureSizeMode = fmPictureSizeModeStretch
        Image3.Picture = LoadPicture("J:\ShareENG\Basanel\DASHBOARD - Yellow\Dashboard - SAT\USA.jpg") 'USA flag
        Image3.PictureSizeMode = fmPictureSizeModeStretch


End Sub




Private Sub Image2_Click() 'Hebrew Language


UserForm1.Label24.Caption = "1"


UserForm1.Caption = "יצירת תוכנית להרצה במכונת איכול"

ComboBox1.Text = "בחר מוצר"


Label1.Caption = "יצירת תוכנית חדשה והעלאה למכונת איכול אוטומטית"

Label9.Caption = "הכנס מספר מנה"

Label8.Caption = "הכנס מספר פרוסות שהולכים לרוץ במכונה"

Search_Button.Caption = "חיפוש"

Reset_Button.Caption = "איפוס"

CommandButton5.Caption = "הרצת טסט"

CommandButton6.Caption = "הראה זמני הרצה אחרונים"

Label2.Caption = "סוג מוצר:"

Label3.Caption = "עובי נחושת [מיקרון]:"

Label4.Caption = "קצת איכול [מיקרון/דקה]:"

Label5.Caption = "עובי רזיסט [מיקרון]:"

Label6.Caption = "רוחב האלמנט [מיקרון]:"

CommandButton1.Caption = "חשב זמן איכול הנדרש [ש']"

CommandButton2.Caption = "צור תוכנית הרצה"

Label21.Caption = "עובי נחושת התחלתי [מיקרון]:"

Label20.Caption = "עובי נחושת סופי [מיקרון]:"

CommandButton4.Caption = "חשב קצב איכול [מיקרון/דקה]"

Label22.Caption = "ממוצע:"

Label23.Caption = "סטיית תקן:"



Frame1.Label19.Caption = "מוצר"
Frame1.Label17.Caption = "גודל"
Frame1.Label16.Caption = "ערך חשמלי"
Frame1.Label15.Caption = "תהליך"



End Sub




Private Sub Image3_Click() 'English Language


UserForm1.Label24.Caption = "0"


UserForm1.Caption = "Upload recipe to SAT"

ComboBox1.Text = "Select Product"


Label1.Caption = "Create and upload recipe to SAT Copper Tool"

Label9.Caption = "Enter Lot #"

Label8.Caption = "How many wafers will run in Ch 2?"

Search_Button.Caption = "Search"

Reset_Button.Caption = "Reset"

CommandButton5.Caption = "Run TEST"

CommandButton6.Caption = "Show last etch time [s]:"

Label2.Caption = "Product/Etch type:"

Label3.Caption = "Cu thickness [um]:"

Label4.Caption = "Etch rate [um/min]:"

Label5.Caption = "PR width [um]:"

Label6.Caption = "Target element width [um]:"

CommandButton1.Caption = "Calculate required etch time [s]:"

CommandButton2.Caption = "Create recipe"

Label21.Caption = "Initial copper thickness [um]:"

Label20.Caption = "Final copper thickness [um]:"

CommandButton4.Caption = "Calculate ER [um/min]"

Label22.Caption = "Mean:"

Label23.Caption = "STD:"


Frame1.Label19.Caption = "Product"
Frame1.Label17.Caption = "Size"
Frame1.Label16.Caption = "Electrical value"
Frame1.Label15.Caption = "Proc."



End Sub


