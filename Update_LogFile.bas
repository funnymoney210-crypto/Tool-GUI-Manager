Attribute VB_Name = "Update_LogFile"
Sub Update_Log_File()


Dim Log_ws As Worksheet '\\avx-files\share\ShareENG\Yellow\Etch process\Etch Process.xls
Dim ws As Worksheet     ' Thisworkkbook
Dim Description As String
Dim cols As Variant
Dim maxRow As Long
Dim colLetter As String
Dim WWID As Long
Dim WWNAME As String
Dim Last_Row As Long
Dim lastRow As Long
Dim operator_Num As String
Dim opName As String



'Get WWID parameter
'WWNAME
'--------------------------------------------------------------------------------------------
IsValid = False

    Do While Not IsValid
        operator_Num = InputBox("Please fill the operator's number")
        If operator_Num = "" Then Exit Sub ' User pressed cancel or left blank
        
        Select Case operator_Num
            Case "155": opName = "ברוך": IsValid = True
            Case "303": opName = "נועם": IsValid = True
            Case "503": opName = "רן": IsValid = True
            Case "705": opName = "איליה": IsValid = True
            Case "1313": opName = "אלי": IsValid = True
            'Case "1333": opName = "דימטרי מ.": IsValid = True
            'Case "1406": opName = "מקסים נ.": IsValid = True
            Case "1528": opName = "אולגה": IsValid = True
            Case "1532": opName = "איליה ס.": IsValid = True
            'Case "1709": opName = "אלכס פ.": IsValid = True
            'Case "2111": opName = "יהודה ש.": IsValid = True
            Case "1234": opName = "": IsValid = True
            Case Else
                MsgBox "מספר לא מזוהה, נסה שוב.", vbExclamation
        End Select
    Loop
    
'--------------------------------------------------------------------------------------------



Workbooks.Open ("\\avx-files\share\ShareENG\Yellow\Etch process\Etch Process.xls")


Set Log_ws = Workbooks("Etch Process.xls").Worksheets("מעקב מנות SAT נחושת")
Set ws = ThisWorkbook.Worksheets("Sheet1")

Description = ws.Cells(2, 8) 'Cu Etch = 58;Refresh = 6;Lot = 850725_8;Cu_Thick = 1;Couplers;DB0603N2140AN00F;603;2140;196.5

Log_ws.Select


lastRow = ThisWorkbook.Worksheets("Log_file").Cells(ThisWorkbook.Worksheets("Log_file").Rows.Count, 1).End(xlUp).Row ' נותן שורה אחרונה בלשונית Log_file
lastRow = lastRow + 1

'Find the maximum Last_Row from columns "A", "B", "C", "D", "E", "G", "I", "J", "K", "L", "M", "N", "O", "P"
    cols = Array("A", "B", "C", "D", "E", "G", "I", "J", "K", "L", "M", "N", "O", "P")
    maxRow = 0

    For i = LBound(cols) To UBound(cols)
        colLetter = cols(i)
        lastR = Log_ws.Cells(Log_ws.Rows.Count, colLetter).End(xlUp).Row
        'Debug.Print "Last row in column " & colLetter & ": " & lastR
        If lastR > maxRow Then maxRow = lastR
    Next i
Last_Row = maxRow + 1







Log_ws.Cells(Last_Row, 1).Value = Date
Log_ws.Cells(Last_Row, 4).Value = Format(Time, "hh:mm")

ThisWorkbook.Worksheets("Log_file").Cells(lastRow, 1).Value = Date
ThisWorkbook.Worksheets("Log_file").Cells(lastRow, 2).Value = Format(Time, "hh:mm")
ThisWorkbook.Worksheets("Log_file").Cells(lastRow, 4).Value = Description
ThisWorkbook.Worksheets("Log_file").Cells(lastRow, 3).Value = opName



If Description = "Copper Etch 30sec" Then

Log_ws.Cells(Last_Row, 5).Value = "test"
Log_ws.Cells(Last_Row, 11).Value = opName
Log_ws.Cells(Last_Row, 7).Value = 1
ThisWorkbook.Worksheets("Log_file").Cells(lastRow, 4).Value = Description
ThisWorkbook.Worksheets("Log_file").Cells(lastRow, 3).Value = opName
GoTo SAVE_Log_ws

End If




'Cu Etch = 58;Refresh = 6;Lot = 850725_8;Cu_Thick = 1;Couplers;DB0603N2140AN00F;603;2140;196.5

                                                                       
        ' Split the cell value by semicolon
        splitValues = Split(Description, ";")
        
        ' Column P = column 16
        ' Place the split values in the adjacent columns startinf from column P
        For i = LBound(splitValues) To UBound(splitValues)
        
        If i = 0 Then '"Cu Etch = "
        temp = splitValues(i)
        temp = Replace(temp, "Cu Etch = ", "")      'Delete all the "Cu Etch = " "Refresh = " "Lot = " "Cu_Thick = "
        
        Log_ws.Cells(Last_Row, 9).Value = temp
        ThisWorkbook.Worksheets("Log_file").Cells(lastRow, 5).Value = temp
        
        End If
        
        If i = 1 Then '"Refresh = "
        temp = splitValues(i)
        temp = Replace(temp, "Refresh = ", "")    'Delete all the "Cu Etch = " "Refresh = " "Lot = " "Cu_Thick = "
        
        Log_ws.Cells(Last_Row, 16).Value = temp
        ThisWorkbook.Worksheets("Log_file").Cells(lastRow, 6).Value = temp
        End If
        
        If i = 2 Then '"Lot = "
        temp = splitValues(i)
        temp = Replace(temp, "Lot = ", "")          'Delete all the "Cu Etch = " "Refresh = " "Lot = " "Cu_Thick = "
        temp = Replace(temp, "$M", "")          ' Delete "$M"
        Log_ws.Cells(Last_Row, 7).Value = Split(temp, "_")(1)        'This uses the Split function to divide the string at _ and takes the second part (index 1) and put in column פר במנה.
        ThisWorkbook.Worksheets("Log_file").Cells(lastRow, 7).Value = Split(temp, "_")(1)
        
        Log_ws.Cells(Last_Row, 5).Value = Split(temp, "_")(0)
        ThisWorkbook.Worksheets("Log_file").Cells(lastRow, 8).Value = Split(temp, "_")(0)
        
        
        End If
        
        If i = 3 Then '"Cu_Thick = "
        temp = splitValues(i)
        temp = Replace(temp, "Cu_Thick = ", "")    'Delete all the "Cu Etch = " "Refresh = " "Lot = " "Cu_Thick = "
        Log_ws.Cells(Last_Row, 13).Value = temp & "micron"
        ThisWorkbook.Worksheets("Log_file").Cells(lastRow, 9).Value = temp & "micron"
        End If
        
        If i = 4 Then 'Product name
        Log_ws.Cells(Last_Row, 17).Value = splitValues(i)
        ThisWorkbook.Worksheets("Log_file").Cells(lastRow, 10).Value = splitValues(i)
        End If
        
        If i = 5 Then 'ESN
        Log_ws.Cells(Last_Row, 18).Value = splitValues(i)
        ThisWorkbook.Worksheets("Log_file").Cells(lastRow, 11).Value = splitValues(i)
        End If
        
        If i = 6 Then 'Size
        Log_ws.Cells(Last_Row, 19).Value = splitValues(i)
        ThisWorkbook.Worksheets("Log_file").Cells(lastRow, 12) = "'" & CStr(splitValues(i))
        End If
        
        If i = 7 Then 'ערך
        Log_ws.Cells(Last_Row, 20).Value = splitValues(i)
        ThisWorkbook.Worksheets("Log_file").Cells(lastRow, 13).Value = splitValues(i)
        End If
        
        If i = 8 Then 'מספר שלב
        Log_ws.Cells(Last_Row, 21) = CStr(splitValues(i))
        Log_ws.Cells(Last_Row, 21).NumberFormat = "0.0"
        ThisWorkbook.Worksheets("Log_file").Cells(lastRow, 14).Value = "'" & CStr(splitValues(i))
        'ThisWorkbook.Worksheets("Log_file").Cells(lastRow, 14).NumberFormat = "0.0"
        
                ' Search for the value in column B in Sheet "RPQC06V1"
            Set foundCell = ThisWorkbook.Worksheets("RPQC06V1").Columns("B").Find(What:=Log_ws.Cells(Last_Row, 21).Text, LookIn:=xlValues, LookAt:=xlWhole)
            
                If Not foundCell Is Nothing Then
                    Log_ws.Cells(Last_Row, 22).Value = foundCell.Offset(0, -1).Value  ' Value from column A of the same row
                    ThisWorkbook.Worksheets("Log_file").Cells(lastRow, 15).Value = foundCell.Offset(0, -1).Value
                Else
                   Log_ws.Cells(Last_Row, 22).Value = ""
                   ThisWorkbook.Worksheets("Log_file").Cells(lastRow, 15).Value = ""
                   
                End If
            

        End If
        
        
      
        Next i
                                                           
                                                           
Log_ws.Cells(Last_Row, 9).Value = Log_ws.Cells(Last_Row, 9).Value + Log_ws.Cells(Last_Row, 16).Value
Log_ws.Cells(Last_Row, 11).Value = opName ' Operator name

ThisWorkbook.Worksheets("Log_file").Cells(lastRow, 5).Value = ThisWorkbook.Worksheets("Log_file").Cells(lastRow, 5).Value + ThisWorkbook.Worksheets("Log_file").Cells(lastRow, 6).Value
                                                           
SAVE_Log_ws: '\\avx-files\share\ShareENG\Yellow\Etch process\Etch Process.xls

Application.DisplayAlerts = False
'Workbooks("Etch Process.xls").Save
Workbooks("Etch Process.xls").Close SaveChanges:=True
ThisWorkbook.Save
Application.DisplayAlerts = True

End Sub
