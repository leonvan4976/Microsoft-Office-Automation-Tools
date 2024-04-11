Attribute VB_Name = "Module1"
Option Explicit
Public myDictionary As Object
Public Mydate As String
Public MydateDisplay As String
Sub OpenOrderAutomation()
    Dim num As Double
    Dim tarCell As String
    Dim cell As Range
    Dim targetCell As Range
    Dim lastRow As Long
    
    Dim sheetName As String
    sheetName = ActiveSheet.Name
    
    ' Set today's date
    Mydate = Date$
    MydateDisplay = Replace(Date$, "-", "/")
    
     'Initialize dictionary if its empty
    If myDictionary Is Nothing Then
        Set myDictionary = CreateObject("Scripting.Dictionary")
    End If
    
    ReadExcel ("\\SERVER2\Tech\UPS_Reference\UPS_CSV_EXPORT.csv")
    ReadPDFWithPyPDF2
    ReadQuickbookReport
    
    'Initialize the lastRow
    lastRow = Cells(Rows.Count, "B").End(xlUp).Row
    'MsgBox lastRow
    
    'Cells("410", "B").value = "999"
    
    ' Loop through each cell in column B
    For Each cell In Range("B1:B" & lastRow)
        ' Add cell value to total
        ' Convert the cell value to a string
        Dim stringValue As String
        stringValue = CStr(cell.value)
        
        If myDictionary.Exists(stringValue) Then
            'MsgBox "po # found in dictionary!"
            'targetCell = Range("C" + cell.Row)
            num = CDbl(Cells(cell.Row, "D").value)
            'MsgBox num
            If num <> 0 Then
                If sheetName = "Online" Then
                    Cells(cell.Row, "D").Formula = "=" & num & "-" & num
                    Cells(cell.Row, "E").Formula = "SHIPPED"
                    Cells(cell.Row, "H").value = MydateDisplay
                    Cells(cell.Row, "I").value = myDictionary(stringValue)(1)
                    Cells(cell.Row, "J").value = myDictionary(stringValue)(0)
                Else
                    Cells(cell.Row, "D").Formula = "=" & num & "-" & num
                    Cells(cell.Row, "F").value = MydateDisplay
                    Cells(cell.Row, "H").value = myDictionary(stringValue)(0)
                    Cells(cell.Row, "G").value = myDictionary(stringValue)(1)
                End If
            End If
            'Cells(cell.Row, "C").Formula = "=" & num & " - " & num
            'MsgBox Cells(cell.Row, "C").value
            
        End If
        
    Next cell
    
    
    MsgBox "Open Order automation complete. Please manually check the result and adjust imcomplete partial order if applicable."
    

    ' Display the total in a message box
    'MsgBox "Sum of Column C: " & total
End Sub
Function ReadExcel(filePath As String) As String
    Dim ExcelApp As Object
    Dim ExcelWorkbook As Object
    Dim ExcelWorksheet As Object
    Dim ExcelFile As String
    Dim ExcelData As String
    
    
    ' Define the path to your Excel file
    ExcelFile = filePath

    ' Create Excel and Outlook objects
    Set ExcelApp = CreateObject("Excel.Application")

    ' Open the Excel file
    Set ExcelWorkbook = ExcelApp.Workbooks.Open(ExcelFile)
    Set ExcelWorksheet = ExcelWorkbook.Sheets(1)

    ' Extract data from Excel
    Dim i As Integer
    Dim key As String, value As String
    For i = 1 To 1000
        'End if no more PO
        If ExcelWorksheet.Cells(i, "B").value = "" Then
            Exit For
        End If
        'Check if PO# exists, skip otherwise
        If ExcelWorksheet.Cells(i, "A").value = "" Then
            GoTo SkipIteration
        End If
        'MsgBox ExcelWorksheet.Cells(i, "A").value & " " & ExcelWorksheet.Cells(i, "B").value
        'Put the data in the dictionary
        key = ExcelWorksheet.Cells(i, "A").value
        value = ExcelWorksheet.Cells(i, "B").value
        
        'Hanlde combine orders
        Dim SplitKey() As String
        'Check if reference number already exist in keys
        If myDictionary.Exists(key) Then
            If InStr(1, key, ", ", vbTextCompare) > 0 Then
                ' Handle combine orders
                SplitKey = Split(key, ", ")
                myDictionary(SplitKey(0))(0) = myDictionary(SplitKey(0))(0) & " / " & value
                myDictionary(SplitKey(1))(0) = myDictionary(SplitKey(1))(0) & " / " & value
            Else
                'Avoid duplicate tracking number
                If InStr(1, myDictionary(key)(0), value, vbTextCompare) <= 0 Then
                    myDictionary(key)(0) = myDictionary(key)(0) & " / " & value
                End If
            End If
        Else
            If InStr(1, key, ", ", vbTextCompare) > 0 Then
                SplitKey = Split(key, ", ")
                
                ' get length of splitKey: UBound(Splitkey)
                Dim j As Integer
                For j = 0 To UBound(SplitKey)
                    myDictionary.Add SplitKey(j), Array(value, "0")
                Next j
            Else
                myDictionary.Add key, Array(value, "0")
            End If
        End If
        
SkipIteration:
    Next i
    
    
    
    ' Close Excel
    ExcelWorkbook.Close False
    ExcelApp.Quit

    ' Release objects
    Set ExcelWorksheet = Nothing
    Set ExcelWorkbook = Nothing
    Set ExcelApp = Nothing
    
    
End Function

Sub ReadPDFWithPyPDF2()
    Dim pythonScriptPath As String
    Dim cmd As String
    
    ' Path to the modified Python script
    pythonScriptPath = "\\SERVER2\Tech\USPS_Reference\ExtractPDF.py"
    
    ' Command to run Python script
    cmd = "pythonw.exe " & pythonScriptPath
    
    ' Execute the Python script from VBA
    Call Shell(cmd, vbNormalFocus)
    
    ' Wait for the Python script to finish (adjust the delay if necessary)
    'Application.Wait Now + TimeValue("00:00:03") ' Wait for 3 seconds
    
    ReadExcel ("\\SERVER2\Tech\USPS_Reference\USPS_Excel.csv")
    
    
End Sub

Sub ReadQuickbookReport()
    Dim xlApp As Object
    Dim xlWorkbook As Object
    Dim xlWorksheet As Object
    Dim excelFilePath As String
    Dim numRows As Long
    
    'MsgBox "Mydate is " + Mydate
    
    ' Set the path to your Excel file
    excelFilePath = "\\SERVER2\Tech\Daily_Quickbook_Report\" + Mydate + " Daily Quickbook Report.xlsx"
    
    ' Create a new instance of Excel application
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False ' You can set this to True to make Excel visible
    
    If (Dir(excelFilePath) <> "") Then
        ' Open the Excel workbook
        Set xlWorkbook = xlApp.Workbooks.Open(excelFilePath)
        ' Specify the worksheet you want to work with
        Set xlWorksheet = xlWorkbook.Sheets("Sheet1") ' Change "Sheet1" to your sheet name
        
        ' Get the last row number with data in column G
        numRows = xlWorksheet.Cells(xlWorksheet.Rows.Count, "G").End(xlUp).Row
                    
        Dim i As Integer
        Dim exist As Boolean
        Dim formattedDate As String
        exist = False
        'MsgBox "MydateDisplay is " & MydateDisplay
        formattedDate = Format(DateSerial(Year(CDate(MydateDisplay)), Month(CDate(MydateDisplay)), Day(CDate(MydateDisplay))), "m/d/yyyy")
        MsgBox formattedDate
        'MsgBox xlWorksheet.Cells(527, "G").value
        'Gotta fix Mydate (e.g 12/01/2023 -> 12/1/2023)
        For i = 1 To numRows
            If xlWorksheet.Cells(i, "G").value = formattedDate And myDictionary.Exists(xlWorksheet.Cells(i, "K").value) Then
                ' Convert the cell value to a string
                Dim stringValue As String
                stringValue = CStr(xlWorksheet.Cells(i, "K").value)

                ' Retrieve the array from the dictionary
                Dim retrievedArray() As Variant
                retrievedArray = myDictionary(stringValue)
                
                ' Change the second element of the array
                retrievedArray(1) = xlWorksheet.Cells(i, "I").value
                
                ' Update the array in the dictionary with the modified array
                myDictionary(stringValue) = retrievedArray
                
            End If
            Next i
        
        ' Close Excel without saving changes
        xlWorkbook.Close False
        
        
        'Test if dictionary contains right value
        Dim key As Variant
        Dim output As String
        output = "Values in the dictionary:" & vbCrLf
        
        ' Iterate through the dictionary and print values
        For Each key In myDictionary.Keys
            output = output & "Key: " & key & ", Value: " & myDictionary(key)(0) & ", " & myDictionary(key)(1) & vbCrLf
        Next key
        'MsgBox output
        
        ' Display the number of rows in a message box
        'MsgBox "The number of rows in the worksheet is: " & numRows
    Else
        ' Workbook failed to open
        MsgBox "Failed to open the workbook. Make sure to export/update the current daily quickbook report with today's date."
    End If
        
    ' Close Excel without saving changes
    xlApp.Quit
    ' Release Excel objects from memory
    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApp = Nothing

End Sub



