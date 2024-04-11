Attribute VB_Name = "Module1"
Option Explicit
Dim myDictionary As Object

Sub ReplyMSG()
    Dim MyEmail As MailItem
    'Set MyEmail = Application.CreateItem(olMailItem)
    Dim olReply As MailItem
    Dim olRecip As Recipient
    Dim Mydate As String
    Dim strBody As String, reBody As String
    Dim lines() As String
    Dim searchCc As String
    Dim foundCc As String
    Dim splitResult() As String
    Dim foundAddress As String
    Dim i As Integer, j As Integer, k As Integer
    Dim olRecipient As Object
    Dim arrLines As String
    
    'Initialize dictionary if its empty
    If myDictionary Is Nothing Then
        InitializeDictionary
    End If
    
    ' Set today's date
    Mydate = Date$
    Mydate = Replace(Mydate, "-", "/")
    
    For Each MyEmail In Application.ActiveExplorer.Selection
        ' Check if the selected item is an email
        If TypeOf MyEmail Is Outlook.MailItem Then
        
            'find PO # in subject
            'MsgBox MyEmail.subject
            Dim subject As String
            Dim splitPoNum() As String
            Dim PoNum As Variant
            Dim TrackingNum As Variant
            
            subject = MyEmail.subject
            splitPoNum = Split(subject, "# ")
            PoNum = splitPoNum(1)
            If myDictionary.Exists(PoNum) Then
                'MsgBox PoNum & " " & myDictionary(PoNum)
                TrackingNum = myDictionary(PoNum)
                
            Else
                ReadExcel ("\\SERVER2\Tech\UPS_Reference\UPS_CSV_EXPORT.csv")
                ReadPDFWithPyPDF2
                
                If myDictionary.Exists(PoNum) Then
                    'MsgBox PoNum & " " & myDictionary(PoNum)
                    TrackingNum = myDictionary(PoNum)
                Else
                    MsgBox "Tracking Number Not Found"
                End If
            End If
    
            Set olReply = MyEmail.Reply
            strBody = MyEmail.htmlBody
                                                                                       
            reBody = MyEmail.Body
            ' Remove all original recipients
            For Each olRecipient In olReply.Recipients
                olRecipient.Delete
            Next olRecipient
            
            'Define the string you want to search for
            searchCc = "Cc"
    
            'Split the email body into lines
            lines = Split(reBody, vbCrLf)
            
            'Loop through the lines to find the first line containing the word "Cc"
            For i = LBound(lines) To UBound(lines)
                If InStr(1, lines(i), searchCc, vbTextCompare) > 0 Then
                    foundCc = lines(i)
                    Exit For ' Exit the loop when the first matching line is found
                End If
            Next i
            
            'Check if the search string was found
            If Len(foundCc) > 0 Then
                'Display the found line in a message box (you can modify this part)
                splitResult = Split(foundCc, ";")
                Dim splitResultLength As Integer
                splitResultLength = UBound(splitResult) - LBound(splitResult) + 1
                If splitResultLength = 3 Then
                    If splitResult(1) = splitResult(2) Then
                        Set olRecip = olReply.Recipients.Add(splitResult(1))
                    Else
                        If splitResult(2) = " jgallardo@tesla.com" Then
                            Set olRecip = olReply.Recipients.Add(splitResult(1))
                        Else
                            Set olRecip = olReply.Recipients.Add(splitResult(1))
                            Set olRecip = olReply.Recipients.Add(splitResult(2))
                        End If
                    End If
                ElseIf splitResultLength = 2 Then
                    Set olRecip = olReply.Recipients.Add(splitResult(1))
                Else
                    MsgBox "Cannot find the recipient address"
                End If
    
            Else
                MsgBox "The search string '" & searchCc & "' was not found in the email body."
            End If
            
            
            Dim SignatureHTML As String
            SignatureHTML = "<p style='color:#2F5496;'>" & "MPCM Logistics Team" & "<br>" & "MPCM, INC" & "<br>" & "115 Phelan Ave, Suite 6" & "<br>" & "San Jose, CA 95112" & "</p>"
            With olReply
                .htmlBody = "<p style='font-size: 12pt; font-family: Times New Roman;'>" & "Dear Customer, " & "<br><br>" & "Your order was shipped on " & Mydate & "." & "<br><br>" & "TK#: " & TrackingNum & "<br><br>" & "Thank you for your order." & "<br>" & SignatureHTML & strBody & "</p>"
                .CC = "mnguyen@mpcmfg.com"
                .BCC = "xtran@mpcmfg.com;orders@mpcmfg.com"
                .BodyFormat = olFormatHTML
                
                .Display
            End With
            
        Else
            MsgBox "Please select an email."
        End If
        
    Next MyEmail
    
    
    Set MyEmail = Nothing
    Set olReply = Nothing
    Set olRecip = Nothing
End Sub
Function InitializeDictionary()
    Set myDictionary = CreateObject("Scripting.Dictionary")
End Function
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
        
        If myDictionary.Exists(key) Then
            If InStr(1, key, ", ", vbTextCompare) > 0 Then
                SplitKey = Split(key, ", ")
                myDictionary(SplitKey(0)) = myDictionary(SplitKey(0)) & " / " & value
                myDictionary(SplitKey(1)) = myDictionary(SplitKey(1)) & " / " & value
            Else
                'If InStr(1, myDictionary(key), value, vbTextCompare) = 0 Then
                If InStr(1, myDictionary(key), value, vbTextCompare) <= 0 Then
                    myDictionary(key) = myDictionary(key) & " / " & value
                End If
            End If
            
        Else
            If InStr(1, key, ", ", vbTextCompare) > 0 Then
                SplitKey = Split(key, ", ")
                myDictionary.Add SplitKey(0), value
                myDictionary.Add SplitKey(1), value
            Else
                myDictionary.Add key, value
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


