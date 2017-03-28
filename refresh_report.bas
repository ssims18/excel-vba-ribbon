Attribute VB_Name = "refresh_report"
' Provide the invoice age category label
'
' Param: invoice_age - int - age of the record
' Return: age text label - string - category of invoice age
Public Function getAging(invoice_age) As String
    
    If invoice_age < 31 Then
        getAging = "0-30"
    ElseIf invoice_age >= 31 And invoice_age <= 60 Then
        getAging = "31-60"
    ElseIf invoice_age >= 61 And invoice_age <= 90 Then
        getAging = "61-90"
    ElseIf invoice_age > 90 Then
        getAging = "Over 90"
    End If

End Function
' Take the business line param and return the corresponding
' report sheet to be read for processing
'
' Param: BusinessLine - string
' Return: invoice sheet name - string - !! Must match imported worksheet name !!
Public Function getInvoiceSheet(BusinessLine) As String
    
    If BusinessLine = "service1" Then
        getInvoiceSheet = "service1_invoices"
    ElseIf BusinessLine = "service2" Then
        getInvoiceSheet = "service2_invoices"
    ElseIf BusinessLine = "service3" Then
        getInvoiceSheet = "service3_invoices"
    ElseIf BusinessLine = "service4" Then
        getInvoiceSheet = "service4_invoices"
    ' Default output should error out
    Else
        getInvoiceSheet = "ERROR-UNDEFINED SHEET"
    End If
    
End Function
' Return the appropriate principal reference for the business line
'
' Param: BusinessLine - string
' Return: formula text expression with appropriate service line values
Public Function getPrincipalLookup(BusinessLine, LastRow1) As String

    If BusinessLine = "service2" Then
        getPrincipalLookup = "=IF(J" & LastRow1 + 1 & " = ""Service Line 2 Name"", ""Service Line 2 Name"", IFERROR(VLOOKUP(A" & LastRow1 + 1 & ",'AO-AM'!A:E, MATCH(K$1,'AO-AM'!$A$1:$G$1,0), FALSE),0))"
        
    ElseIf BusinessLine = "service3" Then
        getPrincipalLookup = "=IF(J" & LastRow1 + 1 & " = ""Service Line 3 Name"", ""Service Line 3 Name"", IFERROR(VLOOKUP(A" & LastRow1 + 1 & ",'AO-AM'!A:E, MATCH(K$1,'AO-AM'!$A$1:$G$1,0), FALSE),0))"
        
    ElseIf BusinessLine = "service4" Then
        getPrincipalLookup = "=IF(J" & LastRow1 + 1 & " = ""Service Line 4 Name"", ""Service Line 4 Name"", IFERROR(VLOOKUP(A" & LastRow1 + 1 & ",'AO-AM'!A:E, MATCH(K$1,'AO-AM'!$A$1:$G$1,0), FALSE),0))"
    
    ' default to service1 lookup
    Else
        getPrincipalLookup = "=IFERROR(VLOOKUP(A" & LastRow1 + 1 & ",'AO-AM'!A:E, MATCH(K$1,'AO-AM'!$A$1:$G$1,0), FALSE),0)"
    
    End If


End Function
' Return the appropriate AO reference for the business line
'
' Param: BusinessLine - string
' Return: formula text expression with appropriate service line values
Public Function getAOLookup(BusinessLine, LastRow1) As String

    If BusinessLine = "service2" Then
        getAOLookup = "=IF(J" & LastRow1 + 1 & " = ""Service Line 2 Name"", ""Service Line 2 Name"", IFERROR(VLOOKUP(A" & LastRow1 + 1 & ",'AO-AM'!A:G, 6, FALSE),0))"
    ElseIf BusinessLine = "service3" Then
        getAOLookup = "=IF(J" & LastRow1 + 1 & " = ""Service Line 3 Name"", ""Service Line 3 Name"", IFERROR(VLOOKUP(A" & LastRow1 + 1 & ",'AO-AM'!A:G, 6, FALSE),0))"
    ElseIf BusinessLine = "service4" Then
        getAOLookup = "=IF(J" & LastRow1 + 1 & " = ""Service Line 4 Name"", ""Service Line 4 Name"", IFERROR(VLOOKUP(A" & LastRow1 + 1 & ",'AO-AM'!A:G, 6, FALSE),0))"
    ' default to service1 lookup
    Else
        getAOLookup = "=IFERROR(VLOOKUP(A" & LastRow1 + 1 & ",'AO-AM'!A:E, MATCH(K$1,'AO-AM'!$A$1:$G$1,0), FALSE),0)"
    End If

End Function
' Return the appropriate AM reference for the business line
'
' Param: BusinessLine - string
' Return: formula text expression with appropriate service line values
Public Function getAMLookup(BusinessLine, LastRow1) As String

    If BusinessLine = "service2" Then
        getAMLookup = "=IF(J" & LastRow1 + 1 & " = ""Service Line 2 Name"", ""Service Line 2 Name"", IFERROR(VLOOKUP(A" & LastRow1 + 1 & ",'AO-AM'!A:G, 7, FALSE),0))"
    ElseIf BusinessLine = "service3" Then
        getAMLookup = "=IF(J" & LastRow1 + 1 & " = ""Service Line 3 Name"", ""Service Line 3 Name"", IFERROR(VLOOKUP(A" & LastRow1 + 1 & ",'AO-AM'!A:G, 7, FALSE),0))"
    ElseIf BusinessLine = "service4" Then
        getAMLookup = "=IF(J" & LastRow1 + 1 & " = ""Service Line 4 Name"", ""Service Line 4 Name"", IFERROR(VLOOKUP(A" & LastRow1 + 1 & ",'AO-AM'!A:G, 7, FALSE),0))"
    ' default to service1 lookup
    Else
        getAMLookup = "=IFERROR(VLOOKUP(A" & LastRow1 + 1 & ",'AO-AM'!A:E, MATCH(K$1,'AO-AM'!$A$1:$G$1,0), FALSE),0)"
    End If

End Function


' Get Invoices for a business line.
' Expectes to be called from modRibbonControl
' Param: BusinessLine
' exmaple: service1, service2, service3, service4

Sub GetInvoices(BusinessLine)
   
    ' Turn off screen updating, and then open the target workbook.
    'Application.ScreenUpdating = False
    
    ' Dev
    'Dim BusinessLine As String
    'BusinessLine = "service4"

    Dim row As Integer
    Dim LastRow As Long
    Dim LastRow1 As Long
    
    Dim ClientRow As Variant
    Dim client_cd As String
    Dim client_name As String
    Dim invoice As String
    Dim OnAcctAmt As String
    Dim Balance As String
    Dim InvoiceDte As String
    Dim invoice_age As Integer
    
   
    ' Sheet for exported report data
    Dim rptInvoices As Worksheet
    Set rptInvoices = ThisWorkbook.Sheets(getInvoiceSheet(BusinessLine))
    
    
    ' Sheet for writing template data
    Dim AR_template As Worksheet
    Set AR_template = ThisWorkbook.Sheets("Advantage AR")
    
    ' Sheet for client vertical VLookUp
    Dim lookupSheet As Worksheet
    Set lookupSheet = ThisWorkbook.Sheets("Client-Vertical")
    
    ' Sheet for AO/AM VLookUp
    Dim aoSheet As Worksheet
    Set aoSheet = ThisWorkbook.Sheets("AO-AM")
    
    ' Get last row of exported report
    With rptInvoices
        LastRow = .Range("A1").SpecialCells(xlCellTypeLastCell).row
    End With
    
    ' parameters
    ' header row 3, data starts 5
    ' columns C - T
    
    ' criteria
    ' record row if invoice number present
    ' OR
    ' if numerical/currency value in P, then "on account"
    
    ' Grab requested data from active rows in report
    ' Last two rows are summary, omit from process
    
    'Worksheet
    rptInvoices.Select
    
    row = 5
    Do While row <= (LastRow - 2)

        'Start at column C
        Range("C" & row).Select
       
        'Get Client Name
        ' split on " - "
        If Selection.Cells(1, 1).Value <> "" Then
            ClientRow = Range("C" & row).Value
           
            'parse combined client code, client name
            If InStr(1, Range("C" & row).Value, " - ") > 0 Then
                ClientRow = Split(Range("C" & row).Value, " - ")
                client_cd = Trim(ClientRow(0))
                client_name = Trim(ClientRow(1))
            End If
        End If
        
        invoice = Replace(Replace(Range("D" & row).Value, "-00", ""), "*", "")
        
        OnAcctAmt = Range("P" & row).Value
        Balance = Range("R" & row).Value
        InvoiceDte = Range("T" & row).Value 'format as datetime
        ' days since invoice date
        If IsDate(InvoiceDte) = True Then
            invoice_age = DateDiff("d", CDate(InvoiceDte), Now())
        Else
            invoice_age = 0
        End If
        
        ' write to data tab in template (Advantage AR)
        ' if criteria is met
        With AR_template
            LastRow1 = .Range("A1").SpecialCells(xlCellTypeLastCell).row
        End With
        
        If invoice <> "" Or IsNumeric(OnAcctAmt) = True Then
            With AR_template
                ' next open row = LastRow1 + 1
                
                'Client ID
                AR_template.Cells(LastRow1 + 1, 1) = client_cd
                'Client Name
                AR_template.Cells(LastRow1 + 1, 2) = client_name
                'Invoice Number & 'Balance
                If invoice = "" Then
                    AR_template.Cells(LastRow1 + 1, 3) = "On Account"
                    AR_template.Cells(LastRow1 + 1, 7) = CCur(OnAcctAmt)    'Open Balance
                Else
                    AR_template.Cells(LastRow1 + 1, 3) = invoice
                    AR_template.Cells(LastRow1 + 1, 7) = CCur(Balance)    'Open Balance
                End If
                'Date
                AR_template.Cells(LastRow1 + 1, 4) = CDate(InvoiceDte)
                'Age
                AR_template.Cells(LastRow1 + 1, 5) = invoice_age
                'get aging range
                AR_template.Cells(LastRow1 + 1, 6) = getAging(invoice_age)


                'postage formula
                AR_template.Cells(LastRow1 + 1, 8).Formula = "=IFERROR(VLOOKUP(C" & LastRow1 + 1 & ",'Advantage Postage'!A:G,7,FALSE),0)"
                
                'Net Balance
                AR_template.Cells(LastRow1 + 1, 9).Formula = "=G" & LastRow1 + 1 & "-H" & LastRow1 + 1 & ""
                
                'Vertical --> vlookup
                On Error Resume Next
					' run VLookUp and insert result 
                    'AR_template.Cells(LastRow1 + 1, 10) = Application.WorksheetFunction.VLookup(AR_template.Cells(LastRow1 + 1, 1), lookupSheet.Range("A:C"), 3, False)
					' insert VLookUp as formula
                    AR_template.Cells(LastRow1 + 1, 10).Formula = "=IFERROR(VLOOKUP(A" & LastRow1 + 1 & ", 'Client-Vertical'!A:C, 3, FALSE),0)"
                
                'Principal
                On Error Resume Next
                    AR_template.Cells(LastRow1 + 1, 11).Formula = getPrincipalLookup(BusinessLine, LastRow1)
                'AO
                On Error Resume Next
                    AR_template.Cells(LastRow1 + 1, 12).Formula = getAOLookup(BusinessLine, LastRow1)
                'AM
                On Error Resume Next
                    AR_template.Cells(LastRow1 + 1, 13).Formula = getAMLookup(BusinessLine, LastRow1)
                'Net
                AR_template.Cells(LastRow1 + 1, 15).Formula = "=E" & LastRow1 + 1 & "*I" & LastRow1 + 1 & ""
                'Open
                AR_template.Cells(LastRow1 + 1, 17).Formula = "=E" & LastRow1 + 1 & "*G" & LastRow1 + 1 & ""
            End With
        End If
        
        'advance to next row
        row = row + 1
    Loop
   
    'hide screen updating
    'Application.ScreenUpdating = True
    
    're-hide raw data
    rptInvoices.Visible = False
    
    'Save progress
    ThisWorkbook.Save
End Sub


' Advantage Postage
' Reformat the Advantage Postage report, and add any requested formulas
' Expected sheet name: postage
Sub GetPostage(control As IRibbonControl)

    'Turn off screen updating, and then open the target workbook.
    Application.ScreenUpdating = False
    
    Dim row As Integer
    Dim LastRow As Long
    Dim LastRow1 As Long
    Dim templateLastRow As Long
        
    Dim ClientRow As Variant
    Dim client_cd As String
    Dim client_name As String
    Dim ARInv As String
    Dim ARInvDate As String
    Dim JobNbr As String
    Dim JobDesc As String
    Dim FncCode1 As String
    Dim FncDesc1 As String
    Dim BilledAmt As String
    
    ' Sheet for exported report data
    Dim rptPstg As Worksheet
    Set rptPstg = ThisWorkbook.Sheets("postage")
    
    With rptPstg
            LastRow = .Range("A1").SpecialCells(xlCellTypeLastCell).row
    End With
    
    ' Sheet for writing template data
    Dim pstg_template As Worksheet
    Set pstg_template = ThisWorkbook.Sheets("Advantage Postage")
   
   ' process report export
    rptPstg.Select
    row = 3
    Do While row < (LastRow + 1)

        'Start at column G
        Range("I" & row).Select
        
        If Range("I" & row).Value <> "" Then
            ARInv = Replace(Replace(Range("I" & row).Value, "-00", ""), "*", "")
        End If
        If Range("K" & row).Value <> "" Then
            JobNbr = Range("K" & row).Value
            JobDesc = Range("L" & row).Value
        End If
        
        'Get values
        'ARInv = Replace(Replace(Range("I" & row).Value, "-00", ""), "*", "")
        ARInvDate = Range("J" & row).Value
        'JobNbr = Range("K" & row).Value
        'JobDesc = Range("L" & row).Value
        FncCode1 = Range("M" & row).Value
        FncDesc1 = Range("N" & row).Value
        BilledAmt = Range("O" & row).Value
        
        With pstg_template
            LastRow1 = .Range("A1").SpecialCells(xlCellTypeLastCell).row
        End With
        
        ' write to template (Advantage Postage)
        If FncCode1 = "postag" Or FncCode1 = "zpost" Then
            With pstg_template
                'Invoice
                pstg_template.Cells(LastRow1 + 1, 1) = ARInv
                'Invoice Date
                pstg_template.Cells(LastRow1 + 1, 2) = CDate(ARInvDate)
                'Job Number
                pstg_template.Cells(LastRow1 + 1, 3) = JobNbr
                'Job Description
                pstg_template.Cells(LastRow1 + 1, 4) = JobDesc
                'Function Code
                pstg_template.Cells(LastRow1 + 1, 5) = FncCode1
                'Function Desc
                pstg_template.Cells(LastRow1 + 1, 6) = FncDesc1
                'Billed Amount
                pstg_template.Cells(LastRow1 + 1, 7) = CCur(BilledAmt)
            End With
        End If
        
        ' advance to next row
        row = row + 1
    Loop
    
    ' enable screen updating
    Application.ScreenUpdating = True
    
    ' re-hide raw data
    rptPstg.Visible = False
    
    ' Save progress
    ThisWorkbook.Save
    
    ' Alert that process has finished
    MsgBox "Postage Imported.", vbOKOnly, "Progress"
    
End Sub

' Method to check if report worksheet is in the
' list of front facing sheets to clear
' called from ClearTemplate to help manage the sheet name IF bocks
Public Function isMainReportSheet(wrkSheet As String) As Boolean
    
    'create array of sheets to check
    Dim arrWrkSheets As Variant
    arrWrkSheets = Array("Advantage AR", "Advantage Postage", "AO-AM")
    
    'default return = false
    isRequestedSheet = False
    
    For x = LBound(arrWrkSheets) To UBound(arrWrkSheets)
        If arrWrkSheets(x) = wrkSheet Then
            isRequestedSheet = True
            Exit Function
        End If
    Next x

End Function

' Method to check if supporting report worksheet is in the
' list of supporting sheets to clear
' called from ClearTemplate to help manage the sheet name IF bocks
Public Function isSupportingSheet(wrkSheet As String) As Boolean
    
    'create array of sheets to check
    Dim arrWrkSheets As Variant
    arrWrkSheets = Array("service1_invoices", "service2_invoices", "service3_invoices", "service4_invoices", "postage")
    
    ' default return = false
    isSupportingSheet = False
    
    For x = LBound(arrWrkSheets) To UBound(arrWrkSheets)
        If arrWrkSheets(x) = wrkSheet Then
            isSupportingSheet = True
            Exit Function
        End If
    Next x

End Function
' Delete exported report data, AR, and Postage reports
' Clears the report template before importing new data
Sub ClearTemplate(control As IRibbonControl)
    
    'Turn off screen updating, and then open the target workbook.
    Application.ScreenUpdating = False
    
    'set worksheets
    Dim AR_template As Worksheet

    'find last rows
    Dim intRow
    Dim intLastRow
    
    For Each ws In ThisWorkbook.Sheets
        Set AR_template = ThisWorkbook.Sheets(ws.Name)
    
        'unhide sheets for working
        If AR_template.Visible = False Then
            AR_template.Visible = True
        End If
        'select sheet for work
        AR_template.Select
        
        'find last row number
        With AR_template
            intLastRow = .Range("A1").SpecialCells(xlCellTypeLastCell).row
        End With
        
        'If isMainReportSheet(ws.Name) Then
        If ws.Name = "Advantage AR" Or ws.Name = "Advantage Postage" Or ws.Name = "AO-AM" Then
            'Delete the range from the 2nd row
            Rows("2:" + CStr(intLastRow) + "").Delete
            
            'place cursor on top cell
            AR_template.Range("A2").Select
        ElseIf isSupportingSheet(ws.Name) Then
            'Delete range from the first row
            Rows("1:" + CStr(intLastRow) + "").Delete
            
            'place cursor on top cell
            AR_template.Range("A1").Select
        End If
    Next
    
    'hide screen updating
    Application.ScreenUpdating = True
    
    'Application.AutoCalculate = False
    
    'Save progress
    ThisWorkbook.Save
    
    MsgBox "Data Cleared!", vbOKOnly

End Sub
' Add subtotal formulas to the Advantage AR sheet
' for reference checks
Sub addSubTotal(control As IRibbonControl)
   
    ' Sheet for writing template data
    Dim LastRow As Long
    Dim subTtlRow As Long
    Dim AR_template As Worksheet
    Set AR_template = ThisWorkbook.Sheets("Advantage AR")
    
    With AR_template
        LastRow = .Range("A1").SpecialCells(xlCellTypeLastCell).row
    End With
    
    subTtlRow = LastRow + 2
    
    AR_template.Cells(subTtlRow, 7).Select
    
    AR_template.Cells(subTtlRow, 7).Formula = "=SUBTOTAL(9,G2:G" & LastRow & ")"
    AR_template.Cells(subTtlRow, 9).Formula = "=SUBTOTAL(9,I2:I" & LastRow & ")"
    AR_template.Cells(subTtlRow, 14).Formula = "=SUBTOTAL(9,O2:O" & LastRow & ")"
    AR_template.Cells(subTtlRow, 16).Formula = "=SUBTOTAL(9,Q2:Q" & LastRow & ")"
        
    AR_template.Cells(LastRow + 4, 14).Formula = "=TEXT(O" & subTtlRow & "/I" & subTtlRow & "/100, ""0.00%"")"
    AR_template.Cells(LastRow + 4, 16).Formula = "=TEXT(Q" & subTtlRow & "/G" & subTtlRow & "/100, ""0.00%"")"

End Sub


