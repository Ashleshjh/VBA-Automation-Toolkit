Option Explicit

Sub Generate_Live_Formula_Salary_Sheets()

    Dim masterPath As Variant, company1AttPath As Variant, company2AttPath As Variant
    Dim wbMaster As Workbook, wbFinal As Workbook
    Dim monthStr As String, monthDays As Variant
    
    ' --- 1. GET MONTH & DAYS ---
    monthStr = InputBox("Enter the Month and Year for the title:" & vbCrLf & "(Example: August-2026)", "Report Month")
    If monthStr = "" Then Exit Sub
    
    monthDays = Application.InputBox("Enter the total days in this payroll month (e.g., 30 or 31):", "Month Days", 31, Type:=1)
    If monthDays = False Or monthDays <= 0 Then Exit Sub
    
    ' --- 2. SELECT THE FILES ---
    masterPath = Application.GetOpenFilename("Excel Files (*.xls*), *.xls*", , "1. Select the MASTER SALARY Sheet")
    If masterPath = False Then Exit Sub
    
    company1AttPath = Application.GetOpenFilename("Excel Files (*.xls*), *.xls*", , "2. Select the Company A Attendance Sheet")
    If company1AttPath = False Then Exit Sub
    
    company2AttPath = Application.GetOpenFilename("Excel Files (*.xls*), *.xls*", , "3. Select the Company B Attendance Sheet")
    If company2AttPath = False Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' --- 3. OPEN MASTER FILE & CREATE FINAL WORKBOOK ---
    Set wbMaster = Workbooks.Open(masterPath)
    Set wbFinal = Workbooks.Add
    
    While wbFinal.Sheets.count > 1
        wbFinal.Sheets(2).Delete
    Wend
    
    ' --- 4. PROCESS BOTH COMPANIES WITH LIVE FORMULAS ---
    Call Build_Formula_Tab(wbMaster, wbFinal, company1AttPath, "CompanyA", "Company A Pvt. Ltd.", monthStr, monthDays)
    Call Build_Formula_Tab(wbMaster, wbFinal, company2AttPath, "CompanyB", "Company B Pvt. Ltd.", monthStr, monthDays)
    
    On Error Resume Next
    If wbFinal.Sheets.count > 2 Then wbFinal.Sheets("Sheet1").Delete
    On Error GoTo 0
    
    wbMaster.Close False
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Salary Sheets generated successfully with Live Formulas!", vbInformation

End Sub

' ==============================================================================
' HELPER SUB: BUILDS TAB AND INJECTS EXCEL FORMULAS
' ==============================================================================
Private Sub Build_Formula_Tab(wbMaster As Workbook, wbFinal As Workbook, attPath As Variant, sheetName As String, compTitle As String, monthStr As String, monthDays As Variant)

    Dim wbAtt As Workbook, wsAtt As Worksheet
    Dim wsMaster As Worksheet, wsFinal As Worksheet
    Dim dictAtt As Object
    Dim i As Long, finalRow As Long
    Dim lastRowAtt As Long, lastRowMaster As Long
    Dim cName As String, cPaidDays As Double
    Dim colName As Long, colPaidDays As Long
    
    Set wsMaster = wbMaster.Sheets(1)
    Set dictAtt = CreateObject("Scripting.Dictionary")
    dictAtt.CompareMode = 1
    
    ' --- A. LOAD ATTENDANCE INTO MEMORY ---
    Set wbAtt = Workbooks.Open(attPath)
    Set wsAtt = wbAtt.Sheets("Attendance")
    lastRowAtt = wsAtt.Cells(wsAtt.Rows.count, "B").End(xlUp).Row
    
    On Error Resume Next
    colName = wsAtt.Rows(1).Find("Name of Employees", LookAt:=xlWhole).Column
    colPaidDays = wsAtt.Rows(1).Find("Total Paid Days", LookAt:=xlWhole).Column
    On Error GoTo 0
    
    If colName = 0 Then colName = 2
    If colPaidDays = 0 Then colPaidDays = 12
    
    For i = 2 To lastRowAtt
        cName = Trim(wsAtt.Cells(i, colName).Value)
        cPaidDays = Val(wsAtt.Cells(i, colPaidDays).Value)
        If cName <> "" And Not dictAtt.Exists(cName) Then
            dictAtt.Add cName, cPaidDays
        End If
    Next i
    wbAtt.Close False
    
    ' --- B. BUILD THE FINAL SHEET & HEADERS ---
    Set wsFinal = wbFinal.Sheets.Add(After:=wbFinal.Sheets(wbFinal.Sheets.count))
    wsFinal.Name = sheetName
    
    wsFinal.Cells(1, 1).Value = compTitle
    wsFinal.Cells(1, 1).Font.Size = 14
    wsFinal.Cells(1, 1).Font.Bold = True
    
    wsFinal.Cells(2, 1).Value = "Salary & Wages- " & monthStr
    wsFinal.Cells(2, 1).Font.Size = 12
    wsFinal.Cells(2, 1).Font.Bold = True
    
    With wsFinal.Cells(2, 5)
        .Value = monthDays
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(200, 200, 200)
        .Borders.LineStyle = xlContinuous
    End With
    
    Dim headers As Variant
    headers = Array("SL No", "Employee Name", "Designation", "Fixed Salary", "PAID DAYS", _
                    "EARNED SALARY", "OT Hours", "OT Payable", "Total Payable", "PT", _
                    "SALARY ADVANCE", "Invoice No", "Purchase Deductions", "TOTAL DEDUCTIONS", _
                    "Total Paid", "MODE OF PAYMENT")
                    
    wsFinal.Range("A3").Resize(1, 16).Value = headers
    wsFinal.Rows(3).Font.Bold = True
    wsFinal.Rows(3).HorizontalAlignment = xlCenter
    wsFinal.Rows(3).VerticalAlignment = xlCenter
    wsFinal.Rows(3).WrapText = True
    
    ' --- C. INJECT DATA & LIVE FORMULAS ---
    lastRowMaster = wsMaster.Cells(wsMaster.Rows.count, "B").End(xlUp).Row
    finalRow = 4
    
    Dim vBasic As Double, vFixed As Double
    Dim vOTHours As Double, vAdvance As Double, vPurch As Double
    
    For i = 2 To lastRowMaster
        cName = Trim(wsMaster.Cells(i, "B").Value)
        
        ' Check if employee exists in this specific company's attendance
        If cName <> "" And dictAtt.Exists(cName) Then
            
            ' 1. Static Inputs from Master
            wsFinal.Cells(finalRow, 1).Value = finalRow - 3 ' SL No
            wsFinal.Cells(finalRow, 2).Value = cName
            wsFinal.Cells(finalRow, 3).Value = wsMaster.Cells(i, "E").Value ' Designation
            
            vFixed = Val(wsMaster.Cells(i, "K").Value)
            wsFinal.Cells(finalRow, 4).Value = vFixed ' Fixed Salary
            
            wsFinal.Cells(finalRow, 5).Value = dictAtt(cName) ' Paid Days
            
            vOTHours = Val(wsMaster.Cells(i, "L").Value)
            If vOTHours > 0 Then wsFinal.Cells(finalRow, 7).Value = vOTHours Else wsFinal.Cells(finalRow, 7).Value = "-"
            
            vAdvance = Val(wsMaster.Cells(i, "O").Value)
            If vAdvance > 0 Then wsFinal.Cells(finalRow, 11).Value = vAdvance Else wsFinal.Cells(finalRow, 11).Value = "-"
            
            wsFinal.Cells(finalRow, 12).Value = "" ' Invoice No
            
            vPurch = Val(wsMaster.Cells(i, "P").Value)
            If vPurch > 0 Then wsFinal.Cells(finalRow, 13).Value = vPurch Else wsFinal.Cells(finalRow, 13).Value = "-"
            
            wsFinal.Cells(finalRow, 16).Value = wsMaster.Cells(i, "T").Value ' Mode of Payment
            
            ' 2. INJECT LIVE EXCEL FORMULAS
            vBasic = Val(wsMaster.Cells(i, "F").Value) ' Grab Basic internally for the OT formula
            
            ' Earned Salary: =ROUND(Fixed / MonthDays * PaidDays, 0)
            wsFinal.Cells(finalRow, 6).Formula = "=ROUND(D" & finalRow & "/$E$2*E" & finalRow & ", 0)"
            
            ' OT Payable: = Basic / MonthDays / 8 * OT Hours
            If vOTHours > 0 Then
                wsFinal.Cells(finalRow, 8).Formula = "=ROUND(" & vBasic & "/$E$2/8*G" & finalRow & ", 0)"
            Else
                wsFinal.Cells(finalRow, 8).Value = "-"
            End If
            
            ' Total Payable: =Earned + OT
            wsFinal.Cells(finalRow, 9).Formula = "=ROUND(F" & finalRow & "+N(H" & finalRow & "), 0)"
            
            ' PT: =IF(TotalPayable > 24999, 200, 0)
            wsFinal.Cells(finalRow, 10).Formula = "=IF(I" & finalRow & ">24999, 200, 0)"
            
            ' Total Deductions: =SUM(PT, Advance, Purchase)
            wsFinal.Cells(finalRow, 14).Formula = "=SUM(J" & finalRow & ", K" & finalRow & ", M" & finalRow & ")"
            
            ' Total Paid: =TotalPayable - TotalDeductions
            wsFinal.Cells(finalRow, 15).Formula = "=I" & finalRow & "-N" & finalRow
            
            finalRow = finalRow + 1
        End If
    Next i
    
    ' --- D. TOTALS ROW ---
    wsFinal.Cells(finalRow, 2).Value = "TOTAL"
    wsFinal.Cells(finalRow, 2).Font.Bold = True
    
    Dim sumCols As Variant, c As Variant
    sumCols = Array(4, 5, 6, 8, 9, 10, 11, 13, 14, 15)
    
    For Each c In sumCols
        wsFinal.Cells(finalRow, c).Formula = "=SUM(" & Split(wsFinal.Cells(1, c).Address, "$")(1) & "4:" & Split(wsFinal.Cells(1, c).Address, "$")(1) & finalRow - 1 & ")"
    Next c
    
    With wsFinal.Range(wsFinal.Cells(finalRow, 1), wsFinal.Cells(finalRow, 16))
        .Interior.Color = RGB(146, 208, 80)
        .Font.Bold = True
    End With
    
    ' --- E. FORMATTING ---
    With wsFinal.Range(wsFinal.Cells(3, 1), wsFinal.Cells(finalRow, 16))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .ColorIndex = 1
    End With
    
    wsFinal.Range("D4:D" & finalRow).NumberFormat = "#,##0"
    wsFinal.Range("F4:F" & finalRow).NumberFormat = "#,##0"
    wsFinal.Range("H4:J" & finalRow).NumberFormat = "#,##0"
    wsFinal.Range("N4:O" & finalRow).NumberFormat = "#,##0"
    wsFinal.Range("E4:E" & finalRow).NumberFormat = "0.00"
    
    wsFinal.UsedRange.Columns.AutoFit
    wsFinal.Columns("A").ColumnWidth = 5
    wsFinal.Activate
    ActiveWindow.DisplayGridlines = False

End Sub