Option Explicit

Sub GSTR2B_Consolidation_Click()

    Dim folderPath As String
    Dim FSO As Object, targetFldr As Variant
    Dim fileObj As Object
    Dim wbSource As Workbook, wbFinal As Workbook
    Dim wsCons As Worksheet, wsSource As Worksheet, wsPivot As Worksheet
    Dim destRow As Long, srcLastRow As Long
    Dim sName As Variant
    Dim monthDate As Date
    Dim i As Long, r As Long, tgtCol As Long, srcCol As Long

    Dim pc As PivotCache, pt As PivotTable, pvtRange As Range
    Dim valFields As Variant, v As Variant
    Dim savePath As Variant

    ' Variables for dynamic year tracking
    Dim minDate As Date, maxDate As Date
    Dim fYearStr As String
    minDate = DateSerial(2100, 1, 1) ' Set artificially high to start
    maxDate = DateSerial(1900, 1, 1) ' Set artificially low to start

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select your GSTR-2B Folder"
        If .Show = -1 Then
            folderPath = .SelectedItems(1) & "\"
        Else
            Exit Sub
        End If
    End With

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set targetFldr = FSO.GetFolder(folderPath)

    Set wbFinal = Workbooks.Add
    Set wsCons = wbFinal.Sheets(1)
    wsCons.Name = "Purchase 2B"

    While wbFinal.Sheets.Count > 1
        wbFinal.Sheets(2).Delete
    Wend

    wsCons.Cells(1, 1).Value = "Company Name Private Limited"
    wsCons.Cells(1, 1).Font.Bold = True
    ' The sub-title in A2 will be dynamically updated at the end of the script
    wsCons.Cells(2, 1).Font.Bold = True

    Dim headers() As Variant
    headers = Array("GSTIN of supplier", "Trade/Legal name", "Invoice number", "Invoice type", _
                    "Invoice Date", "Invoice Value(?)", "Place of supply", "Supply Attract Reverse Charge", _
                    "Taxable Value (?)", "Integrated Tax(?)", "Central Tax(?)", "State/UT Tax(?)", _
                    "Cess(?)", "GSTR-1/IFF/GSTR-5 Period", "GSTR-1/IFF/GSTR-5 Filing Date", "ITC Availability", _
                    "Reason", "Applicable % of Tax Rate", "Source", "IRN", "IRN Date", "Month")

    For i = 0 To UBound(headers)
        wsCons.Cells(4, i + 1).Value = headers(i)
        wsCons.Cells(4, i + 1).Font.Bold = True
    Next i

    destRow = 5

    For Each fileObj In targetFldr.Files
        If InStr(1, fileObj.Name, ".xls", vbTextCompare) > 0 And Left(fileObj.Name, 2) <> "~$" Then

            On Error Resume Next
            Dim m As Integer, y As Integer
            m = CInt(Left(fileObj.Name, 2))
            y = CInt(Mid(fileObj.Name, 3, 4))
            monthDate = DateSerial(y, m, 1)

            ' Track the oldest and newest dates dynamically
            If monthDate < minDate Then minDate = monthDate
            If monthDate > maxDate Then maxDate = monthDate
            On Error GoTo 0

            Set wbSource = Nothing
            On Error Resume Next
            Set wbSource = Workbooks.Open(fileObj.Path, ReadOnly:=True, UpdateLinks:=False)
            On Error GoTo 0

            If Not wbSource Is Nothing Then
                Dim sheetNames As Variant
                sheetNames = Array("B2B", "B2B-CDNR", "B2BA", "B2B-CDNRA")

                For Each sName In sheetNames
                    Set wsSource = Nothing
                    On Error Resume Next
                    Set wsSource = wbSource.Sheets(sName)
                    On Error GoTo 0

                    If Not wsSource Is Nothing Then
                        Dim startRow As Long
                        Dim colMap As Variant

                        If sName = "B2B" Then
                            startRow = 7
                            colMap = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21)

                        ElseIf sName = "B2B-CDNR" Then
                            startRow = 7
                            colMap = Array(0, 1, 2, 3, 4, 6, 7, 8, 9, 10, 11, 12, 13, 14, 21, 22, 23, 24, 25, 26, 27, 28)

                        ElseIf sName = "B2BA" Then
                            startRow = 8
                            colMap = Array(0, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 22, 23, 24, 25, 26, 0, 0, 0)

                        ElseIf sName = "B2B-CDNRA" Then
                            startRow = 8
                            colMap = Array(0, 4, 5, 6, 7, 9, 10, 11, 12, 13, 14, 15, 16, 17, 24, 25, 26, 27, 28, 0, 0, 0)
                        End If

                        Dim reduceCol As Long
                        reduceCol = 0
                        If sName = "B2BA" Or sName = "B2B-CDNRA" Then
                            Dim headCol As Long
                            For headCol = 10 To 25
                                If InStr(1, wsSource.Cells(startRow - 1, headCol).Value, "reduced", vbTextCompare) > 0 Then
                                    reduceCol = headCol
                                    Exit For
                                End If
                            Next headCol
                        End If

                        Dim lastCell As Range
                        Set lastCell = wsSource.Cells.Find(What:="*", LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
                        If Not lastCell Is Nothing Then
                            srcLastRow = lastCell.Row
                        Else
                            srcLastRow = 0
                        End If

                        If srcLastRow >= startRow Then
                            For r = startRow To srcLastRow

                                Dim isBlank As Boolean, isTotal As Boolean
                                isBlank = True
                                isTotal = False

                                Dim checkCol As Long
                                For checkCol = 1 To 10
                                    Dim cellTxt As String
                                    cellTxt = Trim(CStr(wsSource.Cells(r, checkCol).Value))
                                    If cellTxt <> "" Then isBlank = False
                                    If UCase(cellTxt) = "TOTAL" Or UCase(cellTxt) = "GRAND TOTAL" Then isTotal = True
                                Next checkCol

                                If isTotal Then Exit For
                                If isBlank Then GoTo SkipRow

                                Dim isCredit As Boolean
                                Dim isReduced As Boolean
                                isCredit = False
                                isReduced = False

                                If sName = "B2B-CDNR" Then
                                    If InStr(1, UCase(Trim(CStr(wsSource.Cells(r, 4).Value))), "CREDIT") > 0 Or Left(UCase(Trim(CStr(wsSource.Cells(r, 4).Value))), 1) = "C" Then isCredit = True
                                ElseIf sName = "B2B-CDNRA" Then
                                    If InStr(1, UCase(Trim(CStr(wsSource.Cells(r, 7).Value))), "CREDIT") > 0 Or Left(UCase(Trim(CStr(wsSource.Cells(r, 7).Value))), 1) = "C" Then isCredit = True
                                End If

                                If sName = "B2BA" Or sName = "B2B-CDNRA" Then
                                    If reduceCol > 0 Then
                                        If Left(UCase(Trim(CStr(wsSource.Cells(r, reduceCol).Value))), 1) = "Y" Then isReduced = True
                                    End If
                                End If

                                For tgtCol = 1 To 21
                                    srcCol = colMap(tgtCol)

                                    If srcCol > 0 Then
                                        Dim val As Variant
                                        val = wsSource.Cells(r, srcCol).Value

                                        If IsNumeric(val) And val <> "" Then
                                            If tgtCol = 6 Or tgtCol = 9 Or tgtCol = 10 Or tgtCol = 11 Or tgtCol = 12 Or tgtCol = 13 Then
                                                If (isCredit Or isReduced) And val > 0 Then
                                                    val = val * -1
                                                End If
                                            End If
                                        End If

                                        wsCons.Cells(destRow, tgtCol).Value = val
                                    End If
                                Next tgtCol

                                ' ========================================================
                                ' NEW LOGIC: HIGHLIGHT B2BA / B2B-CDNRA AMENDMENTS
                                ' ========================================================
                                If sName = "B2BA" Or sName = "B2B-CDNRA" Then
                                    ' Applies a light orange/yellow color across all 22 columns for this row
                                    wsCons.Range(wsCons.Cells(destRow, 1), wsCons.Cells(destRow, 22)).Interior.Color = RGB(255, 235, 156)
                                End If

                                wsCons.Cells(destRow, 22).Value = monthDate
                                destRow = destRow + 1

SkipRow:
                            Next r
                        End If
                    End If
                Next sName

                wbSource.Close False
            Else
                MsgBox "Excel blocked VBA from opening: " & fileObj.Name, vbExclamation
            End If
        End If
    Next fileObj

    wsCons.Rows(4).Font.Bold = True
    wsCons.Columns(22).NumberFormat = "mmm-yy"
    wsCons.UsedRange.Columns.AutoFit

    ' ========================================================
    ' NEW LOGIC: APPLY ALL BORDERS TO DATA RANGE
    ' ========================================================
    If destRow > 5 Then
        With wsCons.Range(wsCons.Cells(4, 1), wsCons.Cells(destRow - 1, 22))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.ColorIndex = 1 ' Black border color
        End With
    End If

    If destRow > 5 Then
        Set wsPivot = wbFinal.Sheets.Add(Before:=wsCons)
        wsPivot.Name = "Pivot_Report"

        Set pvtRange = wsCons.Range(wsCons.Cells(4, 1), wsCons.Cells(destRow - 1, UBound(headers) + 1))
        Set pc = wbFinal.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pvtRange)
        Set pt = pc.CreatePivotTable(TableDestination:=wsPivot.Range("A3"), TableName:="PVT_Report")

        With pt.PivotFields("Month")
            .Orientation = xlRowField
            .Position = 1
        End With

        valFields = Array("Taxable Value (?)", "Integrated Tax(?)", "Central Tax(?)", "State/UT Tax(?)")

        For Each v In valFields
            On Error Resume Next
            With pt.PivotFields(v)
                .Orientation = xlDataField
                .Function = xlSum
                .NumberFormat = "#,##0"
                .Caption = "Sum of " & Replace(v, "(?)", "")
            End With
            On Error GoTo 0
        Next v

        On Error Resume Next
        With pt.PivotFields("ITC Availability")
            .Orientation = xlPageField
            .Position = 1
        End With
        With pt.PivotFields("Supply Attract Reverse Charge")
            .Orientation = xlPageField
            .Position = 2
        End With
        With pt.PivotFields("Invoice type")
            .Orientation = xlPageField
            .Position = 3
        End With
        On Error GoTo 0

        pt.RowAxisLayout xlTabularRow
        wsPivot.Columns("A:E").AutoFit
    End If

    ' ========================================================
    ' DYNAMIC YEAR LOGIC & SAVE AS PROMPT
    ' ========================================================
    If maxDate >= minDate Then
        ' Generate year string (e.g., "2025-26" or just "2025")
        If Year(minDate) = Year(maxDate) Then
            fYearStr = CStr(Year(minDate))
        Else
            fYearStr = Year(minDate) & "-" & Right(CStr(Year(maxDate)), 2)
        End If
        ' Update the actual sheet title dynamically
        wsCons.Cells(2, 1).Value = "GSTR-2B Report from " & Format(minDate, "mmm-yy") & " to " & Format(maxDate, "mmm-yy")
    Else
        fYearStr = "Year" ' Fallback if no files were found
    End If

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    savePath = Application.GetSaveAsFilename( _
        InitialFileName:="GST Reconciliation of " & fYearStr & ".xlsx", _
        FileFilter:="Excel Workbook (*.xlsx), *.xlsx", _
        Title:="Save Consolidated GST Report As")

    If savePath <> False Then
        Application.DisplayAlerts = False
        wbFinal.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
        Application.DisplayAlerts = True
        MsgBox "Data successfully extracted and saved as 'GST Reconciliation of " & fYearStr & ".xlsx'!", vbInformation
    Else
        MsgBox "Extraction complete! (Note: You canceled the save dialog, so the file is unsaved).", vbInformation
    End If

End Sub
