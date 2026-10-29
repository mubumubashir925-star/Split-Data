VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SPLIT_DATA 
   Caption         =   "SPLIT_DATA"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8040
   OleObjectBlob   =   "SPLIT_DATA 22.10.26.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SPLIT_DATA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub GetPath_Click()

End Sub

Private Sub CMD_Add_Coloumn_Click()
    Dim ws As Worksheet
    Dim headerValue As String
    Dim placeholders As String

    Set ws = ActiveSheet ' Use the active sheet
    
    headerValue = ColoumnPickerCB.Value ' Get the selected header value
    
    If headerValue <> "" Then
        placeholders = "<" & headerValue & ">"
        TextBox1.Value = TextBox1.Value & placeholders
    End If
End Sub

Private Sub CMD_Folder_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    Dim activeWorkbookPath As String
    
    ' Get the folder path of the active workbook
    activeWorkbookPath = ActiveWorkbook.Path

    With fd
        .Title = "Select Folder"
        
        ' Set the initial folder to the active workbook's folder path
        If activeWorkbookPath <> "" Then
            .InitialFileName = activeWorkbookPath
        End If
        
        If .Show = -1 Then
            lastSelectedFolder = .SelectedItems(1) & "\"
            TextBox1.Value = lastSelectedFolder
        End If
    End With
End Sub


Private Sub CMD_Generate_Click()
    Dim ws As Worksheet
    Dim pasteCol As String
    Dim rw As Long
    Dim generatedPath As String
    Dim header As Range
    Dim placeholder As String
    Dim fso As Object
    Dim missingFileCount As Long
    Dim cellText As String
    Dim totalRows As Long
    
    pasteCol = TargetColoumnCB.Value
    Set ws = ActiveSheet
    
    If pasteCol = "" Then
        MsgBox "Please select a column to paste the file path.", vbExclamation
        Exit Sub
    End If
    
    totalRows = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    ' Initialize Safe Progress Bar
    SetProgress 0, totalRows - 1, True
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    missingFileCount = 0
    
    Application.ScreenUpdating = False
    
    For rw = 2 To totalRows
        generatedPath = TextBox1.Value
        
        For Each header In ws.Rows(1).Cells
            If Trim(header.Value & "") <> "" Then
                placeholder = "<" & header.Value & ">"
                
                ' THIS IS THE FINAL FIX – NO MORE "#" EVER
                cellText = SafeCellText(ws.Cells(rw, header.Column))
                
                generatedPath = Replace(generatedPath, placeholder, cellText)
            End If
        Next header
        
        With ws.Cells(rw, Application.Match(pasteCol, ws.Rows(1), 0))
            .Value = generatedPath
            If fso.FileExists(generatedPath) Then
                .Font.Color = vbBlue
            Else
                .Font.Color = vbRed
                missingFileCount = missingFileCount + 1
            End If
        End With
        
        ' Update Safe Progress Bar
        SetProgress rw - 1, totalRows - 1, True
        DoEvents
    Next rw
    
    SetProgress 0, 0, False
    Application.ScreenUpdating = True
    Set fso = Nothing
    
    MsgBox IIf(missingFileCount = 0, "All files found!", missingFileCount & " file(s) not found."), _
           IIf(missingFileCount = 0, vbInformation, vbExclamation), "Finished"
End Sub


'--------------------------------------------------------------
' THE ONLY FUNCTION YOU NEED – NEVER RETURNS "#" OR "##"
'--------------------------------------------------------------
Private Function SafeCellText(c As Range) As String
    Dim tempWidth As Double
    
    If c Is Nothing Then
        SafeCellText = ""
        Exit Function
    End If
    
    ' Save current column width
    tempWidth = c.EntireColumn.ColumnWidth
    
    ' Temporarily make column wide enough so Excel can calculate real .Text
    If tempWidth < 15 Then c.EntireColumn.ColumnWidth = 20
    
    ' Now .Text is 100% reliable
    On Error Resume Next
    SafeCellText = c.Text
    On Error GoTo 0
    
    ' If still empty but cell has a value ? fallback
    If SafeCellText = "" Or Left(SafeCellText, 1) = "#" Then
        If Not IsEmpty(c.Value) Then
            SafeCellText = CStr(c.Value)
        Else
            SafeCellText = ""
        End If
    End If
    
    ' Restore original width instantly (no visual change)
    If tempWidth < 15 Then c.EntireColumn.ColumnWidth = tempWidth
    
End Function


Private Sub CMD_Refresh_Click()
    Call RefreshComboBoxes
End Sub

Private Sub ColoumnPickerCB_Change()

End Sub

Private Sub DltHiddenBlank_Click()
    Dim ws As Worksheet
    Dim total As Long, processed As Long
    Dim r As Long, c As Long
    Dim lastRow As Long, lastCol As Long
    Dim isBlank As Boolean
    
    Application.ScreenUpdating = False
    On Error GoTo CleanExit
    
    '--- Setup Safe Progress Bar ---
    total = ActiveWorkbook.Worksheets.Count
    SetProgress 0, total * 3, True
    processed = 0
    
    '====================================================================
    ' 1. DELETE ALL HIDDEN OR BLANK SHEETS
    '====================================================================
    For Each ws In ActiveWorkbook.Worksheets
        processed = processed + 1
        SetProgress processed, total * 3, True
        
        ' Check if sheet is hidden OR completely blank
        Dim isHidden As Boolean
        Dim isSheetBlank As Boolean
        
        isHidden = (ws.Visible <> xlSheetVisible)
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        isSheetBlank = (lastRow = 0 Or lastCol = 0) Or _
                       (Application.WorksheetFunction.CountA(ws.Cells) = 0)
        
        If isHidden Or isSheetBlank Then
            On Error Resume Next
            ws.Delete
            On Error GoTo 0
        End If
    Next ws
    
    '====================================================================
    ' 2. ON EVERY REMAINING VISIBLE SHEET: DELETE HIDDEN + BLANK ROWS/COLS
    '====================================================================
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            processed = processed + 1
            SetProgress processed, total * 3, True
            
            On Error Resume Next
            ws.Rows.Hidden = False
            ws.Columns.Hidden = False
            On Error GoTo 0
            
            '--- Delete BLANK or HIDDEN ROWS (bottom-up) ---
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
            For r = lastRow To 1 Step -1
                isBlank = (Application.WorksheetFunction.CountA(ws.Rows(r)) = 0)
                If isBlank Or ws.Rows(r).Hidden Then
                    ws.Rows(r).Delete Shift:=xlUp
                End If
            Next r
            
            '--- Delete BLANK or HIDDEN COLUMNS (right-to-left) ---
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            For c = lastCol To 1 Step -1
                isBlank = (Application.WorksheetFunction.CountA(ws.Columns(c)) = 0)
                If isBlank Or ws.Columns(c).Hidden Then
                    ws.Columns(c).Delete Shift:=xlToLeft
                End If
            Next c
        End If
    Next ws
    
CleanExit:
    SetProgress 0, 0, False
    Application.ScreenUpdating = True
    
    If Err.Number = 0 Then
        MsgBox "Cleanup Complete!" & vbCrLf & _
               "• All hidden sheets deleted" & vbCrLf & _
               "• All blank sheets deleted" & vbCrLf & _
               "• All hidden & blank rows/columns removed", _
               vbInformation
    Else
        MsgBox "Error: " & Err.Description, vbCritical
    End If
End Sub

Private Sub Split_Order_Click()
    Dim ws As Worksheet
    Dim totalRows As Long
    Dim i As Long
    Dim targetSheet As Worksheet
    Dim part1Sheet As Worksheet
    Dim part2Sheet As Worksheet
    Dim progress As Integer
    Dim halfPoint As Long

    Set ws = ActiveSheet
    ws.Name = "Original"

    totalRows = ws.UsedRange.Rows.Count
    halfPoint = WorksheetFunction.Ceiling((totalRows - 1) / 2, 1) + 1

    Set part1Sheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    part1Sheet.Name = "Part 1"
    Set part2Sheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    part2Sheet.Name = "Part 2"

    ws.Rows(1).Copy Destination:=part1Sheet.Rows(1)
    ws.Rows(1).Copy Destination:=part2Sheet.Rows(1)

    ' Initialize Safe Progress Bar
    SetProgress 0, totalRows, True

    For i = 2 To totalRows
        If i <= halfPoint Then
            Set targetSheet = part1Sheet
        Else
            Set targetSheet = part2Sheet
        End If
        ws.Rows(i).Copy Destination:=targetSheet.Rows(targetSheet.UsedRange.Rows.Count + 1)
        
        ' Update Safe Progress Bar
        progress = i
        SetProgress progress, totalRows, True
    Next i

    SetProgress 0, 0, False
End Sub

Private Sub ConvertToSingleLine_Click()
    Dim cell As Range
    Dim lines As Variant
    Dim i As Long
    Dim result As String
    Dim processRange As Range
    Dim processedCount As Long
    Dim totalCells As Long
    
    If ActiveSheet Is Nothing Then
        MsgBox "No worksheet is active. Please select a worksheet and try again.", vbExclamation, "No Active Worksheet"
        Exit Sub
    End If
    
    If TypeName(Application.Selection) <> "Range" Then
        MsgBox "Please select a column or a range of cells first.", vbExclamation, "Selection Error"
        Exit Sub
    ElseIf Application.Selection.Cells.Count <= 1 Then
        MsgBox "You only selected a single cell. Please select an entire column or highlight multiple cells to convert.", vbExclamation, "Selection Error"
        Exit Sub
    End If
    
    Set processRange = Application.Intersect(Application.Selection, ActiveSheet.UsedRange)
    
    If processRange Is Nothing Then
        MsgBox "There is no data in your selected cells to convert.", vbInformation, "No Data"
        Exit Sub
    End If

    ' Setup Safe Progress Bar
    totalCells = processRange.Cells.Count
    SetProgress 0, totalCells, True
    processedCount = 0
    
    Application.ScreenUpdating = False
    
    On Error GoTo ErrorHandler
    For Each cell In processRange
        If Not IsEmpty(cell) Then
            If VarType(cell.Value) = vbString And Not IsError(cell.Value) Then
                lines = VBA.Split(cell.Value, Chr(10))
                result = ""
                For i = LBound(lines) To UBound(lines)
                    If Trim(lines(i)) <> "" Then
                        result = result & lines(i)
                        If i < UBound(lines) Then
                            result = result & ", "
                        End If
                    End If
                Next i
                cell.Value = result
            End If
        End If
        
        processedCount = processedCount + 1
        SetProgress processedCount, totalCells, True
    Next cell
    
    Application.ScreenUpdating = True
    
    SetProgress 0, 0, False
    MsgBox "Conversion completed successfully!" & vbCrLf & _
           "Processed cells: " & processedCount, vbInformation, "Done"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    SetProgress 0, 0, False
    MsgBox "An error occurred while processing the text in cell " & cell.Address & ". Ensure the selected cells contain text with line breaks.", vbCritical, "Processing Error"
End Sub

Private Sub Frame1_Click()

End Sub


Private Sub Save_TXT_Click()
    Dim ws As Worksheet
    Dim rng As Range
    Dim filePath As String
    Dim outputText As String
    Dim cellValue As Variant
    Dim i As Integer, j As Integer
    Dim defaultFileName As String
    Dim FolderPath As String
    Dim response As VbMsgBoxResult

    Set ws = ActiveSheet
    Set rng = ws.UsedRange
    rng.Columns.AutoFit

    FolderPath = ActiveWorkbook.Path
    If FolderPath = "" Then
        MsgBox "Please save the workbook first.", vbExclamation
        Exit Sub
    End If

    defaultFileName = Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1) & ".txt"
    filePath = Application.GetSaveAsFilename(InitialFileName:=FolderPath & "\" & defaultFileName, _
                                             fileFilter:="Text Files (*.txt), *.txt")
    If filePath = "False" Then Exit Sub

    If Dir(filePath) <> "" Then
        If MsgBox("File exists. Replace?", vbYesNo) = vbNo Then Exit Sub
    End If

    For i = 1 To rng.Rows.Count
        For j = 1 To rng.Columns.Count
            cellValue = rng.Cells(i, j).Text

            If cellValue = "" Or IsError(rng.Cells(i, j)) Then
                If IsDate(rng.Cells(i, j).Value) Then
                    cellValue = Format(rng.Cells(i, j).Value, rng.Cells(i, j).NumberFormat)
                Else
                    cellValue = rng.Cells(i, j).Value
                End If
            End If

            If Not IsNull(cellValue) Then
                cellValue = CStr(cellValue)
                cellValue = Replace(cellValue, vbCrLf, " ")
                cellValue = Replace(cellValue, vbLf, " ")
                cellValue = Replace(cellValue, vbCr, " ")
                cellValue = Trim(cellValue)
            Else
                cellValue = ""
            End If

            If j < rng.Columns.Count Then
                outputText = outputText & cellValue & vbTab
            Else
                outputText = outputText & cellValue
            End If
        Next j
        If i < rng.Rows.Count Then outputText = outputText & vbCrLf
    Next i

    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(filePath, True, True)
    ts.Write outputText
    ts.Close

    MsgBox "Exported successfully: " & filePath, vbInformation
End Sub


Private Sub Label3_Click()

End Sub


Private Sub CompineVertically_Click()
    Dim ws As Worksheet
    Dim combinedSheet As Worksheet
    Dim lastRow As Long
    Dim lastColumn As Long
    Dim nextRow As Long
    Dim totalSheets As Long
    Dim processedSheets As Long
    Dim findRange As Range

    Set combinedSheet = ActiveWorkbook.Sheets.Add(After:= _
             ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    combinedSheet.Name = "Combined Sheet"
    nextRow = 1

    ' Initialize Safe Progress Bar
    totalSheets = ActiveWorkbook.Worksheets.Count - 1 ' Exclude the combined sheet
    processedSheets = 0
    SetProgress 0, totalSheets, True

    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> combinedSheet.Name Then
            
            ' --- Check if the sheet is empty to avoid Error 91 ---
            On Error Resume Next
            Set findRange = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
            On Error GoTo 0
            
            ' Only process if findRange is not Nothing (sheet is NOT empty)
            If Not findRange Is Nothing Then
                lastRow = findRange.row
                lastColumn = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
                
                ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastColumn)).Copy combinedSheet.Cells(nextRow, 1)
                nextRow = combinedSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row + 1
            End If
            ' -----------------------------------------------------
            
            ' Update Safe Progress Bar
            processedSheets = processedSheets + 1
            SetProgress processedSheets, totalSheets, True
        End If
    Next ws

    SetProgress 0, 0, False
End Sub


Private Sub CompineHorizontally_Click()
    Dim ws As Worksheet
    Dim combinedSheet As Worksheet
    Dim lastRow As Long
    Dim lastColumn As Long
    Dim nextColumn As Long
    Dim totalSheets As Long
    Dim processedSheets As Long
    Dim findRange As Range

    Set combinedSheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    combinedSheet.Name = "Combined Sheet"
    nextColumn = 1

    ' Initialize Safe Progress Bar
    totalSheets = ActiveWorkbook.Worksheets.Count - 1 ' Exclude the combined sheet
    processedSheets = 0
    SetProgress 0, totalSheets, True

    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> combinedSheet.Name And ws.Name <> "Original" Then
            
            ' --- Check if the sheet is empty to avoid Error 91 ---
            On Error Resume Next
            Set findRange = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
            On Error GoTo 0
            
            ' Only process if findRange is not Nothing (sheet is NOT empty)
            If Not findRange Is Nothing Then
                lastRow = findRange.row
                lastColumn = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
                
                ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastColumn)).Copy combinedSheet.Cells(1, nextColumn)
                nextColumn = combinedSheet.Cells(1, combinedSheet.Columns.Count).End(xlToLeft).Column + 1
            End If
            ' -----------------------------------------------------

            ' Update Safe Progress Bar
            processedSheets = processedSheets + 1
            SetProgress processedSheets, totalSheets, True
        End If
    Next ws

    SetProgress 0, 0, False
End Sub

Private Sub Label1_Click()

End Sub

Private Sub SplitAsComboCount_Click()
    Dim ws As Worksheet
    Dim totalRows As Long
    Dim i As Long
    Dim targetSheet As Worksheet
    Dim part1Sheet As Worksheet
    Dim part2Sheet As Worksheet
    Dim part3Sheet As Worksheet
    Dim part4Sheet As Worksheet
    Dim progress As Integer
    Dim splitCount As Integer
    Dim baseRows As Long
    Dim halfPoint As Long
    Dim extraRows As Long
    Dim part1Rows As Long
    Dim part2Rows As Long
    Dim part3Rows As Long
    
    ' Check if controls exist
    If Not HasControl(Me, "AltSplit") Or Not HasControl(Me, "OrderSplit") Then
        MsgBox "Option buttons AltSplit or OrderSplit not found. Please check UserForm controls.", vbCritical
        Exit Sub
    End If
    If Not HasControl(Me, "SplitCountCombo") Then
        MsgBox "SplitCountCombo not found. Please check UserForm controls.", vbCritical
        Exit Sub
    End If
    
    If Not Me.AltSplit.Value And Not Me.OrderSplit.Value Then
        MsgBox "Please select a splitting method.", vbExclamation
        Exit Sub
    End If
    
    Select Case Me.SplitCountCombo.Value
        Case "Two Parts"
            splitCount = 2
        Case "Four Parts"
            splitCount = 4
        Case Else
            MsgBox "Please select the number of parts (Two Parts or Four Parts) in SplitCountCombo.", vbExclamation
            Exit Sub
    End Select
    
    On Error Resume Next
    Set ws = ActiveSheet
    If ws Is Nothing Then
        MsgBox "No active worksheet found. Please select a worksheet.", vbCritical
        Exit Sub
    End If
    ws.Name = "Original"
    On Error GoTo 0
    
    totalRows = ws.UsedRange.Rows.Count
    If totalRows < 2 Then
        MsgBox "Worksheet has insufficient rows (must have at least one header and one data row).", vbExclamation
        SetProgress 0, 0, False
        Exit Sub
    End If
    
    Set part1Sheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    part1Sheet.Name = "Part 1"
    Set part2Sheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    part2Sheet.Name = "Part 2"
    If splitCount = 4 Then
        Set part3Sheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
        part3Sheet.Name = "Part 3"
        Set part4Sheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
        part4Sheet.Name = "Part 4"
    End If
    
    ws.Rows(1).Copy Destination:=part1Sheet.Rows(1)
    ws.Rows(1).Copy Destination:=part2Sheet.Rows(1)
    If splitCount = 4 Then
        ws.Rows(1).Copy Destination:=part3Sheet.Rows(1)
        ws.Rows(1).Copy Destination:=part4Sheet.Rows(1)
    End If
    
    ' Initialize Safe Progress Bar
    SetProgress 0, totalRows, True
    
    If Me.AltSplit.Value Then
        For i = 2 To totalRows
            If splitCount = 2 Then
                If i Mod 2 = 0 Then
                    Set targetSheet = part1Sheet
                Else
                    Set targetSheet = part2Sheet
                End If
            Else
                Select Case (i - 2) Mod 4
                    Case 0
                        Set targetSheet = part1Sheet
                    Case 1
                        Set targetSheet = part2Sheet
                    Case 2
                        Set targetSheet = part3Sheet
                    Case 3
                        Set targetSheet = part4Sheet
                End Select
            End If
            ws.Rows(i).Copy Destination:=targetSheet.Rows(targetSheet.UsedRange.Rows.Count + 1)
            
            progress = i
            SetProgress progress, totalRows, True
        Next i
    ElseIf Me.OrderSplit.Value Then
        If splitCount = 2 Then
            halfPoint = WorksheetFunction.Ceiling((totalRows - 1) / 2, 1) + 1
            For i = 2 To totalRows
                If i <= halfPoint Then
                    Set targetSheet = part1Sheet
                Else
                    Set targetSheet = part2Sheet
                End If
                ws.Rows(i).Copy Destination:=targetSheet.Rows(targetSheet.UsedRange.Rows.Count + 1)
                
                progress = i
                SetProgress progress, totalRows, True
            Next i
        Else
            baseRows = Int((totalRows - 1) / 4)
            extraRows = (totalRows - 1) Mod 4
            part1Rows = baseRows + IIf(extraRows >= 1, 1, 0)
            part2Rows = baseRows + IIf(extraRows >= 2, 1, 0)
            part3Rows = baseRows + IIf(extraRows >= 3, 1, 0)
            For i = 2 To totalRows
                If i <= part1Rows + 1 Then
                    Set targetSheet = part1Sheet
                ElseIf i <= (part1Rows + part2Rows + 1) Then
                    Set targetSheet = part2Sheet
                ElseIf i <= (part1Rows + part2Rows + part3Rows + 1) Then
                    Set targetSheet = part3Sheet
                Else
                    Set targetSheet = part4Sheet
                End If
                ws.Rows(i).Copy Destination:=targetSheet.Rows(targetSheet.UsedRange.Rows.Count + 1)
                
                progress = i
                SetProgress progress, totalRows, True
            Next i
        End If
    End If
    
    SetProgress 0, 0, False
End Sub

' Helper function to check if a control exists
Private Function HasControl(frm As Object, controlName As String) As Boolean
    On Error Resume Next
    HasControl = Not frm.Controls(controlName) Is Nothing
    On Error GoTo 0
End Function

' =========================================================
' SPLIT DATA PRO - MODERN NATIVE USERFORM THEME
' Designed to preserve existing controls and event handlers.
' Colors use native VBA RGB values; no external dependencies.
' =========================================================
Private Sub ApplyModernTheme()
    On Error Resume Next

    Dim ctl As Object
    Dim themeFont As String
    themeFont = "Segoe UI"

    Me.Caption = "Split Data Pro  |  Excel Data Tools"
    Me.BackColor = RGB(246, 248, 252)

    For Each ctl In Me.Controls
        Err.Clear
        Select Case TypeName(ctl)
            Case "CommandButton"
                ctl.Font.Name = themeFont
                ctl.Font.Size = 9
                ctl.Font.Bold = True
                ctl.BackColor = RGB(235, 241, 250)
                ctl.ForeColor = RGB(35, 57, 92)
                ctl.TakeFocusOnClick = False
                ctl.ControlTipText = "Click to " & Replace(Replace(ctl.Name, "CMD_", ""), "_", " ")
            Case "TextBox", "ComboBox"
                ctl.Font.Name = themeFont
                ctl.Font.Size = 10
                ctl.BackColor = RGB(255, 255, 255)
                ctl.ForeColor = RGB(30, 41, 59)
                ctl.SpecialEffect = 2  ' fmSpecialEffectSunken
            Case "Label"
                ctl.Font.Name = themeFont
                ctl.ForeColor = RGB(71, 85, 105)
                ctl.BackStyle = 0  ' fmBackStyleTransparent
            Case "Frame"
                ctl.Font.Name = themeFont
                ctl.Font.Bold = True
                ctl.ForeColor = RGB(37, 62, 102)
                ctl.BackColor = RGB(255, 255, 255)
            Case "OptionButton", "CheckBox"
                ctl.Font.Name = themeFont
                ctl.ForeColor = RGB(51, 65, 85)
                ctl.BackColor = RGB(246, 248, 252)
        End Select
    Next ctl

    ' Emphasize the action buttons without altering their positions.
    StylePrimaryButton "Split"
    StylePrimaryButton "CMD_Generate"
    StylePrimaryButton "SplitAsComboCount"
    StylePrimaryButton "CompineVertically"
    StylePrimaryButton "CompineHorizontally"

    Me.Repaint
    On Error GoTo 0
End Sub

Private Sub StylePrimaryButton(ByVal controlName As String)
    On Error Resume Next
    With Me.Controls(controlName)
        .BackColor = RGB(37, 99, 235)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    On Error GoTo 0
End Sub

Private Sub UserForm_Initialize()
    ApplyModernTheme
    Dim ws As Worksheet
    Dim col As Range
    Set ws = ActiveSheet ' Use the active sheet

    ' Populate ColoumnPickerCB and TargetColoumnCB with header values from the first row
    For Each col In ws.Rows(1).Cells ' Assuming headers are in the first row
        If col.Value <> "" Then
            ColoumnPickerCB.AddItem col.Value
            TargetColoumnCB.AddItem col.Value
        End If
    Next col
    SplitCountCombo.AddItem "Two Parts"
    SplitCountCombo.AddItem "Four Parts"
    SplitCountCombo.Value = "Two Parts" ' Optional default
End Sub

Private Sub RefreshComboBoxes()
    Dim ws As Worksheet
    Dim col As Range
    Set ws = ActiveSheet ' Use the active sheet
    
    ' Clear existing items
    ColoumnPickerCB.Clear
    TargetColoumnCB.Clear
    
    ' Populate ColoumnPickerCB and TargetColoumnCB with header values from the first row
    For Each col In ws.Rows(1).Cells ' Assuming headers are in the first row
        If col.Value <> "" Then
            ColoumnPickerCB.AddItem col.Value
            TargetColoumnCB.AddItem col.Value
        End If
    Next col
End Sub

Sub SplitIntoSheets(rowCount As Long)
    Dim ws As Worksheet
    Dim partWs As Worksheet
    Dim lastRow As Long
    Dim headerRange As Range
    Dim partIndex As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim progress As Integer

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set ws = ActiveSheet
    ws.Name = "Original"
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    Set headerRange = ws.Rows(1)
    partIndex = 1
    startRow = 2

    ' Initialize Safe Progress Bar
    SetProgress 0, lastRow, True

    Do While startRow <= lastRow
        endRow = startRow + rowCount - 1
        If endRow > lastRow Then endRow = lastRow

        Set partWs = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
        
        If Not partWs Is Nothing Then
            partWs.Name = "Part" & partIndex
            headerRange.Copy Destination:=partWs.Rows(1)
            ws.Rows(startRow & ":" & endRow).Copy Destination:=partWs.Rows(2)

            startRow = endRow + 1
            partIndex = partIndex + 1

            ' Update Safe Progress Bar
            progress = startRow
            If progress > lastRow Then progress = lastRow
            SetProgress progress, lastRow, True
        Else
            MsgBox "Failed to create a new worksheet.", vbCritical, "Error"
            Exit Do
        End If
    Loop

    SetProgress 0, 0, False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Sub SplitIntoWorkbooks(rowCount As Long)
    Dim ws As Worksheet
    Dim partWs As Worksheet
    Dim lastRow As Long
    Dim headerRange As Range
    Dim partIndex As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim newWorkbook As Workbook
    Dim FolderPath As String
    Dim FileName As String
    Dim sheet As Worksheet
    Dim progress As Integer
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Set ws = ActiveSheet
    ws.Name = "Original"
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    Set headerRange = ws.Rows(1)
    partIndex = 1
    startRow = 2
    FolderPath = ActiveWorkbook.Path
    
    ' Initialize Safe Progress Bar
    SetProgress 0, lastRow, True
    
    Do While startRow <= lastRow
        endRow = startRow + rowCount - 1
        If endRow > lastRow Then endRow = lastRow
        
        Set newWorkbook = Workbooks.Add
        
        Application.DisplayAlerts = False
        For Each sheet In newWorkbook.Sheets
            If sheet.Name <> "Sheet1" Then
                sheet.Delete
            End If
        Next sheet
        Application.DisplayAlerts = True
        
        Set partWs = newWorkbook.Sheets("Sheet1")
        headerRange.Copy Destination:=partWs.Rows(1)
        ws.Rows(startRow & ":" & endRow).Copy Destination:=partWs.Rows(2)
        
        FileName = FolderPath & "\Part " & partIndex & ".xlsx"
        newWorkbook.SaveAs FileName
        newWorkbook.Close
        
        startRow = endRow + 1
        partIndex = partIndex + 1
        
        ' Update Safe Progress Bar
        progress = startRow
        If progress > lastRow Then progress = lastRow
        SetProgress progress, lastRow, True
    Loop
    
    SetProgress 0, 0, False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub


Private Sub Split_Click()
    Dim rowCount As Long
    
    If No_of_rows.Value = "" Or Not IsNumeric(No_of_rows.Value) Then
        MsgBox "Please enter a valid number of rows.", vbExclamation, "Input Error"
        Exit Sub
    End If
    
    rowCount = CLng(No_of_rows.Value)
    
    If to_sheet.Value = False And to_workbook.Value = False Then
        MsgBox "Please select an option.", vbExclamation, "Selection Error"
        Exit Sub
    End If
    
    If to_sheet.Value = True Then
        Call SplitIntoSheets(rowCount)
    ElseIf to_workbook.Value = True Then
        Call SplitIntoWorkbooks(rowCount)
    End If
    
End Sub

' =========================================================
' UNIVERSAL SAFE PROGRESS BAR HELPER
' =========================================================
Private Sub SetProgress(ByVal currentVal As Long, ByVal maxVal As Long, Optional ByVal isVisible As Boolean = True)
    On Error Resume Next ' Prevents the "missing" error from stopping the macro
    
    Dim pb As Object
    Set pb = Me.Controls("ProgressBar1")
    
    If Not pb Is Nothing Then
        pb.Visible = isVisible
        If isVisible And maxVal > 0 Then
            pb.Max = maxVal
            ' Ensure value doesn't exceed max or drop below min
            If currentVal > maxVal Then currentVal = maxVal
            If currentVal < 0 Then currentVal = 0
            pb.Value = currentVal
        End If
    End If
    
    On Error GoTo 0
End Sub

Private Sub UserForm_Click()

End Sub

