Attribute VB_Name = "DelaxToolsExtra"

Option Explicit

' DelaxTools change-history event state.
' DelaxTools v166
' Module-level declarations must be placed before all procedures.
Private gDxaEvents As DelaxToolsAppEvents
Private gDxaSessionId As String


#If VBA7 Then
Private Type DxaDataBlob
    cbData As Long
    pbData As LongPtr
End Type

Private Declare PtrSafe Function DxaCryptProtectData Lib "crypt32.dll" Alias "CryptProtectData" (ByRef pDataIn As DxaDataBlob, ByVal szDataDescr As LongPtr, ByVal pOptionalEntropy As LongPtr, ByVal pvReserved As LongPtr, ByVal pPromptStruct As LongPtr, ByVal dwFlags As Long, ByRef pDataOut As DxaDataBlob) As Long
Private Declare PtrSafe Function DxaCryptUnprotectData Lib "crypt32.dll" Alias "CryptUnprotectData" (ByRef pDataIn As DxaDataBlob, ByVal ppszDataDescr As LongPtr, ByVal pOptionalEntropy As LongPtr, ByVal pvReserved As LongPtr, ByVal pPromptStruct As LongPtr, ByVal dwFlags As Long, ByRef pDataOut As DxaDataBlob) As Long
Private Declare PtrSafe Sub DxaCopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As LongPtr)
Private Declare PtrSafe Function DxaLocalFree Lib "kernel32" Alias "LocalFree" (ByVal hMem As LongPtr) As LongPtr
#Else
Private Type DxaDataBlob
    cbData As Long
    pbData As Long
End Type

Private Declare Function DxaCryptProtectData Lib "crypt32.dll" Alias "CryptProtectData" (ByRef pDataIn As DxaDataBlob, ByVal szDataDescr As Long, ByVal pOptionalEntropy As Long, ByVal pvReserved As Long, ByVal pPromptStruct As Long, ByVal dwFlags As Long, ByRef pDataOut As DxaDataBlob) As Long
Private Declare Function DxaCryptUnprotectData Lib "crypt32.dll" Alias "CryptUnprotectData" (ByRef pDataIn As DxaDataBlob, ByVal ppszDataDescr As Long, ByVal pOptionalEntropy As Long, ByVal pvReserved As Long, ByVal pPromptStruct As Long, ByVal dwFlags As Long, ByRef pDataOut As DxaDataBlob) As Long
Private Declare Sub DxaCopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function DxaLocalFree Lib "kernel32" Alias "LocalFree" (ByVal hMem As Long) As Long
#End If



' DelaxTools v108

' 自動アップデート機能は含めていません。

' 追加機能はExcel内VBAとして実行します。



Public Sub DxaCreateHolidaySheet(ByVal control As Object)

    On Error GoTo EH

    Dim yText As String

    yText = InputBox("休日一覧を作成する年を入力してください。", "休日シート作成", CStr(Year(Date)))

    If Len(Trim$(yText)) = 0 Then Exit Sub

    If Not IsNumeric(yText) Then

        MsgBox "年は数値で入力してください。", vbExclamation, "DelaxTools"

        Exit Sub

    End If



    Dim y As Long

    y = CLng(yText)

    If y < 1900 Or y > 2100 Then

        MsgBox "1900～2100の範囲で入力してください。", vbExclamation, "DelaxTools"

        Exit Sub

    End If



    Dim wb As Workbook

    Set wb = ActiveWorkbook

    If wb Is Nothing Then Exit Sub



    Dim sheetName As String

    sheetName = "休日" & CStr(y)



    Application.ScreenUpdating = False

    Application.DisplayAlerts = False

    Dim ws As Worksheet

    On Error Resume Next

    Set ws = wb.Worksheets(sheetName)

    On Error GoTo EH

    If Not ws Is Nothing Then ws.Delete

    Application.DisplayAlerts = True



    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))

    ws.Name = sheetName



    Dim items As Object

    Set items = CreateObject("Scripting.Dictionary")

    AddJapaneseHolidays y, items



    ws.Range("A1:C1").Value = Array("日付", "曜日", "休日名")

    ws.Range("A1:C1").Font.Bold = True



    Dim keys As Variant

    keys = items.Keys

    SortDateKeys keys



    Dim r As Long, k As Variant

    r = 2

    For Each k In keys

        ws.Cells(r, 1).Value = CDate(k)

        ws.Cells(r, 2).Value = JapaneseWeekday(CDate(k))

        ws.Cells(r, 3).Value = items(k)

        r = r + 1

    Next



    ws.Columns("A:A").NumberFormatLocal = "yyyy/mm/dd"

    ws.Columns("A:C").AutoFit

    ws.Range("A1:C1").AutoFilter

    ws.Activate

    ws.Range("A1").Select

    Application.ScreenUpdating = True

    MsgBox CStr(y) & "年の休日一覧を作成しました。", vbInformation, "DelaxTools"

    Exit Sub

EH:

    Application.DisplayAlerts = True

    Application.ScreenUpdating = True

    MsgBox "休日シート作成でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "DelaxTools"

End Sub



Public Sub DxaAllSheetsZoom100(ByVal control As Object)

    On Error GoTo EH

    Dim wb As Workbook

    Set wb = ActiveWorkbook

    If wb Is Nothing Then Exit Sub



    Dim activeName As String

    activeName = ActiveSheet.Name

    Dim ws As Worksheet

    Application.ScreenUpdating = False

    For Each ws In wb.Worksheets

        If ws.Visible = xlSheetVisible Then

            ws.Activate

            ActiveWindow.Zoom = 100

        End If

    Next

    wb.Worksheets(activeName).Activate

    Application.ScreenUpdating = True

    MsgBox "全シートの倍率を100%にしました。", vbInformation, "DelaxTools"

    Exit Sub

EH:

    Application.ScreenUpdating = True

    MsgBox "全シート倍率100%でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "DelaxTools"

End Sub



Public Sub DxaHalfToFull(ByVal control As Object)

    ConvertSelectionAscii True

End Sub



Public Sub DxaFullToHalf(ByVal control As Object)

    ConvertSelectionAscii False

End Sub



Public Sub DxaAutoFitActiveSheetColumns(ByVal control As Object)

    On Error GoTo EH

    Dim ws As Worksheet

    Set ws = ActiveSheet

    If ws Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then

        ws.Columns.AutoFit

    Else

        ws.UsedRange.Columns.AutoFit

    End If

    Application.ScreenUpdating = True

    MsgBox "実行シートの列幅を自動調整しました。", vbInformation, "DelaxTools"

    Exit Sub

EH:

    Application.ScreenUpdating = True

    MsgBox "列幅自動調整でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "DelaxTools"

End Sub



Public Sub DxaAutoFitActiveSheetRows(ByVal control As Object)

    On Error GoTo EH

    Dim ws As Worksheet

    Set ws = ActiveSheet

    If ws Is Nothing Then Exit Sub



    Application.ScreenUpdating = False

    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then

        ws.Rows.AutoFit

    Else

        ws.UsedRange.Rows.AutoFit

    End If

    Application.ScreenUpdating = True

    MsgBox "実行シートの行高さを自動調整しました。", vbInformation, "DelaxTools"

    Exit Sub

EH:

    Application.ScreenUpdating = True

    MsgBox "行高さ自動調整でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "DelaxTools"

End Sub



Private Sub ConvertSelectionAscii(ByVal halfToFull As Boolean)

    On Error GoTo EH

    If TypeName(Selection) = "Nothing" Then Exit Sub

    Application.ScreenUpdating = False



    Dim rng As Range, c As Range

    On Error Resume Next

    Set rng = Selection.SpecialCells(xlCellTypeConstants)

    On Error GoTo EH

    If Not rng Is Nothing Then

        For Each c In rng.Cells

            If Not IsError(c.Value) Then

                c.Value = ConvertAsciiText(CStr(c.Value), halfToFull)

            End If

        Next

    End If



    Dim sr As ShapeRange, shp As Shape

    On Error Resume Next

    Set sr = Selection.ShapeRange

    On Error GoTo EH

    If Not sr Is Nothing Then

        For Each shp In sr

            ConvertShapeText shp, halfToFull

        Next

    End If



    Application.ScreenUpdating = True

    If halfToFull Then

        MsgBox "選択範囲の半角英数字を全角に変換しました。", vbInformation, "DelaxTools"

    Else

        MsgBox "選択範囲の全角英数字を半角に変換しました。", vbInformation, "DelaxTools"

    End If

    Exit Sub

EH:

    Application.ScreenUpdating = True

    MsgBox "文字変換でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "DelaxTools"

End Sub



Private Sub ConvertShapeText(ByVal shp As Shape, ByVal halfToFull As Boolean)

    On Error Resume Next

    If shp.TextFrame2.HasText Then shp.TextFrame2.TextRange.Text = ConvertAsciiText(shp.TextFrame2.TextRange.Text, halfToFull)

    If shp.TextFrame.Characters.Count > 0 Then shp.TextFrame.Characters.Text = ConvertAsciiText(shp.TextFrame.Characters.Text, halfToFull)

End Sub



Private Function ConvertAsciiText(ByVal s As String, ByVal halfToFull As Boolean) As String

    Dim i As Long, code As Long, ch As String, out As String

    For i = 1 To Len(s)

        ch = Mid$(s, i, 1)

        code = AscW(ch)

        If halfToFull Then

            If (code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Then

                out = out & ChrW$(code + &HFEE0)

            Else

                out = out & ch

            End If

        Else

            If (code >= &HFF10 And code <= &HFF19) Or (code >= &HFF21 And code <= &HFF3A) Or (code >= &HFF41 And code <= &HFF5A) Then

                out = out & ChrW$(code - &HFEE0)

            Else

                out = out & ch

            End If

        End If

    Next

    ConvertAsciiText = out

End Function



Private Sub AddJapaneseHolidays(ByVal y As Long, ByVal d As Object)

    AddHoliday d, DateSerial(y, 1, 1), "元日"

    AddHoliday d, NthMonday(y, 1, 2), "成人の日"

    AddHoliday d, DateSerial(y, 2, 11), "建国記念の日"

    If y >= 2020 Then AddHoliday d, DateSerial(y, 2, 23), "天皇誕生日"

    AddHoliday d, VernalEquinox(y), "春分の日"

    AddHoliday d, DateSerial(y, 4, 29), "昭和の日"

    AddHoliday d, DateSerial(y, 5, 3), "憲法記念日"

    AddHoliday d, DateSerial(y, 5, 4), "みどりの日"

    AddHoliday d, DateSerial(y, 5, 5), "こどもの日"

    AddHoliday d, NthMonday(y, 7, 3), "海の日"

    AddHoliday d, DateSerial(y, 8, 11), "山の日"

    AddHoliday d, NthMonday(y, 9, 3), "敬老の日"

    AddHoliday d, AutumnalEquinox(y), "秋分の日"

    AddHoliday d, NthMonday(y, 10, 2), "スポーツの日"

    AddHoliday d, DateSerial(y, 11, 3), "文化の日"

    AddHoliday d, DateSerial(y, 11, 23), "勤労感謝の日"

    AddSubstituteHolidays y, d

    AddCitizensHolidays y, d

End Sub



Private Sub AddHoliday(ByVal d As Object, ByVal dt As Date, ByVal name As String)

    Dim key As String

    key = Format$(dt, "yyyy/mm/dd")

    If d.Exists(key) Then

        If InStr(1, d(key), name, vbTextCompare) = 0 Then d(key) = d(key) & " / " & name

    Else

        d.Add key, name

    End If

End Sub



Private Sub AddSubstituteHolidays(ByVal y As Long, ByVal d As Object)

    Dim keys As Variant, k As Variant, dt As Date, subDt As Date

    keys = d.Keys

    For Each k In keys

        dt = CDate(k)

        If Year(dt) = y And Weekday(dt, vbSunday) = vbSunday Then

            subDt = DateAdd("d", 1, dt)

            Do While d.Exists(Format$(subDt, "yyyy/mm/dd"))

                subDt = DateAdd("d", 1, subDt)

            Loop

            If Year(subDt) = y Then AddHoliday d, subDt, "振替休日"

        End If

    Next

End Sub



Private Sub AddCitizensHolidays(ByVal y As Long, ByVal d As Object)

    Dim dt As Date

    For dt = DateSerial(y, 1, 2) To DateSerial(y, 12, 30)

        If Not d.Exists(Format$(dt, "yyyy/mm/dd")) Then

            If d.Exists(Format$(DateAdd("d", -1, dt), "yyyy/mm/dd")) And d.Exists(Format$(DateAdd("d", 1, dt), "yyyy/mm/dd")) Then

                If Weekday(dt, vbSunday) <> vbSunday Then AddHoliday d, dt, "国民の休日"

            End If

        End If

    Next

End Sub



Private Function NthMonday(ByVal y As Long, ByVal m As Long, ByVal n As Long) As Date

    Dim dt As Date

    dt = DateSerial(y, m, 1)

    Do While Weekday(dt, vbMonday) <> 1

        dt = DateAdd("d", 1, dt)

    Loop

    NthMonday = DateAdd("d", (n - 1) * 7, dt)

End Function



Private Function VernalEquinox(ByVal y As Long) As Date

    Dim dayNum As Long

    If y <= 2099 Then

        dayNum = Int(20.8431 + 0.242194 * (y - 1980) - Int((y - 1980) / 4))

    Else

        dayNum = Int(21.851 + 0.242194 * (y - 1980) - Int((y - 1980) / 4))

    End If

    VernalEquinox = DateSerial(y, 3, dayNum)

End Function



Private Function AutumnalEquinox(ByVal y As Long) As Date

    Dim dayNum As Long

    If y <= 2099 Then

        dayNum = Int(23.2488 + 0.242194 * (y - 1980) - Int((y - 1980) / 4))

    Else

        dayNum = Int(24.2488 + 0.242194 * (y - 1980) - Int((y - 1980) / 4))

    End If

    AutumnalEquinox = DateSerial(y, 9, dayNum)

End Function



Private Function JapaneseWeekday(ByVal dt As Date) As String

    JapaneseWeekday = Choose(Weekday(dt, vbSunday), "日", "月", "火", "水", "木", "金", "土")

End Function



Private Sub SortDateKeys(ByRef keys As Variant)

    Dim i As Long, j As Long, tmp As Variant

    For i = LBound(keys) To UBound(keys) - 1

        For j = i + 1 To UBound(keys)

            If CDate(keys(i)) > CDate(keys(j)) Then

                tmp = keys(i)

                keys(i) = keys(j)

                keys(j) = tmp

            End If

        Next

    Next

End Sub





'============================================================

' v95 追加機能

'============================================================

Public Sub DxaCreateSheetIndex(ByVal control As Object)

    On Error GoTo EH

    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub

    Dim indexName As String
    indexName = "シート一覧"

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim wsIndex As Worksheet
    On Error Resume Next
    Set wsIndex = wb.Worksheets(indexName)
    On Error GoTo EH
    If Not wsIndex Is Nothing Then wsIndex.Delete

    Application.DisplayAlerts = True

    Set wsIndex = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    wsIndex.Name = indexName

    wsIndex.Range("A1:C1").Value = Array("No", "シート名", "表示状態")
    wsIndex.Range("A1:C1").Font.Bold = True
    wsIndex.Range("A1:C1").Interior.Color = RGB(221, 235, 247)

    Dim r As Long
    r = 2

    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Name <> wsIndex.Name Then
            wsIndex.Cells(r, 1).Value = r - 1
            wsIndex.Cells(r, 2).Value = ws.Name
            wsIndex.Hyperlinks.Add Anchor:=wsIndex.Cells(r, 2), Address:="", SubAddress:="'" & Replace(ws.Name, "'", "''") & "'!A1", TextToDisplay:=ws.Name
            Select Case ws.Visible
                Case xlSheetVisible
                    wsIndex.Cells(r, 3).Value = "表示"
                Case xlSheetHidden
                    wsIndex.Cells(r, 3).Value = "非表示"
                Case xlSheetVeryHidden
                    wsIndex.Cells(r, 3).Value = "VeryHidden"
                Case Else
                    wsIndex.Cells(r, 3).Value = CStr(ws.Visible)
            End Select
            r = r + 1
        End If
    Next ws

    wsIndex.Columns("A:C").AutoFit
    wsIndex.Range("A1:C1").AutoFilter
    wsIndex.Activate
    wsIndex.Range("A1").Select

    Application.ScreenUpdating = True
    MsgBox "シート一覧を作成しました。", vbInformation, "DelaxTools"
    Exit Sub
EH:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "シート一覧でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "DelaxTools"
End Sub

Public Sub DxaBacklogGroupByParent(ByVal control As Object)
    DxaBacklogGroupByParentCore
End Sub

Private Sub DxaBacklogGroupByParentCore()
    On Error GoTo EH

    Dim wb As Workbook
    Dim wsTarget As Worksheet
    Dim wsParent As Worksheet
    Dim parentDict As Object
    Dim lastRowTarget As Long
    Dim lastRowParent As Long
    Dim rowIndex As Long
    Dim cellValue As String
    Dim parentSource As String

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "対象ブックを開いてから実行してください。", vbExclamation, "DelaxTools"
        Exit Sub
    End If

    Set wsTarget = ActiveSheet
    If wsTarget Is Nothing Then
        MsgBox "対象シートを選択してから実行してください。", vbExclamation, "DelaxTools"
        Exit Sub
    End If

    Set parentDict = CreateObject("Scripting.Dictionary")
    Set wsParent = DxaFindWorksheetInWorkbook(wb, "親課題一覧")

    If Not wsParent Is Nothing Then
        lastRowParent = wsParent.Cells(wsParent.Rows.Count, "A").End(xlUp).Row
        For rowIndex = 1 To lastRowParent
            cellValue = DxaBacklogIssueKeyText(wsParent.Cells(rowIndex, "A"))
            If Len(cellValue) > 0 Then parentDict(cellValue) = True
        Next rowIndex
        parentSource = "親課題一覧"
    Else
        DxaCollectBacklogParentCandidates wsTarget, parentDict
        parentSource = "ガントシート自動判定"
    End If

    If parentDict.Count = 0 Then
        MsgBox "親課題が見つかりませんでした。" & vbCrLf & _
               "親課題一覧シートを作成するか、Backlogガント出力シートを選択してから実行してください。", _
               vbExclamation, "DelaxTools"
        Exit Sub
    End If

    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row
    If lastRowTarget < 5 Then
        MsgBox "グループ化対象の行が見つかりませんでした。", vbExclamation, "DelaxTools"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    On Error Resume Next
    wsTarget.Rows.ClearOutline
    On Error GoTo EH

    wsTarget.Range("A5:A" & lastRowTarget).IndentLevel = 0

    rowIndex = 5
    Do While rowIndex <= lastRowTarget
        cellValue = DxaBacklogIssueKeyText(wsTarget.Cells(rowIndex, "A"))

        If parentDict.Exists(cellValue) Then
            rowIndex = DxaGroupOneBacklogParent(wsTarget, parentDict, rowIndex, lastRowTarget)
        Else
            rowIndex = rowIndex + 1
        End If
    Loop

    Application.ScreenUpdating = True
    MsgBox "親課題でグループ化しました。" & vbCrLf & "親課題の判定方法: " & parentSource, vbInformation, "DelaxTools"
    Exit Sub

EH:
    Application.ScreenUpdating = True
    MsgBox "親課題でグループ化中にエラーが発生しました。" & vbCrLf & _
           "対象ブック・対象シート・親課題一覧シートを確認してください。" & vbCrLf & _
           Err.Description, vbExclamation, "DelaxTools"
End Sub

Private Function DxaGroupOneBacklogParent(ByVal wsTarget As Worksheet, ByVal parentDict As Object, ByVal parentRow As Long, ByVal lastRowTarget As Long) As Long
    Dim nextParentRow As Long
    Dim searchRow As Long
    Dim startChild As Long
    Dim endChild As Long
    Dim i As Long
    Dim keyText As String

    nextParentRow = 0

    With wsTarget.Cells(parentRow, "A").Font
        .Bold = True
        .Size = Application.Max(8, .Size + 4)
    End With

    With wsTarget.Cells(parentRow, "C").Font
        .Bold = True
        .Size = Application.Max(8, .Size + 4)
    End With

    For searchRow = parentRow + 1 To lastRowTarget
        keyText = DxaBacklogIssueKeyText(wsTarget.Cells(searchRow, "A"))
        If parentDict.Exists(keyText) Then
            nextParentRow = searchRow
            Exit For
        End If
    Next searchRow

    startChild = parentRow + 1
    If nextParentRow > 0 Then
        endChild = nextParentRow - 1
    Else
        endChild = lastRowTarget
    End If

    If startChild <= endChild Then
        wsTarget.Rows(startChild & ":" & endChild).Group
        For i = startChild To endChild
            wsTarget.Cells(i, "A").IndentLevel = wsTarget.Cells(i, "A").IndentLevel + 1
        Next i
    End If

    If nextParentRow > 0 Then
        DxaGroupOneBacklogParent = nextParentRow
    Else
        DxaGroupOneBacklogParent = lastRowTarget + 1
    End If
End Function

Private Sub DxaCollectBacklogParentCandidates(ByVal ws As Worksheet, ByVal parentDict As Object)
    Dim headerRow As Long
    Dim firstDataRow As Long
    Dim lastRow As Long
    Dim r As Long
    Dim keyText As String

    headerRow = DxaBacklogFindHeaderRow(ws)
    If headerRow = 0 Then headerRow = 4
    firstDataRow = headerRow + 1
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For r = firstDataRow To lastRow
        keyText = DxaBacklogIssueKeyText(ws.Cells(r, "A"))
        If Len(keyText) > 0 Then
            If DxaLooksLikeBacklogParentRow(ws, r, firstDataRow, lastRow) Then
                parentDict(keyText) = True
            End If
        End If
    Next r
End Sub

Private Function DxaLooksLikeBacklogParentRow(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal firstDataRow As Long, ByVal lastRow As Long) As Boolean
    Dim assigneeText As String
    Dim plannedText As String
    Dim actualText As String
    Dim startValue As Variant
    Dim dueValue As Variant
    Dim durationDays As Long

    assigneeText = Trim$(CStr(ws.Cells(rowNo, "G").Value))
    plannedText = Trim$(CStr(ws.Cells(rowNo, "J").Value))
    actualText = Trim$(CStr(ws.Cells(rowNo, "K").Value))
    startValue = ws.Cells(rowNo, "H").Value
    dueValue = ws.Cells(rowNo, "I").Value

    If Len(assigneeText) > 0 Then Exit Function
    If Len(plannedText) > 0 Or Len(actualText) > 0 Then Exit Function

    If IsDate(startValue) And IsDate(dueValue) Then
        durationDays = CLng(CDate(dueValue) - CDate(startValue))
        If durationDays >= 2 Then
            DxaLooksLikeBacklogParentRow = True
            Exit Function
        End If
    End If

    If DxaNextRowsLookLikeChildren(ws, rowNo, lastRow) Then
        DxaLooksLikeBacklogParentRow = True
    End If
End Function

Private Function DxaNextRowsLookLikeChildren(ByVal ws As Worksheet, ByVal parentRow As Long, ByVal lastRow As Long) As Boolean
    Dim r As Long
    Dim checkLimit As Long
    Dim childCount As Long

    checkLimit = parentRow + 5
    If checkLimit > lastRow Then checkLimit = lastRow

    For r = parentRow + 1 To checkLimit
        If Len(DxaBacklogIssueKeyText(ws.Cells(r, "A"))) > 0 Then
            If Len(Trim$(CStr(ws.Cells(r, "G").Value))) > 0 Or Len(Trim$(CStr(ws.Cells(r, "J").Value))) > 0 Or Len(Trim$(CStr(ws.Cells(r, "K").Value))) > 0 Then
                childCount = childCount + 1
            End If
        End If
    Next r

    DxaNextRowsLookLikeChildren = (childCount >= 1)
End Function

Private Function DxaBacklogIssueKeyText(ByVal cell As Range) As String
    On Error Resume Next
    If cell.Hyperlinks.Count > 0 Then
        DxaBacklogIssueKeyText = Trim$(CStr(cell.Hyperlinks(1).TextToDisplay))
    ElseIf cell.HasFormula Then
        DxaBacklogIssueKeyText = Trim$(CStr(cell.Text))
    Else
        DxaBacklogIssueKeyText = Trim$(CStr(cell.Value))
    End If
End Function

Private Function DxaBacklogFindHeaderRow(ByVal ws As Worksheet) As Long
    Dim r As Long
    For r = 1 To 20
        If Trim$(CStr(ws.Cells(r, "A").Value)) = "キー" _
           And Trim$(CStr(ws.Cells(r, "C").Value)) = "件名" Then
            DxaBacklogFindHeaderRow = r
            Exit Function
        End If
    Next r
    DxaBacklogFindHeaderRow = 0
End Function

Private Function DxaFindWorksheetInWorkbook(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set DxaFindWorksheetInWorkbook = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

' フォルダツリー作成：選択したフォルダ配下のフォルダ/ファイル構造を、現在のシートへ文字列ベースで挿入します。
' 参考動作：RelaxAppsの「フォルダー ツリー作成」同様、フォルダを選択して取得し、Excel上にツリーを貼り付けます。
Public Sub DxaCreateFolderTreeWithFolderPicker(ByVal control As Object)
    On Error GoTo ErrHandler

    Dim rootFolder As String
    rootFolder = DxaPickSourceFolder("ツリーを作成するフォルダを選択してください")
    If Len(rootFolder) = 0 Then Exit Sub

    Dim includeFilesAnswer As VbMsgBoxResult
    includeFilesAnswer = MsgBox("ファイルもツリーに含めますか？" & vbCrLf & _
                                "はい: フォルダ＋ファイルを出力" & vbCrLf & _
                                "いいえ: フォルダのみ出力", _
                                vbQuestion + vbYesNoCancel, "フォルダツリー")
    If includeFilesAnswer = vbCancel Then Exit Sub

    Dim includeFiles As Boolean
    includeFiles = (includeFilesAnswer = vbYes)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(rootFolder) Then
        MsgBox "選択したフォルダが見つかりません。" & vbCrLf & rootFolder, vbExclamation, "フォルダツリー"
        Exit Sub
    End If

    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("フォルダツリー")
    On Error GoTo ErrHandler
    If Not ws Is Nothing Then ws.Delete

    Application.DisplayAlerts = True

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "フォルダツリー"

    Dim rowNo As Long
    rowNo = 1

    Dim colTree As Long
    Dim colType As Long
    Dim colPath As Long
    Dim colModified As Long
    Dim colSize As Long
    colTree = 1
    colType = 2
    colPath = 3
    colModified = 4
    colSize = 5

    ws.Cells(rowNo, colTree).Value = "ツリー"
    ws.Cells(rowNo, colType).Value = "種別"
    ws.Cells(rowNo, colPath).Value = "パス"
    ws.Cells(rowNo, colModified).Value = "更新日時"
    ws.Cells(rowNo, colSize).Value = "サイズ(KB)"
    ws.Range(ws.Cells(rowNo, colTree), ws.Cells(rowNo, colSize)).Font.Bold = True
    ws.Range(ws.Cells(rowNo, colTree), ws.Cells(rowNo, colSize)).Interior.Color = RGB(221, 235, 247)
    rowNo = rowNo + 1

    Dim root As Object
    Set root = fso.GetFolder(rootFolder)

    DxaWriteTreeLine ws, rowNo, colTree, colType, colPath, colModified, colSize, _
                     "■ " & root.Name, "フォルダ", root.Path, root.DateLastModified, "", root.Path
    rowNo = rowNo + 1

    Dim folderCount As Long
    Dim fileCount As Long
    folderCount = 1
    fileCount = 0

    DxaOutputFolderTree ws, rowNo, colTree, colType, colPath, colModified, colSize, root, "", includeFiles, folderCount, fileCount

    ws.Range(ws.Cells(1, colTree), ws.Cells(rowNo - 1, colSize)).Columns.AutoFit
    ws.Range("A1:E1").AutoFilter
    ws.Activate
    ws.Range("A1").Select

    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "フォルダツリーを作成しました。" & vbCrLf & _
           "取得元: " & rootFolder & vbCrLf & _
           "出力先: フォルダツリー シート" & vbCrLf & _
           "フォルダ: " & folderCount & " 件" & vbCrLf & _
           "ファイル: " & fileCount & " 件", vbInformation, "フォルダツリー"
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "フォルダツリー作成でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "フォルダツリー"
End Sub

Private Sub DxaOutputFolderTree(ByVal ws As Worksheet, ByRef rowNo As Long, _
                                ByVal colTree As Long, ByVal colType As Long, ByVal colPath As Long, ByVal colModified As Long, ByVal colSize As Long, _
                                ByVal folderObj As Object, ByVal prefix As String, ByVal includeFiles As Boolean, _
                                ByRef folderCount As Long, ByRef fileCount As Long)
    On Error GoTo AccessDenied

    Dim folderPaths As Variant
    Dim filePaths As Variant
    folderPaths = DxaSortedFolderPaths(folderObj)
    If includeFiles Then filePaths = DxaSortedFilePaths(folderObj)

    Dim folderTotal As Long
    Dim fileTotal As Long
    folderTotal = DxaVariantItemCount(folderPaths)
    If includeFiles Then fileTotal = DxaVariantItemCount(filePaths) Else fileTotal = 0

    Dim i As Long
    For i = 1 To folderTotal
        Dim fPath As String
        fPath = CStr(folderPaths(i))

        Dim childFolder As Object
        Set childFolder = CreateObject("Scripting.FileSystemObject").GetFolder(fPath)

        Dim isLastFolderItem As Boolean
        isLastFolderItem = (i = folderTotal And fileTotal = 0)

        Dim branch As String
        If isLastFolderItem Then branch = "└─ " Else branch = "├─ "

        DxaWriteTreeLine ws, rowNo, colTree, colType, colPath, colModified, colSize, _
                         prefix & branch & "□ " & childFolder.Name, "フォルダ", childFolder.Path, childFolder.DateLastModified, "", childFolder.Path
        rowNo = rowNo + 1
        folderCount = folderCount + 1

        Dim nextPrefix As String
        If isLastFolderItem Then nextPrefix = prefix & "    " Else nextPrefix = prefix & "│  "
        DxaOutputFolderTree ws, rowNo, colTree, colType, colPath, colModified, colSize, childFolder, nextPrefix, includeFiles, folderCount, fileCount
    Next i

    If includeFiles Then
        For i = 1 To fileTotal
            Dim filePath As String
            filePath = CStr(filePaths(i))

            Dim fileObj As Object
            Set fileObj = CreateObject("Scripting.FileSystemObject").GetFile(filePath)

            Dim fileBranch As String
            If i = fileTotal Then fileBranch = "└─ " Else fileBranch = "├─ "

            DxaWriteTreeLine ws, rowNo, colTree, colType, colPath, colModified, colSize, _
                             prefix & fileBranch & "・ " & fileObj.Name, "ファイル", fileObj.Path, fileObj.DateLastModified, DxaFormatKb(fileObj.Size), fileObj.Path
            rowNo = rowNo + 1
            fileCount = fileCount + 1
        Next i
    End If
    Exit Sub

AccessDenied:
    DxaWriteTreeLine ws, rowNo, colTree, colType, colPath, colModified, colSize, _
                     prefix & "└─ [アクセス不可] " & folderObj.Name, "エラー", folderObj.Path, "", "", folderObj.Path
    rowNo = rowNo + 1
End Sub

Private Sub DxaWriteTreeLine(ByVal ws As Worksheet, ByVal rowNo As Long, _
                             ByVal colTree As Long, ByVal colType As Long, ByVal colPath As Long, ByVal colModified As Long, ByVal colSize As Long, _
                             ByVal treeText As String, ByVal typeText As String, ByVal pathText As String, ByVal modifiedValue As Variant, ByVal sizeText As String, ByVal linkPath As String)
    ws.Cells(rowNo, colTree).Value = treeText
    ws.Cells(rowNo, colType).Value = typeText
    ws.Cells(rowNo, colPath).Value = pathText
    If Len(CStr(modifiedValue)) > 0 Then ws.Cells(rowNo, colModified).Value = modifiedValue
    ws.Cells(rowNo, colSize).Value = sizeText

    On Error Resume Next
    ws.Hyperlinks.Add Anchor:=ws.Cells(rowNo, colTree), Address:=linkPath, TextToDisplay:=treeText
    On Error GoTo 0
End Sub

Private Function DxaSortedFolderPaths(ByVal folderObj As Object) As Variant
    Dim col As Collection
    Set col = New Collection

    Dim f As Object
    For Each f In folderObj.SubFolders
        col.Add CStr(f.Path)
    Next f

    DxaSortedFolderPaths = DxaCollectionToSortedArray(col)
End Function

Private Function DxaSortedFilePaths(ByVal folderObj As Object) As Variant
    Dim col As Collection
    Set col = New Collection

    Dim f As Object
    For Each f In folderObj.Files
        col.Add CStr(f.Path)
    Next f

    DxaSortedFilePaths = DxaCollectionToSortedArray(col)
End Function

Private Function DxaCollectionToSortedArray(ByVal col As Collection) As Variant
    If col.Count = 0 Then
        DxaCollectionToSortedArray = Empty
        Exit Function
    End If

    Dim arr() As String
    ReDim arr(1 To col.Count)

    Dim i As Long
    For i = 1 To col.Count
        arr(i) = CStr(col(i))
    Next i

    DxaSortStringArray arr
    DxaCollectionToSortedArray = arr
End Function

Private Sub DxaSortStringArray(ByRef arr() As String)
    Dim i As Long
    Dim j As Long
    Dim tmp As String

    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If StrComp(arr(i), arr(j), vbTextCompare) > 0 Then
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next j
    Next i
End Sub

Private Function DxaVariantItemCount(ByVal values As Variant) As Long
    On Error GoTo EmptyValue
    If IsEmpty(values) Then
        DxaVariantItemCount = 0
    Else
        DxaVariantItemCount = UBound(values) - LBound(values) + 1
    End If
    Exit Function
EmptyValue:
    DxaVariantItemCount = 0
End Function

Private Function DxaFormatKb(ByVal bytes As Currency) As String
    DxaFormatKb = Format$(CDbl(bytes) / 1024#, "#,##0.0")
End Function

Private Function DxaPickSourceFolder(ByVal titleText As String) As String
    On Error GoTo Fallback
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = titleText
        .AllowMultiSelect = False
        .InitialFileName = DxaDefaultFolder() & Application.PathSeparator
        If .Show <> -1 Then
            DxaPickSourceFolder = ""
        Else
            DxaPickSourceFolder = .SelectedItems(1)
        End If
    End With
    Exit Function
Fallback:
    DxaPickSourceFolder = InputBox(titleText & vbCrLf & "フォルダのパスを入力してください。", "フォルダ選択", DxaDefaultFolder())
End Function

Public Sub DxaCreateFileList(ByVal control As Object)
    On Error GoTo ErrHandler

    Dim rootFolder As String
    rootFolder = DxaPickSourceFolder("ファイル一覧を作成するフォルダを選択してください")
    If Len(rootFolder) = 0 Then Exit Sub

    Dim includeSubFoldersAnswer As VbMsgBoxResult
    includeSubFoldersAnswer = MsgBox("サブフォルダ内のファイルも一覧に含めますか？" & vbCrLf & _
                                     "はい: サブフォルダを含める" & vbCrLf & _
                                     "いいえ: 選択フォルダ直下のみ", _
                                     vbQuestion + vbYesNoCancel, "ファイル一覧")
    If includeSubFoldersAnswer = vbCancel Then Exit Sub

    Dim includeSubFolders As Boolean
    includeSubFolders = (includeSubFoldersAnswer = vbYes)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(rootFolder) Then
        MsgBox "選択したフォルダが見つかりません。" & vbCrLf & rootFolder, vbExclamation, "ファイル一覧"
        Exit Sub
    End If

    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("ファイル一覧")
    On Error GoTo ErrHandler
    If Not ws Is Nothing Then ws.Delete

    Application.DisplayAlerts = True
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "ファイル一覧"

    ws.Range("A1:H1").Value = Array("No", "ファイル名", "拡張子", "フォルダ", "フルパス", "サイズ(KB)", "更新日時", "作成日時")
    ws.Range("A1:H1").Font.Bold = True
    ws.Range("A1:H1").Interior.Color = RGB(221, 235, 247)

    Dim rowNo As Long
    rowNo = 2

    Dim fileCount As Long
    fileCount = 0

    Dim root As Object
    Set root = fso.GetFolder(rootFolder)
    DxaOutputFileList ws, rowNo, root, includeSubFolders, fileCount

    ws.Columns("A:H").AutoFit
    ws.Range("A1:H1").AutoFilter
    ws.Activate
    ws.Range("A1").Select

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "ファイル一覧を作成しました。" & vbCrLf & _
           "対象フォルダ: " & rootFolder & vbCrLf & _
           "ファイル数: " & fileCount & " 件", vbInformation, "ファイル一覧"
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "ファイル一覧でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "ファイル一覧"
End Sub

Private Sub DxaOutputFileList(ByVal ws As Worksheet, ByRef rowNo As Long, ByVal folderObj As Object, _
                              ByVal includeSubFolders As Boolean, ByRef fileCount As Long)
    On Error GoTo AccessDenied

    Dim filePaths As Variant
    filePaths = DxaSortedFilePaths(folderObj)

    Dim i As Long
    Dim fileObj As Object
    For i = 1 To DxaVariantItemCount(filePaths)
        Set fileObj = CreateObject("Scripting.FileSystemObject").GetFile(CStr(filePaths(i)))
        fileCount = fileCount + 1
        DxaWriteFileListLine ws, rowNo, fileCount, fileObj
        rowNo = rowNo + 1
    Next i

    If includeSubFolders Then
        Dim folderPaths As Variant
        folderPaths = DxaSortedFolderPaths(folderObj)
        For i = 1 To DxaVariantItemCount(folderPaths)
            DxaOutputFileList ws, rowNo, CreateObject("Scripting.FileSystemObject").GetFolder(CStr(folderPaths(i))), includeSubFolders, fileCount
        Next i
    End If

    Exit Sub
AccessDenied:
    ' 権限がないフォルダや一時的に参照できないファイルは処理を継続します。
    Err.Clear
End Sub

Private Sub DxaWriteFileListLine(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal no As Long, ByVal fileObj As Object)
    On Error Resume Next

    ws.Cells(rowNo, 1).Value = no
    ws.Cells(rowNo, 2).Value = fileObj.Name
    ws.Hyperlinks.Add Anchor:=ws.Cells(rowNo, 2), Address:=fileObj.Path, TextToDisplay:=fileObj.Name
    ws.Cells(rowNo, 3).Value = DxaFileExtension(fileObj.Name)
    ws.Cells(rowNo, 4).Value = fileObj.ParentFolder.Path
    ws.Cells(rowNo, 5).Value = fileObj.Path
    ws.Hyperlinks.Add Anchor:=ws.Cells(rowNo, 5), Address:=fileObj.Path, TextToDisplay:=fileObj.Path
    ws.Cells(rowNo, 6).Value = Round(CDbl(fileObj.Size) / 1024, 1)
    ws.Cells(rowNo, 7).Value = fileObj.DateLastModified
    ws.Cells(rowNo, 8).Value = fileObj.DateCreated
End Sub

Private Function DxaFileExtension(ByVal fileName As String) As String
    Dim p As Long
    p = InStrRev(fileName, ".")
    If p > 0 And p < Len(fileName) Then
        DxaFileExtension = Mid$(fileName, p + 1)
    Else
        DxaFileExtension = ""
    End If
End Function

Public Sub DxaExportVbaWithFolderPicker(ByVal control As Object)

    On Error GoTo EH



    Dim wb As Workbook

    Set wb = ActiveWorkbook

    If wb Is Nothing Then Exit Sub



    Dim selectedFolder As String

    selectedFolder = DxaPickOutputFolder("VBAエクスポート先フォルダを選択してください")

    If Len(selectedFolder) = 0 Then Exit Sub



    Dim exportFolder As String

    exportFolder = selectedFolder & Application.PathSeparator & "VBAExport_" & DxaSafeFileName(DxaWorkbookBaseName(wb.Name)) & "_" & Format(Now, "yyyymmdd_hhnnss")

    If Dir(exportFolder, vbDirectory) = "" Then MkDir exportFolder



    Dim vbProj As Object

    Set vbProj = wb.VBProject



    Dim comp As Object

    Dim ext As String

    Dim exportPath As String

    Dim count As Long



    For Each comp In vbProj.VBComponents

        ext = DxaVbComponentExtension(CLng(comp.Type))

        exportPath = exportFolder & Application.PathSeparator & DxaSafeFileName(CStr(comp.Name)) & ext

        comp.Export exportPath

        count = count + 1

    Next comp



    MsgBox "VBAソースをエクスポートしました。" & vbCrLf & _

           "出力先: " & exportFolder & vbCrLf & _

           "出力数: " & CStr(count), vbInformation, "DelaxTools"

    Exit Sub

EH:

    MsgBox "VBAエクスポートでエラーが発生しました。" & vbCrLf & _

           "Excelの『VBAプロジェクト オブジェクト モデルへのアクセスを信頼する』が必要です。" & vbCrLf & _

           Err.Description, vbExclamation, "DelaxTools"

End Sub



Private Function DxaPickOutputFolder(ByVal titleText As String) As String

    On Error GoTo Fallback

    Dim fd As FileDialog

    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    With fd

        .Title = titleText

        .AllowMultiSelect = False

        .InitialFileName = DxaDefaultFolder() & Application.PathSeparator

        If .Show <> -1 Then

            DxaPickOutputFolder = ""

        Else

            DxaPickOutputFolder = .SelectedItems(1)

        End If

    End With

    Exit Function

Fallback:

    DxaPickOutputFolder = InputBox(titleText & vbCrLf & "フォルダのパスを入力してください。", "フォルダ選択", DxaDefaultFolder())

End Function



Private Function DxaDefaultFolder() As String

    On Error Resume Next

    If Len(ActiveWorkbook.Path) > 0 Then

        DxaDefaultFolder = ActiveWorkbook.Path

    Else

        DxaDefaultFolder = CreateObject("WScript.Shell").SpecialFolders("Desktop")

    End If

    If Len(DxaDefaultFolder) = 0 Then DxaDefaultFolder = CurDir$

End Function



Private Function DxaVbComponentExtension(ByVal componentType As Long) As String

    Select Case componentType

        Case 1

            DxaVbComponentExtension = ".bas"

        Case 2

            DxaVbComponentExtension = ".cls"

        Case 3

            DxaVbComponentExtension = ".frm"

        Case 100

            DxaVbComponentExtension = ".cls"

        Case Else

            DxaVbComponentExtension = ".txt"

    End Select

End Function



Private Function DxaWorkbookBaseName(ByVal fileName As String) As String

    Dim p As Long

    p = InStrRev(fileName, ".")

    If p > 1 Then

        DxaWorkbookBaseName = Left$(fileName, p - 1)

    Else

        DxaWorkbookBaseName = fileName

    End If

End Function



Private Function DxaSafeFileName(ByVal text As String) As String

    Dim s As String

    s = text

    s = Replace(s, "\", "_")

    s = Replace(s, "/", "_")

    s = Replace(s, ":", "_")

    s = Replace(s, "*", "_")

    s = Replace(s, "?", "_")

    s = Replace(s, """", "_")

    s = Replace(s, "<", "_")

    s = Replace(s, ">", "_")

    s = Replace(s, "|", "_")

    If Len(Trim$(s)) = 0 Then s = "Export"

    DxaSafeFileName = s

End Function



' ============================================================
' 変更履歴作成支援
' - 元ブックにはシートを追加しません。
' - 変更前状態は外部一時ファイルへ自動保存します。
' - 変更履歴作成時だけ、その一時ファイルを読み込んで比較します。
' - 対象ブックを閉じたとき、またはExcel終了時に一時ファイルを削除します。
' ============================================================

Public Sub Auto_Open()
    On Error Resume Next
    DxaInitChangeHistoryEvents
End Sub

Public Sub Auto_Close()
    On Error Resume Next
    DxaDeleteCurrentSessionSnapshots
End Sub

Public Sub DxaInitChangeHistoryEvents()
    On Error Resume Next
    If Len(gDxaSessionId) = 0 Then gDxaSessionId = Format$(Now, "yyyymmddhhnnss") & "_" & CStr(Int(Rnd() * 1000000))
    If gDxaEvents Is Nothing Then
        Set gDxaEvents = New DelaxToolsAppEvents
        gDxaEvents.Init Application
    End If

    DxaCleanupOldChangeSnapshots

    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If Not DxaIsWorkbookExcluded(wb) Then
            DxaEnsureSnapshotForWorkbook wb
        End If
    Next
End Sub

Public Sub DxaEnsureSnapshotForWorkbook(ByVal wb As Workbook)
    On Error GoTo EH
    If wb Is Nothing Then Exit Sub
    If DxaIsWorkbookExcluded(wb) Then Exit Sub
    If Len(gDxaSessionId) = 0 Then DxaInitChangeHistoryEvents

    Dim path As String
    path = DxaSnapshotPathForWorkbook(wb)
    If Len(path) = 0 Then Exit Sub
    If DxaFileExists(path) Then Exit Sub

    Dim text As String
    text = DxaBuildSnapshotText(wb)
    DxaWriteTextUtf8 path, text
    Exit Sub
EH:
End Sub

Public Sub DxaDeleteSnapshotForWorkbook(ByVal wb As Workbook)
    On Error Resume Next
    If wb Is Nothing Then Exit Sub
    If Len(gDxaSessionId) = 0 Then Exit Sub
    Dim path As String
    path = DxaSnapshotPathForWorkbook(wb)
    If Len(path) > 0 Then
        If DxaFileExists(path) Then Kill path
    End If
End Sub

Public Sub DxaCreateChangeHistory(ByVal control As Object)
    On Error GoTo EH
    DxaInitChangeHistoryEvents

    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    If DxaIsWorkbookExcluded(wb) Then
        MsgBox "変更履歴作成の対象ブックを開いてから実行してください。", vbExclamation, "DelaxTools"
        Exit Sub
    End If

    Dim snapshotPath As String
    snapshotPath = DxaSnapshotPathForWorkbook(wb)
    If Len(snapshotPath) = 0 Or Not DxaFileExists(snapshotPath) Then
        DxaEnsureSnapshotForWorkbook wb
        MsgBox "変更前状態が未作成だったため、現在の状態を自動保存しました。編集後に再度『変更履歴作成』を実行してください。" & vbCrLf & vbCrLf & _
               "※元ブックにはシートを追加していません。", vbInformation, "DelaxTools"
        Exit Sub
    End If

    Dim oldMap As Object
    Set oldMap = DxaReadSnapshotMap(snapshotPath)

    Dim curMap As Object
    Set curMap = DxaBuildSnapshotMap(wb)

    Dim details As Collection
    Set details = DxaCompareSnapshotMaps(oldMap, curMap)

    If details.Count = 0 Then
        MsgBox "変更は検出されませんでした。", vbInformation, "DelaxTools"
        Exit Sub
    End If

    DxaOutputChangeHistoryWorkbook wb, details
    MsgBox "変更履歴貼付用ブックを作成しました。" & vbCrLf & _
           "元ブックにはシートを追加していません。", vbInformation, "DelaxTools"
    Exit Sub
EH:
    MsgBox "変更履歴作成でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "DelaxTools"
End Sub

Private Function DxaIsWorkbookExcluded(ByVal wb As Workbook) As Boolean
    On Error Resume Next
    If wb Is Nothing Then DxaIsWorkbookExcluded = True: Exit Function
    If wb.IsAddin Then DxaIsWorkbookExcluded = True: Exit Function
    If LCase$(wb.Name) = "dexcelassist.xlam" Then DxaIsWorkbookExcluded = True: Exit Function
    If LCase$(wb.Name) Like "変更履歴出力_*" Then DxaIsWorkbookExcluded = True: Exit Function
    If LCase$(wb.Name) Like "book*" And wb.Path = "" Then
        ' 新規ブックも対象にはできますが、誤検知を避けるため開いた直後の空ブックは除外します。
        If wb.Worksheets.Count = 1 And Application.WorksheetFunction.CountA(wb.Worksheets(1).Cells) = 0 Then
            DxaIsWorkbookExcluded = True
            Exit Function
        End If
    End If
End Function

Private Function DxaChangeSnapshotDir() As String
    Dim p As String
    p = Environ$("APPDATA") & "\DelaxTools\ChangeSnapshots"
    DxaEnsureFolder p
    DxaChangeSnapshotDir = p
End Function

Private Function DxaSnapshotPathForWorkbook(ByVal wb As Workbook) As String
    If wb Is Nothing Then Exit Function
    If Len(gDxaSessionId) = 0 Then gDxaSessionId = Format$(Now, "yyyymmddhhnnss") & "_" & CStr(Int(Rnd() * 1000000))

    Dim keyText As String
    If Len(wb.FullName) > 0 Then
        keyText = wb.FullName
    Else
        keyText = wb.Name
    End If

    DxaSnapshotPathForWorkbook = DxaChangeSnapshotDir() & "\" & gDxaSessionId & "_" & DxaSafeFileName(DxaWorkbookBaseName(wb.Name)) & "_" & DxaSimpleHash(keyText) & ".tsv"
End Function

Private Function DxaBuildSnapshotText(ByVal wb As Workbook) As String
    Dim m As Object
    Set m = DxaBuildSnapshotMap(wb)

    Dim sb As String
    sb = "Key" & vbTab & "Sheet" & vbTab & "Address" & vbTab & "Row" & vbTab & "Column" & vbTab & "Header" & vbTab & "Item" & vbTab & "Formula" & vbTab & "Value" & vbTab & "Text" & vbTab & "Link" & vbTab & "Comment" & vbCrLf

    Dim k As Variant
    For Each k In m.Keys
        sb = sb & CStr(k) & vbTab & m(k) & vbCrLf
    Next

    DxaBuildSnapshotText = sb
End Function

Private Function DxaBuildSnapshotMap(ByVal wb As Workbook) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            DxaAddSheetSnapshot dict, ws
        End If
    Next

    Set DxaBuildSnapshotMap = dict
End Function

Private Sub DxaAddSheetSnapshot(ByVal dict As Object, ByVal ws As Worksheet)
    On Error GoTo EH
    Dim ur As Range
    Set ur = ws.UsedRange
    If ur Is Nothing Then Exit Sub

    Dim r1 As Long, c1 As Long, r2 As Long, c2 As Long
    r1 = ur.Row
    c1 = ur.Column
    r2 = ur.Row + ur.Rows.Count - 1
    c2 = ur.Column + ur.Columns.Count - 1

    Dim r As Long, c As Long
    Dim cell As Range
    For r = r1 To r2
        For c = c1 To c2
            Set cell = ws.Cells(r, c)
            If DxaCellHasSnapshotValue(cell) Then
                Dim key As String
                key = ws.Name & "!" & cell.Address(False, False)

                Dim headerText As String
                headerText = DxaGetHeaderText(ws, c)

                Dim itemText As String
                itemText = DxaGetRowItemText(ws, r)

                dict(key) = DxaJoinSnapshotFields(Array( _
                    DxaEsc(ws.Name), _
                    DxaEsc(cell.Address(False, False)), _
                    CStr(r), _
                    CStr(c), _
                    DxaEsc(headerText), _
                    DxaEsc(itemText), _
                    DxaEsc(DxaCellFormulaText(cell)), _
                    DxaEsc(DxaCellValueText(cell)), _
                    DxaEsc(DxaCellDisplayText(cell)), _
                    DxaEsc(DxaCellLinkText(cell)), _
                    DxaEsc(DxaCellCommentText(cell)) _
                ))
            End If
        Next c
    Next r
    Exit Sub
EH:
End Sub

Private Function DxaCellHasSnapshotValue(ByVal cell As Range) As Boolean
    On Error Resume Next
    If Len(DxaCellFormulaText(cell)) > 0 Then DxaCellHasSnapshotValue = True: Exit Function
    If Len(DxaCellValueText(cell)) > 0 Then DxaCellHasSnapshotValue = True: Exit Function
    If Len(DxaCellLinkText(cell)) > 0 Then DxaCellHasSnapshotValue = True: Exit Function
    If Len(DxaCellCommentText(cell)) > 0 Then DxaCellHasSnapshotValue = True: Exit Function
End Function

Private Function DxaCellFormulaText(ByVal cell As Range) As String
    On Error Resume Next
    If cell.HasFormula Then DxaCellFormulaText = CStr(cell.Formula)
End Function

Private Function DxaCellValueText(ByVal cell As Range) As String
    On Error Resume Next
    If IsError(cell.Value) Then
        DxaCellValueText = cell.Text
    Else
        DxaCellValueText = CStr(cell.Value)
    End If
End Function

Private Function DxaCellDisplayText(ByVal cell As Range) As String
    On Error Resume Next
    DxaCellDisplayText = CStr(cell.Text)
End Function

Private Function DxaCellLinkText(ByVal cell As Range) As String
    On Error Resume Next
    If cell.Hyperlinks.Count > 0 Then
        DxaCellLinkText = cell.Hyperlinks(1).Address
        If Len(cell.Hyperlinks(1).SubAddress) > 0 Then
            DxaCellLinkText = DxaCellLinkText & "#" & cell.Hyperlinks(1).SubAddress
        End If
    End If
End Function

Private Function DxaCellCommentText(ByVal cell As Range) As String
    On Error Resume Next
    If Not cell.Comment Is Nothing Then DxaCellCommentText = cell.Comment.Text
End Function

Private Function DxaGetHeaderText(ByVal ws As Worksheet, ByVal col As Long) As String
    On Error Resume Next
    Dim s As String
    s = Trim$(CStr(ws.Cells(1, col).Text))
    If Len(s) = 0 And col > 1 Then s = Trim$(CStr(ws.Cells(2, col).Text))
    If Len(s) = 0 Then s = DxaColumnLetter(col) & "列"
    DxaGetHeaderText = s
End Function

Private Function DxaGetRowItemText(ByVal ws As Worksheet, ByVal rowNo As Long) As String
    On Error Resume Next
    Dim s As String
    If rowNo > 1 Then
        s = Trim$(CStr(ws.Cells(rowNo, 1).Text))
        If Len(s) = 0 Then s = Trim$(CStr(ws.Cells(rowNo, 2).Text))
    End If
    If Len(s) = 0 Then s = CStr(rowNo) & "行目"
    DxaGetRowItemText = s
End Function

Private Function DxaReadSnapshotMap(ByVal path As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim text As String
    text = DxaReadTextUtf8(path)

    Dim lines As Variant
    lines = Split(text, vbLf)

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim lineText As String
        lineText = Replace(CStr(lines(i)), vbCr, "")
        If i > 0 And Len(lineText) > 0 Then
            Dim parts As Variant
            parts = Split(lineText, vbTab)
            If UBound(parts) >= 11 Then
                dict(parts(0)) = DxaJoinSnapshotFields(Array(parts(1), parts(2), parts(3), parts(4), parts(5), parts(6), parts(7), parts(8), parts(9), parts(10), parts(11)))
            End If
        End If
    Next

    Set DxaReadSnapshotMap = dict
End Function

Private Function DxaCompareSnapshotMaps(ByVal oldMap As Object, ByVal curMap As Object) As Collection
    Dim details As New Collection
    Dim k As Variant

    For Each k In oldMap.Keys
        If Not curMap.Exists(k) Then
            details.Add DxaBuildChangeDetail(CStr(k), "削除", oldMap(k), "")
        ElseIf DxaSnapshotComparableText(oldMap(k)) <> DxaSnapshotComparableText(curMap(k)) Then
            details.Add DxaBuildChangeDetail(CStr(k), DxaDetectChangeType(oldMap(k), curMap(k)), oldMap(k), curMap(k))
        End If
    Next

    For Each k In curMap.Keys
        If Not oldMap.Exists(k) Then
            details.Add DxaBuildChangeDetail(CStr(k), "追加", "", curMap(k))
        End If
    Next

    Set DxaCompareSnapshotMaps = details
End Function

Private Function DxaSnapshotComparableText(ByVal recordText As String) As String
    Dim f As Variant
    f = DxaSplitSnapshotFields(recordText)
    If UBound(f) < 10 Then
        DxaSnapshotComparableText = recordText
    Else
        DxaSnapshotComparableText = f(6) & Chr$(30) & f(7) & Chr$(30) & f(9) & Chr$(30) & f(10)
    End If
End Function

Private Function DxaDetectChangeType(ByVal oldRecord As String, ByVal newRecord As String) As String
    Dim o As Variant, n As Variant
    o = DxaSplitSnapshotFields(oldRecord)
    n = DxaSplitSnapshotFields(newRecord)

    If UBound(o) >= 10 And UBound(n) >= 10 Then
        If o(6) <> n(6) Then DxaDetectChangeType = "数式変更": Exit Function
        If o(7) <> n(7) Then DxaDetectChangeType = "値変更": Exit Function
        If o(9) <> n(9) Then DxaDetectChangeType = "リンク変更": Exit Function
        If o(10) <> n(10) Then DxaDetectChangeType = "コメント変更": Exit Function
    End If
    DxaDetectChangeType = "変更"
End Function

Private Function DxaBuildChangeDetail(ByVal key As String, ByVal changeType As String, ByVal oldRecord As String, ByVal newRecord As String) As Variant
    Dim baseRecord As String
    If Len(newRecord) > 0 Then baseRecord = newRecord Else baseRecord = oldRecord

    Dim b As Variant, o As Variant, n As Variant
    b = DxaSplitSnapshotFields(baseRecord)
    If Len(oldRecord) > 0 Then o = DxaSplitSnapshotFields(oldRecord) Else o = DxaEmptySnapshotFields()
    If Len(newRecord) > 0 Then n = DxaSplitSnapshotFields(newRecord) Else n = DxaEmptySnapshotFields()

    Dim oldValue As String, newValue As String
    oldValue = DxaDisplayValueFromFields(o)
    newValue = DxaDisplayValueFromFields(n)

    DxaBuildChangeDetail = Array( _
        key, _
        changeType, _
        DxaUnesc(CStr(b(0))), _
        DxaUnesc(CStr(b(1))), _
        CLng(Val(b(2))), _
        CLng(Val(b(3))), _
        DxaUnesc(CStr(b(4))), _
        DxaUnesc(CStr(b(5))), _
        oldValue, _
        newValue _
    )
End Function

Private Function DxaEmptySnapshotFields() As Variant
    DxaEmptySnapshotFields = Array("", "", "0", "0", "", "", "", "", "", "", "")
End Function

Private Function DxaDisplayValueFromFields(ByVal f As Variant) As String
    If UBound(f) < 10 Then Exit Function
    If Len(DxaUnesc(CStr(f(6)))) > 0 Then
        DxaDisplayValueFromFields = DxaUnesc(CStr(f(6)))
    ElseIf Len(DxaUnesc(CStr(f(7)))) > 0 Then
        DxaDisplayValueFromFields = DxaUnesc(CStr(f(7)))
    ElseIf Len(DxaUnesc(CStr(f(9)))) > 0 Then
        DxaDisplayValueFromFields = DxaUnesc(CStr(f(9)))
    ElseIf Len(DxaUnesc(CStr(f(10)))) > 0 Then
        DxaDisplayValueFromFields = DxaUnesc(CStr(f(10)))
    End If
End Function

Private Sub DxaOutputChangeHistoryWorkbook(ByVal sourceWb As Workbook, ByVal details As Collection)
    Dim outWb As Workbook
    Set outWb = Workbooks.Add(xlWBATWorksheet)

    Dim wsSummary As Worksheet
    Set wsSummary = outWb.Worksheets(1)
    wsSummary.Name = "変更履歴貼付用"

    Dim wsDetail As Worksheet
    Set wsDetail = outWb.Worksheets.Add(After:=wsSummary)
    wsDetail.Name = "変更詳細"

    DxaWriteChangeDetailSheet sourceWb, wsDetail, details
    DxaWriteChangeSummarySheet wsSummary, details

    wsSummary.Activate
End Sub

Private Sub DxaWriteChangeDetailSheet(ByVal sourceWb As Workbook, ByVal ws As Worksheet, ByVal details As Collection)
    ws.Range("A1:J1").Value = Array("No", "対象シート", "セル", "行", "列見出し", "対象", "変更種別", "変更前", "変更後", "変更内容")
    ws.Range("A1:J1").Font.Bold = True

    Dim i As Long
    For i = 1 To details.Count
        Dim d As Variant
        d = details(i)
        ws.Cells(i + 1, 1).Value = i
        ws.Cells(i + 1, 2).Value = d(2)
        ws.Cells(i + 1, 3).Value = d(3)
        ws.Cells(i + 1, 4).Value = d(4)
        ws.Cells(i + 1, 5).Value = d(6)
        ws.Cells(i + 1, 6).Value = d(7)
        ws.Cells(i + 1, 7).Value = d(1)
        ws.Cells(i + 1, 8).Value = d(8)
        ws.Cells(i + 1, 9).Value = d(9)
        ws.Cells(i + 1, 10).Value = DxaBuildDetailText(d)

        On Error Resume Next
        If Len(sourceWb.FullName) > 0 Then
            ws.Hyperlinks.Add Anchor:=ws.Cells(i + 1, 3), Address:=sourceWb.FullName, SubAddress:="'" & d(2) & "'!" & d(3), TextToDisplay:=d(3)
        End If
        On Error GoTo 0
    Next

    ws.Columns("A:J").AutoFit
    ws.Range("A1:J1").AutoFilter
End Sub

Private Sub DxaWriteChangeSummarySheet(ByVal ws As Worksheet, ByVal details As Collection)
    ws.Range("A1:E1").Value = Array("No", "変更日", "対象シート", "対象", "変更内容")
    ws.Range("A1:E1").Font.Bold = True

    Dim groups As Object
    Set groups = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To details.Count
        Dim d As Variant
        d = details(i)
        Dim gKey As String
        gKey = d(2) & "|" & CStr(d(4))
        If Not groups.Exists(gKey) Then
            groups(gKey) = DxaNewSummaryGroup(d)
        Else
            groups(gKey) = DxaAppendSummaryGroup(groups(gKey), d)
        End If
    Next

    Dim r As Long
    r = 2
    Dim k As Variant
    For Each k In groups.Keys
        Dim g As Variant
        g = Split(CStr(groups(k)), Chr$(31))
        ws.Cells(r, 1).Value = r - 1
        ws.Cells(r, 2).Value = Date
        ws.Cells(r, 3).Value = g(0)
        ws.Cells(r, 4).Value = g(2)
        ws.Cells(r, 5).Value = DxaBuildSummaryText(g)
        r = r + 1
    Next

    ws.Columns("A:E").AutoFit
    ws.Range("A1:E1").AutoFilter
End Sub

Private Function DxaNewSummaryGroup(ByVal d As Variant) As String
    DxaNewSummaryGroup = d(2) & Chr$(31) & CStr(d(4)) & Chr$(31) & d(7) & Chr$(31) & d(6) & Chr$(31) & d(1) & Chr$(31) & d(8) & Chr$(31) & d(9)
End Function

Private Function DxaAppendSummaryGroup(ByVal groupText As String, ByVal d As Variant) As String
    Dim g As Variant
    g = Split(groupText, Chr$(31))
    g(3) = DxaAppendUniqueText(CStr(g(3)), CStr(d(6)))
    g(4) = DxaAppendUniqueText(CStr(g(4)), CStr(d(1)))
    If Len(CStr(g(5))) = 0 Then g(5) = d(8)
    If Len(CStr(g(6))) = 0 Then g(6) = d(9)
    DxaAppendSummaryGroup = Join(g, Chr$(31))
End Function

Private Function DxaBuildSummaryText(ByVal g As Variant) As String
    Dim target As String
    target = CStr(g(2))
    If Len(Trim$(target)) = 0 Then target = CStr(g(1)) & "行目"

    Dim headers As String
    headers = CStr(g(3))

    Dim types As String
    types = CStr(g(4))

    Dim oldSample As String
    oldSample = CStr(g(5))

    Dim newSample As String
    newSample = CStr(g(6))

    If InStr(types, "追加") > 0 And InStr(types, "変更") = 0 And InStr(types, "削除") = 0 Then
        DxaBuildSummaryText = target & "に「" & DxaShortText(newSample) & "」を追加。"
    ElseIf InStr(types, "削除") > 0 And InStr(types, "変更") = 0 And InStr(types, "追加") = 0 Then
        DxaBuildSummaryText = target & "の「" & DxaShortText(oldSample) & "」を削除。"
    ElseIf DxaCountList(headers) = 1 And Len(oldSample) > 0 And Len(newSample) > 0 Then
        DxaBuildSummaryText = target & "の" & headers & "を「" & DxaShortText(oldSample) & "」から「" & DxaShortText(newSample) & "」に変更。"
    Else
        DxaBuildSummaryText = target & "の" & headers & "を変更。"
    End If
End Function

Private Function DxaBuildDetailText(ByVal d As Variant) As String
    Select Case CStr(d(1))
        Case "追加"
            DxaBuildDetailText = d(7) & "の" & d(6) & "に「" & DxaShortText(d(9)) & "」を追加。"
        Case "削除"
            DxaBuildDetailText = d(7) & "の" & d(6) & "から「" & DxaShortText(d(8)) & "」を削除。"
        Case Else
            DxaBuildDetailText = d(7) & "の" & d(6) & "を「" & DxaShortText(d(8)) & "」から「" & DxaShortText(d(9)) & "」に変更。"
    End Select
End Function

Private Function DxaShortText(ByVal text As String) As String
    Dim s As String
    s = Replace(CStr(text), vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Trim$(s)
    If Len(s) > 80 Then s = Left$(s, 80) & "..."
    DxaShortText = s
End Function

Private Function DxaAppendUniqueText(ByVal baseText As String, ByVal addText As String) As String
    If Len(Trim$(addText)) = 0 Then
        DxaAppendUniqueText = baseText
    ElseIf Len(Trim$(baseText)) = 0 Then
        DxaAppendUniqueText = addText
    ElseIf InStr(1, "," & baseText & ",", "," & addText & ",", vbTextCompare) > 0 Then
        DxaAppendUniqueText = baseText
    Else
        DxaAppendUniqueText = baseText & "," & addText
    End If
End Function

Private Function DxaCountList(ByVal text As String) As Long
    If Len(Trim$(text)) = 0 Then Exit Function
    DxaCountList = UBound(Split(text, ",")) + 1
End Function

Private Function DxaJoinSnapshotFields(ByVal values As Variant) As String
    Dim i As Long
    Dim s As String
    For i = LBound(values) To UBound(values)
        If i > LBound(values) Then s = s & vbTab
        s = s & CStr(values(i))
    Next
    DxaJoinSnapshotFields = s
End Function

Private Function DxaSplitSnapshotFields(ByVal recordText As String) As Variant
    DxaSplitSnapshotFields = Split(CStr(recordText), vbTab)
End Function

Private Function DxaEsc(ByVal text As String) As String
    Dim s As String
    s = CStr(text)
    s = Replace(s, "\", "\\")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    s = Replace(s, Chr$(30), "\u001e")
    DxaEsc = s
End Function

Private Function DxaUnesc(ByVal text As String) As String
    Dim s As String
    s = CStr(text)

    Dim i As Long
    Dim ch As String
    Dim nx As String
    Dim result As String

    i = 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If ch = "\" And i < Len(s) Then
            nx = Mid$(s, i + 1, 1)
            Select Case nx
                Case "n"
                    result = result & vbLf
                    i = i + 2
                Case "t"
                    result = result & vbTab
                    i = i + 2
                Case "\"
                    result = result & "\"
                    i = i + 2
                Case "u"
                    If Mid$(s, i + 1, 5) = "u001e" Then
                        result = result & Chr$(30)
                        i = i + 6
                    Else
                        result = result & ch
                        i = i + 1
                    End If
                Case Else
                    result = result & ch
                    i = i + 1
            End Select
        Else
            result = result & ch
            i = i + 1
        End If
    Loop

    DxaUnesc = result
End Function

Private Sub DxaWriteTextUtf8(ByVal path As String, ByVal text As String)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText text
    stm.SaveToFile path, 2
    stm.Close
End Sub

Private Function DxaReadTextUtf8(ByVal path As String) As String
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile path
    DxaReadTextUtf8 = stm.ReadText
    stm.Close
End Function

Private Function DxaFileExists(ByVal path As String) As Boolean
    On Error Resume Next
    DxaFileExists = (Len(Dir$(path, vbNormal + vbHidden + vbSystem)) > 0)
End Function

Private Sub DxaEnsureFolder(ByVal folderPath As String)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
End Sub

Private Sub DxaCleanupOldChangeSnapshots()
    On Error Resume Next
    Dim dirPath As String
    dirPath = DxaChangeSnapshotDir()
    Dim f As String
    f = Dir$(dirPath & "\*.tsv")
    Do While Len(f) > 0
        Dim p As String
        p = dirPath & "\" & f
        If DateDiff("d", FileDateTime(p), Now) >= 1 Then Kill p
        f = Dir$()
    Loop
End Sub

Private Sub DxaDeleteCurrentSessionSnapshots()
    On Error Resume Next
    If Len(gDxaSessionId) = 0 Then Exit Sub
    Dim dirPath As String
    dirPath = DxaChangeSnapshotDir()
    Dim f As String
    f = Dir$(dirPath & "\" & gDxaSessionId & "_*.tsv")
    Do While Len(f) > 0
        Kill dirPath & "\" & f
        f = Dir$()
    Loop
End Sub

Private Function DxaSimpleHash(ByVal text As String) As String
    Dim h As Double
    Dim i As Long
    Dim code As Long
    h = 5381
    For i = 1 To Len(text)
        code = AscW(Mid$(text, i, 1))
        If code < 0 Then code = code + 65536
        h = h * 33 + code
        h = h - Fix(h / 2147483647#) * 2147483647#
    Next
    DxaSimpleHash = Hex$(CLng(h))
End Function

Private Function DxaColumnLetter(ByVal col As Long) As String
    DxaColumnLetter = Split(Cells(1, col).Address(False, False), "1")(0)
End Function

'============================================================
' 表記揺れチェック
'============================================================
Public Sub DxaCheckNotationVariants(ByVal control As Object)
    On Error GoTo EH

    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub

    Dim groups As Object
    Set groups = DxaBuildNotationGroups()
    If groups.Count = 0 Then
        MsgBox "表記揺れチェック用の辞書が空です。", vbExclamation, "DelaxTools"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.StatusBar = "DelaxTools: 表記揺れをチェックしています..."

    Dim records As Collection
    Set records = New Collection

    Dim counts As Object
    Set counts = CreateObject("Scripting.Dictionary")

    Dim found As Object
    Set found = CreateObject("Scripting.Dictionary")

    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        DxaScanNotationWorksheet ws, groups, records, counts, found
    Next

    Dim inconsistent As Object
    Set inconsistent = DxaBuildInconsistentNotationGroups(groups, found)

    DxaOutputNotationCheckWorkbook wb, groups, records, counts, inconsistent

    Application.StatusBar = False
    Application.ScreenUpdating = True

    If inconsistent.Count = 0 Then
        MsgBox "表記揺れは検出されませんでした。結果ブックを作成しました。", vbInformation, "DelaxTools"
    Else
        MsgBox "表記揺れチェックが完了しました。検出グループ数: " & inconsistent.Count, vbInformation, "DelaxTools"
    End If
    Exit Sub
EH:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "表記揺れチェックでエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "DelaxTools"
End Sub

Private Function DxaBuildNotationGroups() As Object
    Dim groups As Object
    Set groups = CreateObject("Scripting.Dictionary")

    DxaAddNotationGroup groups, "server", "サーバー", "サーバー", "サーバ", "server", "Server", "SERVER"
    DxaAddNotationGroup groups, "user", "ユーザー", "ユーザー", "ユーザ", "user", "User", "USER"
    DxaAddNotationGroup groups, "computer", "コンピューター", "コンピューター", "コンピュータ", "PC", "パソコン"
    DxaAddNotationGroup groups, "printer", "プリンター", "プリンター", "プリンタ"
    DxaAddNotationGroup groups, "folder", "フォルダー", "フォルダー", "フォルダ"
    DxaAddNotationGroup groups, "browser", "ブラウザー", "ブラウザー", "ブラウザ"
    DxaAddNotationGroup groups, "driver", "ドライバー", "ドライバー", "ドライバ"
    DxaAddNotationGroup groups, "viewer", "ビューアー", "ビューアー", "ビューア"
    DxaAddNotationGroup groups, "parameter", "パラメーター", "パラメーター", "パラメータ"
    DxaAddNotationGroup groups, "member", "メンバー", "メンバー", "メンバ"
    DxaAddNotationGroup groups, "data", "データ", "データ", "データー"
    DxaAddNotationGroup groups, "database", "データベース", "データベース", "DB", "ＤＢ"
    DxaAddNotationGroup groups, "id", "ID", "ID", "ＩＤ", "Id", "id"
    DxaAddNotationGroup groups, "api", "API", "API", "ＡＰＩ", "Api", "api"
    DxaAddNotationGroup groups, "url", "URL", "URL", "ＵＲＬ", "Url", "url"
    DxaAddNotationGroup groups, "csv", "CSV", "CSV", "ＣＳＶ", "Csv", "csv"
    DxaAddNotationGroup groups, "pdf", "PDF", "PDF", "ＰＤＦ", "Pdf", "pdf"
    DxaAddNotationGroup groups, "excel", "Excel", "Excel", "EXCEL", "エクセル"
    DxaAddNotationGroup groups, "mail", "メール", "メール", "Eメール", "E-Mail", "e-mail", "Email", "email"
    DxaAddNotationGroup groups, "login", "ログイン", "ログイン", "ログオン", "サインイン"
    DxaAddNotationGroup groups, "logout", "ログアウト", "ログアウト", "ログオフ", "サインアウト"
    DxaAddNotationGroup groups, "password", "パスワード", "パスワード", "PW", "ＰＷ", "Password", "password"
    DxaAddNotationGroup groups, "message", "メッセージ", "メッセージ", "メッセージー", "MSG", "ＭＳＧ"
    DxaAddNotationGroup groups, "error", "エラー", "エラー", "エラ－", "ERROR", "Error", "error"
    DxaAddNotationGroup groups, "backup", "バックアップ", "バックアップ", "バックUP", "バックアップデータ"
    DxaAddNotationGroup groups, "master", "マスター", "マスター", "マスタ"
    DxaAddNotationGroup groups, "manager", "マネージャー", "マネージャー", "マネージャ"
    DxaAddNotationGroup groups, "center", "センター", "センター", "センタ"
    DxaAddNotationGroup groups, "check", "チェック", "チェック", "確認"
    DxaAddNotationGroup groups, "delete", "削除", "削除", "消去", "削る"
    DxaAddNotationGroup groups, "update", "更新", "更新", "アップデート", "修正"
    DxaAddNotationGroup groups, "create", "作成", "作成", "生成", "作る"
    DxaAddNotationGroup groups, "register", "登録", "登録", "追加"
    DxaAddNotationGroup groups, "output", "出力", "出力", "エクスポート", "Export", "export"
    DxaAddNotationGroup groups, "input", "入力", "入力", "インポート", "Import", "import"

    DxaLoadUserNotationDictionary groups
    Set DxaBuildNotationGroups = groups
End Function

Private Sub DxaAddNotationGroup(ByVal groups As Object, ByVal groupId As String, ByVal preferred As String, ParamArray variants() As Variant)
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("preferred") = CStr(preferred)
    d("variants") = DxaSortVariantsByLength(variants)
    If groups.Exists(CStr(groupId)) Then
        Set groups(CStr(groupId)) = d
    Else
        groups.Add CStr(groupId), d
    End If
End Sub

Private Function DxaSortVariantsByLength(ByVal variants As Variant) As Variant
    Dim arr() As String
    Dim i As Long
    ReDim arr(LBound(variants) To UBound(variants))
    For i = LBound(variants) To UBound(variants)
        arr(i) = CStr(variants(i))
    Next

    Dim j As Long, tmp As String
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If Len(arr(j)) > Len(arr(i)) Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next
    Next
    DxaSortVariantsByLength = arr
End Function

Private Sub DxaLoadUserNotationDictionary(ByVal groups As Object)
    On Error Resume Next
    Dim path As String
    path = DxaNotationDictionaryPath()
    If Len(Dir$(path)) = 0 Then Exit Sub

    Dim text As String
    text = DxaReadTextUtf8(path)
    If Len(text) = 0 Then Exit Sub

    Dim lines As Variant
    lines = Split(Replace(text, vbCrLf, vbLf), vbLf)

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim line As String
        line = Trim$(CStr(lines(i)))
        If Len(line) = 0 Then GoTo ContinueLine
        If Left$(line, 1) = "#" Then GoTo ContinueLine

        Dim parts As Variant
        parts = Split(line, ",")
        If UBound(parts) >= 1 Then
            Dim preferred As String
            preferred = Trim$(CStr(parts(0)))
            If Len(preferred) > 0 Then
                Dim v() As Variant
                ReDim v(0 To UBound(parts))
                Dim j As Long
                v(0) = preferred
                For j = 1 To UBound(parts)
                    v(j) = Trim$(CStr(parts(j)))
                Next
                DxaAddNotationGroup groups, "user_" & CStr(i), preferred, v
            End If
        End If
ContinueLine:
    Next
End Sub

Private Function DxaNotationDictionaryPath() As String
    DxaNotationDictionaryPath = Environ$("APPDATA") & "\DelaxTools\notation_variants.csv"
End Function

Private Sub DxaScanNotationWorksheet(ByVal ws As Worksheet, ByVal groups As Object, ByVal records As Collection, ByVal counts As Object, ByVal found As Object)
    On Error Resume Next
    If Application.WorksheetFunction.CountA(ws.Cells) > 0 Then
        Dim rng As Range
        Set rng = ws.UsedRange
        If Not rng Is Nothing Then
            Dim c As Range
            For Each c In rng.Cells
                If Not IsError(c.Value) Then
                    Dim text As String
                    text = CStr(c.Value)
                    If Len(text) > 0 Then DxaCollectNotationHits groups, records, counts, found, ws.Name, c.Address(False, False), "セル", text
                End If
            Next
        End If
    End If

    Dim shp As Shape
    For Each shp In ws.Shapes
        Dim s As String
        s = ""
        On Error Resume Next
        If shp.TextFrame2.HasText Then s = shp.TextFrame2.TextRange.Text
        If Len(s) = 0 Then
            If shp.TextFrame.HasText Then s = shp.TextFrame.Characters.Text
        End If
        On Error GoTo 0
        If Len(s) > 0 Then DxaCollectNotationHits groups, records, counts, found, ws.Name, shp.Name, "図形", s
    Next
End Sub

Private Sub DxaCollectNotationHits(ByVal groups As Object, ByVal records As Collection, ByVal counts As Object, ByVal found As Object, ByVal sheetName As String, ByVal location As String, ByVal kind As String, ByVal text As String)
    Dim gKey As Variant
    For Each gKey In groups.Keys
        Dim g As Object
        Set g = groups(gKey)
        Dim vars As Variant
        vars = g("variants")

        Dim i As Long
        For i = LBound(vars) To UBound(vars)
            Dim variantText As String
            variantText = CStr(vars(i))
            If Len(variantText) > 0 Then
                If DxaContainsNotationVariant(text, variantText) Then
                    Dim countKey As String
                    countKey = CStr(gKey) & Chr$(31) & variantText
                    If counts.Exists(countKey) Then counts(countKey) = CLng(counts(countKey)) + 1 Else counts(countKey) = 1
                    found(countKey) = True
                    records.Add Array(CStr(gKey), variantText, CStr(g("preferred")), sheetName, location, kind, DxaShortText(text))
                End If
            End If
        Next
    Next
End Sub

Private Function DxaContainsNotationVariant(ByVal text As String, ByVal variantText As String) As Boolean
    Dim pos As Long
    pos = 1
    Do
        pos = InStr(pos, text, variantText, vbTextCompare)
        If pos = 0 Then Exit Function
        If DxaIsValidNotationHit(text, variantText, pos) Then
            DxaContainsNotationVariant = True
            Exit Function
        End If
        pos = pos + Len(variantText)
    Loop
End Function

Private Function DxaIsValidNotationHit(ByVal text As String, ByVal variantText As String, ByVal pos As Long) As Boolean
    Dim beforeCh As String, afterCh As String
    beforeCh = "": afterCh = ""
    If pos > 1 Then beforeCh = Mid$(text, pos - 1, 1)
    If pos + Len(variantText) <= Len(text) Then afterCh = Mid$(text, pos + Len(variantText), 1)

    If DxaIsAsciiWord(variantText) Then
        If DxaIsAsciiWordChar(beforeCh) Then Exit Function
        If DxaIsAsciiWordChar(afterCh) Then Exit Function
    End If

    If Right$(variantText, 1) <> "ー" And Right$(variantText, 1) <> "ｰ" Then
        If afterCh = "ー" Or afterCh = "ｰ" Then Exit Function
    End If

    DxaIsValidNotationHit = True
End Function

Private Function DxaIsAsciiWord(ByVal s As String) As Boolean
    Dim i As Long, code As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        code = AscW(Mid$(s, i, 1))
        If Not ((code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Or code = 95) Then Exit Function
    Next
    DxaIsAsciiWord = True
End Function

Private Function DxaIsAsciiWordChar(ByVal s As String) As Boolean
    If Len(s) = 0 Then Exit Function
    DxaIsAsciiWordChar = DxaIsAsciiWord(Left$(s, 1))
End Function

Private Function DxaBuildInconsistentNotationGroups(ByVal groups As Object, ByVal found As Object) As Object
    Dim inconsistent As Object
    Set inconsistent = CreateObject("Scripting.Dictionary")

    Dim gKey As Variant
    For Each gKey In groups.Keys
        Dim g As Object
        Set g = groups(gKey)
        Dim vars As Variant
        vars = g("variants")
        Dim foundCount As Long
        Dim i As Long
        For i = LBound(vars) To UBound(vars)
            If found.Exists(CStr(gKey) & Chr$(31) & CStr(vars(i))) Then foundCount = foundCount + 1
        Next
        If foundCount >= 2 Then inconsistent(CStr(gKey)) = True
    Next

    Set DxaBuildInconsistentNotationGroups = inconsistent
End Function

Private Sub DxaOutputNotationCheckWorkbook(ByVal sourceWb As Workbook, ByVal groups As Object, ByVal records As Collection, ByVal counts As Object, ByVal inconsistent As Object)
    Dim outWb As Workbook
    Set outWb = Workbooks.Add(xlWBATWorksheet)

    Dim wsSummary As Worksheet
    Set wsSummary = outWb.Worksheets(1)
    wsSummary.Name = "表記揺れチェック"

    Dim wsDetail As Worksheet
    Set wsDetail = outWb.Worksheets.Add(After:=wsSummary)
    wsDetail.Name = "検出詳細"

    DxaWriteNotationSummarySheet wsSummary, groups, counts, inconsistent
    DxaWriteNotationDetailSheet sourceWb, wsDetail, records, inconsistent

    wsSummary.Activate
End Sub

Private Sub DxaWriteNotationSummarySheet(ByVal ws As Worksheet, ByVal groups As Object, ByVal counts As Object, ByVal inconsistent As Object)
    ws.Range("A1:F1").Value = Array("No", "推奨表記", "検出表記", "件数", "判定", "備考")
    ws.Range("A1:F1").Font.Bold = True

    Dim r As Long
    r = 2

    If inconsistent.Count = 0 Then
        ws.Cells(r, 1).Value = 1
        ws.Cells(r, 5).Value = "表記揺れなし"
        ws.Cells(r, 6).Value = "同一グループ内で複数表記は検出されませんでした。"
    Else
        Dim gKey As Variant
        For Each gKey In groups.Keys
            If inconsistent.Exists(CStr(gKey)) Then
                Dim g As Object
                Set g = groups(gKey)
                Dim vars As Variant
                vars = g("variants")
                Dim i As Long
                For i = LBound(vars) To UBound(vars)
                    Dim countKey As String
                    countKey = CStr(gKey) & Chr$(31) & CStr(vars(i))
                    If counts.Exists(countKey) Then
                        ws.Cells(r, 1).Value = r - 1
                        ws.Cells(r, 2).Value = CStr(g("preferred"))
                        ws.Cells(r, 3).Value = CStr(vars(i))
                        ws.Cells(r, 4).Value = CLng(counts(countKey))
                        If StrComp(CStr(vars(i)), CStr(g("preferred")), vbTextCompare) = 0 Then
                            ws.Cells(r, 5).Value = "推奨表記"
                        Else
                            ws.Cells(r, 5).Value = "揺れ候補"
                        End If
                        ws.Cells(r, 6).Value = "推奨表記に統一するか確認してください。"
                        r = r + 1
                    End If
                Next
            End If
        Next
    End If

    ws.Columns("A:F").AutoFit
    ws.Range("A1:F1").AutoFilter
End Sub

Private Sub DxaWriteNotationDetailSheet(ByVal sourceWb As Workbook, ByVal ws As Worksheet, ByVal records As Collection, ByVal inconsistent As Object)
    ws.Range("A1:H1").Value = Array("No", "対象シート", "場所", "種別", "検出表記", "推奨表記", "周辺テキスト", "確認結果")
    ws.Range("A1:H1").Font.Bold = True

    Dim r As Long
    r = 2

    Dim i As Long
    For i = 1 To records.Count
        Dim rec As Variant
        rec = records(i)
        If inconsistent.Exists(CStr(rec(0))) Then
            ws.Cells(r, 1).Value = r - 1
            ws.Cells(r, 2).Value = rec(3)
            ws.Cells(r, 3).Value = rec(4)
            ws.Cells(r, 4).Value = rec(5)
            ws.Cells(r, 5).Value = rec(1)
            ws.Cells(r, 6).Value = rec(2)
            ws.Cells(r, 7).Value = rec(6)
            ws.Cells(r, 8).Value = "要確認"

            On Error Resume Next
            If CStr(rec(5)) = "セル" And Len(sourceWb.FullName) > 0 Then
                ws.Hyperlinks.Add Anchor:=ws.Cells(r, 3), Address:=sourceWb.FullName, SubAddress:="'" & rec(3) & "'!" & rec(4), TextToDisplay:=rec(4)
            End If
            On Error GoTo 0
            r = r + 1
        End If
    Next

    If r = 2 Then
        ws.Cells(2, 1).Value = 1
        ws.Cells(2, 8).Value = "表記揺れなし"
    End If

    ws.Columns("A:H").AutoFit
    ws.Range("A1:H1").AutoFilter
End Sub

' ============================================================
' 重いExcel診断
' 元ブックにはシートを追加せず、診断結果を別ブックに出力します。
' ============================================================
Public Sub DxaDiagnoseHeavyWorkbook(ByVal control As Object)
    On Error GoTo EH

    Dim srcWb As Workbook
    Set srcWb = ActiveWorkbook
    If srcWb Is Nothing Then Exit Sub
    If srcWb.Name = ThisWorkbook.Name Then
        MsgBox "診断対象のブックをアクティブにしてから実行してください。", vbExclamation, "DelaxTools"
        Exit Sub
    End If

    Dim reportWb As Workbook
    Dim wsSummary As Worksheet
    Dim wsDetail As Worksheet

    Application.ScreenUpdating = False
    Application.StatusBar = "DelaxTools: 重いExcel診断を実行しています..."

    Set reportWb = Application.Workbooks.Add(xlWBATWorksheet)
    Set wsSummary = reportWb.Worksheets(1)
    wsSummary.Name = "重いExcel診断"
    Set wsDetail = reportWb.Worksheets.Add(After:=wsSummary)
    wsDetail.Name = "診断詳細"

    DxaPrepareHeavySummarySheet wsSummary, srcWb
    DxaPrepareHeavyDetailSheet wsDetail

    Dim detailRow As Long
    detailRow = 2

    Dim totalFormula As Double
    Dim totalVolatile As Double
    Dim totalFormatConditions As Double
    Dim totalValidations As Double
    Dim totalShapes As Double
    Dim totalPictures As Double
    Dim totalHyperlinks As Double
    Dim totalComments As Double
    Dim totalPivotTables As Double
    Dim totalTables As Double
    Dim totalHiddenSheets As Double
    Dim totalBloatedUsedRange As Double
    Dim totalLargeSheets As Double

    Dim ws As Worksheet
    For Each ws In srcWb.Worksheets
        Application.StatusBar = "DelaxTools: 重いExcel診断中 - " & ws.Name

        Dim lastRow As Long
        Dim lastCol As Long
        Dim hasData As Boolean
        hasData = DxaGetActualLastCell(ws, lastRow, lastCol)

        Dim usedRows As Double
        Dim usedCols As Double
        Dim usedCells As Double
        usedRows = 0
        usedCols = 0
        usedCells = 0
        On Error Resume Next
        usedRows = CDbl(ws.UsedRange.Rows.Count)
        usedCols = CDbl(ws.UsedRange.Columns.Count)
        usedCells = CDbl(ws.UsedRange.CountLarge)
        On Error GoTo EH

        Dim actualCells As Double
        If hasData Then
            actualCells = CDbl(lastRow) * CDbl(lastCol)
        Else
            actualCells = 0
        End If

        Dim formulaCount As Double
        formulaCount = DxaCountSpecialCells(ws, xlCellTypeFormulas)
        totalFormula = totalFormula + formulaCount

        Dim volatileCount As Double
        volatileCount = DxaCountVolatileFormulas(ws)
        totalVolatile = totalVolatile + volatileCount

        Dim fcCount As Double
        fcCount = DxaCountFormatConditions(ws)
        totalFormatConditions = totalFormatConditions + fcCount

        Dim validationCount As Double
        validationCount = DxaCountSpecialCells(ws, xlCellTypeAllValidation)
        totalValidations = totalValidations + validationCount

        Dim shapeCount As Double
        shapeCount = DxaSafeShapeCount(ws)
        totalShapes = totalShapes + shapeCount

        Dim pictureCount As Double
        pictureCount = DxaSafePictureCount(ws)
        totalPictures = totalPictures + pictureCount

        Dim hyperlinkCount As Double
        hyperlinkCount = DxaSafeHyperlinkCount(ws)
        totalHyperlinks = totalHyperlinks + hyperlinkCount

        Dim commentCount As Double
        commentCount = DxaSafeCommentCount(ws)
        totalComments = totalComments + commentCount

        Dim pivotCount As Double
        pivotCount = DxaSafePivotCount(ws)
        totalPivotTables = totalPivotTables + pivotCount

        Dim tableCount As Double
        tableCount = DxaSafeTableCount(ws)
        totalTables = totalTables + tableCount

        If ws.Visible <> xlSheetVisible Then totalHiddenSheets = totalHiddenSheets + 1

        Dim usedRangeStatus As String
        Dim usedRangeReason As String
        usedRangeStatus = "OK"
        usedRangeReason = "使用範囲に大きな異常は見つかりません。"
        If usedCells >= 1000000# Then
            usedRangeStatus = "注意"
            usedRangeReason = "UsedRangeが大きいです。不要な行列に書式が残っている可能性があります。"
            totalLargeSheets = totalLargeSheets + 1
        End If
        If hasData Then
            If (usedRows > lastRow + 1000) Or (usedCols > lastCol + 20) Then
                usedRangeStatus = "警告"
                usedRangeReason = "実データ範囲よりUsedRangeが広いです。未使用範囲のリセット候補です。"
                totalBloatedUsedRange = totalBloatedUsedRange + 1
            End If
        End If

        DxaWriteHeavyDetail wsDetail, detailRow, "シート概要", ws.Name, "表示状態", DxaSheetVisibleText(ws), DxaStatusBySheetVisible(ws), "非表示/VeryHiddenシートが不要であれば表示または削除を検討してください。"
        DxaWriteHeavyDetail wsDetail, detailRow, "使用範囲", ws.Name, "UsedRange", "行=" & CStr(usedRows) & ", 列=" & CStr(usedCols) & ", セル=" & Format$(usedCells, "#,##0"), usedRangeStatus, usedRangeReason
        DxaWriteHeavyDetail wsDetail, detailRow, "実データ範囲", ws.Name, "最終セル", IIf(hasData, "行=" & CStr(lastRow) & ", 列=" & CStr(lastCol), "データなし"), "情報", "UsedRangeと実データ範囲の差が大きい場合、Excelが重くなる原因になります。"
        DxaWriteHeavyDetail wsDetail, detailRow, "数式", ws.Name, "数式セル数", Format$(formulaCount, "#,##0"), DxaStatusByNumber(formulaCount, 10000, 50000), "数式が多い場合は計算方式、不要数式、値貼り付けを検討してください。"
        DxaWriteHeavyDetail wsDetail, detailRow, "揮発性関数", ws.Name, "推定件数", Format$(volatileCount, "#,##0"), DxaStatusByNumber(volatileCount, 1, 100), "NOW/TODAY/RAND/OFFSET/INDIRECTなどは再計算負荷が高くなる場合があります。"
        DxaWriteHeavyDetail wsDetail, detailRow, "条件付き書式", ws.Name, "件数", Format$(fcCount, "#,##0"), DxaStatusByNumber(fcCount, 100, 1000), "コピー貼り付けで条件付き書式が増殖していないか確認してください。"
        DxaWriteHeavyDetail wsDetail, detailRow, "入力規則", ws.Name, "対象セル数", Format$(validationCount, "#,##0"), DxaStatusByNumber(validationCount, 5000, 50000), "入力規則が大量に複製されると動作が重くなる場合があります。"
        DxaWriteHeavyDetail wsDetail, detailRow, "図形/画像", ws.Name, "図形=" & Format$(shapeCount, "#,##0"), "画像=" & Format$(pictureCount, "#,##0"), DxaStatusByNumber(shapeCount, 100, 500), "不要な図形、透明画像、貼り付け画像が残っていないか確認してください。"
        DxaWriteHeavyDetail wsDetail, detailRow, "リンク/コメント", ws.Name, "リンク=" & Format$(hyperlinkCount, "#,##0"), "コメント=" & Format$(commentCount, "#,##0"), DxaStatusByNumber(hyperlinkCount + commentCount, 200, 1000), "不要なリンク、コメント、メモが残っていないか確認してください。"
        DxaWriteHeavyDetail wsDetail, detailRow, "集計オブジェクト", ws.Name, "ピボット=" & Format$(pivotCount, "#,##0"), "テーブル=" & Format$(tableCount, "#,##0"), "情報", "ピボットやテーブルが多い場合は更新範囲やキャッシュを確認してください。"
    Next ws

    Dim externalLinkCount As Double
    externalLinkCount = DxaCountExternalLinks(srcWb)

    Dim nameCount As Double
    Dim brokenNameCount As Double
    Dim externalNameCount As Double
    DxaCountNames srcWb, nameCount, brokenNameCount, externalNameCount

    Dim styleCount As Double
    styleCount = DxaSafeStyleCount(srcWb)

    Dim fileSizeText As String
    Dim fileSizeMB As Double
    fileSizeMB = DxaWorkbookFileSizeMB(srcWb)
    If fileSizeMB >= 0 Then
        fileSizeText = Format$(fileSizeMB, "0.00") & " MB"
    Else
        fileSizeText = "未保存または取得不可"
    End If

    Dim r As Long
    r = 5
    DxaWriteHeavySummary wsSummary, r, "ファイルサイズ", DxaStatusByFileSize(fileSizeMB), fileSizeText, "ファイルサイズが大きい場合は画像、条件付き書式、未使用範囲、不要スタイルを確認してください。"
    DxaWriteHeavySummary wsSummary, r, "シート数", DxaStatusByNumber(srcWb.Worksheets.Count, 30, 80), CStr(srcWb.Worksheets.Count), "シート数が多い場合は不要シートや非表示シートを確認してください。"
    DxaWriteHeavySummary wsSummary, r, "非表示シート数", DxaStatusByNumber(totalHiddenSheets, 1, 10), Format$(totalHiddenSheets, "#,##0"), "不要な非表示/VeryHiddenシートがないか確認してください。"
    DxaWriteHeavySummary wsSummary, r, "UsedRange肥大候補", DxaStatusByNumber(totalBloatedUsedRange, 1, 5), Format$(totalBloatedUsedRange, "#,##0"), "実データ範囲よりUsedRangeが広いシートは、未使用範囲リセット候補です。"
    DxaWriteHeavySummary wsSummary, r, "大規模UsedRangeシート", DxaStatusByNumber(totalLargeSheets, 1, 5), Format$(totalLargeSheets, "#,##0"), "使用範囲が非常に大きいシートは重くなる原因です。"
    DxaWriteHeavySummary wsSummary, r, "数式セル数", DxaStatusByNumber(totalFormula, 50000, 200000), Format$(totalFormula, "#,##0"), "数式が多い場合、値貼り付け・計算範囲見直しを検討してください。"
    DxaWriteHeavySummary wsSummary, r, "揮発性関数推定数", DxaStatusByNumber(totalVolatile, 1, 100), Format$(totalVolatile, "#,##0"), "揮発性関数は再計算負荷が高いため、必要性を確認してください。"
    DxaWriteHeavySummary wsSummary, r, "条件付き書式数", DxaStatusByNumber(totalFormatConditions, 500, 3000), Format$(totalFormatConditions, "#,##0"), "条件付き書式が増殖している場合は整理してください。"
    DxaWriteHeavySummary wsSummary, r, "入力規則対象セル数", DxaStatusByNumber(totalValidations, 10000, 100000), Format$(totalValidations, "#,##0"), "入力規則が広範囲に設定されている場合は範囲を見直してください。"
    DxaWriteHeavySummary wsSummary, r, "図形数", DxaStatusByNumber(totalShapes, 200, 1000), Format$(totalShapes, "#,##0"), "不要な図形や透明オブジェクトがないか確認してください。"
    DxaWriteHeavySummary wsSummary, r, "画像数", DxaStatusByNumber(totalPictures, 50, 200), Format$(totalPictures, "#,##0"), "画像が多い場合は圧縮や不要画像削除を検討してください。"
    DxaWriteHeavySummary wsSummary, r, "外部リンク数", DxaStatusByNumber(externalLinkCount, 1, 10), Format$(externalLinkCount, "#,##0"), "不要な外部リンクが残っていないか確認してください。"
    DxaWriteHeavySummary wsSummary, r, "名前定義数", DxaStatusByNumber(nameCount, 200, 1000), Format$(nameCount, "#,##0"), "不要な名前定義が増えていないか確認してください。"
    DxaWriteHeavySummary wsSummary, r, "参照切れ名前定義数", DxaStatusByNumber(brokenNameCount, 1, 10), Format$(brokenNameCount, "#,##0"), "#REF!を含む名前定義は削除候補です。"
    DxaWriteHeavySummary wsSummary, r, "外部参照名前定義数", DxaStatusByNumber(externalNameCount, 1, 10), Format$(externalNameCount, "#,##0"), "名前定義内の外部参照は外部リンク警告の原因になる場合があります。"
    DxaWriteHeavySummary wsSummary, r, "スタイル数", DxaStatusByNumber(styleCount, 500, 2000), Format$(styleCount, "#,##0"), "不要スタイルが増殖している場合、ファイル肥大化の原因になる場合があります。"

    wsSummary.Columns("A:D").AutoFit
    wsDetail.Columns("A:G").AutoFit
    wsSummary.Range("A4:D4").AutoFilter
    wsDetail.Range("A1:G1").AutoFilter
    wsSummary.Activate
    wsSummary.Range("A1").Select

    Application.StatusBar = False
    Application.ScreenUpdating = True

    MsgBox "重いExcel診断が完了しました。結果は別ブックに出力しました。", vbInformation, "DelaxTools"
    Exit Sub

EH:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "重いExcel診断でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "DelaxTools"
End Sub

Private Sub DxaPrepareHeavySummarySheet(ByVal ws As Worksheet, ByVal srcWb As Workbook)
    ws.Range("A1").Value = "重いExcel診断"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16
    ws.Range("A2").Value = "対象ブック"
    ws.Range("B2").Value = srcWb.Name
    ws.Range("A3").Value = "診断日時"
    ws.Range("B3").Value = Now
    ws.Range("B3").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    ws.Range("A4:D4").Value = Array("診断項目", "判定", "件数/値", "推奨対応")
    ws.Range("A4:D4").Font.Bold = True
End Sub

Private Sub DxaPrepareHeavyDetailSheet(ByVal ws As Worksheet)
    ws.Range("A1:G1").Value = Array("No", "カテゴリ", "シート名", "対象", "値1", "判定", "推奨対応")
    ws.Range("A1:G1").Font.Bold = True
End Sub

Private Sub DxaWriteHeavySummary(ByVal ws As Worksheet, ByRef rowNo As Long, ByVal itemName As String, ByVal statusText As String, ByVal valueText As String, ByVal recommendation As String)
    ws.Cells(rowNo, 1).Value = itemName
    ws.Cells(rowNo, 2).Value = statusText
    ws.Cells(rowNo, 3).Value = valueText
    ws.Cells(rowNo, 4).Value = recommendation
    DxaApplyStatusColor ws.Cells(rowNo, 2), statusText
    rowNo = rowNo + 1
End Sub

Private Sub DxaWriteHeavyDetail(ByVal ws As Worksheet, ByRef rowNo As Long, ByVal category As String, ByVal sheetName As String, ByVal targetName As String, ByVal valueText As String, ByVal statusText As String, ByVal recommendation As String)
    ws.Cells(rowNo, 1).Value = rowNo - 1
    ws.Cells(rowNo, 2).Value = category
    ws.Cells(rowNo, 3).Value = sheetName
    ws.Cells(rowNo, 4).Value = targetName
    ws.Cells(rowNo, 5).Value = valueText
    ws.Cells(rowNo, 6).Value = statusText
    ws.Cells(rowNo, 7).Value = recommendation
    DxaApplyStatusColor ws.Cells(rowNo, 6), statusText
    rowNo = rowNo + 1
End Sub

Private Sub DxaApplyStatusColor(ByVal cell As Range, ByVal statusText As String)
    On Error Resume Next
    Select Case statusText
        Case "警告"
            cell.Interior.Color = RGB(255, 199, 206)
            cell.Font.Color = RGB(156, 0, 6)
        Case "注意"
            cell.Interior.Color = RGB(255, 235, 156)
            cell.Font.Color = RGB(156, 101, 0)
        Case "OK"
            cell.Interior.Color = RGB(198, 239, 206)
            cell.Font.Color = RGB(0, 97, 0)
        Case Else
            cell.Interior.Color = RGB(217, 225, 242)
    End Select
End Sub

Private Function DxaStatusByNumber(ByVal value As Double, ByVal cautionThreshold As Double, ByVal warningThreshold As Double) As String
    If value >= warningThreshold Then
        DxaStatusByNumber = "警告"
    ElseIf value >= cautionThreshold Then
        DxaStatusByNumber = "注意"
    Else
        DxaStatusByNumber = "OK"
    End If
End Function

Private Function DxaStatusByFileSize(ByVal mb As Double) As String
    If mb < 0 Then
        DxaStatusByFileSize = "情報"
    ElseIf mb >= 50 Then
        DxaStatusByFileSize = "警告"
    ElseIf mb >= 10 Then
        DxaStatusByFileSize = "注意"
    Else
        DxaStatusByFileSize = "OK"
    End If
End Function

Private Function DxaStatusBySheetVisible(ByVal ws As Worksheet) As String
    If ws.Visible = xlSheetVisible Then
        DxaStatusBySheetVisible = "OK"
    Else
        DxaStatusBySheetVisible = "注意"
    End If
End Function

Private Function DxaSheetVisibleText(ByVal ws As Worksheet) As String
    Select Case ws.Visible
        Case xlSheetVisible
            DxaSheetVisibleText = "表示"
        Case xlSheetHidden
            DxaSheetVisibleText = "非表示"
        Case xlSheetVeryHidden
            DxaSheetVisibleText = "VeryHidden"
        Case Else
            DxaSheetVisibleText = CStr(ws.Visible)
    End Select
End Function

Private Function DxaGetActualLastCell(ByVal ws As Worksheet, ByRef lastRow As Long, ByRef lastCol As Long) As Boolean
    On Error GoTo EH
    Dim c As Range
    Set c = ws.Cells.Find(What:="*", After:=ws.Range("A1"), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    If c Is Nothing Then
        lastRow = 0
        lastCol = 0
        DxaGetActualLastCell = False
        Exit Function
    End If
    lastRow = c.Row
    Set c = ws.Cells.Find(What:="*", After:=ws.Range("A1"), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    lastCol = c.Column
    DxaGetActualLastCell = True
    Exit Function
EH:
    lastRow = 0
    lastCol = 0
    DxaGetActualLastCell = False
End Function

Private Function DxaCountSpecialCells(ByVal ws As Worksheet, ByVal cellType As Long) As Double
    On Error GoTo EH
    Dim rng As Range
    Set rng = ws.UsedRange.SpecialCells(cellType)
    DxaCountSpecialCells = CDbl(rng.CountLarge)
    Exit Function
EH:
    DxaCountSpecialCells = 0
End Function

Private Function DxaCountFormatConditions(ByVal ws As Worksheet) As Double
    On Error GoTo EH
    DxaCountFormatConditions = CDbl(ws.Cells.FormatConditions.Count)
    Exit Function
EH:
    DxaCountFormatConditions = 0
End Function

Private Function DxaSafeShapeCount(ByVal ws As Worksheet) As Double
    On Error Resume Next
    DxaSafeShapeCount = CDbl(ws.Shapes.Count)
End Function

Private Function DxaSafePictureCount(ByVal ws As Worksheet) As Double
    On Error GoTo EH
    Dim shp As Shape
    Dim n As Double
    For Each shp In ws.Shapes
        If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then n = n + 1
    Next shp
    DxaSafePictureCount = n
    Exit Function
EH:
    DxaSafePictureCount = 0
End Function

Private Function DxaSafeHyperlinkCount(ByVal ws As Worksheet) As Double
    On Error Resume Next
    DxaSafeHyperlinkCount = CDbl(ws.Hyperlinks.Count)
End Function

Private Function DxaSafeCommentCount(ByVal ws As Worksheet) As Double
    On Error Resume Next
    Dim n As Double
    n = 0
    n = n + CDbl(ws.Comments.Count)
    Err.Clear
    n = n + CDbl(ws.CommentsThreaded.Count)
    DxaSafeCommentCount = n
End Function

Private Function DxaSafePivotCount(ByVal ws As Worksheet) As Double
    On Error Resume Next
    DxaSafePivotCount = CDbl(ws.PivotTables.Count)
End Function

Private Function DxaSafeTableCount(ByVal ws As Worksheet) As Double
    On Error Resume Next
    DxaSafeTableCount = CDbl(ws.ListObjects.Count)
End Function

Private Function DxaSafeStyleCount(ByVal wb As Workbook) As Double
    On Error Resume Next
    DxaSafeStyleCount = CDbl(wb.Styles.Count)
End Function

Private Function DxaWorkbookFileSizeMB(ByVal wb As Workbook) As Double
    On Error GoTo EH
    If Len(wb.FullName) = 0 Then
        DxaWorkbookFileSizeMB = -1
    ElseIf Len(Dir$(wb.FullName)) = 0 Then
        DxaWorkbookFileSizeMB = -1
    Else
        DxaWorkbookFileSizeMB = CDbl(FileLen(wb.FullName)) / 1024# / 1024#
    End If
    Exit Function
EH:
    DxaWorkbookFileSizeMB = -1
End Function

Private Function DxaCountExternalLinks(ByVal wb As Workbook) As Double
    On Error GoTo EH
    Dim links As Variant
    links = wb.LinkSources(Type:=xlLinkTypeExcelLinks)
    If IsEmpty(links) Then
        DxaCountExternalLinks = 0
    Else
        DxaCountExternalLinks = CDbl(UBound(links) - LBound(links) + 1)
    End If
    Exit Function
EH:
    DxaCountExternalLinks = 0
End Function

Private Sub DxaCountNames(ByVal wb As Workbook, ByRef nameCount As Double, ByRef brokenNameCount As Double, ByRef externalNameCount As Double)
    On Error Resume Next
    Dim nm As Name
    Dim refText As String
    nameCount = 0
    brokenNameCount = 0
    externalNameCount = 0
    For Each nm In wb.Names
        nameCount = nameCount + 1
        Err.Clear
        refText = nm.RefersTo
        If InStr(1, refText, "#REF!", vbTextCompare) > 0 Then brokenNameCount = brokenNameCount + 1
        If InStr(1, refText, "[", vbTextCompare) > 0 And InStr(1, refText, "]", vbTextCompare) > 0 Then externalNameCount = externalNameCount + 1
    Next nm
End Sub

Private Function DxaCountVolatileFormulas(ByVal ws As Worksheet) As Double
    On Error GoTo EH
    Dim rng As Range
    Set rng = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    If rng Is Nothing Then
        DxaCountVolatileFormulas = 0
        Exit Function
    End If

    Dim volatileWords As Variant
    volatileWords = Array("NOW(", "TODAY(", "RAND(", "RANDBETWEEN(", "OFFSET(", "INDIRECT(", "CELL(", "INFO(")

    Dim total As Double
    Dim word As Variant
    For Each word In volatileWords
        total = total + DxaCountFormulaFindHits(rng, CStr(word))
    Next word
    DxaCountVolatileFormulas = total
    Exit Function
EH:
    DxaCountVolatileFormulas = 0
End Function

Private Function DxaCountFormulaFindHits(ByVal rng As Range, ByVal token As String) As Double
    On Error GoTo EH
    Dim firstAddress As String
    Dim c As Range
    Dim n As Double
    Set c = rng.Find(What:=token, After:=rng.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    If c Is Nothing Then
        DxaCountFormulaFindHits = 0
        Exit Function
    End If
    firstAddress = c.Address(External:=True)
    Do
        n = n + 1
        Set c = rng.FindNext(c)
        If c Is Nothing Then Exit Do
    Loop While c.Address(External:=True) <> firstAddress
    DxaCountFormulaFindHits = n
    Exit Function
EH:
    DxaCountFormulaFindHits = 0
End Function

'============================================================
' Backlogガントチャート支援機能 v107
' Backlogからエクスポートしたガントチャートを見やすく整形します。
' 想定形式：A～L列が課題情報、M列以降が日付ガント
'============================================================
Public Sub DxaBacklogFormatGantt(ByVal control As Object)
    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub

    Dim headerRow As Long, dataFirstRow As Long, lastRow As Long, lastCol As Long, dateStartCol As Long
    If Not DxaBacklogDetectLayout(ws, headerRow, dataFirstRow, lastRow, lastCol, dateStartCol) Then Exit Sub

    Application.ScreenUpdating = False
    Application.StatusBar = "DelaxTools: Backlogガントを整形しています..."

    DxaBacklogFormatIssueColumns ws, headerRow, dataFirstRow, lastRow, lastCol, dateStartCol
    DxaBacklogFormatDateColumns ws, headerRow, dataFirstRow, lastRow, lastCol, dateStartCol
    DxaBacklogHighlightRows ws, dataFirstRow, lastRow, lastCol
    DxaBacklogFreezeGantt ws, dataFirstRow, dateStartCol

    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Backlogガント整形が完了しました。", vbInformation, "DelaxTools"
    Exit Sub
EH:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Backlogガント整形でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "DelaxTools"
End Sub

Public Sub DxaBacklogCreateGanttSummary(ByVal control As Object)
    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub

    Dim headerRow As Long, dataFirstRow As Long, lastRow As Long, lastCol As Long, dateStartCol As Long
    If Not DxaBacklogDetectLayout(ws, headerRow, dataFirstRow, lastRow, lastCol, dateStartCol) Then Exit Sub

    Application.ScreenUpdating = False
    Application.StatusBar = "DelaxTools: Backlogガントサマリーを作成しています..."

    Dim outWs As Worksheet
    Set outWs = DxaBacklogRecreateSheet(ws.Parent, "Backlogガントサマリー")
    DxaBacklogWriteSummary ws, outWs, dataFirstRow, lastRow

    outWs.Activate
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Backlogガントサマリーを作成しました。", vbInformation, "DelaxTools"
    Exit Sub
EH:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Backlogガントサマリー作成でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "DelaxTools"
End Sub

Public Sub DxaBacklogCreateDelayList(ByVal control As Object)
    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub

    Dim headerRow As Long, dataFirstRow As Long, lastRow As Long, lastCol As Long, dateStartCol As Long
    If Not DxaBacklogDetectLayout(ws, headerRow, dataFirstRow, lastRow, lastCol, dateStartCol) Then Exit Sub

    Application.ScreenUpdating = False
    Application.StatusBar = "DelaxTools: Backlog遅延一覧を作成しています..."

    Dim outWs As Worksheet
    Set outWs = DxaBacklogRecreateSheet(ws.Parent, "Backlog遅延一覧")
    DxaBacklogWriteDelayList ws, outWs, dataFirstRow, lastRow

    outWs.Activate
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Backlog遅延一覧を作成しました。", vbInformation, "DelaxTools"
    Exit Sub
EH:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Backlog遅延一覧作成でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "DelaxTools"
End Sub

Public Sub DxaBacklogCreateMeetingView(ByVal control As Object)
    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub

    Dim headerRow As Long, dataFirstRow As Long, lastRow As Long, lastCol As Long, dateStartCol As Long
    If Not DxaBacklogDetectLayout(ws, headerRow, dataFirstRow, lastRow, lastCol, dateStartCol) Then Exit Sub

    Application.ScreenUpdating = False
    Application.StatusBar = "DelaxTools: Backlog会議用ビューを作成しています..."

    Dim outWs As Worksheet
    Set outWs = DxaBacklogRecreateSheet(ws.Parent, "Backlog会議用")
    DxaBacklogWriteMeetingView ws, outWs, dataFirstRow, lastRow

    outWs.Activate
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Backlog会議用ビューを作成しました。", vbInformation, "DelaxTools"
    Exit Sub
EH:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Backlog会議用ビュー作成でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "DelaxTools"
End Sub

Public Sub DxaBacklogCreateAssigneeLoad(ByVal control As Object)
    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub

    Dim headerRow As Long, dataFirstRow As Long, lastRow As Long, lastCol As Long, dateStartCol As Long
    If Not DxaBacklogDetectLayout(ws, headerRow, dataFirstRow, lastRow, lastCol, dateStartCol) Then Exit Sub

    Application.ScreenUpdating = False
    Application.StatusBar = "DelaxTools: Backlog担当者別負荷を作成しています..."

    Dim outWs As Worksheet
    Set outWs = DxaBacklogRecreateSheet(ws.Parent, "Backlog担当者別負荷")
    DxaBacklogWriteAssigneeLoad ws, outWs, dataFirstRow, lastRow

    outWs.Activate
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Backlog担当者別負荷を作成しました。", vbInformation, "DelaxTools"
    Exit Sub
EH:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Backlog担当者別負荷作成でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "DelaxTools"
End Sub

Private Function DxaBacklogDetectLayout(ByVal ws As Worksheet, ByRef headerRow As Long, ByRef dataFirstRow As Long, ByRef lastRow As Long, ByRef lastCol As Long, ByRef dateStartCol As Long) As Boolean
    Dim r As Long
    headerRow = 0
    For r = 1 To 20
        If InStr(1, CStr(ws.Cells(r, 1).Value), "キー", vbTextCompare) > 0 And _
           InStr(1, CStr(ws.Cells(r, 3).Value), "件名", vbTextCompare) > 0 Then
            headerRow = r
            Exit For
        End If
    Next

    If headerRow = 0 Then headerRow = 4
    dataFirstRow = headerRow + 1
    dateStartCol = 13

    lastRow = DxaBacklogLastUsedRow(ws)
    lastCol = DxaBacklogLastUsedCol(ws)
    If lastRow < dataFirstRow Or lastCol < 12 Then
        MsgBox "Backlogガント出力形式を判定できませんでした。A～L列に課題情報があるシートをアクティブにして実行してください。", vbExclamation, "DelaxTools"
        Exit Function
    End If

    If lastCol < dateStartCol Then lastCol = dateStartCol
    DxaBacklogDetectLayout = True
End Function

Private Function DxaBacklogLastUsedRow(ByVal ws As Worksheet) As Long
    Dim c As Range
    On Error Resume Next
    Set c = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0
    If c Is Nothing Then DxaBacklogLastUsedRow = 1 Else DxaBacklogLastUsedRow = c.Row
End Function

Private Function DxaBacklogLastUsedCol(ByVal ws As Worksheet) As Long
    Dim c As Range
    On Error Resume Next
    Set c = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0
    If c Is Nothing Then DxaBacklogLastUsedCol = 1 Else DxaBacklogLastUsedCol = c.Column
End Function

Private Sub DxaBacklogFormatIssueColumns(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal dataFirstRow As Long, ByVal lastRow As Long, ByVal lastCol As Long, ByVal dateStartCol As Long)
    With ws
        .Rows(headerRow).Font.Bold = True
        .Rows(headerRow).Interior.Color = RGB(217, 225, 242)
        .Range(.Cells(headerRow, 1), .Cells(headerRow, lastCol)).AutoFilter
        .Columns("A:A").ColumnWidth = 14
        .Columns("B:B").ColumnWidth = 9
        .Columns("C:C").ColumnWidth = 55
        .Columns("C:C").WrapText = True
        .Columns("D:F").ColumnWidth = 12
        .Columns("G:G").ColumnWidth = 14
        .Columns("H:I").ColumnWidth = 12
        .Columns("H:I").NumberFormatLocal = "yyyy/mm/dd"
        .Columns("J:K").ColumnWidth = 10
        .Columns("L:L").ColumnWidth = 10
        .Rows(dataFirstRow & ":" & lastRow).RowHeight = 24
        .Range(.Cells(headerRow, 1), .Cells(lastRow, 12)).Borders.LineStyle = xlContinuous
    End With
End Sub

Private Sub DxaBacklogFormatDateColumns(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal dataFirstRow As Long, ByVal lastRow As Long, ByVal lastCol As Long, ByVal dateStartCol As Long)
    Dim c As Long
    For c = dateStartCol To lastCol
        ws.Columns(c).ColumnWidth = 3.5
        Dim d As Date
        If DxaBacklogColumnDate(ws, c, headerRow, d) Then
            If Weekday(d, vbMonday) >= 6 Then
                ws.Range(ws.Cells(1, c), ws.Cells(lastRow, c)).Interior.Color = RGB(242, 242, 242)
            End If
            If DateValue(d) = Date Then
                ws.Range(ws.Cells(1, c), ws.Cells(lastRow, c)).Interior.Color = RGB(255, 242, 204)
                ws.Range(ws.Cells(1, c), ws.Cells(lastRow, c)).Borders(xlEdgeLeft).Weight = xlThick
                ws.Range(ws.Cells(1, c), ws.Cells(lastRow, c)).Borders(xlEdgeRight).Weight = xlThick
            End If
            If Day(d) = 1 Then
                ws.Range(ws.Cells(1, c), ws.Cells(lastRow, c)).Borders(xlEdgeLeft).Weight = xlMedium
            End If
        End If
    Next
    ws.Range(ws.Cells(1, dateStartCol), ws.Cells(headerRow, lastCol)).HorizontalAlignment = xlCenter
End Sub

Private Function DxaBacklogColumnDate(ByVal ws As Worksheet, ByVal colNo As Long, ByVal headerRow As Long, ByRef d As Date) As Boolean
    Dim r As Long
    For r = 1 To headerRow
        If IsDate(ws.Cells(r, colNo).Value) Then
            d = CDate(ws.Cells(r, colNo).Value)
            DxaBacklogColumnDate = True
            Exit Function
        End If
    Next
End Function

Private Sub DxaBacklogHighlightRows(ByVal ws As Worksheet, ByVal dataFirstRow As Long, ByVal lastRow As Long, ByVal lastCol As Long)
    Dim r As Long
    For r = dataFirstRow To lastRow
        If Len(Trim$(CStr(ws.Cells(r, 1).Value))) = 0 And Len(Trim$(CStr(ws.Cells(r, 3).Value))) = 0 Then GoTo ContinueRow

        Dim statusText As String
        statusText = CStr(ws.Cells(r, 12).Value)

        If DxaBacklogIsCompleted(statusText) Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 12)).Interior.Color = RGB(217, 217, 217)
        ElseIf DxaBacklogIsOverdue(ws.Cells(r, 9).Value, statusText) Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 12)).Interior.Color = RGB(255, 199, 206)
        ElseIf DxaBacklogIsDueWithin(ws.Cells(r, 9).Value, statusText, 3) Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 12)).Interior.Color = RGB(244, 176, 132)
        ElseIf DxaBacklogIsDueWithin(ws.Cells(r, 9).Value, statusText, 7) Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 12)).Interior.Color = RGB(255, 242, 204)
        End If

        If Len(Trim$(CStr(ws.Cells(r, 7).Value))) = 0 Then
            ws.Cells(r, 7).Interior.Color = RGB(255, 199, 206)
        End If
        If Not IsDate(ws.Cells(r, 8).Value) Then
            ws.Cells(r, 8).Interior.Color = RGB(255, 199, 206)
        End If
        If Not IsDate(ws.Cells(r, 9).Value) Then
            ws.Cells(r, 9).Interior.Color = RGB(255, 199, 206)
        ElseIf IsDate(ws.Cells(r, 8).Value) Then
            If CDate(ws.Cells(r, 9).Value) < CDate(ws.Cells(r, 8).Value) Then
                ws.Cells(r, 9).Interior.Color = RGB(255, 0, 0)
                ws.Cells(r, 9).Font.Color = RGB(255, 255, 255)
            End If
        End If
ContinueRow:
    Next
End Sub

Private Sub DxaBacklogFreezeGantt(ByVal ws As Worksheet, ByVal dataFirstRow As Long, ByVal dateStartCol As Long)
    On Error Resume Next
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Cells(dataFirstRow, dateStartCol).Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
End Sub

Private Function DxaBacklogIsCompleted(ByVal statusText As String) As Boolean
    Dim s As String
    s = Trim$(statusText)
    DxaBacklogIsCompleted = (InStr(1, s, "完了", vbTextCompare) > 0 Or InStr(1, s, "終了", vbTextCompare) > 0 Or InStr(1, s, "Closed", vbTextCompare) > 0)
End Function

Private Function DxaBacklogIsOverdue(ByVal dueValue As Variant, ByVal statusText As String) As Boolean
    If DxaBacklogIsCompleted(statusText) Then Exit Function
    If Not IsDate(dueValue) Then Exit Function
    DxaBacklogIsOverdue = (DateValue(CDate(dueValue)) < Date)
End Function

Private Function DxaBacklogIsDueWithin(ByVal dueValue As Variant, ByVal statusText As String, ByVal days As Long) As Boolean
    If DxaBacklogIsCompleted(statusText) Then Exit Function
    If Not IsDate(dueValue) Then Exit Function
    Dim d As Date
    d = DateValue(CDate(dueValue))
    DxaBacklogIsDueWithin = (d >= Date And d <= DateAdd("d", days, Date))
End Function

Private Function DxaBacklogDueStatus(ByVal dueValue As Variant, ByVal statusText As String) As String
    If DxaBacklogIsCompleted(statusText) Then
        DxaBacklogDueStatus = "完了"
    ElseIf Not IsDate(dueValue) Then
        DxaBacklogDueStatus = "期限日未設定"
    ElseIf DateValue(CDate(dueValue)) < Date Then
        DxaBacklogDueStatus = "期限超過"
    ElseIf DateValue(CDate(dueValue)) <= DateAdd("d", 3, Date) Then
        DxaBacklogDueStatus = "期限3日以内"
    ElseIf DateValue(CDate(dueValue)) <= DateAdd("d", 7, Date) Then
        DxaBacklogDueStatus = "期限7日以内"
    Else
        DxaBacklogDueStatus = "通常"
    End If
End Function

Private Function DxaBacklogOverdueDays(ByVal dueValue As Variant, ByVal statusText As String) As Long
    If Not DxaBacklogIsOverdue(dueValue, statusText) Then Exit Function
    DxaBacklogOverdueDays = DateDiff("d", DateValue(CDate(dueValue)), Date)
End Function

Private Function DxaBacklogAssignee(ByVal ws As Worksheet, ByVal rowNo As Long) As String
    DxaBacklogAssignee = Trim$(CStr(ws.Cells(rowNo, 7).Value))
    If Len(DxaBacklogAssignee) = 0 Then DxaBacklogAssignee = "未担当"
End Function

Private Function DxaBacklogRecreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets(sheetName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = sheetName
    Set DxaBacklogRecreateSheet = ws
End Function

Private Sub DxaBacklogWriteSummary(ByVal srcWs As Worksheet, ByVal outWs As Worksheet, ByVal dataFirstRow As Long, ByVal lastRow As Long)
    Dim total As Long, completed As Long, processing As Long, notStarted As Long, overdue As Long, due3 As Long, due7 As Long
    Dim noAssignee As Long, noStart As Long, noDue As Long, invalidDue As Long
    Dim statusCounts As Object
    Set statusCounts = CreateObject("Scripting.Dictionary")

    Dim r As Long
    For r = dataFirstRow To lastRow
        If Len(Trim$(CStr(srcWs.Cells(r, 1).Value))) = 0 And Len(Trim$(CStr(srcWs.Cells(r, 3).Value))) = 0 Then GoTo ContinueRow
        total = total + 1
        Dim st As String
        st = Trim$(CStr(srcWs.Cells(r, 12).Value))
        If Len(st) = 0 Then st = "状態未設定"
        If statusCounts.Exists(st) Then statusCounts(st) = CLng(statusCounts(st)) + 1 Else statusCounts.Add st, 1
        If DxaBacklogIsCompleted(st) Then completed = completed + 1
        If InStr(1, st, "処理中", vbTextCompare) > 0 Then processing = processing + 1
        If InStr(1, st, "未対応", vbTextCompare) > 0 Then notStarted = notStarted + 1
        If DxaBacklogIsOverdue(srcWs.Cells(r, 9).Value, st) Then overdue = overdue + 1
        If DxaBacklogIsDueWithin(srcWs.Cells(r, 9).Value, st, 3) Then due3 = due3 + 1
        If DxaBacklogIsDueWithin(srcWs.Cells(r, 9).Value, st, 7) Then due7 = due7 + 1
        If Len(Trim$(CStr(srcWs.Cells(r, 7).Value))) = 0 Then noAssignee = noAssignee + 1
        If Not IsDate(srcWs.Cells(r, 8).Value) Then noStart = noStart + 1
        If Not IsDate(srcWs.Cells(r, 9).Value) Then noDue = noDue + 1
        If IsDate(srcWs.Cells(r, 8).Value) And IsDate(srcWs.Cells(r, 9).Value) Then
            If CDate(srcWs.Cells(r, 9).Value) < CDate(srcWs.Cells(r, 8).Value) Then invalidDue = invalidDue + 1
        End If
ContinueRow:
    Next

    outWs.Range("A1:B1").Value = Array("項目", "件数")
    outWs.Range("A1:B1").Font.Bold = True
    outWs.Cells(2, 1).Value = "全課題数": outWs.Cells(2, 2).Value = total
    outWs.Cells(3, 1).Value = "完了": outWs.Cells(3, 2).Value = completed
    outWs.Cells(4, 1).Value = "処理中": outWs.Cells(4, 2).Value = processing
    outWs.Cells(5, 1).Value = "未対応": outWs.Cells(5, 2).Value = notStarted
    outWs.Cells(6, 1).Value = "期限超過": outWs.Cells(6, 2).Value = overdue
    outWs.Cells(7, 1).Value = "期限3日以内": outWs.Cells(7, 2).Value = due3
    outWs.Cells(8, 1).Value = "期限7日以内": outWs.Cells(8, 2).Value = due7
    outWs.Cells(9, 1).Value = "担当者未設定": outWs.Cells(9, 2).Value = noAssignee
    outWs.Cells(10, 1).Value = "開始日未設定": outWs.Cells(10, 2).Value = noStart
    outWs.Cells(11, 1).Value = "期限日未設定": outWs.Cells(11, 2).Value = noDue
    outWs.Cells(12, 1).Value = "期限日が開始日より前": outWs.Cells(12, 2).Value = invalidDue

    outWs.Range("D1:E1").Value = Array("状態", "件数")
    outWs.Range("D1:E1").Font.Bold = True
    Dim rowOut As Long
    rowOut = 2
    Dim k As Variant
    For Each k In statusCounts.Keys
        outWs.Cells(rowOut, 4).Value = CStr(k)
        outWs.Cells(rowOut, 5).Value = CLng(statusCounts(k))
        rowOut = rowOut + 1
    Next

    outWs.Columns("A:E").AutoFit
    outWs.Range("A1:E1").Interior.Color = RGB(217, 225, 242)
End Sub

Private Sub DxaBacklogWriteDelayList(ByVal srcWs As Worksheet, ByVal outWs As Worksheet, ByVal dataFirstRow As Long, ByVal lastRow As Long)
    outWs.Range("A1:I1").Value = Array("No", "課題キー", "件名", "担当者", "状態", "開始日", "期限日", "超過日数", "確認")
    outWs.Range("A1:I1").Font.Bold = True
    outWs.Range("A1:I1").Interior.Color = RGB(217, 225, 242)

    Dim rowOut As Long
    rowOut = 2
    Dim r As Long
    For r = dataFirstRow To lastRow
        Dim st As String
        st = CStr(srcWs.Cells(r, 12).Value)
        If DxaBacklogIsOverdue(srcWs.Cells(r, 9).Value, st) Then
            outWs.Cells(rowOut, 1).Value = rowOut - 1
            outWs.Cells(rowOut, 2).Value = srcWs.Cells(r, 1).Value
            outWs.Cells(rowOut, 3).Value = srcWs.Cells(r, 3).Value
            outWs.Cells(rowOut, 4).Value = DxaBacklogAssignee(srcWs, r)
            outWs.Cells(rowOut, 5).Value = st
            outWs.Cells(rowOut, 6).Value = srcWs.Cells(r, 8).Value
            outWs.Cells(rowOut, 7).Value = srcWs.Cells(r, 9).Value
            outWs.Cells(rowOut, 8).Value = DxaBacklogOverdueDays(srcWs.Cells(r, 9).Value, st)
            outWs.Cells(rowOut, 9).Value = "元行へ移動"
            On Error Resume Next
            outWs.Hyperlinks.Add Anchor:=outWs.Cells(rowOut, 9), Address:="", SubAddress:="'" & srcWs.Name & "'!A" & r, TextToDisplay:="元行へ移動"
            On Error GoTo 0
            rowOut = rowOut + 1
        End If
    Next

    If rowOut = 2 Then outWs.Cells(2, 1).Value = "期限超過課題はありません。"
    outWs.Columns("A:I").AutoFit
    outWs.Range("A1:I1").AutoFilter
End Sub

Private Sub DxaBacklogWriteMeetingView(ByVal srcWs As Worksheet, ByVal outWs As Worksheet, ByVal dataFirstRow As Long, ByVal lastRow As Long)
    outWs.Range("A1:J1").Value = Array("No", "課題キー", "件名", "担当者", "状態", "開始日", "期限日", "遅延状況", "予定時間", "実績時間")
    outWs.Range("A1:J1").Font.Bold = True
    outWs.Range("A1:J1").Interior.Color = RGB(217, 225, 242)

    Dim rowOut As Long
    rowOut = 2
    Dim r As Long
    For r = dataFirstRow To lastRow
        If Len(Trim$(CStr(srcWs.Cells(r, 1).Value))) = 0 And Len(Trim$(CStr(srcWs.Cells(r, 3).Value))) = 0 Then GoTo ContinueRow
        Dim st As String
        st = CStr(srcWs.Cells(r, 12).Value)
        If DxaBacklogIsCompleted(st) Then GoTo ContinueRow

        outWs.Cells(rowOut, 1).Value = rowOut - 1
        outWs.Cells(rowOut, 2).Value = srcWs.Cells(r, 1).Value
        outWs.Cells(rowOut, 3).Value = srcWs.Cells(r, 3).Value
        outWs.Cells(rowOut, 4).Value = DxaBacklogAssignee(srcWs, r)
        outWs.Cells(rowOut, 5).Value = st
        outWs.Cells(rowOut, 6).Value = srcWs.Cells(r, 8).Value
        outWs.Cells(rowOut, 7).Value = srcWs.Cells(r, 9).Value
        outWs.Cells(rowOut, 8).Value = DxaBacklogDueStatus(srcWs.Cells(r, 9).Value, st)
        outWs.Cells(rowOut, 9).Value = srcWs.Cells(r, 10).Value
        outWs.Cells(rowOut, 10).Value = srcWs.Cells(r, 11).Value
        On Error Resume Next
        outWs.Hyperlinks.Add Anchor:=outWs.Cells(rowOut, 2), Address:="", SubAddress:="'" & srcWs.Name & "'!A" & r, TextToDisplay:=CStr(srcWs.Cells(r, 1).Value)
        On Error GoTo 0
        rowOut = rowOut + 1
ContinueRow:
    Next

    If rowOut = 2 Then outWs.Cells(2, 1).Value = "会議用に表示する未完了課題はありません。"
    outWs.Columns("A:J").AutoFit
    outWs.Range("A1:J1").AutoFilter
End Sub

Private Sub DxaBacklogWriteAssigneeLoad(ByVal srcWs As Worksheet, ByVal outWs As Worksheet, ByVal dataFirstRow As Long, ByVal lastRow As Long)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim r As Long
    For r = dataFirstRow To lastRow
        If Len(Trim$(CStr(srcWs.Cells(r, 1).Value))) = 0 And Len(Trim$(CStr(srcWs.Cells(r, 3).Value))) = 0 Then GoTo ContinueRow
        Dim assignee As String
        assignee = DxaBacklogAssignee(srcWs, r)
        If Not dict.Exists(assignee) Then dict.Add assignee, Array(0#, 0#, 0#, 0#, 0#, 0#)

        Dim a As Variant
        a = dict(assignee)
        a(0) = CDbl(a(0)) + 1
        a(1) = CDbl(a(1)) + DxaBacklogToDouble(srcWs.Cells(r, 10).Value)
        a(2) = CDbl(a(2)) + DxaBacklogToDouble(srcWs.Cells(r, 11).Value)
        If Not DxaBacklogIsCompleted(CStr(srcWs.Cells(r, 12).Value)) Then a(3) = CDbl(a(3)) + 1
        If DxaBacklogIsOverdue(srcWs.Cells(r, 9).Value, CStr(srcWs.Cells(r, 12).Value)) Then a(4) = CDbl(a(4)) + 1
        If DxaBacklogIsDueWithin(srcWs.Cells(r, 9).Value, CStr(srcWs.Cells(r, 12).Value), 7) Then a(5) = CDbl(a(5)) + 1
        dict(assignee) = a
ContinueRow:
    Next

    outWs.Range("A1:G1").Value = Array("担当者", "課題数", "予定時間", "実績時間", "未完了", "期限超過", "期限7日以内")
    outWs.Range("A1:G1").Font.Bold = True
    outWs.Range("A1:G1").Interior.Color = RGB(217, 225, 242)

    Dim rowOut As Long
    rowOut = 2
    Dim k As Variant
    For Each k In dict.Keys
        Dim v As Variant
        v = dict(k)
        outWs.Cells(rowOut, 1).Value = CStr(k)
        outWs.Cells(rowOut, 2).Value = v(0)
        outWs.Cells(rowOut, 3).Value = v(1)
        outWs.Cells(rowOut, 4).Value = v(2)
        outWs.Cells(rowOut, 5).Value = v(3)
        outWs.Cells(rowOut, 6).Value = v(4)
        outWs.Cells(rowOut, 7).Value = v(5)
        rowOut = rowOut + 1
    Next

    outWs.Columns("A:G").AutoFit
    outWs.Range("A1:G1").AutoFilter
End Sub

Private Function DxaBacklogToDouble(ByVal v As Variant) As Double
    On Error GoTo EH
    If IsNumeric(v) Then DxaBacklogToDouble = CDbl(v)
    Exit Function
EH:
    DxaBacklogToDouble = 0
End Function

 ' ============================================================
' BITS勤怠表取得
' BITS Remote Time-card Systemへログインし、管理一覧から対象者の勤怠一覧へ遷移して、
' 日付・出勤時刻・退勤時刻を取得します。
' 認証情報はユーザー確認後に保存できます。保存時のパスワードはWindows DPAPIでユーザー単位に暗号化します。
' ============================================================
Public Sub DxaImportTimecardNormalWork(ByVal control As Object)
    ' 通常勤務：退勤時刻は 17:30～18:14 を 17:30 として扱い、18:15以降は15分単位で切り捨てます。
    DxaImportTimecardFromWebCore 2, "通常勤務"
End Sub

Public Sub DxaImportTimecardShiftWork(ByVal control As Object)
    ' シフト勤務：退勤時刻は常に15分単位で切り捨てます。
    DxaImportTimecardFromWebCore 1, "シフト勤務"
End Sub

Public Sub DxaConfigureTimecardTargetName(ByVal control As Object)
    ' 勤怠取得時に優先検索する対象者名を事前保存します。
    ' v157: 対象者設定画面から不要な説明枠とタイトルのDelaxTools表記を削除します。
    On Error GoTo EH

    DxaOpenTimecardTargetSettingsSheet
    Exit Sub

EH:
    MsgBox "勤怠取得対象者名の設定でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "DelaxTools"
End Sub

Public Sub DxaOpenTimecardTargetSettingsSheet()
    On Error GoTo EH

    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        Set wb = Workbooks.Add
    End If

    Dim ws As Worksheet
    Set ws = DxaGetOrCreateTimecardTargetSettingsSheet(wb)

    DxaRenderTimecardTargetSettingsSheet ws
    ws.Activate
    ws.Range("B8").Select
    Exit Sub

EH:
    MsgBox "勤怠取得対象者設定画面を開けませんでした。" & vbCrLf & Err.Description, vbExclamation, "DelaxTools"
End Sub

Public Sub DxaSaveTimecardTargetNamesFromSheet()
    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = DxaGetTimecardTargetSettingsSheet()
    If ws Is Nothing Then
        MsgBox "勤怠対象者設定シートが見つかりません。対象者設定ボタンから開き直してください。", vbExclamation, "DelaxTools"
        Exit Sub
    End If

    Dim targetNames As Collection
    Set targetNames = New Collection

    Dim r As Long
    For r = 8 To 57
        Call DxaAddTimecardTargetName(targetNames, CStr(ws.Cells(r, 2).Value))
    Next

    If targetNames.Count = 0 Then
        If MsgBox("対象者名が入力されていません。保存済み対象者名をすべて削除しますか？", vbQuestion + vbYesNo, "DelaxTools 勤怠対象者設定") = vbYes Then
            If DxaDeleteTimecardSavedTargetName() Then
                ws.Range("B8:B57").ClearContents
                MsgBox "保存済み対象者名を削除しました。", vbInformation, "DelaxTools"
            Else
                MsgBox "保存済み対象者名の削除に失敗しました。", vbExclamation, "DelaxTools"
            End If
        End If
        Exit Sub
    End If

    If DxaSaveTimecardSavedTargetName(DxaJoinTimecardTargetNames(targetNames)) Then
        DxaRenderTimecardTargetSettingsSheet ws
        MsgBox "勤怠取得対象者名を保存しました。" & vbCrLf & vbCrLf & DxaJoinTimecardTargetNamesForMessage(DxaJoinTimecardTargetNames(targetNames)), vbInformation, "DelaxTools"
    Else
        MsgBox "勤怠取得対象者名の保存に失敗しました。", vbExclamation, "DelaxTools"
    End If
    Exit Sub

EH:
    MsgBox "勤怠取得対象者名の保存でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "DelaxTools"
End Sub

Public Sub DxaClearTimecardTargetNamesFromSheet()
    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = DxaGetTimecardTargetSettingsSheet()
    If ws Is Nothing Then Exit Sub

    If MsgBox("保存済み対象者名をすべて削除しますか？", vbQuestion + vbYesNo, "DelaxTools 勤怠対象者設定") <> vbYes Then Exit Sub

    If DxaDeleteTimecardSavedTargetName() Then
        ws.Range("B8:B57").ClearContents
        MsgBox "保存済み対象者名を削除しました。", vbInformation, "DelaxTools"
    Else
        MsgBox "保存済み対象者名の削除に失敗しました。", vbExclamation, "DelaxTools"
    End If
    Exit Sub

EH:
    MsgBox "勤怠取得対象者名の削除でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "DelaxTools"
End Sub

Private Function DxaGetOrCreateTimecardTargetSettingsSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Name = "勤怠対象者設定" Then
            Set DxaGetOrCreateTimecardTargetSettingsSheet = ws
            Exit Function
        End If
    Next

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "勤怠対象者設定"
    Set DxaGetOrCreateTimecardTargetSettingsSheet = ws
End Function

Private Function DxaGetTimecardTargetSettingsSheet() As Worksheet
    On Error GoTo EH

    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Function

    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Name = "勤怠対象者設定" Then
            Set DxaGetTimecardTargetSettingsSheet = ws
            Exit Function
        End If
    Next
    Exit Function

EH:
    Set DxaGetTimecardTargetSettingsSheet = Nothing
End Function

Private Sub DxaRenderTimecardTargetSettingsSheet(ByVal ws As Worksheet)
    On Error GoTo EH

    Application.ScreenUpdating = False

    ws.Cells.Clear
    DxaDeleteTimecardSettingsShapes ws

    ws.Range("A1:F1").Merge
    ws.Range("A1").Value = "勤怠取得 対象者設定"
    With ws.Range("A1")
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 34

    ws.Range("A3:F5").Merge
    ws.Range("A3").Value = "通常勤務取得・シフト勤務取得で優先検索する対象者を登録します。" & vbCrLf & _
                           "B列に1名ずつ入力して、右側の［保存］を押してください。空白や重複は保存時に整理します。" & vbCrLf & _
                           "取得時は、登録済み対象者の中から番号で選択できます。"
    With ws.Range("A3")
        .Font.Size = 10
        .Font.Color = RGB(64, 64, 64)
        .Interior.Color = RGB(242, 246, 252)
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With

    ws.Range("A7").Value = "No"
    ws.Range("B7").Value = "対象者名"
    ws.Range("C7").Value = "メモ"
    ws.Range("A7:C7").Font.Bold = True
    ws.Range("A7:C7").Font.Color = RGB(255, 255, 255)
    ws.Range("A7:C7").Interior.Color = RGB(68, 114, 196)
    ws.Range("A7:C57").Borders.LineStyle = xlContinuous
    ws.Range("A7:C57").Borders.Color = RGB(210, 220, 235)

    Dim r As Long
    For r = 8 To 57
        ws.Cells(r, 1).Value = r - 7
        ws.Cells(r, 3).Value = ""
    Next

    Dim saved As Collection
    Set saved = DxaParseTimecardTargetNames(DxaLoadTimecardSavedTargetName())

    Dim i As Long
    For i = 1 To saved.Count
        If i > 50 Then Exit For
        ws.Cells(7 + i, 2).Value = CStr(saved(i))
    Next

    ws.Columns("A").ColumnWidth = 6
    ws.Columns("B").ColumnWidth = 28
    ws.Columns("C").ColumnWidth = 32
    ws.Columns("D").ColumnWidth = 3
    ws.Columns("E").ColumnWidth = 16
    ws.Columns("F").ColumnWidth = 16
    ws.Range("A8:A57").HorizontalAlignment = xlCenter
    ws.Range("B8:B57").Interior.Color = RGB(255, 255, 255)
    ws.Range("B8:B57").Font.Size = 11

    ws.Range("E7:F7").Merge
    ws.Range("E7").Value = "操作"
    ws.Range("E7").Font.Bold = True
    ws.Range("E7").Font.Color = RGB(255, 255, 255)
    ws.Range("E7").Interior.Color = RGB(68, 114, 196)
    ws.Range("E7").HorizontalAlignment = xlCenter

    DxaAddTimecardSettingsButton ws, "DxaTimecardSaveButton", "保存", ws.Range("E9").Left, ws.Range("E9").Top, 120, 30, "DxaSaveTimecardTargetNamesFromSheet", RGB(31, 78, 121), RGB(255, 255, 255)
    DxaAddTimecardSettingsButton ws, "DxaTimecardClearButton", "すべて削除", ws.Range("E12").Left, ws.Range("E12").Top, 120, 30, "DxaClearTimecardTargetNamesFromSheet", RGB(128, 128, 128), RGB(255, 255, 255)

    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Range("A7:C57").AutoFilter

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub

EH:
    Application.ScreenUpdating = True
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub DxaDeleteTimecardSettingsShapes(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Shapes("DxaTimecardSaveButton").Delete
    ws.Shapes("DxaTimecardClearButton").Delete
    On Error GoTo 0
End Sub

Private Sub DxaAddTimecardSettingsButton(ByVal ws As Worksheet, ByVal shapeName As String, ByVal caption As String, ByVal buttonLeft As Double, ByVal buttonTop As Double, ByVal buttonWidth As Double, ByVal buttonHeight As Double, ByVal macroName As String, ByVal fillColor As Long, ByVal fontColor As Long)
    On Error GoTo EH

    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(5, buttonLeft, buttonTop, buttonWidth, buttonHeight)
    shp.Name = shapeName
    shp.TextFrame.Characters.Text = caption
    shp.TextFrame.Characters.Font.Bold = True
    shp.TextFrame.Characters.Font.Size = 11
    shp.TextFrame.Characters.Font.Color = fontColor
    shp.TextFrame.HorizontalAlignment = xlHAlignCenter
    shp.TextFrame.VerticalAlignment = xlVAlignCenter
    shp.Fill.ForeColor.RGB = fillColor
    shp.Line.ForeColor.RGB = fillColor
    shp.OnAction = "'" & ThisWorkbook.Name & "'!" & macroName
    Exit Sub

EH:
    ' ボタン作成に失敗しても、設定シート自体は利用できるようにします。
End Sub

Public Sub DxaImportTimecardFromClipboard(ByVal control As Object)
    ' 互換用：旧ボタンから呼ばれた場合は通常勤務としてWeb取得します。
    DxaImportTimecardFromWebCore 2, "通常勤務"
End Sub

Private Sub DxaImportTimecardFromWebCore(ByVal endRoundMode As Long, ByVal workTypeLabel As String)
    On Error GoTo EH

    Dim email As String
    Dim password As String
    Dim loginUrl As String
    Dim managedListUrl As String
    Dim savedEmail As String
    Dim savedPassword As String
    Dim usedSavedCredentials As Boolean
    Dim savedAnswer As VbMsgBoxResult

    loginUrl = "https://staffsvc.bits.co.jp/AttendanceMgt/TimeScreen.jsp"
    managedListUrl = "https://staffsvc.bits.co.jp/AttendanceMgt/GetManagedList"

    If DxaLoadTimecardSavedCredentials(savedEmail, savedPassword) Then
        savedAnswer = MsgBox("保存済みのログイン情報を使用しますか？" & vbCrLf & vbCrLf & _
                             "メールアドレス: " & savedEmail, _
                             vbYesNoCancel + vbQuestion, "DelaxTools 勤怠取得")
        If savedAnswer = vbCancel Then Exit Sub
        If savedAnswer = vbYes Then
            email = savedEmail
            password = savedPassword
            usedSavedCredentials = True
        End If
    End If

    If Len(email) = 0 Then
        email = DxaPromptRequiredText("メールアドレスを入力してください。", "DelaxTools 勤怠取得", savedEmail)
        If Len(email) = 0 Then Exit Sub
    End If

    If Len(password) = 0 Then
        password = DxaPromptRequiredText("パスワードを入力してください。", "DelaxTools 勤怠取得")
        If Len(password) = 0 Then Exit Sub
    End If

    Application.StatusBar = "DelaxTools: 勤怠サイトへ接続しています..."

    Dim displayItems As Collection
    Set displayItems = DxaFetchTimecardDisplayItems(loginUrl, managedListUrl, email, password)

    If displayItems Is Nothing Or displayItems.Count = 0 Then
        MsgBox "勤怠一覧画面を取得できませんでした。" & vbCrLf & vbCrLf & _
               "メールアドレス、パスワード、サイト側の画面構成を確認してください。", vbExclamation, "DelaxTools"
        GoTo CleanExit
    End If

    Dim item As Variant
    Dim html As String
    Dim targetName As String
    Dim records As Collection
    Dim successCount As Long
    Dim failedText As String

    For Each item In displayItems
        targetName = CStr(item(0))
        html = CStr(item(2))

        If Len(Trim$(html)) = 0 Then
            failedText = failedText & targetName & "：勤怠一覧画面を取得できませんでした。" & vbCrLf
        Else
            Set records = DxaParseTimecardRecords(html)
            If records Is Nothing Or records.Count = 0 Then
                failedText = failedText & targetName & "：日付、出勤時刻、退勤時刻を検出できませんでした。" & vbCrLf
            Else
                DxaWriteTimecardRecords records, endRoundMode, targetName
                successCount = successCount + 1
            End If
        End If
    Next

    If successCount = 0 Then
        MsgBox "勤怠一覧画面から日付、出勤時刻、退勤時刻を検出できませんでした。" & vbCrLf & vbCrLf & _
               "対象者の勤怠一覧画面が表示されているか、サイト側のHTML構造を確認してください。" & vbCrLf & vbCrLf & _
               failedText, vbExclamation, "DelaxTools"
        GoTo CleanExit
    End If

    If Len(failedText) > 0 Then
        MsgBox "勤怠取得は一部完了しました。" & vbCrLf & vbCrLf & _
               "取得成功: " & CStr(successCount) & "名" & vbCrLf & vbCrLf & _
               "取得できなかった対象者:" & vbCrLf & failedText, vbExclamation, "DelaxTools"
    End If

    If Not usedSavedCredentials Then
        If MsgBox("今回入力したメールアドレスとパスワードをこのPCに保存しますか？" & vbCrLf & vbCrLf & _
                  "次回から入力を省略できます。" & vbCrLf & _
                  "パスワードはWindowsのユーザー単位で暗号化して保存します。", _
                  vbYesNo + vbQuestion, "DelaxTools 勤怠取得") = vbYes Then
            If DxaSaveTimecardCredentials(email, password) Then
                MsgBox "ログイン情報を保存しました。", vbInformation, "DelaxTools"
            Else
                MsgBox "ログイン情報の保存に失敗しました。" & vbCrLf & _
                       "勤怠取得の結果は出力済みです。", vbExclamation, "DelaxTools"
            End If
        End If
    End If

CleanExit:
    Application.StatusBar = False
    Exit Sub

EH:
    Application.StatusBar = False
    MsgBox "勤怠表取得でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "DelaxTools"
End Sub

Private Function DxaPromptRequiredText(ByVal prompt As String, ByVal title As String, Optional ByVal defaultValue As String = "") As String
    Dim s As String
    s = InputBox(prompt, title, defaultValue)
    s = Trim$(s)
    DxaPromptRequiredText = s
End Function

Private Function DxaLoadTimecardSavedCredentials(ByRef email As String, ByRef password As String) As Boolean
    On Error GoTo EH

    Dim path As String
    path = DxaTimecardCredentialFilePath()
    If Len(path) = 0 Then Exit Function
    If Len(Dir$(path)) = 0 Then Exit Function

    Dim emailEncoded As String
    Dim passwordEncoded As String
    Dim lineText As String
    Dim ff As Integer

    ff = FreeFile
    Open path For Input As #ff
    Do Until EOF(ff)
        Line Input #ff, lineText
        If Left$(lineText, 6) = "Email=" Then
            emailEncoded = Mid$(lineText, 7)
        ElseIf Left$(lineText, 9) = "Password=" Then
            passwordEncoded = Mid$(lineText, 10)
        End If
    Loop
    Close #ff

    If Len(emailEncoded) = 0 Or Len(passwordEncoded) = 0 Then Exit Function

    Dim emailBytes() As Byte
    Dim protectedPasswordBytes() As Byte
    Dim passwordBytes() As Byte

    emailBytes = DxaBase64ToBytes(emailEncoded)
    protectedPasswordBytes = DxaBase64ToBytes(passwordEncoded)
    passwordBytes = DxaDpapiUnprotectBytes(protectedPasswordBytes)

    email = DxaUtf16BytesToString(emailBytes)
    password = DxaUtf16BytesToString(passwordBytes)

    DxaLoadTimecardSavedCredentials = (Len(email) > 0 And Len(password) > 0)
    Exit Function

EH:
    On Error Resume Next
    If ff <> 0 Then Close #ff
    DxaLoadTimecardSavedCredentials = False
End Function

Private Function DxaSaveTimecardCredentials(ByVal email As String, ByVal password As String) As Boolean
    On Error GoTo EH

    Dim path As String
    path = DxaTimecardCredentialFilePath()
    If Len(path) = 0 Then Exit Function

    Dim emailBytes() As Byte
    Dim passwordBytes() As Byte
    Dim protectedPasswordBytes() As Byte

    emailBytes = DxaStringToUtf16Bytes(email)
    passwordBytes = DxaStringToUtf16Bytes(password)
    protectedPasswordBytes = DxaDpapiProtectBytes(passwordBytes)

    If DxaByteArrayLength(emailBytes) = 0 Then Exit Function
    If DxaByteArrayLength(protectedPasswordBytes) = 0 Then Exit Function

    Dim ff As Integer
    ff = FreeFile
    Open path For Output As #ff
    Print #ff, "Version=1"
    Print #ff, "Email=" & DxaBytesToBase64(emailBytes)
    Print #ff, "Password=" & DxaBytesToBase64(protectedPasswordBytes)
    Close #ff

    DxaSaveTimecardCredentials = True
    Exit Function

EH:
    On Error Resume Next
    If ff <> 0 Then Close #ff
    DxaSaveTimecardCredentials = False
End Function

Private Function DxaTimecardCredentialFilePath() As String
    On Error GoTo EH

    Dim appData As String
    appData = Environ$("APPDATA")
    If Len(appData) = 0 Then Exit Function

    Dim folderPath As String
    folderPath = appData & "\DelaxTools"
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then MkDir folderPath

    DxaTimecardCredentialFilePath = folderPath & "\timecard_credentials.dat"
    Exit Function

EH:
    DxaTimecardCredentialFilePath = ""
End Function

Private Function DxaLoadTimecardSavedTargetName() As String
    On Error GoTo EH

    Dim path As String
    path = DxaTimecardTargetNameFilePath()
    If Len(path) = 0 Then Exit Function
    If Len(Dir$(path)) = 0 Then Exit Function

    Dim targetNames As Collection
    Set targetNames = New Collection

    Dim lineText As String
    Dim ff As Integer

    ff = FreeFile
    Open path For Input As #ff
    Do Until EOF(ff)
        Line Input #ff, lineText
        If Left$(lineText, 11) = "TargetName=" Then
            Call DxaAddTimecardTargetName(targetNames, DxaUtf16BytesToString(DxaBase64ToBytes(Mid$(lineText, 12))))
        ElseIf Left$(lineText, 10) = "TargetName" And InStr(1, lineText, "=", vbTextCompare) > 0 Then
            Call DxaAddTimecardTargetName(targetNames, DxaUtf16BytesToString(DxaBase64ToBytes(Mid$(lineText, InStr(1, lineText, "=", vbTextCompare) + 1))))
        End If
    Loop
    Close #ff

    DxaLoadTimecardSavedTargetName = DxaJoinTimecardTargetNames(targetNames)
    Exit Function

EH:
    On Error Resume Next
    If ff <> 0 Then Close #ff
    DxaLoadTimecardSavedTargetName = ""
End Function

Private Function DxaSaveTimecardSavedTargetName(ByVal targetName As String) As Boolean
    On Error GoTo EH

    Dim targetNames As Collection
    Set targetNames = DxaParseTimecardTargetNames(targetName)
    If targetNames.Count = 0 Then Exit Function

    Dim path As String
    path = DxaTimecardTargetNameFilePath()
    If Len(path) = 0 Then Exit Function

    Dim ff As Integer
    ff = FreeFile
    Open path For Output As #ff
    Print #ff, "Version=2"

    Dim i As Long
    For i = 1 To targetNames.Count
        Dim targetBytes() As Byte
        targetBytes = DxaStringToUtf16Bytes(CStr(targetNames(i)))
        If DxaByteArrayLength(targetBytes) > 0 Then
            Print #ff, "TargetName" & CStr(i) & "=" & DxaBytesToBase64(targetBytes)
        End If
    Next

    Close #ff

    DxaSaveTimecardSavedTargetName = True
    Exit Function

EH:
    On Error Resume Next
    If ff <> 0 Then Close #ff
    DxaSaveTimecardSavedTargetName = False
End Function

Private Function DxaDeleteTimecardSavedTargetName() As Boolean
    On Error GoTo EH

    Dim path As String
    path = DxaTimecardTargetNameFilePath()
    If Len(path) = 0 Then Exit Function
    If Len(Dir$(path)) > 0 Then Kill path

    DxaDeleteTimecardSavedTargetName = True
    Exit Function

EH:
    DxaDeleteTimecardSavedTargetName = False
End Function

Private Function DxaParseTimecardTargetNames(ByVal targetText As String) As Collection
    On Error GoTo EH

    Dim result As Collection
    Set result = New Collection

    Dim s As String
    s = CStr(targetText)
    s = Replace$(s, vbCrLf, vbLf)
    s = Replace$(s, vbCr, vbLf)
    s = Replace$(s, "、", vbLf)
    s = Replace$(s, "，", vbLf)
    s = Replace$(s, ",", vbLf)
    s = Replace$(s, "；", vbLf)
    s = Replace$(s, ";", vbLf)
    s = Replace$(s, "／", vbLf)
    s = Replace$(s, "/", vbLf)
    s = Replace$(s, "｜", vbLf)
    s = Replace$(s, "|", vbLf)

    Dim parts As Variant
    parts = Split(s, vbLf)

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Call DxaAddTimecardTargetName(result, CStr(parts(i)))
    Next

    Set DxaParseTimecardTargetNames = result
    Exit Function

EH:
    Set DxaParseTimecardTargetNames = New Collection
End Function

Private Sub DxaAddTimecardTargetName(ByVal targetNames As Collection, ByVal targetName As String)
    On Error Resume Next

    Dim normalizedName As String
    normalizedName = DxaNormalizeSpaces(Trim$(targetName))
    If Len(normalizedName) = 0 Then Exit Sub

    Dim newKey As String
    newKey = DxaNormalizeNameForMatch(normalizedName)

    Dim i As Long
    For i = 1 To targetNames.Count
        If DxaNormalizeNameForMatch(CStr(targetNames(i))) = newKey Then Exit Sub
    Next

    targetNames.Add normalizedName
End Sub

Private Function DxaJoinTimecardTargetNames(ByVal targetNames As Collection) As String
    On Error GoTo EH

    Dim s As String
    Dim i As Long
    For i = 1 To targetNames.Count
        If Len(s) > 0 Then s = s & vbCrLf
        s = s & CStr(targetNames(i))
    Next

    DxaJoinTimecardTargetNames = s
    Exit Function

EH:
    DxaJoinTimecardTargetNames = ""
End Function

Private Function DxaJoinTimecardTargetNamesInline(ByVal targetNamesText As String) As String
    On Error GoTo EH

    Dim targetNames As Collection
    Set targetNames = DxaParseTimecardTargetNames(targetNamesText)

    Dim s As String
    Dim i As Long
    For i = 1 To targetNames.Count
        If Len(s) > 0 Then s = s & ", "
        s = s & CStr(targetNames(i))
    Next

    DxaJoinTimecardTargetNamesInline = s
    Exit Function

EH:
    DxaJoinTimecardTargetNamesInline = ""
End Function

Private Function DxaJoinTimecardTargetNamesForMessage(ByVal targetNamesText As String) As String
    On Error GoTo EH

    Dim targetNames As Collection
    Set targetNames = DxaParseTimecardTargetNames(targetNamesText)

    Dim s As String
    Dim i As Long
    For i = 1 To targetNames.Count
        If Len(s) > 0 Then s = s & vbCrLf
        s = s & CStr(i) & ": " & CStr(targetNames(i))
    Next

    DxaJoinTimecardTargetNamesForMessage = s
    Exit Function

EH:
    DxaJoinTimecardTargetNamesForMessage = ""
End Function

Private Function DxaTimecardTargetNameCount(ByVal targetNamesText As String) As Long
    On Error GoTo EH

    Dim targetNames As Collection
    Set targetNames = DxaParseTimecardTargetNames(targetNamesText)
    DxaTimecardTargetNameCount = targetNames.Count
    Exit Function

EH:
    DxaTimecardTargetNameCount = 0
End Function

Private Function DxaTimecardTargetNameFilePath() As String
    On Error GoTo EH

    Dim appData As String
    appData = Environ$("APPDATA")
    If Len(appData) = 0 Then Exit Function

    Dim folderPath As String
    folderPath = appData & "\DelaxTools"
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then MkDir folderPath

    DxaTimecardTargetNameFilePath = folderPath & "\timecard_target_name.dat"
    Exit Function

EH:
    DxaTimecardTargetNameFilePath = ""
End Function

Private Function DxaStringToUtf16Bytes(ByVal valueText As String) As Byte()
    Dim bytes() As Byte
    Dim byteCount As Long

    byteCount = LenB(valueText)
    If byteCount <= 0 Then Exit Function

    ReDim bytes(0 To byteCount - 1)
    DxaCopyMemory VarPtr(bytes(0)), StrPtr(valueText), byteCount
    DxaStringToUtf16Bytes = bytes
End Function

Private Function DxaUtf16BytesToString(ByRef bytes() As Byte) As String
    On Error GoTo EH

    Dim byteCount As Long
    byteCount = DxaByteArrayLength(bytes)
    If byteCount <= 0 Then Exit Function

    Dim charCount As Long
    charCount = byteCount \ 2
    If charCount <= 0 Then Exit Function

    Dim valueText As String
    valueText = String$(charCount, vbNullChar)
    DxaCopyMemory StrPtr(valueText), VarPtr(bytes(LBound(bytes))), byteCount
    DxaUtf16BytesToString = valueText
    Exit Function

EH:
    DxaUtf16BytesToString = ""
End Function

Private Function DxaDpapiProtectBytes(ByRef plainBytes() As Byte) As Byte()
    On Error GoTo EH

    Dim byteCount As Long
    byteCount = DxaByteArrayLength(plainBytes)
    If byteCount <= 0 Then Exit Function

    Dim inBlob As DxaDataBlob
    Dim outBlob As DxaDataBlob
    Dim resultBytes() As Byte

    inBlob.cbData = byteCount
    inBlob.pbData = VarPtr(plainBytes(LBound(plainBytes)))

    If DxaCryptProtectData(inBlob, 0, 0, 0, 0, 0, outBlob) <> 0 Then
        If outBlob.cbData > 0 Then
            ReDim resultBytes(0 To outBlob.cbData - 1)
            DxaCopyMemory VarPtr(resultBytes(0)), outBlob.pbData, outBlob.cbData
            DxaDpapiProtectBytes = resultBytes
        End If
        If outBlob.pbData <> 0 Then Call DxaLocalFree(outBlob.pbData)
    End If
    Exit Function

EH:
    On Error Resume Next
    If outBlob.pbData <> 0 Then Call DxaLocalFree(outBlob.pbData)
End Function

Private Function DxaDpapiUnprotectBytes(ByRef protectedBytes() As Byte) As Byte()
    On Error GoTo EH

    Dim byteCount As Long
    byteCount = DxaByteArrayLength(protectedBytes)
    If byteCount <= 0 Then Exit Function

    Dim inBlob As DxaDataBlob
    Dim outBlob As DxaDataBlob
    Dim resultBytes() As Byte

    inBlob.cbData = byteCount
    inBlob.pbData = VarPtr(protectedBytes(LBound(protectedBytes)))

    If DxaCryptUnprotectData(inBlob, 0, 0, 0, 0, 0, outBlob) <> 0 Then
        If outBlob.cbData > 0 Then
            ReDim resultBytes(0 To outBlob.cbData - 1)
            DxaCopyMemory VarPtr(resultBytes(0)), outBlob.pbData, outBlob.cbData
            DxaDpapiUnprotectBytes = resultBytes
        End If
        If outBlob.pbData <> 0 Then Call DxaLocalFree(outBlob.pbData)
    End If
    Exit Function

EH:
    On Error Resume Next
    If outBlob.pbData <> 0 Then Call DxaLocalFree(outBlob.pbData)
End Function

Private Function DxaBytesToBase64(ByRef bytes() As Byte) As String
    On Error GoTo EH

    If DxaByteArrayLength(bytes) = 0 Then Exit Function

    Dim dom As Object
    Dim node As Object

    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    Set node = dom.createElement("b64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = bytes
    DxaBytesToBase64 = Replace$(Replace$(node.Text, vbCr, ""), vbLf, "")
    Exit Function

EH:
    DxaBytesToBase64 = ""
End Function

Private Function DxaBase64ToBytes(ByVal base64Text As String) As Byte()
    On Error GoTo EH

    Dim dom As Object
    Dim node As Object
    Dim bytes() As Byte

    If Len(base64Text) = 0 Then Exit Function

    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    Set node = dom.createElement("b64")
    node.DataType = "bin.base64"
    node.Text = base64Text
    bytes = node.nodeTypedValue
    DxaBase64ToBytes = bytes
    Exit Function

EH:
End Function

Private Function DxaByteArrayLength(ByRef bytes() As Byte) As Long
    On Error GoTo EH
    DxaByteArrayLength = UBound(bytes) - LBound(bytes) + 1
    Exit Function
EH:
    DxaByteArrayLength = 0
End Function

Private Sub DxaWriteTimecardRecords(ByVal records As Collection, ByVal endRoundMode As Long, Optional ByVal targetName As String = "")
    On Error GoTo EH

    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Set wb = Workbooks.Add

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim ws As Worksheet
    Dim sheetName As String
    sheetName = DxaBuildTimecardSheetName(targetName)

    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo EH
    If Not ws Is Nothing Then ws.Delete

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = sheetName

    Application.DisplayAlerts = True

    ws.Range("A1:E1").Value = Array("日付", "出勤時刻", "退勤時刻", "取得出勤時刻", "取得退勤時刻")
    ws.Range("A1:E1").Font.Bold = True
    ws.Range("A1:E1").Interior.Color = RGB(217, 225, 242)
    ws.Columns("A:E").NumberFormatLocal = "@"

    Dim r As Long
    r = 2

    Dim rec As Variant
    Dim rawStart As String
    Dim rawEnd As String
    Dim roundedStart As String
    Dim roundedEnd As String

    For Each rec In records
        rawStart = CStr(rec(1))
        rawEnd = CStr(rec(2))
        roundedStart = DxaRoundTimecardStart(rawStart)
        roundedEnd = DxaRoundTimecardEnd(rawEnd, endRoundMode)

        ws.Cells(r, 1).Value = CStr(rec(0))
        ws.Cells(r, 2).Value = DxaTimecardOutputText(roundedStart)
        ws.Cells(r, 3).Value = DxaTimecardOutputText(roundedEnd)
        ws.Cells(r, 4).Value = DxaTimecardOutputText(rawStart)
        ws.Cells(r, 5).Value = DxaTimecardOutputText(rawEnd)
        r = r + 1
    Next

    ws.Columns("A:E").AutoFit
    ws.Range("A1:E1").AutoFilter
    ws.Activate
    ws.Range("A1").Select

CleanExit:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

EH:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function DxaFetchTimecardDisplayItems(ByVal timeScreenUrl As String, ByVal managedListUrl As String, ByVal email As String, ByVal password As String) As Collection
    On Error GoTo EH

    Dim cookies As String
    Dim firstHtml As String
    Dim loginHtml As String
    Dim loginPageUrl As String
    Dim loginAction As String
    Dim loginMethod As String
    Dim loginBody As String
    Dim loggedInHtml As String
    Dim timeScreenHtml As String
    Dim managerAction As String
    Dim managerBody As String
    Dim managerHtml As String
    Dim listHtml As String

    firstHtml = DxaHttpRequest("GET", timeScreenUrl, "", cookies, "")

    If DxaLooksLikeManagedList(firstHtml) Then
        listHtml = firstHtml
    ElseIf DxaLooksLikeTimeScreen(firstHtml) Then
        timeScreenHtml = firstHtml
    Else
        loginHtml = DxaResolveTimecardLoginHtml(timeScreenUrl, firstHtml, cookies, loginPageUrl)

        If Len(loginHtml) > 0 Then
            loginAction = DxaExtractFormAction(loginHtml)
            If Len(loginAction) = 0 Then loginAction = loginPageUrl
            loginAction = DxaResolveUrl(loginPageUrl, loginAction)

            loginMethod = DxaExtractFormMethod(loginHtml)
            If Len(loginMethod) = 0 Then loginMethod = "POST"

            loginBody = DxaBuildLoginPostBody(loginHtml, email, password)
            loggedInHtml = DxaSubmitTimecardLoginForm(timeScreenUrl, managedListUrl, loginPageUrl, loginHtml, loginMethod, loginAction, loginBody, cookies)

            If DxaLooksLikeManagedList(loggedInHtml) Then
                listHtml = loggedInHtml
            ElseIf DxaLooksLikeTimeScreen(loggedInHtml) Then
                timeScreenHtml = loggedInHtml
            ElseIf Len(DxaTrimHtml(loggedInHtml)) > 0 And Not DxaLooksLikeSessionError(loggedInHtml) And Not DxaLooksLikeLoginForm(loggedInHtml) Then
                timeScreenHtml = DxaHttpRequest("GET", timeScreenUrl, "", cookies, "")
            End If
        Else
            loggedInHtml = firstHtml
        End If
    End If

    If Len(listHtml) = 0 Then
        If Len(timeScreenHtml) = 0 Or Not DxaLooksLikeTimeScreen(timeScreenHtml) Then
            timeScreenHtml = DxaHttpRequest("GET", timeScreenUrl, "", cookies, "")
        End If

        If DxaLooksLikeManagedList(timeScreenHtml) Then
            listHtml = timeScreenHtml
        End If
    End If

    If Len(listHtml) = 0 Or Not DxaLooksLikeManagedList(listHtml) Then
        If Len(timeScreenHtml) = 0 Or Not DxaLooksLikeTimeScreen(timeScreenHtml) Then
            timeScreenHtml = DxaHttpRequest("GET", timeScreenUrl, "", cookies, "")
        End If

        If DxaLooksLikeManagedList(timeScreenHtml) Then
            listHtml = timeScreenHtml
        ElseIf DxaLooksLikeTimeScreen(timeScreenHtml) Then
            managerAction = managedListUrl
            Dim cookiesBeforeManager As String
            cookiesBeforeManager = cookies
            managerHtml = DxaSubmitTimecardManagerButton(timeScreenUrl, managedListUrl, timeScreenHtml, cookies)

            If DxaLooksLikeManagedList(managerHtml) Then
                listHtml = managerHtml
            Else
                cookies = cookiesBeforeManager
                listHtml = DxaHttpRequest("GET", managedListUrl, "", cookies, "", timeScreenUrl)
            End If
        End If
    End If

    If Len(listHtml) = 0 Or Not DxaLooksLikeManagedList(listHtml) Then
        Dim debugPath As String
        debugPath = DxaSaveTimecardDebugHtml("get_managed_list", listHtml)
        If Len(firstHtml) > 0 Then Call DxaSaveTimecardDebugHtml("first_access", firstHtml)
        If Len(loginHtml) > 0 Then Call DxaSaveTimecardDebugHtml("login_form", loginHtml)
        If Len(loggedInHtml) > 0 Then Call DxaSaveTimecardDebugHtml("login_post_result", loggedInHtml)
        If Len(timeScreenHtml) > 0 Then Call DxaSaveTimecardDebugHtml("time_screen_after_login", timeScreenHtml)
        If Len(managerHtml) > 0 Then Call DxaSaveTimecardDebugHtml("manager_post_result", managerHtml)
        Err.Raise vbObjectError + 621, "DelaxTools", "管理一覧画面を取得できませんでした。" & vbCrLf & _
                 "処理順は TimeScreen.jsp → ログイン画面 → ログイン → TimeScreen.jsp → GetManagedList です。" & vbCrLf & _
                 "確認用HTMLを複数保存しました。first_access / login_script / login_form / login_post_result / time_screen_after_login / manager_post_result を確認してください。" & vbCrLf & _
                 "確認用HTML: " & debugPath
    End If

    Dim candidates As Collection
    Set candidates = DxaParseTimecardUserCandidates(listHtml)

    If candidates Is Nothing Or candidates.Count = 0 Then
        Dim debugListPath As String
        debugListPath = DxaSaveTimecardDebugHtml("managed_list", listHtml)
        Err.Raise vbObjectError + 622, "DelaxTools", "管理一覧から勤怠取得対象者を検出できませんでした。" & vbCrLf & _
                 "管理一覧画面は取得できていますが、対象者ボタンのHTML構成が想定と異なる可能性があります。" & vbCrLf & _
                 "確認用HTML: " & debugListPath
    End If

    Dim selectedUsers As Collection
    Set selectedUsers = DxaResolveTimecardSelectedUsers(candidates)
    If selectedUsers Is Nothing Or selectedUsers.Count = 0 Then
        Set DxaFetchTimecardDisplayItems = New Collection
        Exit Function
    End If

    Dim result As Collection
    Set result = New Collection

    Dim item As Variant
    Dim targetName As String
    Dim userValue As String
    Dim html As String
    For Each item In selectedUsers
        targetName = CStr(item(0))
        userValue = CStr(item(1))
        Application.StatusBar = "DelaxTools: " & targetName & " さんの勤怠一覧を取得しています..."
        html = DxaFetchTimecardDisplayHtmlByUserValue(managedListUrl, cookies, userValue, listHtml)
        result.Add Array(targetName, userValue, html)
    Next

    Set DxaFetchTimecardDisplayItems = result
    Exit Function

EH:
    Set DxaFetchTimecardDisplayItems = New Collection
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function DxaFetchTimecardDisplayHtml(ByVal timeScreenUrl As String, ByVal managedListUrl As String, ByVal email As String, ByVal password As String) As String
    On Error GoTo EH

    Dim cookies As String
    Dim firstHtml As String
    Dim loginHtml As String
    Dim loginPageUrl As String
    Dim loginAction As String
    Dim loginMethod As String
    Dim loginBody As String
    Dim loggedInHtml As String
    Dim timeScreenHtml As String
    Dim managerAction As String
    Dim managerBody As String
    Dim managerHtml As String
    Dim listHtml As String

    ' v152:
    ' 1. まず指定URL TimeScreen.jsp にアクセスします。
    ' 2. セッション切れ画面が返った場合は、ログインボタン相当のGET/POST、またはLogin系URLを試してログインフォームを取得します。
    ' 3. ログインフォームのhidden項目は保持しつつ、CHECK_SESSION_ID等のhidden値をメールアドレスで上書きしないようにします。
    ' 4. ログイン後はTimeScreen.jspを再取得し、管理ボタン押下後の遷移先であるGetManagedListを取得します。
    firstHtml = DxaHttpRequest("GET", timeScreenUrl, "", cookies, "")

    If DxaLooksLikeManagedList(firstHtml) Then
        listHtml = firstHtml
    ElseIf DxaLooksLikeTimeScreen(firstHtml) Then
        timeScreenHtml = firstHtml
    Else
        loginHtml = DxaResolveTimecardLoginHtml(timeScreenUrl, firstHtml, cookies, loginPageUrl)

        If Len(loginHtml) > 0 Then
            loginAction = DxaExtractFormAction(loginHtml)
            If Len(loginAction) = 0 Then loginAction = loginPageUrl
            loginAction = DxaResolveUrl(loginPageUrl, loginAction)

            loginMethod = DxaExtractFormMethod(loginHtml)
            If Len(loginMethod) = 0 Then loginMethod = "POST"

            loginBody = DxaBuildLoginPostBody(loginHtml, email, password)
            loggedInHtml = DxaSubmitTimecardLoginForm(timeScreenUrl, managedListUrl, loginPageUrl, loginHtml, loginMethod, loginAction, loginBody, cookies)

            If DxaLooksLikeManagedList(loggedInHtml) Then
                listHtml = loggedInHtml
            ElseIf DxaLooksLikeTimeScreen(loggedInHtml) Then
                timeScreenHtml = loggedInHtml
            ElseIf Len(DxaTrimHtml(loggedInHtml)) > 0 And Not DxaLooksLikeSessionError(loggedInHtml) And Not DxaLooksLikeLoginForm(loggedInHtml) Then
                timeScreenHtml = DxaHttpRequest("GET", timeScreenUrl, "", cookies, "")
            End If
        Else
            loggedInHtml = firstHtml
        End If
    End If

    ' ログイン後は、TimeScreen.jspを再取得してから管理一覧へ進みます。
    If Len(listHtml) = 0 Then
        If Len(timeScreenHtml) = 0 Or Not DxaLooksLikeTimeScreen(timeScreenHtml) Then
            timeScreenHtml = DxaHttpRequest("GET", timeScreenUrl, "", cookies, "")
        End If

        If DxaLooksLikeManagedList(timeScreenHtml) Then
            listHtml = timeScreenHtml
        End If
    End If

    If Len(listHtml) = 0 Or Not DxaLooksLikeManagedList(listHtml) Then
        ' v152: Do not call GetManagedList by direct GET before simulating the manager button.
        ' The real browser flow is TimeScreen.jsp -> POST mainForm to GetManagedList.
        If Len(timeScreenHtml) = 0 Or Not DxaLooksLikeTimeScreen(timeScreenHtml) Then
            timeScreenHtml = DxaHttpRequest("GET", timeScreenUrl, "", cookies, "")
        End If

        If DxaLooksLikeManagedList(timeScreenHtml) Then
            listHtml = timeScreenHtml
        ElseIf DxaLooksLikeTimeScreen(timeScreenHtml) Then
            managerAction = managedListUrl
            Dim cookiesBeforeManager As String
            cookiesBeforeManager = cookies
            managerHtml = DxaSubmitTimecardManagerButton(timeScreenUrl, managedListUrl, timeScreenHtml, cookies)

            If DxaLooksLikeManagedList(managerHtml) Then
                listHtml = managerHtml
            Else
                ' Last fallback only. Restore the cookie state before trying direct GET.
                cookies = cookiesBeforeManager
                listHtml = DxaHttpRequest("GET", managedListUrl, "", cookies, "", timeScreenUrl)
            End If
        End If
    End If

    If Len(listHtml) = 0 Or Not DxaLooksLikeManagedList(listHtml) Then
        Dim debugPath As String
        debugPath = DxaSaveTimecardDebugHtml("get_managed_list", listHtml)
        If Len(firstHtml) > 0 Then Call DxaSaveTimecardDebugHtml("first_access", firstHtml)
        If Len(loginHtml) > 0 Then Call DxaSaveTimecardDebugHtml("login_form", loginHtml)
        If Len(loggedInHtml) > 0 Then Call DxaSaveTimecardDebugHtml("login_post_result", loggedInHtml)
        If Len(timeScreenHtml) > 0 Then Call DxaSaveTimecardDebugHtml("time_screen_after_login", timeScreenHtml)
        If Len(managerHtml) > 0 Then Call DxaSaveTimecardDebugHtml("manager_post_result", managerHtml)
        Err.Raise vbObjectError + 621, "DelaxTools", "管理一覧画面を取得できませんでした。" & vbCrLf & _
                 "処理順は TimeScreen.jsp → ログイン画面 → ログイン → TimeScreen.jsp → GetManagedList です。" & vbCrLf & _
                 "v157では対象者設定画面を簡素化しています。" & vbCrLf & _
                 "確認用HTMLを複数保存しました。first_access / login_script / login_form / login_post_result / time_screen_after_login / manager_post_result を確認してください。" & vbCrLf & _
                 "確認用HTML: " & debugPath
    End If

    Dim candidates As Collection
    Set candidates = DxaParseTimecardUserCandidates(listHtml)

    If candidates Is Nothing Or candidates.Count = 0 Then
        Dim debugListPath As String
        debugListPath = DxaSaveTimecardDebugHtml("managed_list", listHtml)
        Err.Raise vbObjectError + 622, "DelaxTools", "管理一覧から勤怠取得対象者を検出できませんでした。" & vbCrLf & _
                 "管理一覧画面は取得できていますが、対象者ボタンのHTML構成が想定と異なる可能性があります。" & vbCrLf & _
                 "確認用HTML: " & debugListPath
    End If

    Dim userValue As String
    userValue = DxaPromptTimecardUserSelection(candidates)
    If Len(userValue) = 0 Then
        DxaFetchTimecardDisplayHtml = ""
        Exit Function
    End If

    DxaFetchTimecardDisplayHtml = DxaFetchTimecardDisplayHtmlByUserValue(managedListUrl, cookies, userValue, listHtml)
    Exit Function

EH:
    DxaFetchTimecardDisplayHtml = ""
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function DxaLooksLikeTimeScreen(ByVal html As String) As Boolean
    ' v152:
    ' 管理一覧HTMLにも TimeScreen 文字列や ManagerForm.js 内の戻り先情報が含まれるため、
    ' TimeScreen / FUNCTION_TIME_SCREEN の文字列だけでは打刻画面と判定しません。
    ' 先に管理一覧の強い特徴を除外し、そのうえで「管理ボタンを持つ打刻画面」を判定します。
    If Len(DxaTrimHtml(html)) = 0 Then Exit Function
    If DxaLooksLikeSessionError(html) Then Exit Function
    If DxaLooksLikeLoginForm(html) Then Exit Function

    Dim hasManagedListMarker As Boolean
    hasManagedListMarker = (InStr(1, html, "userAttDataBtn", vbTextCompare) > 0) _
                        Or (InStr(1, html, "inputUserDataBtn", vbTextCompare) > 0) _
                        Or (InStr(1, html, "userListArea", vbTextCompare) > 0) _
                        Or ((InStr(1, html, "社員番号", vbTextCompare) > 0) And (InStr(1, html, "氏名", vbTextCompare) > 0))
    If hasManagedListMarker Then Exit Function

    Dim hasManagerButton As Boolean
    hasManagerButton = (InStr(1, html, "name=""manager""", vbTextCompare) > 0) _
                    Or (InStr(1, html, "name='manager'", vbTextCompare) > 0) _
                    Or (InStr(1, html, "id=""manager""", vbTextCompare) > 0) _
                    Or (InStr(1, html, "id='manager'", vbTextCompare) > 0) _
                    Or (InStr(1, html, "value=""manager""", vbTextCompare) > 0) _
                    Or (InStr(1, html, "value='manager'", vbTextCompare) > 0)

    Dim hasTimeScreenOnlyMarker As Boolean
    hasTimeScreenOnlyMarker = (InStr(1, html, "タイムカードを打刻", vbTextCompare) > 0) _
                           Or (InStr(1, html, "class=""workPlace""", vbTextCompare) > 0) _
                           Or (InStr(1, html, "id=""place""", vbTextCompare) > 0) _
                           Or (InStr(1, html, "TimeScreen.js", vbTextCompare) > 0)

    DxaLooksLikeTimeScreen = hasManagerButton Or hasTimeScreenOnlyMarker
End Function

Private Function DxaLooksLikeSessionError(ByVal html As String) As Boolean
    DxaLooksLikeSessionError = (InStr(1, html, "セッションが有効ではありません", vbTextCompare) > 0) _
                            Or (InStr(1, html, "FUNCTION_ERROR", vbTextCompare) > 0) _
                            Or (InStr(1, html, "Error.js", vbTextCompare) > 0)
End Function

Private Function DxaLooksLikeLoginForm(ByVal html As String) As Boolean
    Dim hasPassword As Boolean
    hasPassword = (InStr(1, html, "type=""password""", vbTextCompare) > 0) _
               Or (InStr(1, html, "type='password'", vbTextCompare) > 0) _
               Or (InStr(1, html, "password", vbTextCompare) > 0 And InStr(1, html, "<form", vbTextCompare) > 0) _
               Or (InStr(1, html, "パスワード", vbTextCompare) > 0 And InStr(1, html, "<form", vbTextCompare) > 0)

    Dim hasLoginMarker As Boolean
    hasLoginMarker = (InStr(1, html, "ログイン", vbTextCompare) > 0) _
                  Or (InStr(1, html, "login", vbTextCompare) > 0) _
                  Or (InStr(1, html, "mail", vbTextCompare) > 0) _
                  Or (InStr(1, html, "email", vbTextCompare) > 0) _
                  Or (InStr(1, html, "メール", vbTextCompare) > 0) _
                  Or (InStr(1, html, "ユーザー", vbTextCompare) > 0) _
                  Or (InStr(1, html, "ユーザ", vbTextCompare) > 0)

    DxaLooksLikeLoginForm = hasPassword And hasLoginMarker
End Function

Private Function DxaResolveTimecardLoginHtml(ByVal timeScreenUrl As String, ByVal firstHtml As String, ByRef cookies As String, ByRef loginPageUrl As String) As String
    On Error Resume Next

    If DxaLooksLikeLoginForm(firstHtml) Then
        loginPageUrl = timeScreenUrl
        DxaResolveTimecardLoginHtml = firstHtml
        Exit Function
    End If

    Dim html As String

    ' セッション無効画面の場合は、画面上の「ログイン」ボタンと Error.js の遷移先を優先して解析します。
    html = DxaTryOpenLoginFromSessionError(timeScreenUrl, firstHtml, cookies, loginPageUrl)
    If DxaLooksLikeLoginForm(html) Then
        DxaResolveTimecardLoginHtml = html
        Exit Function
    End If

    ' 固定候補も広めに試します。サイト側のJSP/Servlet名が変わっても追従しやすくします。
    Dim candidates As Variant
    candidates = Array( _
        "/AttendanceMgt/Access", _
        "/AttendanceMgt/Login.jsp", _
        "/AttendanceMgt/Login", _
        "/AttendanceMgt/LoginScreen.jsp", _
        "/AttendanceMgt/LoginScreen", _
        "/AttendanceMgt/LoginInit.jsp", _
        "/AttendanceMgt/LoginInit", _
        "/AttendanceMgt/LoginAction", _
        "/AttendanceMgt/LoginController", _
        "/AttendanceMgt/LoginServlet", _
        "/AttendanceMgt/Index.jsp", _
        "/AttendanceMgt/Index", _
        "/AttendanceMgt/index.jsp", _
        "/AttendanceMgt/index", _
        "/AttendanceMgt/" _
    )

    html = DxaTryOpenLoginCandidateUrls(timeScreenUrl, firstHtml, cookies, loginPageUrl, candidates)
    If DxaLooksLikeLoginForm(html) Then
        DxaResolveTimecardLoginHtml = html
        Exit Function
    End If

    loginPageUrl = timeScreenUrl
    DxaResolveTimecardLoginHtml = ""
End Function

Private Function DxaTryOpenLoginFromSessionError(ByVal timeScreenUrl As String, ByVal errorHtml As String, ByRef cookies As String, ByRef loginPageUrl As String) As String
    On Error Resume Next

    If Not DxaLooksLikeSessionError(errorHtml) Then Exit Function

    Dim html As String

    ' 1. まずはブラウザの通常submitに近い形で、hidden項目のみを送信します。
    html = DxaTrySubmitSessionErrorForm(timeScreenUrl, errorHtml, cookies, loginPageUrl)
    If DxaLooksLikeLoginForm(html) Then
        DxaTryOpenLoginFromSessionError = html
        Exit Function
    End If

    ' 2. Error.js内の action / location.href などからログイン画面のURLを推定します。
    html = DxaTryOpenLoginFromSessionScripts(timeScreenUrl, errorHtml, cookies, loginPageUrl)
    If DxaLooksLikeLoginForm(html) Then
        DxaTryOpenLoginFromSessionError = html
        Exit Function
    End If
End Function

Private Function DxaTrySubmitSessionErrorForm(ByVal timeScreenUrl As String, ByVal errorHtml As String, ByRef cookies As String, ByRef loginPageUrl As String) As String
    On Error Resume Next

    Dim actionUrl As String
    Dim methodName As String
    Dim body As String
    Dim html As String

    actionUrl = DxaExtractFormAction(errorHtml)
    If Len(actionUrl) = 0 Then actionUrl = timeScreenUrl
    actionUrl = DxaResolveUrl(timeScreenUrl, actionUrl)

    methodName = DxaExtractFormMethod(errorHtml)
    If Len(methodName) = 0 Then methodName = "GET"

    ' buttonにnameがない画面では、ブラウザはlogin=loginを送信しません。
    ' そのため最初にhidden項目だけで送ります。
    body = DxaDictionaryToPostBody(DxaExtractInputDictionary(errorHtml))
    html = DxaHttpSubmitForm(methodName, actionUrl, body, cookies)
    If DxaLooksLikeLoginForm(html) Then
        loginPageUrl = actionUrl
        DxaTrySubmitSessionErrorForm = html
        Exit Function
    End If

    ' 互換用：旧実装と同じ login=login 付きも試します。
    body = DxaBuildPostBodyWithButton(errorHtml, "login", "login")
    html = DxaHttpSubmitForm(methodName, actionUrl, body, cookies)
    If DxaLooksLikeLoginForm(html) Then
        loginPageUrl = actionUrl
        DxaTrySubmitSessionErrorForm = html
        Exit Function
    End If

    ' FUNCTION_NAMEをログイン系に差し替えるパターンも試します。
    Dim fnNames As Variant
    fnNames = Array("FUNCTION_LOGIN", "FUNCTION_LOGIN_SCREEN", "FUNCTION_LOGIN_INIT", "FUNCTION_SHOW_LOGIN", "FUNCTION_ERROR_LOGIN", "LOGIN", "login")

    Dim i As Long
    Dim dict As Object
    For i = LBound(fnNames) To UBound(fnNames)
        Set dict = DxaExtractInputDictionary(errorHtml)
        dict("FUNCTION_NAME") = CStr(fnNames(i))
        body = DxaDictionaryToPostBody(dict)

        html = DxaHttpSubmitForm(methodName, actionUrl, body, cookies)
        If DxaLooksLikeLoginForm(html) Then
            loginPageUrl = actionUrl
            DxaTrySubmitSessionErrorForm = html
            Exit Function
        End If
    Next
End Function

Private Function DxaTryOpenLoginFromSessionScripts(ByVal timeScreenUrl As String, ByVal errorHtml As String, ByRef cookies As String, ByRef loginPageUrl As String) As String
    On Error Resume Next

    Dim scriptUrls As Collection
    Set scriptUrls = DxaExtractScriptSrcUrls(timeScreenUrl, errorHtml)
    If scriptUrls Is Nothing Then Exit Function

    Dim allCandidates As Collection
    Set allCandidates = New Collection

    Dim i As Long
    Dim jsUrl As String
    Dim jsText As String
    For i = 1 To scriptUrls.Count
        jsUrl = CStr(scriptUrls(i))
        jsText = DxaHttpRequest("GET", jsUrl, "", cookies, "")
        If Len(jsText) > 0 Then
            Call DxaSaveTimecardDebugHtml("login_script_" & CStr(i), jsText)
            Call DxaAddLoginCandidatesFromScript(allCandidates, timeScreenUrl, jsText)
        End If
    Next

    Dim candidates As Variant
    candidates = DxaCollectionToStringArray(allCandidates)
    If IsEmpty(candidates) Then Exit Function

    DxaTryOpenLoginFromSessionScripts = DxaTryOpenLoginCandidateUrls(timeScreenUrl, errorHtml, cookies, loginPageUrl, candidates)
End Function

Private Function DxaTryOpenLoginCandidateUrls(ByVal timeScreenUrl As String, ByVal baseHtml As String, ByRef cookies As String, ByRef loginPageUrl As String, ByVal candidates As Variant) As String
    On Error Resume Next

    Dim i As Long
    Dim url As String
    Dim html As String
    Dim baseBody As String
    Dim loginBody As String

    baseBody = DxaDictionaryToPostBody(DxaExtractInputDictionary(baseHtml))
    loginBody = DxaBuildPostBodyWithButton(baseHtml, "login", "login")

    For i = LBound(candidates) To UBound(candidates)
        url = DxaResolveUrl(timeScreenUrl, CStr(candidates(i)))
        If Len(url) > 0 Then
            html = DxaHttpRequest("GET", url, "", cookies, "")
            If DxaLooksLikeLoginForm(html) Then
                loginPageUrl = url
                DxaTryOpenLoginCandidateUrls = html
                Exit Function
            End If

            If Len(baseBody) > 0 Then
                html = DxaHttpRequest("GET", DxaAppendQueryString(url, baseBody), "", cookies, "")
                If DxaLooksLikeLoginForm(html) Then
                    loginPageUrl = url
                    DxaTryOpenLoginCandidateUrls = html
                    Exit Function
                End If

                html = DxaHttpRequest("POST", url, baseBody, cookies, "application/x-www-form-urlencoded")
                If DxaLooksLikeLoginForm(html) Then
                    loginPageUrl = url
                    DxaTryOpenLoginCandidateUrls = html
                    Exit Function
                End If
            End If

            If Len(loginBody) > 0 Then
                html = DxaHttpRequest("POST", url, loginBody, cookies, "application/x-www-form-urlencoded")
                If DxaLooksLikeLoginForm(html) Then
                    loginPageUrl = url
                    DxaTryOpenLoginCandidateUrls = html
                    Exit Function
                End If
            End If

            Dim fnNames As Variant
            fnNames = Array("FUNCTION_LOGIN", "FUNCTION_LOGIN_SCREEN", "FUNCTION_LOGIN_INIT", "FUNCTION_SHOW_LOGIN", "FUNCTION_ERROR_LOGIN", "LOGIN", "login")

            Dim k As Long
            Dim dict As Object
            Dim fnBody As String
            For k = LBound(fnNames) To UBound(fnNames)
                Set dict = DxaExtractInputDictionary(baseHtml)
                dict("FUNCTION_NAME") = CStr(fnNames(k))
                fnBody = DxaDictionaryToPostBody(dict)

                html = DxaHttpRequest("GET", DxaAppendQueryString(url, fnBody), "", cookies, "")
                If DxaLooksLikeLoginForm(html) Then
                    loginPageUrl = url
                    DxaTryOpenLoginCandidateUrls = html
                    Exit Function
                End If

                html = DxaHttpRequest("POST", url, fnBody, cookies, "application/x-www-form-urlencoded")
                If DxaLooksLikeLoginForm(html) Then
                    loginPageUrl = url
                    DxaTryOpenLoginCandidateUrls = html
                    Exit Function
                End If
            Next
        End If
    Next

    If Len(html) > 0 Then Call DxaSaveTimecardDebugHtml("login_attempt_last", html)
End Function

Private Function DxaExtractScriptSrcUrls(ByVal baseUrl As String, ByVal html As String) As Collection
    On Error GoTo EH

    Dim result As Collection
    Set result = New Collection

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "<script[^>]*src\s*=\s*([""'])(.*?)\1"

    Dim ms As Object
    Set ms = re.Execute(html)

    Dim m As Object
    For Each m In ms
        Call DxaAddUniqueString(result, DxaResolveUrl(baseUrl, CStr(m.SubMatches(1))))
    Next

    Set DxaExtractScriptSrcUrls = result
    Exit Function
EH:
    Set DxaExtractScriptSrcUrls = New Collection
End Function

Private Sub DxaAddLoginCandidatesFromScript(ByVal candidates As Collection, ByVal baseUrl As String, ByVal jsText As String)
    On Error Resume Next

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = """([^""]+)""|'([^']+)'"

    Dim ms As Object
    Set ms = re.Execute(jsText)

    Call DxaAddFormActionAssignmentsFromScript(candidates, baseUrl, jsText)

    Dim m As Object
    For Each m In ms
        Dim s As String
        s = CStr(m.SubMatches(0))
        If Len(s) = 0 Then s = CStr(m.SubMatches(1))
        s = Trim$(s)

        If DxaLooksLikeLoginUrlCandidate(s) Then
            Call DxaAddUniqueString(candidates, DxaResolveUrl(baseUrl, s))
        End If
    Next

    ' JSが直接URLを文字列で持っていない場合に備えた補助候補です。
    ' Error.jsでは document.mainForm.action = "/AttendanceMgt/Access" でログイン画面へ遷移します。
    Call DxaAddUniqueString(candidates, DxaResolveUrl(baseUrl, "/AttendanceMgt/Access"))
    Call DxaAddUniqueString(candidates, DxaResolveUrl(baseUrl, "/AttendanceMgt/Login"))
    Call DxaAddUniqueString(candidates, DxaResolveUrl(baseUrl, "/AttendanceMgt/Login.jsp"))
    Call DxaAddUniqueString(candidates, DxaResolveUrl(baseUrl, "/AttendanceMgt/LoginScreen"))
    Call DxaAddUniqueString(candidates, DxaResolveUrl(baseUrl, "/AttendanceMgt/LoginScreen.jsp"))
End Sub

Private Function DxaLooksLikeLoginUrlCandidate(ByVal text As String) As Boolean
    Dim s As String
    s = Trim$(text)
    If Len(s) = 0 Then Exit Function
    If Left$(s, 1) = "#" Then Exit Function
    If Left$(s, 1) = "." Then Exit Function
    If InStr(1, s, " ", vbTextCompare) > 0 Then Exit Function
    If InStr(1, s, "function", vbTextCompare) > 0 Then Exit Function
    If InStr(1, s, "click", vbTextCompare) > 0 Then Exit Function

    If InStr(1, s, "AttendanceMgt/Access", vbTextCompare) > 0 _
       Or StrComp(s, "/AttendanceMgt/Access", vbTextCompare) = 0 _
       Or StrComp(s, "Access", vbTextCompare) = 0 Then
        DxaLooksLikeLoginUrlCandidate = True
    ElseIf InStr(1, s, "login", vbTextCompare) > 0 Then
        DxaLooksLikeLoginUrlCandidate = True
    ElseIf InStr(1, s, "Login", vbTextCompare) > 0 Then
        DxaLooksLikeLoginUrlCandidate = True
    ElseIf InStr(1, s, "ログイン", vbTextCompare) > 0 Then
        DxaLooksLikeLoginUrlCandidate = True
    End If
End Function

Private Sub DxaAddUniqueString(ByVal items As Collection, ByVal value As String)
    On Error Resume Next
    Dim s As String
    s = Trim$(value)
    If Len(s) = 0 Then Exit Sub

    Dim i As Long
    For i = 1 To items.Count
        If StrComp(CStr(items(i)), s, vbTextCompare) = 0 Then Exit Sub
    Next
    items.Add s
End Sub

Private Function DxaCollectionToStringArray(ByVal items As Collection) As Variant
    If items Is Nothing Then
        DxaCollectionToStringArray = Empty
        Exit Function
    End If
    If items.Count = 0 Then
        DxaCollectionToStringArray = Empty
        Exit Function
    End If

    Dim arr() As String
    ReDim arr(0 To items.Count - 1)

    Dim i As Long
    For i = 1 To items.Count
        arr(i - 1) = CStr(items(i))
    Next

    DxaCollectionToStringArray = arr
End Function

Private Function DxaHttpSubmitForm(ByVal methodName As String, ByVal actionUrl As String, ByVal body As String, ByRef cookies As String) As String
    If UCase$(Trim$(methodName)) = "GET" Then
        DxaHttpSubmitForm = DxaHttpRequest("GET", DxaAppendQueryString(actionUrl, body), "", cookies, "")
    Else
        DxaHttpSubmitForm = DxaHttpRequest("POST", actionUrl, body, cookies, "application/x-www-form-urlencoded")
    End If
End Function
Private Function DxaSubmitTimecardManagerButton(ByVal timeScreenUrl As String, ByVal managedListUrl As String, ByVal timeScreenHtml As String, ByRef cookies As String) As String
    On Error Resume Next

    Dim candidates As Collection
    Set candidates = New Collection

    ' v148重要:
    ' 管理ボタン以外のURLを絶対に候補へ入れない。
    ' TimeScreen.jsには出勤・退勤・勤務地設定・今月の勤怠一覧などのactionも含まれるため、
    ' それらを機械的に候補へ入れると、別ボタン相当の処理になる危険があります。
    ' ここではユーザー確認済みの管理遷移先 GetManagedList のみを対象にします。
    Call DxaAddUniqueString(candidates, managedListUrl)
    Call DxaAddUniqueString(candidates, DxaResolveUrl(timeScreenUrl, "/AttendanceMgt/GetManagedList"))

    ' TimeScreen.js側で manager ボタンに紐づく GetManagedList が記載されている場合のみ候補へ追加します。
    Dim scriptUrls As Collection
    Set scriptUrls = DxaExtractScriptSrcUrls(timeScreenUrl, timeScreenHtml)

    Dim i As Long
    Dim jsUrl As String
    Dim jsText As String
    If Not scriptUrls Is Nothing Then
        For i = 1 To scriptUrls.Count
            jsUrl = CStr(scriptUrls(i))
            jsText = DxaHttpRequest("GET", jsUrl, "", cookies, "")
            If Len(jsText) > 0 Then
                Call DxaSaveTimecardDebugHtml("time_screen_script_" & CStr(i), jsText)
                Call DxaAddManagerCandidatesFromScript(candidates, timeScreenUrl, jsText)
            End If
        Next
    End If

    Dim baseBody As String
    Dim managerBody As String
    Dim dict As Object

    Set dict = DxaExtractInputDictionary(timeScreenHtml)
    Call DxaNormalizeTimeScreenFormDictionary(dict)
    baseBody = DxaDictionaryToPostBody(dict)

    If Len(baseBody) > 0 Then Call DxaSaveTimecardDebugHtml("manager_post_body", baseBody)

    Dim bodies As Collection
    Set bodies = New Collection
    ' v152:
    ' ユーザーがDevToolsで確認したGetManagedListのForm Dataと同じ項目だけを送信します。
    ' 実ブラウザのPOSTには manager=manager や FUNCTION_NAME差し替えは含まれていなかったため、
    ' 推測パターンは送らず、TimeScreen.jspのmainForm値をそのまま使用します。
    If Len(baseBody) > 0 Then bodies.Add baseBody

    Dim actionUrl As String
    Dim html As String
    Dim lastHtml As String
    Dim j As Long
    Dim trialCookies As String

    For i = 1 To candidates.Count
        actionUrl = CStr(candidates(i))
        If Len(actionUrl) = 0 Then GoTo ContinueCandidate
        If Not DxaLooksLikeManagerUrlCandidate(actionUrl) Then GoTo ContinueCandidate

        For j = 1 To bodies.Count
            trialCookies = cookies
            html = DxaHttpRequest("POST", actionUrl, CStr(bodies(j)), trialCookies, "application/x-www-form-urlencoded", timeScreenUrl)
            If Len(DxaTrimHtml(html)) > 0 Then lastHtml = html
            If DxaLooksLikeManagedList(html) Then
                cookies = trialCookies
                DxaSubmitTimecardManagerButton = html
                Exit Function
            End If
        Next

ContinueCandidate:
    Next

    If Len(lastHtml) > 0 Then Call DxaSaveTimecardDebugHtml("manager_post_result", lastHtml)
    DxaSubmitTimecardManagerButton = lastHtml
End Function

Private Sub DxaAddManagerCandidatesFromScript(ByVal candidates As Collection, ByVal baseUrl As String, ByVal jsText As String)
    On Error Resume Next

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = """([^""]+)""|'([^']+)'"

    Dim ms As Object
    Set ms = re.Execute(jsText)

    Dim m As Object
    For Each m In ms
        Dim s As String
        s = CStr(m.SubMatches(0))
        If Len(s) = 0 Then s = CStr(m.SubMatches(1))
        s = Trim$(s)

        If DxaLooksLikeManagerUrlCandidate(s) Then
            Call DxaAddUniqueString(candidates, DxaResolveUrl(baseUrl, s))
        End If
    Next
End Sub

Private Function DxaLooksLikeManagerUrlCandidate(ByVal text As String) As Boolean
    Dim s As String
    s = Trim$(text)
    If Len(s) = 0 Then Exit Function
    If Left$(s, 1) = "#" Then Exit Function
    If Left$(s, 1) = "." Then Exit Function
    If InStr(1, s, " ", vbTextCompare) > 0 Then Exit Function
    If InStr(1, s, "function", vbTextCompare) > 0 Then Exit Function
    If InStr(1, s, "click", vbTextCompare) > 0 Then Exit Function

    DxaLooksLikeManagerUrlCandidate = (InStr(1, s, "GetManagedList", vbTextCompare) > 0) _
                                   Or (InStr(1, s, "ManagedList", vbTextCompare) > 0) _
                                   Or (InStr(1, s, "manager", vbTextCompare) > 0 And InStr(1, s, "AttendanceMgt", vbTextCompare) > 0)
End Function

Private Function DxaAppendQueryString(ByVal url As String, ByVal body As String) As String
    If Len(body) = 0 Then
        DxaAppendQueryString = url
    ElseIf InStr(1, url, "?", vbTextCompare) > 0 Then
        DxaAppendQueryString = url & "&" & body
    Else
        DxaAppendQueryString = url & "?" & body
    End If
End Function
Private Function DxaFetchTimecardDisplayHtmlByUserValue(ByVal managedListUrl As String, ByRef cookies As String, ByVal userValue As String, ByVal listHtml As String) As String
    On Error GoTo EH

    Dim displayUrl As String
    displayUrl = DxaResolveUrl(managedListUrl, "/AttendanceMgt/DisplayList")

    ' v152:
    ' ManagerForm.js の実動作は、選択された userAttDataBtn.value を hidden の dispUserId にセットし、
    ' mainForm を /AttendanceMgt/DisplayList へ POST する形です。
    ' button 名そのものは Form Data に入らないため、まずこの実ブラウザ相当のPOSTを実行します。
    Dim body As String
    Dim html As String

    body = DxaBuildManagedListUserDisplayBody(listHtml, userValue)
    If Len(body) > 0 Then Call DxaSaveTimecardDebugHtml("display_list_post_body", body)

    html = DxaHttpRequest("POST", displayUrl, body, cookies, "application/x-www-form-urlencoded", managedListUrl)
    If DxaLooksLikeAttendanceDisplay(html) Then
        DxaFetchTimecardDisplayHtmlByUserValue = html
        Exit Function
    End If

    If Len(DxaTrimHtml(html)) > 0 Then Call DxaSaveTimecardDebugHtml("display_list_post_result", html)

    ' Fallbacks for older/alternate HTML structures.
    Dim endpoints As Variant
    endpoints = Array(displayUrl, managedListUrl)

    Dim postNames As Variant
    postNames = Array("dispUserId", "userAttDataBtn", "inputUserDataBtn", "targetUserId", "userId", "id")

    Dim i As Long
    Dim j As Long

    For i = LBound(endpoints) To UBound(endpoints)
        For j = LBound(postNames) To UBound(postNames)
            body = DxaBuildPostBodyWithButton(listHtml, CStr(postNames(j)), userValue)
            html = DxaHttpRequest("POST", CStr(endpoints(i)), body, cookies, "application/x-www-form-urlencoded", managedListUrl)
            If DxaLooksLikeAttendanceDisplay(html) Then
                DxaFetchTimecardDisplayHtmlByUserValue = html
                Exit Function
            End If
        Next
    Next

    html = DxaHttpRequest("GET", displayUrl & "?dispUserId=" & DxaUrlEncode(userValue), "", cookies, "", managedListUrl)
    If DxaLooksLikeAttendanceDisplay(html) Then
        DxaFetchTimecardDisplayHtmlByUserValue = html
        Exit Function
    End If

    html = DxaHttpRequest("GET", displayUrl & "?userAttDataBtn=" & DxaUrlEncode(userValue), "", cookies, "", managedListUrl)
    If DxaLooksLikeAttendanceDisplay(html) Then
        DxaFetchTimecardDisplayHtmlByUserValue = html
        Exit Function
    End If

    Err.Raise vbObjectError + 622, "DelaxTools", "対象者の勤怠一覧画面へ遷移できませんでした。サイト側の送信項目が変更されている可能性があります。" & vbCrLf & _
             "確認用HTMLとして display_list_post_body / display_list_post_result を保存している場合があります。"
    Exit Function

EH:
    DxaFetchTimecardDisplayHtmlByUserValue = ""
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function DxaBuildManagedListUserDisplayBody(ByVal listHtml As String, ByVal userValue As String) As String
    On Error Resume Next

    Dim dict As Object
    Set dict = DxaExtractInputDictionary(listHtml)

    If dict Is Nothing Then Set dict = CreateObject("Scripting.Dictionary")

    dict("dispUserId") = userValue
    If Not dict.Exists("monthFlg") Then dict("monthFlg") = "thisMonth"
    If Not dict.Exists("FUNCTION_NAME") Then dict("FUNCTION_NAME") = "FUNCTION_MANAGER_FORM"

    DxaBuildManagedListUserDisplayBody = DxaDictionaryToPostBody(dict)
End Function

Private Function DxaLooksLikeManagedList(ByVal html As String) As Boolean
    ' v152:
    ' 前回ログでは manager_post_result に管理一覧HTMLが返っていましたが、
    ' 末尾のスクリプト等に TimeScreen 文字列が含まれたため TimeScreen.jsp と誤判定していました。
    ' そのため、管理一覧の強い特徴を先に評価し、該当すれば管理一覧として扱います。
    If Len(DxaTrimHtml(html)) = 0 Then Exit Function
    If DxaLooksLikeSessionError(html) Then Exit Function
    If DxaLooksLikeLoginForm(html) Then Exit Function

    Dim hasUserButton As Boolean
    hasUserButton = (InStr(1, html, "userAttDataBtn", vbTextCompare) > 0) _
                 Or (InStr(1, html, "inputUserDataBtn", vbTextCompare) > 0)

    Dim hasManagedHeaders As Boolean
    hasManagedHeaders = ((InStr(1, html, "社員番号", vbTextCompare) > 0) And (InStr(1, html, "氏名", vbTextCompare) > 0)) _
                     Or ((InStr(1, html, "部門番号", vbTextCompare) > 0) And (InStr(1, html, "勤怠一覧", vbTextCompare) > 0)) _
                     Or ((InStr(1, html, "部署番号", vbTextCompare) > 0) And (InStr(1, html, "勤怠", vbTextCompare) > 0)) _
                     Or ((InStr(1, html, "userListArea", vbTextCompare) > 0) And (InStr(1, html, "<tr", vbTextCompare) > 0)) _
                     Or ((InStr(1, html, "userName", vbTextCompare) > 0) And (InStr(1, html, "<tr", vbTextCompare) > 0))

    Dim hasAttendanceListButtonInTable As Boolean
    hasAttendanceListButtonInTable = (InStr(1, html, "勤怠一覧", vbTextCompare) > 0) _
                                  And (InStr(1, html, "<tr", vbTextCompare) > 0) _
                                  And ((InStr(1, html, "社員番号", vbTextCompare) > 0) _
                                    Or (InStr(1, html, "氏名", vbTextCompare) > 0) _
                                    Or (InStr(1, html, "userAttDataBtn", vbTextCompare) > 0) _
                                    Or (InStr(1, html, "inputUserDataBtn", vbTextCompare) > 0))

    If hasUserButton Or hasManagedHeaders Or hasAttendanceListButtonInTable Then
        DxaLooksLikeManagedList = True
        Exit Function
    End If

    If DxaLooksLikeTimeScreen(html) Then Exit Function
End Function

Private Function DxaSaveTimecardDebugHtml(ByVal stageName As String, ByVal html As String) As String
    On Error GoTo EH
    Dim folderPath As String
    folderPath = Environ$("APPDATA") & "\DelaxTools\logs"
    DxaEnsureFolder folderPath

    Dim path As String
    path = folderPath & "\timecard_" & stageName & "_" & Format$(Now, "yyyymmdd_hhnnss") & ".html"

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText html
    stm.SaveToFile path, 2
    stm.Close

    DxaSaveTimecardDebugHtml = path
    Exit Function
EH:
    DxaSaveTimecardDebugHtml = "(保存失敗)"
End Function

Private Function DxaExtractFormAction(ByVal html As String) As String
    On Error Resume Next
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "<form[^>]*>"

    Dim ms As Object
    Set ms = re.Execute(html)
    If ms.Count > 0 Then DxaExtractFormAction = DxaExtractAttributeValue(CStr(ms(0).Value), "action")
End Function


Private Function DxaExtractFormMethod(ByVal html As String) As String
    On Error Resume Next
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "<form[^>]*>"

    Dim ms As Object
    Set ms = re.Execute(html)
    If ms.Count > 0 Then DxaExtractFormMethod = UCase$(DxaExtractAttributeValue(CStr(ms(0).Value), "method"))
End Function

Private Function DxaSubmitTimecardLoginForm(ByVal timeScreenUrl As String, ByVal managedListUrl As String, ByVal loginPageUrl As String, ByVal loginHtml As String, ByVal preferredMethod As String, ByVal formActionUrl As String, ByVal loginBody As String, ByRef cookies As String) As String
    On Error Resume Next

    Dim candidates As Collection
    Set candidates = New Collection

    Call DxaAddLoginFormSubmitCandidates(candidates, loginPageUrl, loginHtml, cookies)

    ' Browser click on the login button normally changes mainForm.action in LoginForm.js.
    ' Therefore, script-derived actions and common login endpoints must be tried before the blank form action resolved to /Access.
    Call DxaAddUniqueString(candidates, DxaResolveUrl(loginPageUrl, "/AttendanceMgt/Login"))
    Call DxaAddUniqueString(candidates, DxaResolveUrl(loginPageUrl, "/AttendanceMgt/LoginAction"))
    Call DxaAddUniqueString(candidates, DxaResolveUrl(loginPageUrl, "/AttendanceMgt/LoginController"))
    Call DxaAddUniqueString(candidates, DxaResolveUrl(loginPageUrl, "/AttendanceMgt/LoginServlet"))
    Call DxaAddUniqueString(candidates, DxaResolveUrl(loginPageUrl, "/AttendanceMgt/UserLogin"))
    Call DxaAddUniqueString(candidates, DxaResolveUrl(loginPageUrl, "/AttendanceMgt/CheckLogin"))
    Call DxaAddUniqueString(candidates, DxaResolveUrl(loginPageUrl, "/AttendanceMgt/LoginCheck"))
    Call DxaAddUniqueString(candidates, DxaResolveUrl(loginPageUrl, "/AttendanceMgt/LoginAuth"))
    Call DxaAddUniqueString(candidates, DxaResolveUrl(loginPageUrl, "/AttendanceMgt/Auth"))
    If Len(formActionUrl) > 0 Then Call DxaAddUniqueString(candidates, formActionUrl)
    Call DxaAddUniqueString(candidates, DxaResolveUrl(loginPageUrl, "/AttendanceMgt/Access"))

    Dim methods As Variant
    methods = Array(preferredMethod, "POST", "GET")

    Dim i As Long
    Dim j As Long
    Dim actionUrl As String
    Dim methodName As String
    Dim html As String
    Dim checkHtml As String
    Dim listHtml As String
    Dim lastHtml As String

    For i = 1 To candidates.Count
        actionUrl = CStr(candidates(i))
        If Len(actionUrl) = 0 Then GoTo ContinueCandidate

        For j = LBound(methods) To UBound(methods)
            methodName = UCase$(Trim$(CStr(methods(j))))
            If Len(methodName) = 0 Then methodName = "POST"
            If j > LBound(methods) Then
                If UCase$(Trim$(CStr(methods(j)))) = UCase$(Trim$(CStr(methods(j - 1)))) Then GoTo ContinueMethod
            End If

            html = DxaHttpSubmitForm(methodName, actionUrl, loginBody, cookies)
            If Len(DxaTrimHtml(html)) > 0 Then lastHtml = html

            If DxaLooksLikeManagedList(html) Or DxaLooksLikeTimeScreen(html) Then
                DxaSubmitTimecardLoginForm = html
                Exit Function
            End If

            listHtml = DxaHttpRequest("GET", managedListUrl, "", cookies, "")
            If DxaLooksLikeManagedList(listHtml) Then
                DxaSubmitTimecardLoginForm = listHtml
                Exit Function
            End If

            checkHtml = DxaHttpRequest("GET", timeScreenUrl, "", cookies, "")
            If DxaLooksLikeManagedList(checkHtml) Or DxaLooksLikeTimeScreen(checkHtml) Then
                DxaSubmitTimecardLoginForm = checkHtml
                Exit Function
            End If

ContinueMethod:
        Next
ContinueCandidate:
    Next

    If Len(lastHtml) > 0 Then Call DxaSaveTimecardDebugHtml("login_post_result", lastHtml)
    DxaSubmitTimecardLoginForm = lastHtml
End Function

Private Sub DxaAddLoginFormSubmitCandidates(ByVal candidates As Collection, ByVal loginPageUrl As String, ByVal loginHtml As String, ByRef cookies As String)
    On Error Resume Next

    Dim scriptUrls As Collection
    Set scriptUrls = DxaExtractScriptSrcUrls(loginPageUrl, loginHtml)
    If scriptUrls Is Nothing Then Exit Sub

    Dim i As Long
    Dim jsUrl As String
    Dim jsText As String

    For i = 1 To scriptUrls.Count
        jsUrl = CStr(scriptUrls(i))
        jsText = DxaHttpRequest("GET", jsUrl, "", cookies, "")
        If Len(jsText) > 0 Then
            Call DxaSaveTimecardDebugHtml("login_form_script_" & CStr(i), jsText)
            Call DxaAddFormActionAssignmentsFromScript(candidates, loginPageUrl, jsText)
            Call DxaAddLoginCandidatesFromScript(candidates, loginPageUrl, jsText)
        End If
    Next
End Sub

Private Sub DxaAddFormActionAssignmentsFromScript(ByVal candidates As Collection, ByVal baseUrl As String, ByVal jsText As String)
    On Error Resume Next

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "mainForm\.action\s*=\s*([""'])(.*?)\1"

    Dim ms As Object
    Set ms = re.Execute(jsText)

    Dim m As Object
    For Each m In ms
        Dim actionText As String
        actionText = Trim$(CStr(m.SubMatches(1)))
        If Len(actionText) > 0 Then
            If Left$(actionText, 1) = "/" Or LCase$(Left$(actionText, 4)) = "http" Or InStr(1, actionText, "AttendanceMgt", vbTextCompare) > 0 Then
                Call DxaAddUniqueString(candidates, DxaResolveUrl(baseUrl, actionText))
            End If
        End If
    Next
End Sub

Private Function DxaTrimHtml(ByVal html As String) As String
    Dim s As String
    s = Replace$(html, ChrW$(&HFEFF), "")
    s = Replace$(s, vbCr, "")
    s = Replace$(s, vbLf, "")
    s = Replace$(s, vbTab, "")
    DxaTrimHtml = Trim$(s)
End Function

Private Function DxaBuildLoginPostBody(ByVal html As String, ByVal email As String, ByVal password As String) As String
    On Error GoTo EH

    Dim dict As Object
    Set dict = DxaExtractInputDictionary(html)

    Dim emailKey As String
    Dim passKey As String
    emailKey = DxaFindLoginInputName(html, False)
    passKey = DxaFindLoginInputName(html, True)

    If Len(emailKey) > 0 Then
        dict(emailKey) = email
    Else
        ' ログインID欄のnameが取得できない場合のみ、一般的な名前を追加します。
        ' 既存hidden値（CHECK_SESSION_IDなど）は上書きしません。
        dict("Email") = email
        dict("email") = email
        dict("MailAddress") = email
        dict("mailAddress") = email
        dict("UserId") = email
        dict("userId") = email
        dict("LoginId") = email
        dict("loginId") = email
        dict("UserName") = email
        dict("userName") = email
        dict("Account") = email
        dict("account") = email
    End If

    If Len(passKey) > 0 Then
        dict(passKey) = password
    Else
        dict("Password") = password
        dict("password") = password
        dict("PassWord") = password
        dict("passWord") = password
        dict("LoginPassword") = password
        dict("loginPassword") = password
        dict("UserPassword") = password
        dict("userPassword") = password
        dict("Pass") = password
        dict("pass") = password
    End If

    DxaAddLoginSubmitButton dict, html

    DxaBuildLoginPostBody = DxaDictionaryToPostBody(dict)
    Exit Function

EH:
    DxaBuildLoginPostBody = "Email=" & DxaUrlEncode(email) & "&MailAddress=" & DxaUrlEncode(email) & "&Password=" & DxaUrlEncode(password) & "&password=" & DxaUrlEncode(password)
End Function

Private Function DxaFindLoginInputName(ByVal html As String, ByVal wantPassword As Boolean) As String
    On Error GoTo EH

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "<input[^>]*>"

    Dim ms As Object
    Set ms = re.Execute(html)

    Dim m As Object
    For Each m In ms
        Dim tagHtml As String
        Dim nm As String
        Dim typ As String
        Dim lowerName As String

        tagHtml = CStr(m.Value)
        nm = DxaExtractAttributeValue(tagHtml, "name")
        typ = LCase$(DxaExtractAttributeValue(tagHtml, "type"))
        lowerName = LCase$(nm)

        If Len(nm) > 0 Then
            If wantPassword Then
                If typ = "password" Or InStr(1, lowerName, "password", vbTextCompare) > 0 Or InStr(1, lowerName, "pass", vbTextCompare) > 0 Then
                    DxaFindLoginInputName = nm
                    Exit Function
                End If
            Else
                If typ <> "hidden" And typ <> "password" Then
                    If InStr(1, lowerName, "session", vbTextCompare) = 0 _
                       And InStr(1, lowerName, "check", vbTextCompare) = 0 _
                       And InStr(1, lowerName, "function", vbTextCompare) = 0 _
                       And InStr(1, lowerName, "token", vbTextCompare) = 0 _
                       And InStr(1, lowerName, "csrf", vbTextCompare) = 0 _
                       And InStr(1, lowerName, "password", vbTextCompare) = 0 _
                       And InStr(1, lowerName, "pass", vbTextCompare) = 0 Then

                        If typ = "email" _
                           Or InStr(1, lowerName, "mail", vbTextCompare) > 0 _
                           Or InStr(1, lowerName, "email", vbTextCompare) > 0 _
                           Or InStr(1, lowerName, "login", vbTextCompare) > 0 _
                           Or InStr(1, lowerName, "user", vbTextCompare) > 0 _
                           Or InStr(1, lowerName, "account", vbTextCompare) > 0 _
                           Or InStr(1, lowerName, "employee", vbTextCompare) > 0 _
                           Or InStr(1, lowerName, "staff", vbTextCompare) > 0 _
                           Or InStr(1, lowerName, "id", vbTextCompare) > 0 Then
                            DxaFindLoginInputName = nm
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next
    Exit Function

EH:
    DxaFindLoginInputName = ""
End Function

Private Sub DxaAddLoginSubmitButton(ByVal dict As Object, ByVal html As String)
    On Error Resume Next

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "<button[\s\S]*?</button>|<input[^>]*>"

    Dim ms As Object
    Set ms = re.Execute(html)

    Dim m As Object
    For Each m In ms
        Dim tagHtml As String
        Dim lowerTag As String
        Dim nm As String
        Dim val As String

        tagHtml = CStr(m.Value)
        lowerTag = LCase$(tagHtml)

        If InStr(1, lowerTag, "login", vbTextCompare) > 0 Or InStr(1, tagHtml, "ログイン", vbTextCompare) > 0 Then
            nm = DxaExtractAttributeValue(tagHtml, "name")
            If Len(nm) > 0 Then
                val = DxaExtractAttributeValue(tagHtml, "value")
                If Len(val) = 0 Then val = "login"
                dict(nm) = val
                Exit Sub
            End If
        End If
    Next
End Sub
Private Sub DxaNormalizeTimeScreenFormDictionary(ByVal dict As Object)
    On Error Resume Next

    ' TimeScreen.js は画面ロード時に PC/タブレットの場合 terminalFlg=0 を設定します。
    ' VBAのHTTP送信ではJavaScriptが実行されないため、未設定のまま送るとGetManagedList側で
    ' システムエラーになる可能性があります。管理ボタン押下時のブラウザ状態に合わせます。
    If dict Is Nothing Then Exit Sub

    If dict.Exists("terminalFlg") Then
        If Len(Trim$(CStr(dict("terminalFlg")))) = 0 Then dict("terminalFlg") = "0"
    Else
        dict("terminalFlg") = "0"
    End If

    If dict.Exists("managerFlg") Then
        If Len(Trim$(CStr(dict("managerFlg")))) = 0 Then dict("managerFlg") = "2"
    Else
        dict("managerFlg") = "2"
    End If
End Sub

Private Function DxaBuildPostBodyWithButton(ByVal html As String, ByVal buttonName As String, ByVal buttonValue As String) As String
    Dim dict As Object
    Set dict = DxaExtractInputDictionary(html)
    dict(buttonName) = buttonValue
    DxaBuildPostBodyWithButton = DxaDictionaryToPostBody(dict)
End Function

Private Function DxaExtractInputDictionary(ByVal html As String) As Object
    On Error Resume Next

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "<input[^>]*>|<textarea[\s\S]*?</textarea>|<select[\s\S]*?</select>"

    Dim ms As Object
    Set ms = re.Execute(html)

    Dim m As Object
    For Each m In ms
        Dim tagHtml As String
        Dim lowerTag As String
        Dim nm As String
        Dim val As String

        tagHtml = CStr(m.Value)
        lowerTag = LCase$(tagHtml)
        nm = DxaExtractAttributeValue(tagHtml, "name")

        If Len(nm) > 0 Then
            If Left$(lowerTag, 9) = "<textarea" Then
                val = DxaExtractTextareaValue(tagHtml)
            ElseIf Left$(lowerTag, 7) = "<select" Then
                val = DxaExtractSelectedOptionValue(tagHtml)
            Else
                val = DxaExtractAttributeValue(tagHtml, "value")
            End If
            dict(nm) = val
        End If
    Next

    Set DxaExtractInputDictionary = dict
End Function

Private Function DxaExtractTextareaValue(ByVal tagHtml As String) As String
    On Error GoTo EH

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "<textarea[^>]*>([\s\S]*?)</textarea>"

    Dim ms As Object
    Set ms = re.Execute(tagHtml)
    If ms.Count > 0 Then
        DxaExtractTextareaValue = DxaHtmlDecodeBasic(CStr(ms(0).SubMatches(0)))
    Else
        DxaExtractTextareaValue = ""
    End If
    Exit Function
EH:
    DxaExtractTextareaValue = ""
End Function

Private Function DxaExtractSelectedOptionValue(ByVal tagHtml As String) As String
    On Error GoTo EH

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "<option[^>]*>"

    Dim ms As Object
    Set ms = re.Execute(tagHtml)

    Dim m As Object
    Dim firstVal As String
    For Each m In ms
        Dim optionHtml As String
        Dim val As String
        optionHtml = CStr(m.Value)
        val = DxaExtractAttributeValue(optionHtml, "value")
        If Len(firstVal) = 0 Then firstVal = val
        If InStr(1, optionHtml, "selected", vbTextCompare) > 0 Then
            DxaExtractSelectedOptionValue = val
            Exit Function
        End If
    Next

    DxaExtractSelectedOptionValue = firstVal
    Exit Function
EH:
    DxaExtractSelectedOptionValue = ""
End Function

Private Function DxaFindInputNameLike(ByVal dict As Object, ByVal includeWords As Variant, ByVal excludeWords As Variant) As String
    On Error Resume Next

    Dim key As Variant
    For Each key In dict.Keys
        Dim lowerKey As String
        lowerKey = LCase$(CStr(key))

        Dim excluded As Boolean
        Dim w As Variant
        For Each w In excludeWords
            If InStr(1, lowerKey, LCase$(CStr(w)), vbTextCompare) > 0 Then excluded = True
        Next

        If Not excluded Then
            For Each w In includeWords
                If InStr(1, lowerKey, LCase$(CStr(w)), vbTextCompare) > 0 Then
                    DxaFindInputNameLike = CStr(key)
                    Exit Function
                End If
            Next
        End If
    Next
End Function

Private Function DxaDictionaryToPostBody(ByVal dict As Object) As String
    On Error Resume Next

    Dim body As String
    Dim key As Variant
    For Each key In dict.Keys
        If Len(body) > 0 Then body = body & "&"
        body = body & DxaUrlEncode(CStr(key)) & "=" & DxaUrlEncode(CStr(dict(key)))
    Next

    DxaDictionaryToPostBody = body
End Function

Private Function DxaTimecardWebV156Marker() As String
    DxaTimecardWebV156Marker = "v161"
End Function

Private Function DxaParseTimecardUserCandidates(ByVal html As String) As Collection
    On Error GoTo EH

    Dim result As Collection
    Set result = New Collection

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    ' v152:
    ' GetManagedList のHTMLは <tr> が閉じタグではなく次の <tr> で区切られる構造になっているため、
    ' <tr>...</tr> 前提の解析では候補を拾えません。
    ' そのため、まず「4つのtd + userAttDataBtn」を直接検出します。
    Dim reDirect As Object
    Set reDirect = CreateObject("VBScript.RegExp")
    reDirect.Global = True
    reDirect.IgnoreCase = True
    reDirect.Pattern = "<td[^>]*>([\s\S]*?)</td>\s*<td[^>]*>([\s\S]*?)</td>\s*<td[^>]*>([\s\S]*?)</td>\s*<td[^>]*>([\s\S]*?)</td>\s*<td[^>]*>[\s\S]*?<button[^>]*userAttDataBtn[^>]*>"

    Dim msDirect As Object
    Set msDirect = reDirect.Execute(html)

    Dim m As Object
    For Each m In msDirect
        Dim blockHtml As String
        Dim userValue As String
        Dim displayName As String
        Dim rowSummary As String

        blockHtml = CStr(m.Value)
        userValue = DxaExtractButtonValue(blockHtml, "userAttDataBtn")
        If Len(userValue) = 0 Then userValue = DxaExtractFirstButtonValue(blockHtml)

        displayName = Trim$(DxaHtmlDecodeBasic(DxaStripHtmlTags(CStr(m.SubMatches(3)))))
        rowSummary = DxaNormalizeSpaces(DxaStripHtmlTags(blockHtml))

        Call DxaAddTimecardUserCandidate(result, seen, displayName, userValue, rowSummary)
    Next

    If result.Count > 0 Then
        Set DxaParseTimecardUserCandidates = result
        Exit Function
    End If

    ' Fallback 1: valid table row HTML.
    Dim reRow As Object
    Set reRow = CreateObject("VBScript.RegExp")
    reRow.Global = True
    reRow.IgnoreCase = True
    reRow.Pattern = "<tr[\s\S]*?</tr>"

    Dim rows As Object
    Set rows = reRow.Execute(html)

    Dim row As Object
    For Each row In rows
        Dim rowHtml As String
        rowHtml = CStr(row.Value)

        userValue = DxaExtractButtonValue(rowHtml, "userAttDataBtn")
        If Len(userValue) = 0 Then userValue = DxaExtractButtonValue(rowHtml, "inputUserDataBtn")
        If Len(userValue) = 0 Then userValue = DxaExtractFirstButtonValue(rowHtml)

        If Len(userValue) > 0 Then
            displayName = DxaExtractTdTextByClass(rowHtml, "userName")
            If Len(displayName) = 0 Then displayName = DxaExtractTimecardUserNameFromRow(rowHtml)
            rowSummary = DxaNormalizeSpaces(DxaStripHtmlTags(rowHtml))
            Call DxaAddTimecardUserCandidate(result, seen, displayName, userValue, rowSummary)
        End If
    Next

    If result.Count > 0 Then
        Set DxaParseTimecardUserCandidates = result
        Exit Function
    End If

    ' Fallback 2: button位置の直前にある最後のtdを氏名として採用します。
    Dim reButton As Object
    Set reButton = CreateObject("VBScript.RegExp")
    reButton.Global = True
    reButton.IgnoreCase = True
    reButton.Pattern = "<(button|input)[^>]*userAttDataBtn[^>]*>"

    Dim buttons As Object
    Set buttons = reButton.Execute(html)

    For Each m In buttons
        Dim tagHtml As String
        Dim buttonPos As Long
        Dim contextStart As Long
        Dim beforeHtml As String
        Dim contextHtml As String

        tagHtml = CStr(m.Value)
        userValue = DxaExtractAttributeValue(tagHtml, "value")

        buttonPos = CLng(m.FirstIndex) + 1
        contextStart = buttonPos - 1200
        If contextStart < 1 Then contextStart = 1

        beforeHtml = Mid$(html, contextStart, buttonPos - contextStart)
        contextHtml = Mid$(html, contextStart, 1600)
        displayName = DxaExtractLastTdText(beforeHtml)
        rowSummary = DxaNormalizeSpaces(DxaStripHtmlTags(contextHtml))

        Call DxaAddTimecardUserCandidate(result, seen, displayName, userValue, rowSummary)
    Next

    Set DxaParseTimecardUserCandidates = result
    Exit Function

EH:
    Set DxaParseTimecardUserCandidates = New Collection
End Function

Private Sub DxaAddTimecardUserCandidate(ByVal result As Collection, ByVal seen As Object, ByVal displayName As String, ByVal userValue As String, ByVal rowSummary As String)
    On Error Resume Next

    displayName = DxaNormalizeSpaces(displayName)
    userValue = Trim$(CStr(userValue))
    rowSummary = DxaNormalizeSpaces(rowSummary)

    If Len(userValue) = 0 Then Exit Sub
    If Len(displayName) = 0 Then displayName = rowSummary
    If Len(displayName) = 0 Then Exit Sub

    If Not seen Is Nothing Then
        If seen.Exists(userValue) Then Exit Sub
        seen.Add userValue, True
    End If

    result.Add Array(displayName, userValue, rowSummary)
End Sub

Private Function DxaExtractLastTdText(ByVal html As String) As String
    On Error GoTo EH

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "<td[^>]*>([\s\S]*?)</td>"

    Dim ms As Object
    Set ms = re.Execute(html)

    Dim m As Object
    For Each m In ms
        DxaExtractLastTdText = Trim$(DxaHtmlDecodeBasic(DxaStripHtmlTags(CStr(m.SubMatches(0)))))
    Next
    Exit Function

EH:
    DxaExtractLastTdText = ""
End Function

Private Function DxaExtractTimecardUserNameFromRow(ByVal rowHtml As String) As String
    On Error Resume Next

    Dim texts As Collection
    Set texts = DxaExtractTdTexts(rowHtml)

    ' 管理一覧は「部署番号、部署名/会社名、社員番号、氏名、勤怠一覧...」の順で出力される想定です。
    If Not texts Is Nothing Then
        If texts.Count >= 4 Then
            DxaExtractTimecardUserNameFromRow = Trim$(CStr(texts(4)))
            If Len(DxaExtractTimecardUserNameFromRow) > 0 Then Exit Function
        End If
    End If

    Dim rowText As String
    rowText = DxaNormalizeSpaces(DxaStripHtmlTags(rowHtml))
    DxaExtractTimecardUserNameFromRow = rowText
End Function

Private Function DxaExtractTdTexts(ByVal rowHtml As String) As Collection
    On Error GoTo EH

    Dim result As Collection
    Set result = New Collection

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "<td[^>]*>([\s\S]*?)</td>"

    Dim ms As Object
    Set ms = re.Execute(rowHtml)

    Dim m As Object
    For Each m In ms
        result.Add Trim$(DxaHtmlDecodeBasic(DxaStripHtmlTags(CStr(m.SubMatches(0)))))
    Next

    Set DxaExtractTdTexts = result
    Exit Function

EH:
    Set DxaExtractTdTexts = New Collection
End Function

Private Function DxaResolveTimecardSelectedUsers(ByVal candidates As Collection) As Collection
    On Error GoTo EH

    Dim result As Collection
    Set result = New Collection

    Dim savedTargetName As String
    savedTargetName = DxaLoadTimecardSavedTargetName()

    Dim savedTargetNames As Collection
    Set savedTargetNames = DxaParseTimecardTargetNames(savedTargetName)

    Dim selectedValues As Object
    Set selectedValues = CreateObject("Scripting.Dictionary")

    If Not savedTargetNames Is Nothing And savedTargetNames.Count > 0 Then
        Dim i As Long
        Dim targetText As String
        Dim rec As Variant
        Dim failedText As String

        For i = 1 To savedTargetNames.Count
            targetText = CStr(savedTargetNames(i))
            rec = DxaResolveTimecardCandidateForTarget(candidates, targetText)
            If IsArray(rec) Then
                If Not selectedValues.Exists(CStr(rec(1))) Then
                    result.Add rec
                    selectedValues(CStr(rec(1))) = True
                End If
            Else
                failedText = failedText & targetText & vbCrLf
            End If
        Next

        If Len(failedText) > 0 Then
            MsgBox "保存済み対象者名のうち、管理一覧で一致候補を特定できない名前がありました。" & vbCrLf & vbCrLf & _
                   failedText & vbCrLf & _
                   "取得できる対象者のみ処理を続行します。", vbExclamation, "DelaxTools 勤怠取得"
        End If

        If result.Count > 0 Then
            Set DxaResolveTimecardSelectedUsers = result
            Exit Function
        End If
    End If

    Dim userValue As String
    userValue = DxaPromptTimecardUserSelection(candidates)
    If Len(userValue) > 0 Then
        result.Add Array(DxaFindTimecardCandidateNameByValue(candidates, userValue), userValue)
    End If

    Set DxaResolveTimecardSelectedUsers = result
    Exit Function

EH:
    Set DxaResolveTimecardSelectedUsers = New Collection
End Function

Private Function DxaResolveTimecardCandidateForTarget(ByVal candidates As Collection, ByVal targetText As String) As Variant
    On Error GoTo EH

    Dim indexes() As Long
    Dim scores() As Double
    Dim count As Long
    ReDim indexes(1 To candidates.Count)
    ReDim scores(1 To candidates.Count)

    Dim i As Long
    Dim rec As Variant
    Dim score As Double

    For i = 1 To candidates.Count
        rec = candidates(i)
        score = DxaTimecardNameMatchScore(CStr(rec(0)), targetText)
        If score > 0# Then
            count = count + 1
            indexes(count) = i
            scores(count) = score
        End If
    Next

    If count = 0 Then Exit Function

    DxaSortTimecardCandidateIndexes indexes, scores, count

    Dim shouldAutoSelect As Boolean
    If count = 1 Then
        shouldAutoSelect = True
    ElseIf scores(1) >= 1000# Then
        shouldAutoSelect = True
    ElseIf scores(1) >= 800# And scores(1) > scores(2) Then
        shouldAutoSelect = True
    End If

    If shouldAutoSelect Then
        rec = candidates(indexes(1))
        DxaResolveTimecardCandidateForTarget = Array(CStr(rec(0)), CStr(rec(1)))
        Exit Function
    End If

    Dim maxShow As Long
    maxShow = count
    If maxShow > 20 Then maxShow = 20

    Dim prompt As String
    prompt = "保存済み対象者名に一致する候補が複数あります。" & vbCrLf & _
             "保存済み対象者名: " & targetText & vbCrLf & _
             "取得対象者を番号で選択してください。" & vbCrLf & _
             "空欄の場合、この対象者はスキップします。" & vbCrLf & vbCrLf

    For i = 1 To maxShow
        rec = candidates(indexes(i))
        prompt = prompt & CStr(i) & ": " & CStr(rec(0)) & vbCrLf
        If Len(prompt) > 1800 Then Exit For
    Next

    Dim choiceText As String
    choiceText = Trim$(InputBox(prompt, "DelaxTools 勤怠取得", "1"))
    If Len(choiceText) = 0 Then Exit Function
    If Not IsNumeric(choiceText) Then Exit Function

    Dim choiceNo As Long
    choiceNo = CLng(choiceText)
    If choiceNo < 1 Or choiceNo > maxShow Then Exit Function

    rec = candidates(indexes(choiceNo))
    DxaResolveTimecardCandidateForTarget = Array(CStr(rec(0)), CStr(rec(1)))
    Exit Function

EH:
End Function

Private Function DxaFindTimecardCandidateNameByValue(ByVal candidates As Collection, ByVal userValue As String) As String
    On Error GoTo EH

    Dim rec As Variant
    For Each rec In candidates
        If CStr(rec(1)) = CStr(userValue) Then
            DxaFindTimecardCandidateNameByValue = CStr(rec(0))
            Exit Function
        End If
    Next

    DxaFindTimecardCandidateNameByValue = "勤怠一覧"
    Exit Function

EH:
    DxaFindTimecardCandidateNameByValue = "勤怠一覧"
End Function

Private Function DxaBuildTimecardSheetName(ByVal targetName As String) As String
    On Error GoTo EH

    Dim baseName As String
    If Len(Trim$(targetName)) = 0 Then
        baseName = "勤怠一覧"
    Else
        baseName = "勤怠一覧_" & DxaSanitizeWorksheetName(targetName)
    End If

    If Len(baseName) > 31 Then baseName = Left$(baseName, 31)
    If Len(baseName) = 0 Then baseName = "勤怠一覧"
    DxaBuildTimecardSheetName = baseName
    Exit Function

EH:
    DxaBuildTimecardSheetName = "勤怠一覧"
End Function

Private Function DxaSanitizeWorksheetName(ByVal value As String) As String
    Dim s As String
    s = DxaNormalizeSpaces(value)
    s = Replace(s, "\", "")
    s = Replace(s, "/", "")
    s = Replace(s, ":", "")
    s = Replace(s, "?", "")
    s = Replace(s, "*", "")
    s = Replace(s, "[", "")
    s = Replace(s, "]", "")
    s = Replace(s, Chr$(34), "")
    DxaSanitizeWorksheetName = Trim$(s)
End Function

Private Function DxaPromptTimecardUserSelection(ByVal candidates As Collection) As String
    On Error GoTo EH

    Dim savedTargetName As String
    savedTargetName = DxaLoadTimecardSavedTargetName()

    Dim savedTargetNames As Collection
    Set savedTargetNames = DxaParseTimecardTargetNames(savedTargetName)

    Dim filterText As String
    Dim usingSavedTarget As Boolean

    If savedTargetNames.Count = 1 Then
        filterText = CStr(savedTargetNames(1))
        savedTargetName = filterText
        usingSavedTarget = True
    ElseIf savedTargetNames.Count > 1 Then
        filterText = DxaPromptTimecardSavedTargetName(savedTargetNames, usingSavedTarget)
        savedTargetName = filterText
        If Len(filterText) = 0 Then Exit Function
    Else
        filterText = InputBox("取得対象者の氏名の一部を入力してください。" & vbCrLf & _
                              "空白の有無は無視します。空欄の場合は一覧の先頭候補を表示します。", _
                              "DelaxTools 勤怠取得")
        filterText = Trim$(filterText)
    End If

BuildCandidateList:
    Dim indexes() As Long
    Dim scores() As Double
    Dim count As Long
    ReDim indexes(1 To candidates.Count)
    ReDim scores(1 To candidates.Count)

    Dim i As Long
    For i = 1 To candidates.Count
        Dim rec As Variant
        rec = candidates(i)

        Dim score As Double
        If Len(Trim$(filterText)) = 0 Then
            score = 1#
        Else
            score = DxaTimecardNameMatchScore(CStr(rec(0)), filterText)
        End If

        If score > 0# Then
            count = count + 1
            indexes(count) = i
            scores(count) = score
        End If
    Next

    If count = 0 Then
        If usingSavedTarget Then
            MsgBox "保存済み対象者名に一致する候補が見つかりませんでした。" & vbCrLf & vbCrLf & _
                   "保存済み対象者名: " & savedTargetName & vbCrLf & _
                   "氏名の表記が変わっている可能性があります。再入力してください。", vbExclamation, "DelaxTools"
            filterText = InputBox("取得対象者の氏名の一部を入力してください。" & vbCrLf & _
                                  "空白の有無は無視します。", _
                                  "DelaxTools 勤怠取得", savedTargetName)
            filterText = Trim$(filterText)
            usingSavedTarget = False
            If Len(filterText) = 0 Then Exit Function
            GoTo BuildCandidateList
        End If

        MsgBox "入力された氏名に近い候補が見つかりませんでした。" & vbCrLf & _
               "空白の有無や一部の文字で再度試してください。", vbExclamation, "DelaxTools"
        Exit Function
    End If

    DxaSortTimecardCandidateIndexes indexes, scores, count

    Dim maxShow As Long
    maxShow = count
    If maxShow > 40 Then maxShow = 40

    If maxShow = 1 Then
        rec = candidates(indexes(1))
        Dim confirmText As String
        If usingSavedTarget Then
            confirmText = "保存済み対象者名から、以下の対象者を検出しました。" & vbCrLf & vbCrLf & _
                          CStr(rec(0)) & vbCrLf & vbCrLf & _
                          "この対象者の勤怠一覧を取得しますか？"
        Else
            confirmText = "以下の対象者の勤怠一覧を取得しますか？" & vbCrLf & vbCrLf & CStr(rec(0))
        End If

        If MsgBox(confirmText, vbQuestion + vbYesNo, "DelaxTools 勤怠取得") = vbYes Then
            DxaPromptTimecardUserSelection = CStr(rec(1))
        End If
        Exit Function
    End If

    Dim prompt As String
    If usingSavedTarget Then
        prompt = "保存済み対象者名に一致する候補が複数あります。" & vbCrLf & _
                 "保存済み対象者名: " & savedTargetName & vbCrLf & _
                 "取得対象者を番号で選択してください。" & vbCrLf & vbCrLf
    Else
        prompt = "取得対象者を番号で選択してください。" & vbCrLf & _
                 "候補は氏名の部分一致率が高い順に表示します。" & vbCrLf & vbCrLf
    End If

    For i = 1 To maxShow
        rec = candidates(indexes(i))
        prompt = prompt & CStr(i) & ": " & CStr(rec(0)) & vbCrLf
        If Len(prompt) > 1800 Then Exit For
    Next

    Dim choiceText As String
    choiceText = InputBox(prompt, "DelaxTools 勤怠取得")
    choiceText = Trim$(choiceText)
    If Len(choiceText) = 0 Then Exit Function
    If Not IsNumeric(choiceText) Then
        MsgBox "番号を入力してください。", vbExclamation, "DelaxTools"
        Exit Function
    End If

    Dim choiceNo As Long
    choiceNo = CLng(choiceText)
    If choiceNo < 1 Or choiceNo > maxShow Then
        MsgBox "候補一覧に表示されている番号を入力してください。", vbExclamation, "DelaxTools"
        Exit Function
    End If

    rec = candidates(indexes(choiceNo))
    DxaPromptTimecardUserSelection = CStr(rec(1))
    Exit Function

EH:
    DxaPromptTimecardUserSelection = ""
End Function

Private Function DxaPromptTimecardSavedTargetName(ByVal savedTargetNames As Collection, ByRef usingSavedTarget As Boolean) As String
    On Error GoTo EH

    usingSavedTarget = False

    If savedTargetNames Is Nothing Then Exit Function
    If savedTargetNames.Count = 0 Then Exit Function
    If savedTargetNames.Count = 1 Then
        DxaPromptTimecardSavedTargetName = CStr(savedTargetNames(1))
        usingSavedTarget = True
        Exit Function
    End If

    Dim prompt As String
    prompt = "保存済み対象者名が複数あります。" & vbCrLf & _
             "今回取得する対象者を番号で選択してください。" & vbCrLf & _
             "0: 手入力する" & vbCrLf & vbCrLf

    Dim i As Long
    For i = 1 To savedTargetNames.Count
        prompt = prompt & CStr(i) & ": " & CStr(savedTargetNames(i)) & vbCrLf
        If Len(prompt) > 1800 Then Exit For
    Next

    Dim choiceText As String
    choiceText = InputBox(prompt, "DelaxTools 勤怠取得", "1")
    choiceText = Trim$(choiceText)
    If Len(choiceText) = 0 Then Exit Function
    If Not IsNumeric(choiceText) Then
        MsgBox "番号を入力してください。", vbExclamation, "DelaxTools"
        Exit Function
    End If

    Dim choiceNo As Long
    choiceNo = CLng(choiceText)

    If choiceNo = 0 Then
        DxaPromptTimecardSavedTargetName = Trim$(InputBox("取得対象者の氏名の一部を入力してください。" & vbCrLf & _
                                                         "空白の有無は無視します。", _
                                                         "DelaxTools 勤怠取得"))
        usingSavedTarget = False
        Exit Function
    End If

    If choiceNo < 1 Or choiceNo > savedTargetNames.Count Then
        MsgBox "保存済み対象者名一覧に表示されている番号を入力してください。", vbExclamation, "DelaxTools"
        Exit Function
    End If

    DxaPromptTimecardSavedTargetName = CStr(savedTargetNames(choiceNo))
    usingSavedTarget = True
    Exit Function

EH:
    DxaPromptTimecardSavedTargetName = ""
    usingSavedTarget = False
End Function


Private Sub DxaSortTimecardCandidateIndexes(ByRef indexes() As Long, ByRef scores() As Double, ByVal count As Long)
    ' 候補者を一致度の高い順に並べ替えます。
    ' v155: v154で呼び出しのみ存在していたため、コンパイルエラーになっていた不足関数を追加しました。
    On Error GoTo EH

    If count <= 1 Then Exit Sub

    Dim i As Long
    Dim j As Long
    Dim tmpIndex As Long
    Dim tmpScore As Double

    For i = 1 To count - 1
        For j = i + 1 To count
            If scores(j) > scores(i) Then
                tmpScore = scores(i)
                scores(i) = scores(j)
                scores(j) = tmpScore

                tmpIndex = indexes(i)
                indexes(i) = indexes(j)
                indexes(j) = tmpIndex
            ElseIf scores(j) = scores(i) And indexes(j) < indexes(i) Then
                tmpIndex = indexes(i)
                indexes(i) = indexes(j)
                indexes(j) = tmpIndex
            End If
        Next
    Next
    Exit Sub

EH:
    ' 並べ替えに失敗した場合は元の順序で続行します。
End Sub

Private Function DxaTimecardNameMatchScore(ByVal candidateName As String, ByVal filterText As String) As Double
    Dim c As String
    Dim f As String

    c = DxaNormalizeNameForMatch(candidateName)
    f = DxaNormalizeNameForMatch(filterText)

    If Len(f) = 0 Then
        DxaTimecardNameMatchScore = 1#
        Exit Function
    End If

    If c = f Then
        DxaTimecardNameMatchScore = 1000#
        Exit Function
    End If

    If InStr(1, c, f, vbTextCompare) > 0 Then
        DxaTimecardNameMatchScore = 800# + Len(f) / IIf(Len(c) = 0, 1, Len(c))
        Exit Function
    End If

    If InStr(1, f, c, vbTextCompare) > 0 And Len(c) > 0 Then
        DxaTimecardNameMatchScore = 700# + Len(c) / Len(f)
        Exit Function
    End If

    Dim hit As Long
    Dim pos As Long
    Dim i As Long
    For i = 1 To Len(f)
        pos = InStr(1, c, Mid$(f, i, 1), vbTextCompare)
        If pos > 0 Then hit = hit + 1
    Next

    If Len(f) > 0 Then DxaTimecardNameMatchScore = hit / Len(f)
    If DxaTimecardNameMatchScore < 0.5 Then DxaTimecardNameMatchScore = 0#
End Function

Private Function DxaNormalizeNameForMatch(ByVal text As String) As String
    Dim s As String
    s = DxaNormalizeSpaces(text)
    s = Replace(s, " ", "")
    s = Replace(s, vbTab, "")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = LCase$(s)
    DxaNormalizeNameForMatch = s
End Function

Private Function DxaNormalizeName(ByVal text As String) As String
    DxaNormalizeName = DxaNormalizeNameForMatch(text)
End Function

Private Function DxaExtractInputValue(ByVal html As String, ByVal inputName As String) As String
    On Error GoTo EH

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "<input[^>]*name=[""']" & DxaRegexEscape(inputName) & "[""'][^>]*>"

    Dim m As Object
    If re.Test(html) Then
        Set m = re.Execute(html)(0)
        DxaExtractInputValue = DxaExtractAttributeValue(CStr(m.Value), "value")
    End If
    Exit Function

EH:
    DxaExtractInputValue = ""
End Function

Private Function DxaExtractAttributeValue(ByVal tagHtml As String, ByVal attrName As String) As String
    On Error GoTo EH

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True

    ' Quoted attribute: name="value" / name='value'
    re.Pattern = attrName & "\s*=\s*([""'])(.*?)\1"
    If re.Test(tagHtml) Then
        DxaExtractAttributeValue = DxaHtmlDecodeBasic(CStr(re.Execute(tagHtml)(0).SubMatches(1)))
        Exit Function
    End If

    ' Unquoted attribute: name=value
    re.Pattern = attrName & "\s*=\s*([^\s""'>]+)"
    If re.Test(tagHtml) Then
        DxaExtractAttributeValue = DxaHtmlDecodeBasic(CStr(re.Execute(tagHtml)(0).SubMatches(0)))
        Exit Function
    End If

    DxaExtractAttributeValue = ""
    Exit Function

EH:
    DxaExtractAttributeValue = ""
End Function


Private Function DxaUrlOrigin(ByVal url As String) As String
    On Error GoTo EH

    Dim p As Long
    p = InStr(9, url, "/", vbTextCompare)
    If p > 0 Then
        DxaUrlOrigin = Left$(url, p - 1)
    Else
        DxaUrlOrigin = url
    End If
    Exit Function

EH:
    DxaUrlOrigin = ""
End Function

Private Function DxaHttpRequest(ByVal method As String, ByVal url As String, ByVal body As String, ByRef cookies As String, ByVal contentType As String, Optional ByVal refererUrl As String = "") As String
    On Error GoTo EH

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Option(6) = True
    http.Open method, url, False
    http.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36 Edg/148.0.0.0"
    http.SetRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7"
    http.SetRequestHeader "Accept-Language", "ja,en;q=0.9,en-GB;q=0.8,en-US;q=0.7"
    If UCase$(method) = "POST" Then
        http.SetRequestHeader "Cache-Control", "max-age=0"
        http.SetRequestHeader "Upgrade-Insecure-Requests", "1"
        http.SetRequestHeader "Sec-Fetch-Site", "same-origin"
        http.SetRequestHeader "Sec-Fetch-Mode", "navigate"
        http.SetRequestHeader "Sec-Fetch-User", "?1"
        http.SetRequestHeader "Sec-Fetch-Dest", "document"
        http.SetRequestHeader "sec-ch-ua", """Chromium"";v=""148"", ""Microsoft Edge"";v=""148"", ""Not)A;Brand"";v=""99"""
        http.SetRequestHeader "sec-ch-ua-mobile", "?0"
        http.SetRequestHeader "sec-ch-ua-platform", """Windows"""
    Else
        http.SetRequestHeader "Cache-Control", "no-cache"
    End If
    If Len(refererUrl) > 0 Then http.SetRequestHeader "Referer", refererUrl
    If UCase$(method) = "POST" Then
        Dim originUrl As String
        originUrl = DxaUrlOrigin(url)
        If Len(originUrl) > 0 Then http.SetRequestHeader "Origin", originUrl
    End If
    If Len(cookies) > 0 Then http.SetRequestHeader "Cookie", cookies
    If Len(contentType) > 0 Then http.SetRequestHeader "Content-Type", contentType

    If UCase$(method) = "POST" Then
        http.Send body
    Else
        http.Send
    End If

    DxaMergeResponseCookies cookies, http.GetAllResponseHeaders
    DxaHttpRequest = CStr(http.ResponseText)
    Exit Function

EH:
    Err.Raise vbObjectError + 631, "DelaxTools", "Timecard site communication failed." & vbCrLf & Err.Description
End Function

Private Sub DxaMergeResponseCookies(ByRef cookies As String, ByVal headers As String)
    On Error Resume Next

    Dim lines As Variant
    lines = Split(headers, vbCrLf)

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim existing As Variant
    Dim item As Variant
    Dim nameValue As String
    Dim cookieName As String

    If Len(cookies) > 0 Then
        existing = Split(cookies, ";")
        For Each item In existing
            nameValue = Trim$(CStr(item))
            If InStr(1, nameValue, "=", vbTextCompare) > 1 Then
                cookieName = Left$(nameValue, InStr(1, nameValue, "=", vbTextCompare) - 1)
                dict(cookieName) = nameValue
            End If
        Next
    End If

    Dim i As Long
    Dim line As String
    For i = LBound(lines) To UBound(lines)
        line = CStr(lines(i))
        If LCase$(Left$(line, 11)) = "set-cookie:" Then
            nameValue = Trim$(Mid$(line, 12))
            If InStr(1, nameValue, ";", vbTextCompare) > 0 Then nameValue = Left$(nameValue, InStr(1, nameValue, ";", vbTextCompare) - 1)
            If InStr(1, nameValue, "=", vbTextCompare) > 1 Then
                cookieName = Left$(nameValue, InStr(1, nameValue, "=", vbTextCompare) - 1)
                dict(cookieName) = nameValue
            End If
        End If
    Next

    cookies = ""
    Dim key As Variant
    For Each key In dict.Keys
        If Len(cookies) > 0 Then cookies = cookies & "; "
        cookies = cookies & CStr(dict(key))
    Next
End Sub

Private Function DxaFindTimecardUserButtonValue(ByVal html As String, ByVal targetName As String) As String
    ' 互換用。現在は管理一覧から候補を表示して選択する方式を使用します。
    On Error Resume Next
    Dim candidates As Collection
    Set candidates = DxaParseTimecardUserCandidates(html)

    Dim bestScore As Double
    Dim bestValue As String
    Dim i As Long
    For i = 1 To candidates.Count
        Dim rec As Variant
        rec = candidates(i)
        Dim score As Double
        score = DxaTimecardNameMatchScore(CStr(rec(0)), targetName)
        If score > bestScore Then
            bestScore = score
            bestValue = CStr(rec(1))
        End If
    Next

    If bestScore > 0# Then DxaFindTimecardUserButtonValue = bestValue
End Function

Private Function DxaExtractButtonValue(ByVal html As String, ByVal classOrName As String) As String
    On Error Resume Next

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "<(button|input)[^>]*>"

    Dim ms As Object
    Set ms = re.Execute(html)

    Dim m As Object
    For Each m In ms
        Dim tagHtml As String
        tagHtml = CStr(m.Value)
        If InStr(1, tagHtml, classOrName, vbTextCompare) > 0 Then
            DxaExtractButtonValue = DxaExtractAttributeValue(tagHtml, "value")
            If Len(DxaExtractButtonValue) > 0 Then Exit Function
        End If
    Next
End Function

Private Function DxaExtractFirstButtonValue(ByVal html As String) As String
    On Error Resume Next

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "<(button|input)[^>]*>"

    Dim ms As Object
    Set ms = re.Execute(html)

    Dim m As Object
    For Each m In ms
        Dim tagHtml As String
        tagHtml = CStr(m.Value)
        If InStr(1, tagHtml, "userAttDataBtn", vbTextCompare) > 0 Or InStr(1, tagHtml, "inputUserDataBtn", vbTextCompare) > 0 Then
            DxaExtractFirstButtonValue = DxaExtractAttributeValue(tagHtml, "value")
            If Len(DxaExtractFirstButtonValue) > 0 Then Exit Function
        End If
    Next

    If ms.Count > 0 Then DxaExtractFirstButtonValue = DxaExtractAttributeValue(CStr(ms(0).Value), "value")
End Function

Private Function DxaLooksLikeAttendanceDisplay(ByVal html As String) As Boolean
    DxaLooksLikeAttendanceDisplay = (InStr(1, html, "attendanceTime", vbTextCompare) > 0 And InStr(1, html, "retiredTime", vbTextCompare) > 0) _
                                 Or (InStr(1, html, "出勤時間", vbTextCompare) > 0 And InStr(1, html, "退勤時間", vbTextCompare) > 0)
End Function

Private Function DxaResolveUrl(ByVal baseUrl As String, ByVal path As String) As String
    If Len(path) = 0 Then
        DxaResolveUrl = baseUrl
        Exit Function
    End If

    If LCase$(Left$(path, 4)) = "http" Then
        DxaResolveUrl = path
        Exit Function
    End If

    Dim origin As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "^(https?://[^/]+)"

    Dim ms As Object
    Set ms = re.Execute(baseUrl)
    If ms.Count > 0 Then origin = CStr(ms(0).SubMatches(0))

    If Left$(path, 1) = "/" Then
        DxaResolveUrl = origin & path
        Exit Function
    End If

    Dim baseDir As String
    Dim qPos As Long
    qPos = InStr(1, baseUrl, "?", vbTextCompare)
    If qPos > 0 Then baseUrl = Left$(baseUrl, qPos - 1)

    Dim slashPos As Long
    slashPos = InStrRev(baseUrl, "/")
    If slashPos > Len("https://") Then
        baseDir = Left$(baseUrl, slashPos)
        DxaResolveUrl = baseDir & path
    ElseIf Len(origin) > 0 Then
        DxaResolveUrl = origin & "/" & path
    Else
        DxaResolveUrl = path
    End If
End Function
Private Function DxaUrlEncode(ByVal text As String) As String
    On Error GoTo Fallback
    DxaUrlEncode = Application.WorksheetFunction.EncodeURL(text)
    Exit Function
Fallback:
    DxaUrlEncode = Replace$(Replace$(text, " ", "+"), "@", "%40")
End Function

Private Function DxaTimecardEndRoundModeLabel(ByVal mode As Long) As String
    Select Case mode
        Case 1
            DxaTimecardEndRoundModeLabel = "シフト勤務（15分切り捨て）"
        Case 2
            DxaTimecardEndRoundModeLabel = "通常勤務（17:30～18:14は17:30）"
        Case Else
            DxaTimecardEndRoundModeLabel = ""
    End Select
End Function

Private Function DxaRoundTimecardStart(ByVal timeText As String) As String
    Dim totalMinutes As Long
    If Not DxaTimecardTextToMinutes(timeText, totalMinutes) Then
        DxaRoundTimecardStart = timeText
        Exit Function
    End If

    ' 出勤時刻は15分単位で切り上げる。
    ' ただし、分が00の場合だけそのままにする。
    If totalMinutes Mod 60 = 0 Then
        DxaRoundTimecardStart = DxaMinutesToTimecardText(totalMinutes)
    Else
        DxaRoundTimecardStart = DxaMinutesToTimecardText(((totalMinutes \ 15) + 1) * 15)
    End If
End Function

Private Function DxaRoundTimecardEnd(ByVal timeText As String, ByVal mode As Long) As String
    Dim totalMinutes As Long
    If Not DxaTimecardTextToMinutes(timeText, totalMinutes) Then
        DxaRoundTimecardEnd = timeText
        Exit Function
    End If

    Select Case mode
        Case 2
            ' 要望対応：17:30～18:14 の場合は 17:30 として出力する。
            If totalMinutes >= (17 * 60 + 30) And totalMinutes <= (18 * 60 + 14) Then
                DxaRoundTimecardEnd = "17:30"
            Else
                DxaRoundTimecardEnd = DxaMinutesToTimecardText((totalMinutes \ 15) * 15)
            End If
        Case Else
            ' 通常：15分単位で切り捨てる。
            DxaRoundTimecardEnd = DxaMinutesToTimecardText((totalMinutes \ 15) * 15)
    End Select
End Function

Private Function DxaTimecardTextToMinutes(ByVal timeText As String, ByRef totalMinutes As Long) As Boolean
    On Error GoTo EH
    Dim s As String
    s = Trim$(timeText)
    If Len(s) = 0 Then Exit Function

    Dim parts As Variant
    parts = Split(s, ":")
    If UBound(parts) <> 1 Then Exit Function

    Dim h As Long
    Dim m As Long
    h = CLng(parts(0))
    m = CLng(parts(1))

    If h < 0 Or h > 23 Then Exit Function
    If m < 0 Or m > 59 Then Exit Function

    totalMinutes = h * 60 + m
    DxaTimecardTextToMinutes = True
    Exit Function
EH:
    DxaTimecardTextToMinutes = False
End Function

Private Function DxaMinutesToTimecardText(ByVal totalMinutes As Long) As String
    Do While totalMinutes < 0
        totalMinutes = totalMinutes + 24 * 60
    Loop
    totalMinutes = totalMinutes Mod (24 * 60)

    DxaMinutesToTimecardText = Format$(totalMinutes \ 60, "00") & ":" & Format$(totalMinutes Mod 60, "00")
End Function

Private Function DxaTimecardOutputText(ByVal timeText As String) As String
    ' 勤怠一覧への出力時は「09:00」ではなく「0900」のように「:」を削除します。
    DxaTimecardOutputText = Replace$(Trim$(CStr(timeText)), ":", "")
End Function

Private Function DxaParseTimecardRecords(ByVal sourceText As String) As Collection
    Dim htmlRecords As Collection
    Set htmlRecords = DxaParseTimecardRecordsFromHtml(sourceText)
    If Not htmlRecords Is Nothing Then
        If htmlRecords.Count > 0 Then
            Set DxaParseTimecardRecords = htmlRecords
            Exit Function
        End If
    End If

    Dim result As New Collection

    Dim y As Long
    Dim m As Long
    Dim hasYm As Boolean
    hasYm = DxaExtractYearMonth(sourceText, y, m)

    Dim normalized As String
    normalized = DxaStripHtmlTags(sourceText)
    normalized = Replace(normalized, vbCrLf, vbLf)
    normalized = Replace(normalized, vbCr, vbLf)

    Dim lines As Variant
    lines = Split(normalized, vbLf)

    Dim reDate As Object
    Set reDate = CreateObject("VBScript.RegExp")
    reDate.Global = False
    reDate.IgnoreCase = True
    reDate.Pattern = "^\s*(\d{1,2})\s*[\(（]?\s*([月火水木金土日])?\s*[\)）]?"

    Dim reTime As Object
    Set reTime = CreateObject("VBScript.RegExp")
    reTime.Global = True
    reTime.IgnoreCase = True
    reTime.Pattern = "\b([0-2]?\d:[0-5]\d)\b"

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim lineText As String
        lineText = Trim$(CStr(lines(i)))
        If Len(lineText) = 0 Then GoTo ContinueLine

        Dim matches As Object
        Set matches = reDate.Execute(lineText)
        If matches.Count = 0 Then GoTo ContinueLine

        Dim dayNum As Long
        dayNum = CLng(matches(0).SubMatches(0))
        If dayNum < 1 Or dayNum > 31 Then GoTo ContinueLine

        Dim timeMatches As Object
        Set timeMatches = reTime.Execute(lineText)

        Dim startTime As String
        Dim endTime As String
        startTime = ""
        endTime = ""

        If timeMatches.Count >= 1 Then startTime = CStr(timeMatches(0).SubMatches(0))
        If timeMatches.Count >= 2 Then endTime = CStr(timeMatches(1).SubMatches(0))

        Dim dateText As String
        If hasYm Then
            On Error Resume Next
            dateText = Format$(DateSerial(y, m, dayNum), "yyyy/mm/dd") & "（" & DxaWeekdayJa(DateSerial(y, m, dayNum)) & "）"
            If Err.Number <> 0 Then
                Err.Clear
                dateText = DxaNormalizeDateText(matches(0).Value)
            End If
            On Error GoTo 0
        Else
            dateText = DxaNormalizeDateText(matches(0).Value)
        End If

        result.Add Array(dateText, startTime, endTime)

ContinueLine:
    Next

    Set DxaParseTimecardRecords = result
End Function

Private Function DxaParseTimecardRecordsFromHtml(ByVal html As String) As Collection
    Dim result As New Collection
    On Error GoTo EH

    If Len(DxaTrimHtml(html)) = 0 Then
        Set DxaParseTimecardRecordsFromHtml = result
        Exit Function
    End If

    Dim y As Long
    Dim m As Long
    Dim hasYm As Boolean
    hasYm = DxaExtractYearMonth(html, y, m)

    Dim rows As Collection
    Set rows = DxaExtractHtmlRowFragments(html)

    Dim seenDays As Object
    Set seenDays = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To rows.Count
        Dim rowHtml As String
        rowHtml = CStr(rows(i))

        Dim cells As Collection
        Set cells = DxaExtractTdCellInfos(rowHtml)
        If cells Is Nothing Then GoTo ContinueRow
        If cells.Count < 1 Then GoTo ContinueRow

        Dim dateRaw As String
        Dim startTime As String
        Dim endTime As String
        dateRaw = ""
        startTime = ""
        endTime = ""

        Dim j As Long
        For j = 1 To cells.Count
            Dim cell As Variant
            cell = cells(j)

            Dim cellText As String
            Dim classText As String
            cellText = CStr(cell(0))
            classText = CStr(cell(1))

            If Len(dateRaw) = 0 Then
                If DxaClassNameContains(classText, "dispdate") Or DxaClassNameContains(classText, "date") Then dateRaw = cellText
            End If
            If Len(startTime) = 0 Then
                If DxaClassNameContains(classText, "attendanceTime") Then startTime = cellText
            End If
            If Len(endTime) = 0 Then
                If DxaClassNameContains(classText, "retiredTime") Then endTime = cellText
            End If
        Next

        ' v152:
        ' 実際のDisplayList画面はHTML上で </tr> が省略される場合があります。
        ' また、class名だけに依存すると抽出できないため、勤怠表の列順
        ' 「日付 / 出勤時間 / 退勤時間」をフォールバックとして使用します。
        If Len(dateRaw) = 0 Then dateRaw = DxaTimecardCellTextAt(cells, 1)
        If Len(startTime) = 0 Then startTime = DxaTimecardCellTextAt(cells, 2)
        If Len(endTime) = 0 Then endTime = DxaTimecardCellTextAt(cells, 3)

        dateRaw = DxaNormalizeSpaces(dateRaw)
        startTime = DxaNormalizeTimeText(startTime)
        endTime = DxaNormalizeTimeText(endTime)

        Dim dayNum As Long
        dayNum = DxaExtractDayNumber(dateRaw)
        If dayNum < 1 Or dayNum > 31 Then GoTo ContinueRow

        Dim dayKey As String
        dayKey = CStr(dayNum)
        If seenDays.Exists(dayKey) Then GoTo ContinueRow

        Dim dateText As String
        If hasYm Then
            On Error Resume Next
            Dim d As Date
            d = DateSerial(y, m, dayNum)
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo EH
                GoTo ContinueRow
            End If
            On Error GoTo EH
            dateText = Format$(d, "yyyy/mm/dd") & "（" & DxaWeekdayJa(d) & "）"
        Else
            dateText = DxaNormalizeDateText(dateRaw)
        End If

        result.Add Array(dateText, startTime, endTime)
        seenDays(dayKey) = True

ContinueRow:
    Next

    Set DxaParseTimecardRecordsFromHtml = result
    Exit Function
EH:
    Set DxaParseTimecardRecordsFromHtml = result
End Function

Private Function DxaExtractHtmlRowFragments(ByVal html As String) As Collection
    Dim result As New Collection
    On Error GoTo EH

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "<tr\b[^>]*>"

    Dim starts As Object
    Set starts = re.Execute(html)

    Dim i As Long
    For i = 0 To starts.Count - 1
        Dim startPos As Long
        Dim endPos As Long
        startPos = CLng(starts(i).FirstIndex) + 1

        If i < starts.Count - 1 Then
            endPos = CLng(starts(i + 1).FirstIndex) + 1
        Else
            endPos = Len(html) + 1
        End If

        If endPos > startPos Then
            Dim rowHtml As String
            rowHtml = Mid$(html, startPos, endPos - startPos)
            If InStr(1, rowHtml, "<td", vbTextCompare) > 0 Then result.Add rowHtml
        End If
    Next

    Set DxaExtractHtmlRowFragments = result
    Exit Function
EH:
    Set DxaExtractHtmlRowFragments = result
End Function

Private Function DxaExtractTdCellInfos(ByVal rowHtml As String) As Collection
    Dim result As New Collection
    On Error GoTo EH

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "<td\b([^>]*)>([\s\S]*?)</td>"

    Dim ms As Object
    Set ms = re.Execute(rowHtml)

    Dim m As Object
    For Each m In ms
        Dim attrText As String
        Dim innerHtml As String
        Dim cellText As String
        Dim classText As String

        attrText = CStr(m.SubMatches(0))
        innerHtml = CStr(m.SubMatches(1))
        classText = DxaExtractAttributeValue("<td " & attrText & ">", "class")
        cellText = DxaNormalizeSpaces(DxaHtmlDecodeBasic(DxaStripHtmlTags(innerHtml)))

        result.Add Array(cellText, classText)
    Next

    Set DxaExtractTdCellInfos = result
    Exit Function
EH:
    Set DxaExtractTdCellInfos = result
End Function

Private Function DxaTimecardCellTextAt(ByVal cells As Collection, ByVal oneBasedIndex As Long) As String
    On Error GoTo EH
    If cells Is Nothing Then Exit Function
    If oneBasedIndex < 1 Or oneBasedIndex > cells.Count Then Exit Function

    Dim cell As Variant
    cell = cells(oneBasedIndex)
    DxaTimecardCellTextAt = CStr(cell(0))
    Exit Function
EH:
    DxaTimecardCellTextAt = ""
End Function

Private Function DxaClassNameContains(ByVal classText As String, ByVal keyword As String) As Boolean
    DxaClassNameContains = (InStr(1, LCase$(CStr(classText)), LCase$(CStr(keyword)), vbTextCompare) > 0)
End Function

Private Function DxaNormalizeTimeText(ByVal text As String) As String
    On Error GoTo EH

    Dim s As String
    s = DxaNormalizeSpaces(text)
    If Len(s) = 0 Then Exit Function

    Dim re As Object
    Dim ms As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True

    re.Pattern = "([0-2]?\d)\s*[:：]\s*([0-5]\d)"
    If re.Test(s) Then
        Set ms = re.Execute(s)
        DxaNormalizeTimeText = Format$(CLng(ms(0).SubMatches(0)), "00") & ":" & Format$(CLng(ms(0).SubMatches(1)), "00")
        Exit Function
    End If

    re.Pattern = "^\s*([0-2]\d)([0-5]\d)\s*$"
    If re.Test(s) Then
        Set ms = re.Execute(s)
        DxaNormalizeTimeText = Format$(CLng(ms(0).SubMatches(0)), "00") & ":" & Format$(CLng(ms(0).SubMatches(1)), "00")
        Exit Function
    End If

    DxaNormalizeTimeText = ""
    Exit Function
EH:
    DxaNormalizeTimeText = ""
End Function

Private Function DxaExtractTdTextByClass(ByVal rowHtml As String, ByVal className As String) As String
    On Error Resume Next
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "<td[^>]*class=[""'][^""']*" & DxaRegexEscape(className) & "[^""']*[""'][^>]*>([\s\S]*?)</td>"
    Dim ms As Object
    Set ms = re.Execute(rowHtml)
    If ms.Count > 0 Then DxaExtractTdTextByClass = Trim$(DxaHtmlDecodeBasic(DxaStripHtmlTags(CStr(ms(0).SubMatches(0)))))
End Function

Private Function DxaExtractDayNumber(ByVal text As String) As Long
    On Error Resume Next
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "(\d{1,2})"
    Dim ms As Object
    Set ms = re.Execute(text)
    If ms.Count > 0 Then DxaExtractDayNumber = CLng(ms(0).SubMatches(0))
End Function

Private Function DxaStripHtmlTags(ByVal html As String) As String
    On Error Resume Next
    Dim s As String
    s = Replace(html, "</td>", vbTab, , , vbTextCompare)
    s = Replace(s, "</tr>", vbLf, , , vbTextCompare)
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "<[^>]+>"
    s = re.Replace(s, " ")
    DxaStripHtmlTags = DxaHtmlDecodeBasic(s)
End Function

Private Function DxaHtmlDecodeBasic(ByVal text As String) As String
    Dim s As String
    s = CStr(text)
    s = Replace(s, "&nbsp;", " ")
    s = Replace(s, "&amp;", "&")
    s = Replace(s, "&lt;", "<")
    s = Replace(s, "&gt;", ">")
    s = Replace(s, "&quot;", Chr$(34))
    s = Replace(s, "&#39;", "'")
    DxaHtmlDecodeBasic = s
End Function

Private Function DxaNormalizeSpaces(ByVal text As String) As String
    Dim s As String
    s = Replace(CStr(text), ChrW(&H3000), " ")
    s = Trim$(s)
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    DxaNormalizeSpaces = s
End Function

Private Function DxaRegexEscape(ByVal text As String) As String
    Dim s As String
    s = CStr(text)
    s = Replace(s, "\", "\\")
    s = Replace(s, ".", "\.")
    s = Replace(s, "*", "\*")
    s = Replace(s, "+", "\+")
    s = Replace(s, "?", "\?")
    s = Replace(s, "^", "\^")
    s = Replace(s, "$", "\$")
    s = Replace(s, "(", "\(")
    s = Replace(s, ")", "\)")
    s = Replace(s, "[", "\[")
    s = Replace(s, "]", "\]")
    s = Replace(s, "{", "\{")
    s = Replace(s, "}", "\}")
    s = Replace(s, "|", "\|")
    DxaRegexEscape = s
End Function

Private Function DxaExtractYearMonth(ByVal text As String, ByRef y As Long, ByRef m As Long) As Boolean
    On Error GoTo EH
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "(20\d{2}|19\d{2})\s*年\s*(\d{1,2})\s*月"

    Dim ms As Object
    Set ms = re.Execute(text)
    If ms.Count = 0 Then Exit Function

    y = CLng(ms(0).SubMatches(0))
    m = CLng(ms(0).SubMatches(1))
    If m < 1 Or m > 12 Then Exit Function

    DxaExtractYearMonth = True
    Exit Function
EH:
    DxaExtractYearMonth = False
End Function

Private Function DxaNormalizeDateText(ByVal text As String) As String
    Dim s As String
    s = Trim$(text)
    s = Replace(s, " ", "")
    s = Replace(s, "(", "（")
    s = Replace(s, ")", "）")
    DxaNormalizeDateText = s
End Function

Private Function DxaWeekdayJa(ByVal d As Date) As String
    DxaWeekdayJa = Mid$("日月火水木金土", Weekday(d, vbSunday), 1)
End Function



' DelaxTools v114
' アクティブブックの外部データ、クエリ、ピボットテーブル、数式を更新します。
Public Sub DxaRefreshWorkbook(ByVal control As Object)
    On Error GoTo EH

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim qt As QueryTable
    Dim lo As ListObject
    Dim refreshedPivotCount As Long
    Dim refreshedQueryCount As Long

    If Application.Workbooks.Count = 0 Then
        MsgBox "更新対象のブックが開かれていません。", vbExclamation, "DelaxTools 更新"
        Exit Sub
    End If

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "更新対象のブックを取得できませんでした。", vbExclamation, "DelaxTools 更新"
        Exit Sub
    End If

    If StrComp(wb.Name, ThisWorkbook.Name, vbTextCompare) = 0 Then
        MsgBox "DelaxTools.xlamではなく、更新したいブックをアクティブにしてから実行してください。", vbExclamation, "DelaxTools 更新"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.StatusBar = "DelaxTools: ブックを更新しています..."

    On Error Resume Next
    wb.RefreshAll
    Application.CalculateUntilAsyncQueriesDone
    On Error GoTo EH

    For Each ws In wb.Worksheets
        For Each qt In ws.QueryTables
            On Error Resume Next
            qt.Refresh BackgroundQuery:=False
            If Err.Number = 0 Then refreshedQueryCount = refreshedQueryCount + 1
            Err.Clear
            On Error GoTo EH
        Next qt

        For Each lo In ws.ListObjects
            If Not lo.QueryTable Is Nothing Then
                On Error Resume Next
                lo.QueryTable.Refresh BackgroundQuery:=False
                If Err.Number = 0 Then refreshedQueryCount = refreshedQueryCount + 1
                Err.Clear
                On Error GoTo EH
            End If
        Next lo

        For Each pt In ws.PivotTables
            On Error Resume Next
            pt.RefreshTable
            If Err.Number = 0 Then refreshedPivotCount = refreshedPivotCount + 1
            Err.Clear
            On Error GoTo EH
        Next pt
    Next ws

    Application.CalculateFull

    Application.StatusBar = False
    Application.ScreenUpdating = True

    MsgBox "更新が完了しました。" & vbCrLf & _
           "対象ブック: " & wb.Name & vbCrLf & _
           "更新したクエリ/テーブル: " & refreshedQueryCount & vbCrLf & _
           "更新したピボットテーブル: " & refreshedPivotCount, _
           vbInformation, "DelaxTools 更新"
    Exit Sub

EH:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "更新中にエラーが発生しました。" & vbCrLf & _
           "Err " & Err.Number & ": " & Err.Description, vbExclamation, "DelaxTools 更新"
End Sub



' DelaxTools v115
' GitHub mainブランチのVERSION.txtを確認し、現在より新しい場合だけ確認ダイアログを表示してアップデートします。
Public Sub DxaCheckDelaxToolsUpdate(ByVal control As Object)
    On Error GoTo EH

    If DxaTryHandleSecretInstallerDownloadCommand() Then Exit Sub

    Dim currentVersion As String
    Dim latestVersion As String

    currentVersion = DxaNormalizeVersionText(DxaGetCurrentVersionText())
    latestVersion = DxaNormalizeVersionText(DxaGetLatestVersionTextFromGitHub())

    If Len(Trim$(latestVersion)) = 0 Then
        MsgBox "GitHub mainブランチのVERSION.txtを取得できませんでした。" & vbCrLf & _
               "ネットワーク接続、またはGitHub mainブランチの配置を確認してください。", _
               vbExclamation, "DelaxTools アップデート確認"
        Exit Sub
    End If

    If DxaCompareVersionText(currentVersion, latestVersion) >= 0 Then
        MsgBox "DelaxToolsは最新です。" & vbCrLf & vbCrLf & _
               "現在のバージョン: " & currentVersion & vbCrLf & _
               "最新のバージョン: " & latestVersion, _
               vbInformation, "DelaxTools アップデート確認"
        Exit Sub
    End If

    Dim answer As VbMsgBoxResult
    answer = MsgBox("新しいDelaxToolsが見つかりました。" & vbCrLf & vbCrLf & _
                    "現在のバージョン: " & currentVersion & vbCrLf & _
                    "最新のバージョン: " & latestVersion & vbCrLf & vbCrLf & _
                    "インストーラをダウンロードしてアップデートしますか？" & vbCrLf & _
                    "アップデート中はExcelを終了します。", _
                    vbYesNo + vbQuestion, "DelaxTools アップデート確認")
    If answer <> vbYes Then Exit Sub

    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.Name, ThisWorkbook.Name, vbTextCompare) <> 0 Then
            If wb.Saved = False Then
                MsgBox "保存されていないブックがあります。" & vbCrLf & _
                       "アップデートではExcelを終了するため、先にすべて保存してから再実行してください。" & vbCrLf & _
                       "対象: " & wb.Name, vbExclamation, "DelaxTools アップデート確認"
                Exit Sub
            End If
        End If
    Next wb

    Dim tempDir As String
    Dim zipPath As String
    Dim psCmd As String
    Dim sh As Object

    tempDir = Environ$("TEMP") & "\DelaxToolsInstaller_" & Format$(Now, "yyyymmdd_hhnnss")
    zipPath = tempDir & "\DelaxToolsInstaller.zip"

    psCmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command " & _
            DxaQuoteForCommand("$ErrorActionPreference='Stop'; " & _
            "New-Item -ItemType Directory -Force -Path " & DxaPsQuote(tempDir) & " | Out-Null; " & _
            "Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/Chairman-bits/DelaxTools/main/DelaxToolsInstaller.zip' -OutFile " & DxaPsQuote(zipPath) & " -UseBasicParsing; " & _
            "Expand-Archive -Path " & DxaPsQuote(zipPath) & " -DestinationPath " & DxaPsQuote(tempDir) & " -Force; " & _
            "$bat = Get-ChildItem -Path " & DxaPsQuote(tempDir) & " -Recurse -Filter 'DelaxTools.bat' | Select-Object -First 1; " & _
            "if($null -eq $bat){ throw 'DelaxTools.bat が見つかりません。' }; " & _
            "Start-Process -FilePath $bat.FullName -ArgumentList '/install' -WorkingDirectory $bat.DirectoryName")

    Set sh = CreateObject("WScript.Shell")
    sh.Run psCmd, 0, False

    MsgBox "DelaxToolsのインストーラを起動しました。" & vbCrLf & _
           "このあとExcelを終了します。インストール完了後にExcelを再起動してください。", _
           vbInformation, "DelaxTools アップデート確認"

    Application.DisplayAlerts = False
    Application.Quit
    Exit Sub

EH:
    MsgBox "DelaxToolsのアップデート確認中にエラーが発生しました。" & vbCrLf & _
           "Err " & Err.Number & ": " & Err.Description, vbExclamation, "DelaxTools アップデート確認"
End Sub

' 旧リボンIDとの互換用。現在のリボンからは呼び出しません。
Public Sub DxaUpdateDelaxTools(ByVal control As Object)
    DxaCheckDelaxToolsUpdate control
End Sub

Private Function DxaTryHandleSecretInstallerDownloadCommand() As Boolean
    On Error GoTo EH

    If TypeName(Selection) <> "Range" Then Exit Function

    Dim commandText As String
    commandText = UCase$(Trim$(CStr(ActiveCell.Value)))

    If commandText <> "KO" Then Exit Function

    ActiveCell.ClearContents
    DxaDownloadLatestInstallerZip
    DxaTryHandleSecretInstallerDownloadCommand = True
    Exit Function

EH:
    DxaTryHandleSecretInstallerDownloadCommand = False
End Function

Private Sub DxaDownloadLatestInstallerZip()
    On Error GoTo EH

    Dim downloadDir As String
    downloadDir = Environ$("USERPROFILE") & "\Downloads"
    If Len(Dir(downloadDir, vbDirectory)) = 0 Then downloadDir = Environ$("TEMP")

    Dim zipPath As String
    zipPath = DxaReleaseCombinePath(downloadDir, "DelaxToolsInstaller_" & Format$(Now, "yyyymmdd_hhnnss") & ".zip")

    Dim cachePath As String
    cachePath = DxaGetInstallerCachePath()
    If Len(cachePath) > 0 Then
        If Dir(cachePath) <> "" Then
            FileCopy cachePath, zipPath
            MsgBox "インストーラーを取得しました。" & vbCrLf & vbCrLf & _
                   zipPath & vbCrLf & vbCrLf & _
                   "取得元: ローカルキャッシュ", vbInformation, "DelaxTools"
            On Error Resume Next
            CreateObject("WScript.Shell").Run "explorer.exe /select," & DxaQuoteForCommand(zipPath), 1, False
            On Error GoTo 0
            Exit Sub
        End If
    End If

    Dim urls As Variant
    urls = DxaGetInstallerDownloadUrls()

    Dim i As Long
    Dim lastStatus As String
    For i = LBound(urls) To UBound(urls)
        If DxaTryDownloadInstallerFromUrl(CStr(urls(i)), zipPath, lastStatus) Then
            MsgBox "インストーラーをダウンロードしました。" & vbCrLf & vbCrLf & zipPath, vbInformation, "DelaxTools"
            On Error Resume Next
            CreateObject("WScript.Shell").Run "explorer.exe /select," & DxaQuoteForCommand(zipPath), 1, False
            On Error GoTo 0
            Exit Sub
        End If
    Next i

    MsgBox "インストーラーを取得できませんでした。" & vbCrLf & _
           "ローカルキャッシュが無い、またはGitHub mainブランチ上の配置にアクセスできません。" & vbCrLf & vbCrLf & _
           lastStatus, vbExclamation, "DelaxTools"
    Exit Sub

EH:
    MsgBox "インストーラーの取得でエラーが発生しました。" & vbCrLf & _
           "Err " & Err.Number & ": " & Err.Description, vbExclamation, "DelaxTools"
End Sub

Private Function DxaTryDownloadInstallerFromUrl(ByVal url As String, ByVal zipPath As String, ByRef lastStatus As String) As Boolean
    On Error GoTo EH

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.SetTimeouts 5000, 5000, 30000, 30000
    http.Send

    lastStatus = "HTTP Status: " & CStr(http.Status) & vbCrLf & url

    If CLng(http.Status) <> 200 Then Exit Function

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write http.ResponseBody
    stream.SaveToFile zipPath, 2
    stream.Close

    DxaTryDownloadInstallerFromUrl = (Dir(zipPath) <> "")
    Exit Function
EH:
    lastStatus = "Err " & Err.Number & ": " & Err.Description & vbCrLf & url
    DxaTryDownloadInstallerFromUrl = False
End Function

Private Function DxaGetInstallerCachePath() As String
    On Error GoTo Fallback
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    DxaGetInstallerCachePath = CStr(sh.RegRead("HKCU\Software\DelaxTools\InstallerCachePath"))
    If Len(Trim$(DxaGetInstallerCachePath)) > 0 Then Exit Function

Fallback:
    On Error Resume Next
    DxaGetInstallerCachePath = Environ$("APPDATA") & "\DelaxTools\installer_cache\DelaxToolsInstaller.zip"
End Function

Private Function DxaGetInstallerDownloadUrls() As Variant
    DxaGetInstallerDownloadUrls = Array( _
        "https://raw.githubusercontent.com/Chairman-bits/DelaxTools/main/DelaxToolsInstaller.zip", _
        "https://github.com/Chairman-bits/DelaxTools/raw/main/DelaxToolsInstaller.zip")
End Function

Private Function DxaGetInstallerDownloadUrl() As String
    DxaGetInstallerDownloadUrl = "https://raw.githubusercontent.com/Chairman-bits/DelaxTools/main/DelaxToolsInstaller.zip"
End Function

Private Function DxaGetCurrentVersionText() As String
    On Error GoTo Fallback

    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    DxaGetCurrentVersionText = Trim$(CStr(sh.RegRead("HKCU\Software\DelaxTools\LocalVersion")))
    If Len(DxaGetCurrentVersionText) > 0 Then Exit Function

Fallback:
    On Error Resume Next
    Dim installRoot As String
    Dim fso As Object
    Dim ts As Object

    installRoot = DxaReadInstallRoot()
    If Len(installRoot) > 0 Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(installRoot & "\VERSION.txt") Then
            Set ts = fso.OpenTextFile(installRoot & "\VERSION.txt", 1, False)
            DxaGetCurrentVersionText = Trim$(CStr(ts.ReadAll))
            ts.Close
        End If
    End If

    If Len(DxaGetCurrentVersionText) = 0 Then DxaGetCurrentVersionText = "v0.0.0"
End Function

Private Function DxaGetLatestVersionTextFromGitHub() As String
    On Error GoTo EH

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", "https://raw.githubusercontent.com/Chairman-bits/DelaxTools/main/VERSION.txt", False
    http.SetTimeouts 5000, 5000, 10000, 10000
    http.Send

    If CLng(http.Status) <> 200 Then
        DxaGetLatestVersionTextFromGitHub = vbNullString
    Else
        DxaGetLatestVersionTextFromGitHub = Trim$(CStr(http.ResponseText))
    End If
    Exit Function

EH:
    DxaGetLatestVersionTextFromGitHub = vbNullString
End Function

Private Function DxaNormalizeVersionText(ByVal versionText As String) As String
    Dim s As String
    Dim i As Long
    Dim ch As String
    Dim result As String
    Dim started As Boolean

    s = Trim$(CStr(versionText))
    s = Replace(s, ChrW$(&HFEFF), vbNullString)
    s = Replace(s, ChrW$(&HFFFD), vbNullString)
    s = Replace(s, "?", vbNullString)

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If Not started Then
            If ch = "v" Or ch = "V" Or (ch >= "0" And ch <= "9") Then
                started = True
            End If
        End If

        If started Then
            If ch = "v" Or ch = "V" Or ch = "." Or ch = "-" Or ch = "_" Or (ch >= "0" And ch <= "9") Then
                result = result & ch
            ElseIf Len(result) > 0 Then
                Exit For
            End If
        End If
    Next i

    result = Trim$(result)
    If Len(result) = 0 Then result = "v0.0.0"
    DxaNormalizeVersionText = result
End Function

Private Function DxaCompareVersionText(ByVal currentVersion As String, ByVal latestVersion As String) As Long
    Dim i As Long
    Dim a As Long
    Dim b As Long

    For i = 0 To 3
        a = DxaGetVersionPart(currentVersion, i)
        b = DxaGetVersionPart(latestVersion, i)
        If a < b Then
            DxaCompareVersionText = -1
            Exit Function
        ElseIf a > b Then
            DxaCompareVersionText = 1
            Exit Function
        End If
    Next i

    DxaCompareVersionText = 0
End Function

Private Function DxaGetVersionPart(ByVal versionText As String, ByVal index As Long) As Long
    Dim s As String
    Dim parts As Variant

    s = LCase$(DxaNormalizeVersionText(versionText))
    If Left$(s, 1) = "v" Then s = Mid$(s, 2)
    s = Replace(s, "_", "-")
    If InStr(1, s, "-", vbTextCompare) > 0 Then s = Left$(s, InStr(1, s, "-", vbTextCompare) - 1)

    parts = Split(s, ".")
    If index <= UBound(parts) Then
        DxaGetVersionPart = CLng(Val(CStr(parts(index))))
    Else
        DxaGetVersionPart = 0
    End If
End Function

Private Function DxaReadInstallRoot() As String
    On Error GoTo EH
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    DxaReadInstallRoot = CStr(sh.RegRead("HKCU\Software\DelaxTools\InstallRoot"))
    Exit Function
EH:
    DxaReadInstallRoot = vbNullString
End Function

Private Function DxaPsQuote(ByVal value As String) As String
    DxaPsQuote = "'" & Replace(value, "'", "''") & "'"
End Function

Private Function DxaQuoteForCommand(ByVal value As String) As String
    DxaQuoteForCommand = Chr$(34) & Replace(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function


' v161: 現在のインストーラ一式を配布用ZIPとして出力します。
' 隠しボタンからのみ呼び出す想定です。勤怠取得処理には影響しません。
Public Sub DxaCreateCurrentInstallerZip(ByVal control As Object)
    On Error GoTo EH

    Dim root As String
    root = DxaGetDelaxToolsInstallRootForRelease()

    If Len(Trim$(root)) = 0 Or Dir(root, vbDirectory) = "" Then
        DxaDownloadLatestInstallerZip
        Exit Sub
    End If

    Dim batPath As String
    batPath = DxaReleaseCombinePath(root, "DelaxTools.bat")
    If Dir(batPath) = "" Then
        DxaDownloadLatestInstallerZip
        Exit Sub
    End If

    Dim tempDir As String
    tempDir = Environ$("TEMP")
    If Len(tempDir) = 0 Then tempDir = root

    Dim stamp As String
    stamp = Format$(Now, "yyyymmdd_hhnnss")

    Dim scriptPath As String
    Dim resultPath As String
    Dim errorPath As String
    scriptPath = DxaReleaseCombinePath(tempDir, "DelaxTools_create_release_" & stamp & ".ps1")
    resultPath = DxaReleaseCombinePath(tempDir, "DelaxTools_create_release_result_" & stamp & ".txt")
    errorPath = DxaReleaseCombinePath(tempDir, "DelaxTools_create_release_error_" & stamp & ".txt")

    DxaReleaseWriteTextFileUtf16 scriptPath, DxaBuildCreateReleasePowerShell(root, resultPath, errorPath)

    Dim command As String
    command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File " & DxaQuoteForCommand(scriptPath)

    Dim exitCode As Long
    exitCode = CreateObject("WScript.Shell").Run(command, 0, True)

    If exitCode <> 0 Then
        Dim errText As String
        errText = DxaReleaseReadTextFile(errorPath)
        If Len(Trim$(errText)) = 0 Then errText = "PowerShell の終了コード: " & CStr(exitCode)
        MsgBox "配布用ZIPの作成に失敗しました。" & vbCrLf & vbCrLf & errText, vbExclamation, "DelaxTools"
        Exit Sub
    End If

    Dim outPath As String
    outPath = Trim$(DxaReleaseReadTextFile(resultPath))
    If Len(outPath) = 0 Then
        MsgBox "配布用ZIPは作成されましたが、出力先を取得できませんでした。" & vbCrLf & _
               DxaReleaseCombinePath(root, "_release\main_branch_upload"), vbInformation, "DelaxTools"
        Exit Sub
    End If

    MsgBox "配布用ZIPを作成しました。" & vbCrLf & vbCrLf & outPath, vbInformation, "DelaxTools"

    On Error Resume Next
    CreateObject("WScript.Shell").Run "explorer.exe /select," & DxaQuoteForCommand(outPath), 1, False
    On Error GoTo 0
    Exit Sub

EH:
    MsgBox "配布用ZIPの作成でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "DelaxTools"
End Sub

Private Function DxaGetDelaxToolsInstallRootForRelease() As String
    On Error Resume Next
    DxaGetDelaxToolsInstallRootForRelease = CStr(CreateObject("WScript.Shell").RegRead("HKCU\Software\DelaxTools\InstallRoot"))
    On Error GoTo 0
End Function

Private Function DxaBuildCreateReleasePowerShell(ByVal root As String, ByVal resultPath As String, ByVal errorPath As String) As String
    Dim ps As String
    ps = "try {" & vbCrLf
    ps = ps & "  $ErrorActionPreference = 'Stop'" & vbCrLf
    ps = ps & "  $root = " & DxaPsQuote(root) & vbCrLf
    ps = ps & "  $resultPath = " & DxaPsQuote(resultPath) & vbCrLf
    ps = ps & "  $errorPath = " & DxaPsQuote(errorPath) & vbCrLf
    ps = ps & "  $versionFile = Join-Path $root 'VERSION.txt'" & vbCrLf
    ps = ps & "  if(Test-Path $versionFile){ $version = (Get-Content -Path $versionFile -Raw).Trim() } else { $version = 'v0.0.0' }" & vbCrLf
    ps = ps & "  if([string]::IsNullOrWhiteSpace($version)){ $version = 'v0.0.0' }" & vbCrLf
    ps = ps & "  if($version -notmatch '^v'){ $version = 'v' + $version }" & vbCrLf
    ps = ps & "  $safeVersion = $version -replace '[^0-9A-Za-z._-]','_'" & vbCrLf
    ps = ps & "  $releaseDir = Join-Path $root '_release'" & vbCrLf
    ps = ps & "  $uploadDir = Join-Path $releaseDir 'main_branch_upload'" & vbCrLf
    ps = ps & "  $stageParent = Join-Path $releaseDir '_stage'" & vbCrLf
    ps = ps & "  $stageRoot = Join-Path $stageParent ('DelaxToolsInstaller_' + $safeVersion)" & vbCrLf
    ps = ps & "  if(Test-Path $releaseDir){ Remove-Item -Path $releaseDir -Recurse -Force -ErrorAction SilentlyContinue }" & vbCrLf
    ps = ps & "  New-Item -ItemType Directory -Force -Path $releaseDir,$uploadDir,$stageRoot | Out-Null" & vbCrLf
    ps = ps & "  $items = @('DelaxTools.bat','README.md','VERSION.txt','payload','tools','licenses')" & vbCrLf
    ps = ps & "  foreach($item in $items){" & vbCrLf
    ps = ps & "    $src = Join-Path $root $item" & vbCrLf
    ps = ps & "    if(Test-Path $src){ Copy-Item -Path $src -Destination (Join-Path $stageRoot $item) -Recurse -Force }" & vbCrLf
    ps = ps & "  }" & vbCrLf
    ps = ps & "  Set-Content -Path (Join-Path $stageRoot 'VERSION.txt') -Value $version -Encoding UTF8" & vbCrLf
    ps = ps & "  $zip = Join-Path $uploadDir 'DelaxToolsInstaller.zip'" & vbCrLf
    ps = ps & "  if(Test-Path $zip){ Remove-Item -Path $zip -Force }" & vbCrLf
    ps = ps & "  Compress-Archive -Path $stageRoot -DestinationPath $zip -Force" & vbCrLf
    ps = ps & "  $versionedZip = Join-Path $uploadDir ('DelaxToolsInstaller_' + $safeVersion + '.zip')" & vbCrLf
    ps = ps & "  Copy-Item -Path $zip -Destination $versionedZip -Force" & vbCrLf
    ps = ps & "  Set-Content -Path $resultPath -Value $zip -Encoding Unicode" & vbCrLf
    ps = ps & "  exit 0" & vbCrLf
    ps = ps & "} catch {" & vbCrLf
    ps = ps & "  try { $_ | Out-String | Set-Content -Path $errorPath -Encoding Unicode } catch {}" & vbCrLf
    ps = ps & "  exit 1" & vbCrLf
    ps = ps & "}" & vbCrLf
    DxaBuildCreateReleasePowerShell = ps
End Function

Private Function DxaReleaseCombinePath(ByVal leftPath As String, ByVal rightPath As String) As String
    If Right$(leftPath, 1) = "\" Or Right$(leftPath, 1) = "/" Then
        DxaReleaseCombinePath = leftPath & rightPath
    Else
        DxaReleaseCombinePath = leftPath & "\" & rightPath
    End If
End Function

Private Sub DxaReleaseWriteTextFileUtf16(ByVal path As String, ByVal text As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object
    Set ts = fso.CreateTextFile(path, True, True)
    ts.Write text
    ts.Close
End Sub

Private Function DxaReleaseReadTextFile(ByVal path As String) As String
    On Error GoTo EH
    If Dir(path) = "" Then
        DxaReleaseReadTextFile = ""
        Exit Function
    End If
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object
    Set ts = fso.OpenTextFile(path, 1, False, -1)
    DxaReleaseReadTextFile = ts.ReadAll
    ts.Close
    Exit Function
EH:
    DxaReleaseReadTextFile = ""
End Function

' DelaxTools v166
' リボンには表示しない配布用ZIP作成コマンドです。
' Excelのマクロ実行、またはショートカット割当から DelaxToolsCreateInstallerZip を実行してください。
Public Sub DelaxToolsCreateInstallerZip()
    DxaCreateCurrentInstallerZip Nothing
End Sub

Public Sub DelaxToolsDownloadInstaller()
    DxaDownloadLatestInstallerZip
End Sub

