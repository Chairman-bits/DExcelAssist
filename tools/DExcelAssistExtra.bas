Attribute VB_Name = "DExcelAssistExtra"
Option Explicit

' DExcelAssist v84 extra commands.
' Ribbon callbacks use Object instead of IRibbonControl to avoid reference issues.

Public Sub DxaCreateHolidaySheet(ByVal control As Object)
    On Error GoTo EH
    Dim yText As String
    yText = InputBox("休日一覧を作成する年を入力してください。", "休日シート作成", CStr(Year(Date)))
    If Len(Trim$(yText)) = 0 Then Exit Sub
    If Not IsNumeric(yText) Then
        MsgBox "年は数値で入力してください。", vbExclamation, "DExcelAssist"
        Exit Sub
    End If

    Dim y As Long
    y = CLng(yText)
    If y < 1900 Or y > 2100 Then
        MsgBox "1900～2100の範囲で入力してください。", vbExclamation, "DExcelAssist"
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
    MsgBox CStr(y) & "年の休日一覧を作成しました。", vbInformation, "DExcelAssist"
    Exit Sub
EH:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "休日シート作成でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "DExcelAssist"
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
    MsgBox "全シートの倍率を100%にしました。", vbInformation, "DExcelAssist"
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "全シート倍率100%でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "DExcelAssist"
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
    MsgBox "実行シートの列幅を自動調整しました。", vbInformation, "DExcelAssist"
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "列幅自動調整でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "DExcelAssist"
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
    MsgBox "実行シートの行高さを自動調整しました。", vbInformation, "DExcelAssist"
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "行高さ自動調整でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "DExcelAssist"
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
        MsgBox "選択範囲の半角英数字を全角に変換しました。", vbInformation, "DExcelAssist"
    Else
        MsgBox "選択範囲の全角英数字を半角に変換しました。", vbInformation, "DExcelAssist"
    End If
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "文字変換でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "DExcelAssist"
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
