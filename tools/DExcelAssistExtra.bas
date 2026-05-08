Attribute VB_Name = "DExcelAssistExtra"

Option Explicit



' DExcelAssist v99

' 自動アップデート機能は含めていません。

' 追加機能はExcel内VBAとして実行します。



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
    MsgBox "シート一覧を作成しました。", vbInformation, "DExcelAssist"
    Exit Sub
EH:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "シート一覧でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "DExcelAssist"
End Sub

Public Sub DxaBacklogGroupByParent(ByVal control As Object)

    親課題単位でグループ化_Core

End Sub



Public Sub 親課題単位でグループ化_Core()



    On Error GoTo EH



    Dim wsTarget As Worksheet

    Dim wsParent As Worksheet



    Dim lastRowTarget As Long

    Dim lastRowParent As Long



    Dim parentDict As Object



    Dim rowIndex As Long

    Dim groupStartRow As Long

    Dim nextParentRow As Long



    Dim cellValue As String



    Set wsTarget = ActiveSheet

    Set wsParent = ThisWorkbook.Worksheets("親課題一覧")



    Set parentDict = CreateObject("Scripting.Dictionary")



    lastRowParent = wsParent.Cells(wsParent.Rows.Count, "A").End(xlUp).Row



    For rowIndex = 1 To lastRowParent

        cellValue = Trim(CStr(wsParent.Cells(rowIndex, "A").Value))

        If cellValue <> "" Then

            parentDict(cellValue) = True

        End If

    Next rowIndex



    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row



    On Error Resume Next

    wsTarget.Rows.ClearOutline

    On Error GoTo EH



    If lastRowTarget >= 5 Then

        wsTarget.Range("A5:A" & lastRowTarget).IndentLevel = 0

    End If



    rowIndex = 5



    Do While rowIndex <= lastRowTarget



        cellValue = Trim(CStr(wsTarget.Cells(rowIndex, "A").Value))



        If parentDict.Exists(cellValue) Then



            groupStartRow = rowIndex

            nextParentRow = 0



            With wsTarget.Cells(rowIndex, "A").Font

                .Bold = True

                .Size = .Size + 4

            End With



            With wsTarget.Cells(rowIndex, "C").Font

                .Bold = True

                .Size = .Size + 4

            End With



            Dim searchRow As Long

            For searchRow = rowIndex + 1 To lastRowTarget

                cellValue = Trim(CStr(wsTarget.Cells(searchRow, "A").Value))



                If parentDict.Exists(cellValue) Then

                    nextParentRow = searchRow

                    Exit For

                End If

            Next searchRow



            Dim startChild As Long

            Dim endChild As Long



            startChild = groupStartRow + 1



            If nextParentRow > 0 Then

                endChild = nextParentRow - 1

            Else

                endChild = lastRowTarget

            End If



            If startChild <= endChild Then

                wsTarget.Rows(startChild & ":" & endChild).Group



                Dim i As Long

                For i = startChild To endChild

                    With wsTarget.Cells(i, "A")

                        .IndentLevel = .IndentLevel + 1

                    End With

                Next i

            End If



            If nextParentRow > 0 Then

                rowIndex = nextParentRow

            Else

                Exit Do

            End If



        Else

            rowIndex = rowIndex + 1

        End If



    Loop



    MsgBox "グループ化＋インデント完了", vbInformation

    Exit Sub

EH:

    MsgBox "親課題でグループ化中にエラーが発生しました。" & vbCrLf & _

           "対象シートと、このブック内の『親課題一覧』シートを確認してください。" & vbCrLf & _

           Err.Description, vbExclamation, "DExcelAssist"

End Sub






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

           "出力数: " & CStr(count), vbInformation, "DExcelAssist"

    Exit Sub

EH:

    MsgBox "VBAエクスポートでエラーが発生しました。" & vbCrLf & _

           "Excelの『VBAプロジェクト オブジェクト モデルへのアクセスを信頼する』が必要です。" & vbCrLf & _

           Err.Description, vbExclamation, "DExcelAssist"

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

