VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DelaxToolsAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mApp As Excel.Application

Public Sub Init(ByVal targetApp As Excel.Application)
    Set mApp = targetApp
End Sub

Private Sub mApp_WorkbookOpen(ByVal Wb As Workbook)
    ' v172: 重量化防止のため、ブックを開いた時の自動処理は行いません。
End Sub

Private Sub mApp_WorkbookActivate(ByVal Wb As Workbook)
    ' v172: 重量化防止のため、ブック切替時の自動処理は行いません。
End Sub

Private Sub mApp_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    ' v172: 重量化防止のため、セル選択時の常駐処理は行いません。
End Sub

Private Sub mApp_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    ' v172: 重量化防止のため、セル変更時の常駐処理は行いません。
End Sub

Private Sub mApp_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    ' v172: 重量化防止のため、ブック終了時の自動処理は行いません。
End Sub
