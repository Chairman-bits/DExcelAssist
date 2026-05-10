VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DExcelAssistAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents App As Excel.Application

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    On Error Resume Next
    DxaEnsureSnapshotForWorkbook Wb
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    On Error Resume Next
    DxaEnsureSnapshotForWorkbook Wb
End Sub

Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    On Error Resume Next
    DxaDeleteSnapshotForWorkbook Wb
End Sub
