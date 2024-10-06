Attribute VB_Name = "modExportLambdas"
'@IgnoreModule UndeclaredVariable
'@Folder "Lambda.Editor.Exporter.Driver"
Option Explicit

'@Ignore ProcedureNotUsed
Public Sub ExportThisWorkbookLambda()
    ExportLambdas ThisWorkbook.name
End Sub

Public Sub ExportLambdas(Optional ByVal WorkbookName As String = vbNullString)
    Exporter.ExportLambdas WorkbookName
End Sub

Public Sub ExportExportModules()
    Exporter.ExportLambdas "Export Module.xlam"
End Sub

