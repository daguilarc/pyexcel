Attribute VB_Name = "pyexcel_macros"
Option Explicit

' pyexcel_macros
' --------------
' VBA module exposing macro entry points for the Python tools in this
' repo (isolation_forest.py, tabpfn_forecast.py). Requires the xlwings
' Excel add-in to be installed (RunPython is provided by that add-in).
' The workbook holding this module must be saved as .xlsm to keep VBA.
' Tool settings live in PyexcelSettings.xlsx, created beside the
' scripts on first run.
'
' Usage: Alt+F11 -> File -> Import File -> select this .bas file.
' Then Alt+F8 -> RunIsolationForest or RunForecast.

' Full path to the folder containing isolation_forest.py and
' tabpfn_forecast.py, e.g. "C:\Users\you\pyexcel". Leave empty ("") to
' rely on the xlwings PYTHONPATH setting instead.
Const SCRIPT_DIR As String = ""

' Builds and runs the Python command for the given module/function.
' The only place the sys.path logic lives.
Private Sub RunTool(moduleName As String, funcName As String)
    Dim cmd As String

    If Len(SCRIPT_DIR) > 0 Then
        cmd = "import sys; sys.path.insert(0, r'" & SCRIPT_DIR & "'); import " & _
              moduleName & "; " & moduleName & "." & funcName & "()"
    Else
        cmd = "import " & moduleName & "; " & moduleName & "." & funcName & "()"
    End If

    RunPython cmd
End Sub

Public Sub RunIsolationForest()
    RunTool "isolation_forest", "run_isolation_forest"
End Sub

Public Sub RunForecast()
    RunTool "tabpfn_forecast", "run_forecast"
End Sub
