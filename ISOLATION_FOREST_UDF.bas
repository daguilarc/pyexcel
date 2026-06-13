Attribute VB_Name = "ISOLATION_FOREST_UDF"
Option Explicit

' Excel UDF wrapper for Isolation Forest anomaly detection
' Uses RunPython to call the Python function via xlwings add-in
' Returns variable array of anomaly scores

Function IsolationForest(ParamArray args() As Variant) As Variant
    
    On Error GoTo ErrorHandler
    
    ' Get script directory and working directory
    Dim scriptDir As String
    Dim folder As String
    scriptDir = ThisWorkbook.Path
    folder = ActiveWorkbook.Path
    
    ' Parse all arguments - Python will handle parameter detection
    ' Parameters can be: data arrays, contamination (0-100%), monte_carlo_samples (100-10000), inversion (bool), epsilon (bool)
    Dim numArgs As Long
    numArgs = UBound(args) - LBound(args) + 1
    
    If numArgs = 0 Then
        IsolationForest = CVErr(xlErrValue)
        Exit Function
    End If
    
    ' Serialize all arguments (Python will parse them)
    Dim argStrs() As String
    ReDim argStrs(0 To numArgs - 1)
    Dim i As Long
    Dim arg As Variant
    Dim argStr As String
    
    For i = 0 To numArgs - 1
        arg = args(i)
        
        ' Handle different argument types
        If TypeName(arg) = "Range" Then
            ' Range: convert to array and serialize
            argStr = ArrayToPythonList(arg.Value)
        ElseIf IsArray(arg) Then
            ' Array: serialize directly
            argStr = ArrayToPythonList(arg)
        ElseIf IsNumeric(arg) Then
            ' Numeric: pass as number
            argStr = CStr(CDbl(arg))
        ElseIf TypeName(arg) = "Boolean" Then
            ' Boolean: convert to Python True/False
            argStr = IIf(arg, "True", "False")
        Else
            ' Other: try to convert to string
            argStr = CStr(arg)
        End If
        
        argStrs(i) = argStr
    Next i
    
    ' Build Python command with all arguments
    Dim cmd As String
    Dim argsStr As String
    argsStr = Join(argStrs, ", ")
    
    cmd = "import os, sys;" & _
          "os.chdir(r'" & folder & "');" & _
          "sys.path.append(r'" & scriptDir & "');" & _
          "from isolation_forest import ISOLATION_FOREST as py_isolation_forest;" & _
          "result = py_isolation_forest(" & _
          argsStr & _
          ");" & _
          "result"
    
    ' Execute Python and return result
    IsolationForest = RunPython(cmd)
    Exit Function
    
ErrorHandler:
    IsolationForest = CVErr(xlErrValue)
End Function

' Helper function to convert VBA array to Python list string
Private Function ArrayToPythonList(arr As Variant) As String
    On Error GoTo ErrorHandler
    
    Dim i As Long, j As Long
    Dim result As String
    Dim rowStr As String
    Dim hasCols As Boolean
    
    ' Handle non-array (single value)
    If Not IsArray(arr) Then
        ArrayToPythonList = "[" & CDbl(arr) & "]"
        Exit Function
    End If
    
    ' Check if 1D or 2D by testing 2D access
    On Error Resume Next
    Dim testVal As Variant
    testVal = arr(LBound(arr), LBound(arr, 2))
    hasCols = (Err.Number = 0)
    On Error GoTo ErrorHandler
    
    If Not hasCols Then
        ' 1D array (column vector)
        result = "["
        For i = LBound(arr) To UBound(arr)
            If i > LBound(arr) Then result = result & ", "
            result = result & "[" & CDbl(arr(i)) & "]"
        Next i
        result = result & "]"
    Else
        ' 2D array
        result = "["
        For i = LBound(arr, 1) To UBound(arr, 1)
            If i > LBound(arr, 1) Then result = result & ", "
            rowStr = "["
            For j = LBound(arr, 2) To UBound(arr, 2)
                If j > LBound(arr, 2) Then rowStr = rowStr & ", "
                rowStr = rowStr & CDbl(arr(i, j))
            Next j
            rowStr = rowStr & "]"
            result = result & rowStr
        Next i
        result = result & "]"
    End If
    
    ArrayToPythonList = result
    Exit Function
    
ErrorHandler:
    ArrayToPythonList = "[]"
End Function
