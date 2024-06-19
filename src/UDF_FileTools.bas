Attribute VB_Name = "UDF_FileTools"
'''=============================================================================
''' VBA FileTools
''' ---------------------------------------------
''' https://github.com/cristianbuse/VBA-FileTools
''' ---------------------------------------------
''' MIT License
'''
''' Copyright (c) 2012 Ion Cristian Buse
'''
''' Permission is hereby granted, free of charge, to any person obtaining a copy
''' of this software and associated documentation files (the "Software"), to
''' deal in the Software without restriction, including without limitation the
''' rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
''' sell copies of the Software, and to permit persons to whom the Software is
''' furnished to do so, subject to the following conditions:
'''
''' The above copyright notice and this permission notice shall be included in
''' all copies or substantial portions of the Software.
'''
''' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
''' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
''' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
''' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
''' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
''' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
''' IN THE SOFTWARE.
'''=============================================================================

Option Explicit

'*******************************************************************************
''Excel File Tools Module
''Functions in this module make it easy for the user to work with file paths
''   within the Excel interface using the FileTools library
''Functions below are capable to 'spill' on newer Excel versions (Office 365)
'*******************************************************************************

''Important!
'*******************************************************************************
''This module is intended to be used in Microsoft Excel only!
''Call the User-Defined-Functions (UDFs) in this module from Excel Ranges only
''  DO NOT call these functions from VBA! If you need any of the functions below
''  directly in VBA then use their equivalent from the LibFileTools module
'*******************************************************************************

''Requires:
''  - LibFileTools: library module

''Exposed Excel UDFs:
''  - IS_FILE
''  - IS_FOLDER
''  - LOCAL_PATH
''  - REMOTE_PATH

'*******************************************************************************
'Turn the below compiler constant to True if you are using the LibUDFs library
'https://github.com/cristianbuse/VBA-FastExcelUDFs
#Const USE_LIB_FAST_UDFS = False
'*******************************************************************************

Public Function IS_FILE(ByRef filePaths As Variant) As Variant
    Application.Volatile False
    #If USE_LIB_FAST_UDFS Then
        LibUDFs.TriggerFastUDFCalculation
    #End If
    '
    'Only accept 1-Area Ranges. This could alternatively be changed to ignore
    '   the extra Areas by arr = arr.Areas(1).Value2 instead of the 2 lines
    If VBA.TypeName(filePaths) = "Range" Then
        If filePaths.Areas.Count > 1 Then GoTo FailInput
        filePaths = filePaths.Value2
    End If
    '
    On Error GoTo ErrorHandler
    If Not IsArray(filePaths) Then
        IS_FILE = LibFileTools.IsFile(CStr(filePaths))
        Exit Function
    End If
    '
    Dim res() As Boolean
    Dim i As Long
    Dim j As Long
    '
    Select Case GetArrayDimsCount(filePaths)
    Case 1
        ReDim res(LBound(filePaths) To UBound(filePaths))
        For i = LBound(filePaths) To UBound(filePaths)
            res(i) = LibFileTools.IsFile(CStr(filePaths(i)))
        Next i
    Case 2
        ReDim res(LBound(filePaths, 1) To UBound(filePaths, 1) _
                , LBound(filePaths, 2) To UBound(filePaths, 2))
        For i = LBound(filePaths, 1) To UBound(filePaths, 1)
            For j = LBound(filePaths, 2) To UBound(filePaths, 2)
                res(i, j) = LibFileTools.IsFile(CStr(filePaths(i, j)))
            Next j
        Next i
    Case Else
        GoTo FailInput
    End Select
    '
    IS_FILE = res
Exit Function
ErrorHandler:
FailInput:
    IS_FILE = VBA.CVErr(xlErrValue)
End Function

Public Function IS_FOLDER(ByRef folderPaths As Variant) As Variant
    Application.Volatile False
    #If USE_LIB_FAST_UDFS Then
        LibUDFs.TriggerFastUDFCalculation
    #End If
    '
    'Only accept 1-Area Ranges. This could alternatively be changed to ignore
    '   the extra Areas by arr = arr.Areas(1).Value2 instead of the 2 lines
    If VBA.TypeName(folderPaths) = "Range" Then
        If folderPaths.Areas.Count > 1 Then GoTo FailInput
        folderPaths = folderPaths.Value2
    End If
    '
    On Error GoTo ErrorHandler
    If Not IsArray(folderPaths) Then
        IS_FOLDER = LibFileTools.IsFolder(CStr(folderPaths))
        Exit Function
    End If
    '
    Dim res() As Boolean
    Dim i As Long
    Dim j As Long
    '
    Select Case GetArrayDimsCount(folderPaths)
    Case 1
        ReDim res(LBound(folderPaths) To UBound(folderPaths))
        For i = LBound(folderPaths) To UBound(folderPaths)
            res(i) = LibFileTools.IsFolder(CStr(folderPaths(i)))
        Next i
    Case 2
        ReDim res(LBound(folderPaths, 1) To UBound(folderPaths, 1) _
                , LBound(folderPaths, 2) To UBound(folderPaths, 2))
        For i = LBound(folderPaths, 1) To UBound(folderPaths, 1)
            For j = LBound(folderPaths, 2) To UBound(folderPaths, 2)
                res(i, j) = LibFileTools.IsFolder(CStr(folderPaths(i, j)))
            Next j
        Next i
    Case Else
        GoTo FailInput
    End Select
    '
    IS_FOLDER = res
Exit Function
ErrorHandler:
FailInput:
    IS_FOLDER = VBA.CVErr(xlErrValue)
End Function

Public Function LOCAL_PATH(ByRef fullPaths As Variant) As Variant
    Application.Volatile False
    #If USE_LIB_FAST_UDFS Then
        LibUDFs.TriggerFastUDFCalculation
    #End If
    '
    'Only accept 1-Area Ranges. This could alternatively be changed to ignore
    '   the extra Areas by arr = arr.Areas(1).Value2 instead of the 2 lines
    If VBA.TypeName(fullPaths) = "Range" Then
        If fullPaths.Areas.Count > 1 Then GoTo FailInput
        fullPaths = fullPaths.Value2
    End If
    '
    On Error GoTo ErrorHandler
    If Not IsArray(fullPaths) Then
        LOCAL_PATH = LibFileTools.GetLocalPath(CStr(fullPaths))
        Exit Function
    End If
    '
    Dim res() As String
    Dim i As Long
    Dim j As Long
    '
    Select Case GetArrayDimsCount(fullPaths)
    Case 1
        ReDim res(LBound(fullPaths) To UBound(fullPaths))
        For i = LBound(fullPaths) To UBound(fullPaths)
            res(i) = LibFileTools.GetLocalPath(CStr(fullPaths(i)))
        Next i
    Case 2
        ReDim res(LBound(fullPaths, 1) To UBound(fullPaths, 1) _
                , LBound(fullPaths, 2) To UBound(fullPaths, 2))
        For i = LBound(fullPaths, 1) To UBound(fullPaths, 1)
            For j = LBound(fullPaths, 2) To UBound(fullPaths, 2)
                res(i, j) = LibFileTools.GetLocalPath(CStr(fullPaths(i, j)))
            Next j
        Next i
    Case Else
        GoTo FailInput
    End Select
    '
    LOCAL_PATH = res
Exit Function
ErrorHandler:
FailInput:
    LOCAL_PATH = VBA.CVErr(xlErrValue)
End Function

Public Function REMOTE_PATH(ByRef fullPaths As Variant) As Variant
    Application.Volatile False
    #If USE_LIB_FAST_UDFS Then
        LibUDFs.TriggerFastUDFCalculation
    #End If
    '
    'Only accept 1-Area Ranges. This could alternatively be changed to ignore
    '   the extra Areas by arr = arr.Areas(1).Value2 instead of the 2 lines
    If VBA.TypeName(fullPaths) = "Range" Then
        If fullPaths.Areas.Count > 1 Then GoTo FailInput
        fullPaths = fullPaths.Value2
    End If
    '
    On Error GoTo ErrorHandler
    If Not IsArray(fullPaths) Then
        REMOTE_PATH = LibFileTools.GetRemotePath(CStr(fullPaths))
        Exit Function
    End If
    '
    Dim res() As String
    Dim i As Long
    Dim j As Long
    '
    Select Case GetArrayDimsCount(fullPaths)
    Case 1
        ReDim res(LBound(fullPaths) To UBound(fullPaths))
        For i = LBound(fullPaths) To UBound(fullPaths)
            res(i) = LibFileTools.GetRemotePath(CStr(fullPaths(i)))
        Next i
    Case 2
        ReDim res(LBound(fullPaths, 1) To UBound(fullPaths, 1) _
                , LBound(fullPaths, 2) To UBound(fullPaths, 2))
        For i = LBound(fullPaths, 1) To UBound(fullPaths, 1)
            For j = LBound(fullPaths, 2) To UBound(fullPaths, 2)
                res(i, j) = LibFileTools.GetRemotePath(CStr(fullPaths(i, j)))
            Next j
        Next i
    Case Else
        GoTo FailInput
    End Select
    '
    REMOTE_PATH = res
Exit Function
ErrorHandler:
FailInput:
    REMOTE_PATH = VBA.CVErr(xlErrValue)
End Function

'*******************************************************************************
'Returns the Number of dimensions for an input array
'Returns 0 if array is uninitialized or input not an array
'Note that a zero-length array has 1 dimension! Ex. Array() bounds are (0 to -1)
'*******************************************************************************
Private Function GetArrayDimsCount(ByRef arr As Variant) As Long
    Const MAX_DIMENSION As Long = 60 'VB limit
    Dim dimension As Long
    Dim tempBound As Long
    '
    On Error GoTo FinalDimension
    For dimension = 1 To MAX_DIMENSION
        tempBound = LBound(arr, dimension)
    Next dimension
FinalDimension:
    GetArrayDimsCount = dimension - 1
End Function
