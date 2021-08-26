# VBA-FileTools

FileTools is a small VBA library that is useful for interacting with the file system.

## Installation

Just import the following code module in your VBA Project:

* [**LibFileTools.bas**](src/LibFileTools.bas)

## Usage

Here are a couple of demo method calls. Also saved in the [Demo](src/Demo/DemoLibFileTools.bas) module.

```VBA
Option Explicit

Public Sub DemoMain()
    Dim demoFolder As String
    '
    'BrowseForFolder
    #If Mac Then
        demoFolder = InputBox("Please input a valid folder path! Folder should not be restricted", "Folder Path")
    #Else
        demoFolder = BrowseForFolder(dialogTitle:="Please select a valid folder! Folder should not be restricted")
    #End If
    If demoFolder = vbNullString Then Exit Sub
    '
    'A bit of preparation for the demo
    If Not IsFolder(demoFolder) Then
        Debug.Print "Invalid folder selected. Demo cancelled."
        Exit Sub
    Else
        demoFolder = BuildPath(demoFolder, "Demo")
        If Not CreateFolder(demoFolder) Then
            Debug.Print "Folder is restricted. Demo cancelled."
            Exit Sub
        End If
    End If
    Dim fileNum As Long: fileNum = FreeFile
    Dim demoFile As String: demoFile = BuildPath(demoFolder, "demo.txt")
    '
    Open demoFile For Append Access Write Lock Write As fileNum
    Print #fileNum, "[" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "] Running DemoMain"
    Close #fileNum
    '
    Dim filePath As String
    '
    'BrowseForFiles
    #If Mac Then
    #Else
    With BrowseForFiles(dialogTitle:="Select any file!", allowMultiFiles:=False)
        If .Count <> 0 Then filePath = .Item(1)
    End With
    If filePath = vbNullString Then Exit Sub
    Debug.Print "You have selected: " & filePath
    Debug.Print
    '
    Dim collFiles As Collection
    Dim v As Variant
    '
    Do
        Set collFiles = BrowseForFiles(dialogTitle:="Select at least 2 files!", allowMultiFiles:=True)
        If collFiles.Count = 0 Then Exit Sub
    Loop Until collFiles.Count > 1
    For Each v In collFiles
        Debug.Print "You have selected: " & v
    Next v
    Debug.Print
    #End If
    '
    Stop 'You might want to step through code line by line
    '
    'BuildPath
    #If Mac Then
        Debug.Print "Built path: " & BuildPath("/Users/DemoUser/Desktop/Test", "demo.txt")
        Debug.Print "Built path: " & BuildPath("/Users/DemoUser/Desktop/Test/", "demo.txt")
        Debug.Print "Built path: " & BuildPath("//Users//DemoUser/Desktop///Test", "demo.txt")
        Debug.Print "Built path: " & BuildPath("//Users/DemoUser/Desktop//Test", "Demo/demo.txt")
    #Else
        Debug.Print "Built path: " & BuildPath("C:\Users\DemoUser\Desktop\Test", "demo.txt")
        Debug.Print "Built path: " & BuildPath("C:\Users\DemoUser\Desktop\Test\", "demo.txt")
        Debug.Print "Built path: " & BuildPath("C://\Users/\DemoUser\Desktop\\/Test", "demo.txt")
        Debug.Print "Built path: " & BuildPath("C:\\Users\DemoUser\Desktop\\Test", "Demo/demo.txt")
    #End If
    Debug.Print
    '
    'CreateFolder
    Dim folderPath As String
    '
    folderPath = BuildPath(demoFolder, "/1/2/3/4/5/6/7")
    If CreateFolder(folderPath) Then
        Debug.Print "Created folder: " & folderPath
    Else
        Debug.Print "Oops. Cannot create folder: " & folderPath
        Exit Sub
    End If
    Debug.Print
    '
    'CopyFile
    Dim i As Long, j As Long
    Dim subFolder As String
    '
    For i = 1 To 7
        subFolder = subFolder & i & GetPathSeparator()
        For j = 1 To i
            filePath = Replace(demoFile, "demo.txt", subFolder & j & ".txt")
            If CopyFile(demoFile, filePath) Then
                Debug.Print "Copied file: " & filePath
            Else
                Debug.Print "Oops. Cannot copy file: " & filePath
            End If
            #If Mac Then
            #Else
                SetAttr filePath, vbReadOnly + vbHidden + vbSystem
            #End If
        Next j
    Next i
    Debug.Print
    '
    'CopyFolder
    If CopyFolder(BuildPath(demoFolder, "/1"), BuildPath(demoFolder, "/1.Copy")) Then
        Debug.Print "Copied a folder and it's contents"
    Else
        Debug.Print "Oops. Cannot copy folder"
        Exit Sub
    End If
    Debug.Print
    '
    'DeleteFile
    If DeleteFile(demoFile) Then
        Debug.Print "Deleted demo file: " & demoFile
    Else
        Debug.Print "Oops. Cannot delete demo file: " & demoFile
    End If
    Debug.Print
    'DeleteFolder
    '
    If DeleteFolder(BuildPath(demoFolder, "/1.Copy"), True) Then
        Debug.Print "Deleted folder and it's contents"
    Else
        Debug.Print "Oops. Cannot delete folder"
    End If
    Debug.Print
    '
    'FixFileName
    #If Mac Then
        Const wrongFileName As String = "The : and the / are forbidden"
    #Else
        Const wrongFileName As String = "We canot have :*?""<>|\/ and we cannot end in a space or dot ."
    #End If
    Debug.Print "[" & wrongFileName & "] got fixed to [" & FixFileName(wrongFileName) & "]"
    Debug.Print
    '
    'FixPathSeparators
    #If Mac Then
        Debug.Print "Fixed path: " & FixPathSeparators("/Users/DemoUser/Desktop/Test")
        Debug.Print "Fixed path: " & FixPathSeparators("/Users/DemoUser/Desktop/Test/")
        Debug.Print "Fixed path: " & FixPathSeparators("//Users//DemoUser/Desktop///Test")
        Debug.Print "Fixed path: " & FixPathSeparators("//Users/DemoUser/Desktop//Test")
    #Else
        Debug.Print "Fixed path: " & FixPathSeparators("C:\Users\DemoUser\Desktop\Test")
        Debug.Print "Fixed path: " & FixPathSeparators("C:\Users\DemoUser\Desktop\Test\")
        Debug.Print "Fixed path: " & FixPathSeparators("C://\Users/\DemoUser\Desktop\\/Test")
        Debug.Print "Fixed path: " & FixPathSeparators("C:\\Users\DemoUser\Desktop\\Test")
    #End If
    Debug.Print
    '
    'GetFileOwner
    #If Mac Then
    #Else
        filePath = BuildPath(demoFolder, "/1/2/3/2.txt")
        Debug.Print "The owner of: " & filePath & " is " & GetFileOwner(filePath)
        Debug.Print
    #End If
    '
    Dim f As Variant
    '
    'GetFiles
    folderPath = BuildPath(demoFolder, "/1/2/3/4/5")
    Debug.Print "The files in: " & folderPath & " are:"
    For Each f In GetFiles(folderPath, True, True, True)
        Debug.Print f
    Next f
    Debug.Print
    '
    'GetFolders
    Debug.Print "The folders in: " & demoFolder & " are:"
    For Each f In GetFolders(demoFolder, True, True, True)
        Debug.Print f
    Next f
    Debug.Print
    '
    'GetLocalPath
    'GetUNCPath
    #If Mac Then
    #Else
    With BrowseForFiles(dialogTitle:="Please select a file on a mapped network drive", allowMultiFiles:=False)
        If .Count > 0 Then
            filePath = .Item(1)
            Debug.Print "Local path is: " & GetLocalPath(filePath)
            Debug.Print "UNC path is: " & GetUNCPath(filePath)
            Debug.Print
        End If
    End With
    #End If
    '
    'GetPathSeparator
    Debug.Print "The path separator is: " & GetPathSeparator()
    Debug.Print
    '
    'IsFile
    filePath = demoFile
    Debug.Print "This is " & IIf(IsFile(filePath), vbNullString, "not ") & "a file: " & filePath
    filePath = GetFiles(demoFolder, True, True, True).Item(15)
    Debug.Print "This is " & IIf(IsFile(filePath), vbNullString, "not ") & "a file: " & filePath
    Debug.Print
    '
    'IsFolder
    folderPath = GetFolders(demoFolder, True, True, True).Item(5)
    Debug.Print "This is " & IIf(IsFolder(folderPath), vbNullString, "not ") & "a folder: " & folderPath
    folderPath = "Not a folder"
    Debug.Print "This is " & IIf(IsFolder(folderPath), vbNullString, "not ") & "a folder: " & folderPath
    Debug.Print
    '
    'MoveFile
    filePath = GetFiles(demoFolder, True, True, True).Item(10)
    If MoveFile(filePath, demoFile) Then
        Debug.Print "Moved: " & filePath & " to: " & demoFile
    Else
        Debug.Print "Oops. Cannot move file"
    End If
    Debug.Print
    '
    'MoveFolder
    If MoveFolder(BuildPath(demoFolder, "/1/2/3/4"), BuildPath(demoFolder, "////M")) Then
        Debug.Print "Moved a folder and it's contents"
    Else
        Debug.Print "Oops. Failed to move folder"
    End If
    Debug.Print
    '
    Debug.Print "Finished Demo"
    DeleteFolder folderPath:=demoFolder, deleteContents:=True
End Sub
```

## Notes
* No extra library references are needed (e.g. Microsoft Scripting Runtime)
* Works in any host Application (Excel, Word, AutoCAD etc.)
* Works on both Windows and Mac. On Mac, 5 of the methods are not available out of which the ```GetLocalPath``` and ```GetUNCPath``` are not needed as all paths are local on a Mac (network shares are mounted as volumes and OneDrive paths work locally anyway)
* Works in both x32 and x64 application environments

## License
MIT License

Copyright (c) 2012 Ion Cristian Buse

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
