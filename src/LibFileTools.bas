Attribute VB_Name = "LibFileTools"
'''=============================================================================
''' VBA FileTools
'''----------------------------------------------
''' https://github.com/cristianbuse/VBA-FileTools
'''----------------------------------------------
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

'*******************************************************************************
'' Functions in this library module allow easy file system manipulation in VBA
'' regardless of:
''  - the host Application (Excel, Word, AutoCAD etc.)
''  - the operating system (Mac, Windows)
''  - application environment (x32, x64)
'' No extra library references are needed (e.g. Microsoft Scripting Runtime)
''
'' Public/Exposed methods:
''    - BrowseForFiles    (Windows only)
''    - BrowseForFolder   (Windows only)
''    - BuildPath
''    - CopyFile
''    - CopyFolder
''    - CreateFolder
''    - DeleteFile
''    - DeleteFolder
''    - FixFileName
''    - FixPathSeparators
''    - GetFileOwner      (Windows only)
''    - GetFiles
''    - GetFolders
''    - GetLocalPath      (Windows only)
''    - GetPathSeparator
''    - GetUNCPath        (Windows only)
''    - IsFile
''    - IsFolder
''    - MoveFile
''    - MoveFolder
'*******************************************************************************

Option Explicit
Option Private Module

#If Mac Then
#ElseIf VBA7 Then
    Private Declare PtrSafe Function CopyFileA Lib "kernel32" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
    Private Declare PtrSafe Function DeleteFileA Lib "kernel32" (ByVal lpFileName As String) As Long
    Private Declare PtrSafe Function GetFileSecurity Lib "advapi32.dll" Alias "GetFileSecurityA" (ByVal lpFileName As String, ByVal RequestedInformation As Long, pSecurityDescriptor As Byte, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
    Private Declare PtrSafe Function GetSecurityDescriptorOwner Lib "advapi32.dll" (pSecurityDescriptor As Byte, pOwner As LongPtr, lpbOwnerDefaulted As LongPtr) As Long
    Private Declare PtrSafe Function LookupAccountSid Lib "advapi32.dll" Alias "LookupAccountSidA" (ByVal lpSystemName As String, ByVal Sid As LongPtr, ByVal Name As String, cbName As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As LongPtr) As Long
#Else
    Private Declare Function CopyFileA Lib "kernel32" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
    Private Declare Function DeleteFileA Lib "kernel32" (ByVal lpFileName As String) As Long
    Private Declare Function GetFileSecurity Lib "advapi32.dll" Alias "GetFileSecurityA" (ByVal lpFileName As String, ByVal RequestedInformation As Long, pSecurityDescriptor As Byte, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
    Private Declare Function GetSecurityDescriptorOwner Lib "advapi32.dll" (pSecurityDescriptor As Byte, pOwner As Long, lpbOwnerDefaulted As Long) As Long
    Private Declare Function LookupAccountSid Lib "advapi32.dll" Alias "LookupAccountSidA" (ByVal lpSystemName As String, ByVal Sid As Long, ByVal Name As String, cbName As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Long) As Long
#End If

Private Type DRIVE_INFO
    driveName As String
    driveLetter As String
    fileSystem As String
    shareName As String
End Type

Private Type ONEDRIVE_ACCOUNT_INFO
    isBusiness As Boolean
    cid As String
    userFolder As String
    endpointUri As String
End Type

Private Type ONEDRIVE_INFO
    isInitialized As Boolean
    accountsCount As Long
    accounts() As ONEDRIVE_ACCOUNT_INFO
End Type

'*******************************************************************************
'Returns a Collection of file paths by using a FilePicker FileDialog
'*******************************************************************************
#If Mac Then
    'Not implemented
    'Seems achievable via script:
    '   - https://stackoverflow.com/a/15546518/8488913
    '   - https://stackoverflow.com/a/37411960/8488913
#Else
Public Function BrowseForFiles(Optional ByVal initialPath As String _
                             , Optional ByVal dialogTitle As String _
                             , Optional ByVal filterDesc As String _
                             , Optional ByVal filterList As String _
                             , Optional ByVal allowMultiFiles As Boolean = True _
) As Collection
    'In case reference to Microsoft Office X.XX Object Library is missing
    Const dialogTypeFilePicker As Long = 3 'msoFileDialogFilePicker
    Const actionButton As Long = -1
    '
    With Application.FileDialog(dialogTypeFilePicker)
        If dialogTitle <> vbNullString Then .title = dialogTitle
        If initialPath <> vbNullString Then .InitialFileName = initialPath
        If .InitialFileName = vbNullString Then
            Dim app As Object: Set app = Application 'Needs to be late-binded
            Select Case Application.Name
                Case "Microsoft Excel": .InitialFileName = app.ThisWorkbook.Path
                Case "Microsoft Word": .InitialFileName = app.ThisDocument.Path
            End Select
        End If
        '
        .AllowMultiSelect = allowMultiFiles
        .filters.Clear 'Allows all file types
        On Error Resume Next
        .filters.Add filterDesc, filterList
        On Error GoTo 0
        '
        Set BrowseForFiles = New Collection
        If .Show = actionButton Then
            Dim v As Variant
            '
            For Each v In .SelectedItems
                BrowseForFiles.Add v
            Next v
        End If
    End With
End Function
#End If

'*******************************************************************************
'Returns a folder path by using a FolderPicker FileDialog
'*******************************************************************************
#If Mac Then
    'Not implemented
#Else
Public Function BrowseForFolder(Optional ByVal initialPath As String _
                              , Optional ByVal dialogTitle As String _
) As String
    'In case reference to Microsoft Office X.XX Object Library is missing
    Const dialogTypeFolderPicker As Long = 4 'msoFileDialogFolderPicker
    Const actionButton As Long = -1
    '
    With Application.FileDialog(dialogTypeFolderPicker)
        If dialogTitle <> vbNullString Then .title = dialogTitle
        If initialPath <> vbNullString Then .InitialFileName = initialPath
        If .InitialFileName = vbNullString Then
            Dim app As Object: Set app = Application 'Needs to be late-binded
            Select Case Application.Name
                Case "Microsoft Excel": .InitialFileName = app.ThisWorkbook.Path
                Case "Microsoft Word": .InitialFileName = app.ThisDocument.Path
            End Select
        End If
        If .Show = actionButton Then
            .InitialFileName = .SelectedItems.Item(1)
            BrowseForFolder = .InitialFileName
        End If
    End With
End Function
#End If

'*******************************************************************************
'Combines a folder path with a file/folder name or an incomplete path (ex. \a\b)
'*******************************************************************************
Public Function BuildPath(ByVal folderPath As String _
                        , ByVal fsName As String _
) As String
    BuildPath = FixPathSeparators(folderPath & GetPathSeparator() & fsName)
End Function

'*******************************************************************************
'Copies a file. Overwrites existing files unless 'failIfExists' is set to True
'Note that VBA.FileCopy does not copy opened files on Windows but it does on Mac
'If the the destination file already exists and 'failIfExists' is set to False
'   then this method must be able to overwrite the destination file. Rather than
'   failing and then trying again with attribute set to vbNormal this method
'   sets the attribute for the destination path to vbNormal before copying.
'   This is slightly slower than just copying directly but far outperforms two
'   copy operations in the case where the first one fails and the second one is
'   done after setting the destination file attribute to vbNormal.
'*******************************************************************************
Public Function CopyFile(ByVal sourcePath As String _
                       , ByVal destinationPath As String _
                       , Optional ByVal failIfExists As Boolean = False _
) As Boolean
    If sourcePath = vbNullString Then Exit Function
    If destinationPath = vbNullString Then Exit Function
    '
    #If Mac Then
        If failIfExists Then If IsFile(destinationPath) Then Exit Function
        '
        On Error Resume Next
        SetAttr destinationPath, vbNormal 'Too costly to do after failing Copy
        Err.Clear 'Ignore any errors raised by 'SetAttr'
        VBA.FileCopy sourcePath, destinationPath 'Copies opened files as well
        CopyFile = (Err.Number = 0)
        On Error GoTo 0
    #Else
        If Not failIfExists Then
            On Error Resume Next
            SetAttr destinationPath, vbNormal 'Costly to do after failing Copy
            On Error GoTo 0
        End If
        CopyFile = CopyFileA(sourcePath, destinationPath, failIfExists)
    #End If
End Function

'*******************************************************************************
'Copies a folder. Ability to copy all subfolders
'If 'failIfExists' is set to True then this method will fail if any file or
'   subFolder already exists (including the main 'destinationPath')
'If 'ignoreFailedFiles' is set to True then the method continues to copy the
'   remaining files. This is useful when reverting a 'MoveFolder' call across
'   different disk drives. Use with care
'*******************************************************************************
Public Function CopyFolder(ByVal sourcePath As String _
                         , ByVal destinationPath As String _
                         , Optional ByVal includeSubFolders As Boolean = True _
                         , Optional ByVal failIfExists As Boolean = False _
                         , Optional ByVal ignoreFailedFiles As Boolean = False _
) As Boolean
    If Not IsFolder(sourcePath) Then Exit Function
    If Not CreateFolder(destinationPath, failIfExists) Then Exit Function
    '
    Dim fixedSrc As String: fixedSrc = BuildPath(sourcePath, vbNullString)
    Dim fixedDst As String: fixedDst = BuildPath(destinationPath, vbNullString)
    '
    If includeSubFolders Then
        Dim subFolder As Variant
        Dim newFolderPath As String
        '
        For Each subFolder In GetFolders(fixedSrc, True, True, True)
            newFolderPath = Replace(subFolder, fixedSrc, fixedDst)
            If Not CreateFolder(newFolderPath, failIfExists) Then Exit Function
        Next subFolder
    End If
    '
    Dim filePath As Variant
    Dim newFilePath As String
    '
    For Each filePath In GetFiles(fixedSrc, includeSubFolders, True, True)
        newFilePath = Replace(filePath, fixedSrc, fixedDst)
        If Not CopyFile(filePath, newFilePath, failIfExists) Then
            If Not ignoreFailedFiles Then Exit Function
        End If
    Next filePath
    '
    CopyFolder = True
End Function

'*******************************************************************************
'Creates a folder including parent folders if needed
'*******************************************************************************
Public Function CreateFolder(ByVal folderPath As String _
                           , Optional ByVal failIfExists As Boolean = False _
) As Boolean
    If IsFolder(folderPath) Then
        CreateFolder = Not failIfExists
        Exit Function
    End If
    '
    Dim fullPath As String
    '
    fullPath = BuildPath(folderPath, vbNullString)
    fullPath = Left$(fullPath, Len(fullPath) - 1) 'Removing trailing separator
    '
    Dim sepIndex As Long
    Dim collFoldersToCreate As New Collection
    Dim i As Long
    '
    'Note that the same outcome could be achieved via recursivity but this
    '   approach avoids adding extra stack frames to the call stack
    Do
        collFoldersToCreate.Add fullPath
        '
        sepIndex = InStrRev(fullPath, GetPathSeparator())
        If sepIndex < 3 Then Exit Do
        '
        fullPath = Left$(fullPath, sepIndex - 1)
        If IsFolder(fullPath) Then Exit Do
    Loop
    On Error Resume Next
    For i = collFoldersToCreate.Count To 1 Step -1
        MkDir collFoldersToCreate(i)
        If Err.Number <> 0 Then Exit For
    Next i
    CreateFolder = (Err.Number = 0)
    On Error GoTo 0
End Function

'*******************************************************************************
'Deletes a file only. Does not support wildcards * ?
'Rather than failing and then trying again with attribute set to vbNormal this
'   method sets the attribute to normal before deleting. This is slightly slower
'   than just deleting directly but far outperforms two delete operations in the
'   case where the first one fails and the second one is done after setting the
'   file attribute to vbNormal
'*******************************************************************************
Public Function DeleteFile(ByVal filePath As String) As Boolean
    If filePath = vbNullString Then Exit Function
    '
    On Error Resume Next
    SetAttr filePath, vbNormal 'Too costly to do after failing Delete
    On Error GoTo 0
    '
    #If Mac Then
        If Not IsFile(filePath) Then Exit Function 'Avoid 'Kill' on folder
        On Error Resume Next
        Kill filePath
        DeleteFile = (Err.Number = 0)
        On Error GoTo 0
    #Else
        DeleteFile = CBool(DeleteFileA(filePath))
    #End If
End Function

'*******************************************************************************
'Deletes a folder
'If the 'deleteContents' parameter is set to True then all files/folders inside
'   all subfolders will be deleted before attempting to delete the main folder.
'   In this case no attempt is made to roll back any deleted files/folders in
'   case the method fails (ex. after deleting some files/folders the method
'   cannot delete a particular file that is locked/open and so the method stops
'   and returns False without rolling back the already deleted files/folders)
'*******************************************************************************
Public Function DeleteFolder(ByVal folderPath As String _
                           , Optional ByVal deleteContents As Boolean = False _
                           , Optional ByVal failIfMissing As Boolean = False _
) As Boolean
    If folderPath = vbNullString Then Exit Function
    '
    If Not IsFolder(folderPath) Then
        DeleteFolder = Not failIfMissing
        Exit Function
    End If
    '
    On Error Resume Next
    RmDir folderPath 'Assume the folder is empty
    If Err.Number = 0 Then
        DeleteFolder = True
        Exit Function
    End If
    On Error GoTo 0
    If Not deleteContents Then Exit Function
    '
    Dim collFolders As Collection
    Dim i As Long
    '
    Set collFolders = GetFolders(folderPath, True, True, True)
    For i = collFolders.Count To 1 Step -1 'From bottom to top level
        If Not DeleteBottomMostFolder(collFolders.Item(i)) Then Exit Function
    Next i
    '
    DeleteFolder = DeleteBottomMostFolder(folderPath)
End Function

'*******************************************************************************
'Utility for 'DeleteFolder'
'Deletes a folder that can contain files but does NOT contain any other folders
'*******************************************************************************
Private Function DeleteBottomMostFolder(ByVal folderPath As String) As Boolean
    Dim fixedPath As String: fixedPath = BuildPath(folderPath, vbNullString)
    Dim filePath As Variant
    '
    On Error Resume Next
    Kill fixedPath  'Try to batch delete all files (if any)
    Err.Clear       'Kill fails if there are no files so ignore any error
    RmDir fixedPath 'Try to delete folder
    If Err.Number = 0 Then
        DeleteBottomMostFolder = True
        Exit Function
    End If
    On Error GoTo 0
    '
    For Each filePath In GetFiles(fixedPath, False, True, True)
        If Not DeleteFile(filePath) Then Exit Function
    Next filePath
    '
    On Error Resume Next
    RmDir fixedPath
    DeleteBottomMostFolder = (Err.Number = 0)
    On Error GoTo 0
End Function

'*******************************************************************************
'Fixes a file or folder name
'Before creating a file/folder it's useful to fix the name so that the creation
'   does not fail because of forbidden characters, reserved names or other rules
'*******************************************************************************
#If Mac Then
Public Function FixFileName(ByVal nameToFix As String) As String
    FixFileName = Replace(nameToFix, ":", vbNullString)
    FixFileName = Replace(FixFileName, "/", vbNullString)
End Function
#Else
Public Function FixFileName(ByVal nameToFix As String _
                          , Optional ByVal isFATFileSystem As Boolean = False _
) As String
    Dim resultName As String: resultName = nameToFix
    Dim v As Variant
    '
    For Each v In ForbiddenNameChars(addCaret:=isFATFileSystem)
        resultName = Replace(resultName, v, vbNullString)
    Next v
    '
    'Names cannot end in a space or a period character
    Const dotSpace As String = ". "
    Dim nameLen As Long: nameLen = Len(resultName)
    Dim currIndex As Long
    '
    currIndex = nameLen
    If currIndex > 0 Then
        Do While InStr(1, dotSpace, Mid$(resultName, currIndex, 1)) > 0
            currIndex = currIndex - 1
            If currIndex = 0 Then Exit Do
        Loop
    End If
    If currIndex < nameLen Then resultName = Left$(resultName, currIndex)
    '
    If IsReservedName(resultName) Then resultName = vbNullString
    '
    FixFileName = resultName
End Function
#End If

'*******************************************************************************
'Returns a collection of forbidden characters for a file/folder name
'Ability to add the caret ^ char - forbidden on FAT file systems but not on NTFS
'*******************************************************************************
#If Mac Then
#Else
Private Function ForbiddenNameChars(ByVal addCaret As Boolean) As Collection
    Static collForbiddenChars As Collection
    Static hasCaret As Boolean
    '
    If collForbiddenChars Is Nothing Then
        Set collForbiddenChars = New Collection
        Dim v As Variant
        Dim i As Long
        '
        For Each v In Split(":,*,?,"",<,>,|,\,/", ",")
            collForbiddenChars.Add v
        Next v
        For i = 0 To 31 'ASCII control characters and the null character
            collForbiddenChars.Add VBA.Chr$(i)
        Next i
    End If
    If hasCaret And Not addCaret Then
        collForbiddenChars.Remove collForbiddenChars.Count
    ElseIf Not hasCaret And addCaret Then
        collForbiddenChars.Add "^"
    End If
    hasCaret = addCaret
    '
    Set ForbiddenNameChars = collForbiddenChars
End Function
#End If

'*******************************************************************************
'Windows file/folder reserved names: com1 to com9, lpt1 to lpt9, con, nul, prn
'*******************************************************************************
#If Mac Then
#Else
Private Function IsReservedName(ByVal nameToCheck As String) As Boolean
    Static collReservedNames As Collection
    '
    If collReservedNames Is Nothing Then
        Dim v As Variant
        '
        Set collReservedNames = New Collection
        For Each v In Split("com1,com2,com3,com4,com5,com6,com7,com8,com9," _
        & "lpt1,lpt2,lpt3,lpt4,lpt5,lpt6,lpt7,lpt8,lpt9,con,nul,prn", ",")
            collReservedNames.Add v, v
        Next v
    End If
    On Error Resume Next
    collReservedNames.Item nameToCheck
    IsReservedName = (Err.Number = 0)
    On Error GoTo 0
End Function
#End If

'*******************************************************************************
'Fixes path separators for a file/folder path
'Windows example: replace \\, \\\, \\\\, \\//, \/\/\, /, // etc. with a single \
'Note that on a Mac, the network paths (smb:// or afp://) need to be mounted and
'   are only valid via the mounted volumes: /volumes/VolumeName/... unlike on a
'   PC where \\share\data\... is a perfectly valid file/folder path
'*******************************************************************************
Public Function FixPathSeparators(ByVal pathToFix As String) As String
    Static oneSeparator As String
    Static twoSeparators As String
    Dim resultPath As String: resultPath = pathToFix
    '
    If oneSeparator = vbNullString Then
        oneSeparator = GetPathSeparator()
        twoSeparators = oneSeparator & oneSeparator
    End If
    '
    #If Mac Then
    #Else 'Replace forward slashes with back slashes for Windows
        resultPath = Replace(resultPath, "/", oneSeparator)
        Dim isUNC As Boolean: isUNC = Left$(resultPath, 2) = twoSeparators '\\
    #End If
    '
    'Replace repeated separators e.g. replace \\\\\ with \
    Dim previousLength As Long
    Dim currentLength As Long: currentLength = Len(resultPath)
    Do
        previousLength = currentLength
        resultPath = Replace(resultPath, twoSeparators, oneSeparator)
        currentLength = Len(resultPath)
    Loop Until previousLength = currentLength
    '
    #If Mac Then
    #Else
        If isUNC Then resultPath = oneSeparator & resultPath
    #End If
    '
    FixPathSeparators = resultPath
End Function

'*******************************************************************************
'Retrieves the owner name for a file path
'*******************************************************************************
#If Mac Then
#Else
Public Function GetFileOwner(ByVal filePath As String) As String
    Const osi As Long = 1 'OWNER_SECURITY_INFORMATION
    Dim sdSize As Long
    '
    'Get SECURITY_DESCRIPTOR required Buffer Size
    GetFileSecurity filePath, osi, 0, 0&, sdSize
    If sdSize = 0 Then Exit Function
    '
    'Size the SECURITY_DESCRIPTOR buffer
    Dim sd() As Byte: ReDim sd(0 To sdSize - 1)
    '
    'Get SECURITY_DESCRIPTOR buffer
    If GetFileSecurity(filePath, osi, sd(0), sdSize, sdSize) = 0 Then
        Exit Function
    End If
    '
    'Get owner SSID
        Dim pOwner As LongPtr
    #If VBA7 Then
    #Else
        Dim pOwner As Long
    #End If
    If GetSecurityDescriptorOwner(sd(0), pOwner, 0&) = 0 Then Exit Function
    '
    'Get name and domain length
    Dim nameLen As Long, domainLen As Long
    LookupAccountSid vbNullString, pOwner, vbNullString _
                   , nameLen, vbNullString, domainLen, 0&
    If nameLen = 0 Then Exit Function
    '
    'Get name and domain
    Dim owName As String: owName = VBA.Space$(nameLen - 1) '-1 less Null Char
    Dim owDomain As String: owDomain = VBA.Space$(domainLen - 1)
    If LookupAccountSid(vbNullString, pOwner, owName _
                      , nameLen, owDomain, domainLen, 0&) = 0 Then Exit Function
    '
    'Return result
    GetFileOwner = owDomain & GetPathSeparator() & owName
End Function
#End If

'*******************************************************************************
'Returns a Collection of all the files (paths) in a specified folder
'Warning! On Mac the 'Dir' method only accepts the vbHidden and the vbDirectory
'   attributes. However the vbHidden attribute does not work - no hidden files
'   or folders are retrieved regardless if vbHidden is used or not
'On Windows, the vbHidden, and vbSystem attributes work fine with 'Dir' but
'   the vbReadOnly attribute seems to be completely ignored
'*******************************************************************************
Public Function GetFiles(ByVal folderPath As String _
                       , Optional ByVal includeSubFolders As Boolean = False _
                       , Optional ByVal includeHidden As Boolean = False _
                       , Optional ByVal includeSystem As Boolean = False _
) As Collection
    Dim collFiles As New Collection
    Dim fAttribute As VbFileAttribute
    '
    #If Mac Then
        fAttribute = vbNormal
        'Both vbReadOnly and vbSystem are raising errors when used in 'Dir'
        'vbHidden does not raise an error but seems to be ignored entirely
    #Else
        fAttribute = vbReadOnly 'Seems to be ignored entirely anyway
        If includeSystem Then fAttribute = fAttribute + vbSystem
    #End If
    If includeHidden Then fAttribute = fAttribute + vbHidden
    '
    AddFilesTo collFiles, folderPath, fAttribute
    If includeSubFolders Then
        Dim subFolder As Variant
        For Each subFolder In GetFolders(folderPath, True, True, True)
            AddFilesTo collFiles, subFolder, fAttribute
        Next subFolder
    End If
    '
    Set GetFiles = collFiles
End Function

'*******************************************************************************
'Utility for 'GetFiles' method
'Warning! On Mac the 'Dir' method only accepts the vbHidden and the vbDirectory
'   attributes. However the vbHidden attribute does not work - no hidden files
'   or folders are retrieved regardless if vbHidden is used or not
'*******************************************************************************
Private Sub AddFilesTo(ByVal collTarget As Collection _
                     , ByVal folderPath As String _
                     , ByVal fAttribute As VbFileAttribute _
)
    Dim fixedPath As String
    Dim fileName As String
    Dim fullPath As String
    '
    fixedPath = BuildPath(folderPath, vbNullString)
    fileName = Dir(fixedPath, fAttribute)
    Do While fileName <> vbNullString
        collTarget.Add fixedPath & fileName
        fileName = Dir
    Loop
End Sub

'*******************************************************************************
'Returns a Collection of all the subfolders (paths) in a specified folder
'Warning! On Mac the 'Dir' method only accepts the vbHidden and the vbDirectory
'   attributes. However the vbHidden attribute does not work - no hidden files
'   or folders are retrieved regardless if vbHidden is used or not
'On Windows, the vbHidden, and vbSystem attributes work fine with 'Dir'
'*******************************************************************************
Public Function GetFolders(ByVal folderPath As String _
                         , Optional ByVal includeSubFolders As Boolean = False _
                         , Optional ByVal includeHidden As Boolean = False _
                         , Optional ByVal includeSystem As Boolean = False _
) As Collection
    Dim collFolders As New Collection
    Dim fAttribute As VbFileAttribute
    '
    fAttribute = vbDirectory
    #If Mac Then
        'vbSystem is raising an error when used in 'Dir'
        'vbHidden does not raise an error but seems to be ignored entirely
    #Else
        If includeSystem Then fAttribute = fAttribute + vbSystem
    #End If
    If includeHidden Then fAttribute = fAttribute + vbHidden
    '
    AddFoldersTo collFolders, folderPath, includeSubFolders, fAttribute
    Set GetFolders = collFolders
End Function

'*******************************************************************************
'Utility for 'GetFolders' method
'Returning a Collection and then adding the elements of that collection to
'   another collection higher up in the stack frame is simply inefficient and
'   unnecessary when doing recursion. Instead this method adds the elements
'   directly in the final collection instance ('collTarget'). Top-down approach
'Because 'Dir' does not allow recursive calls to 'Dir', a temporary collection
'   is used to get all the subfolders (only if 'includeSubFolders' is True).
'   The temporary collection is then iterated in order to get the subfolders for
'   each of the initial subfolders
'Warning! On Mac the 'Dir' method only accepts the vbHidden and the vbDirectory
'   attributes. However the vbHidden attribute does not work - no hidden files
'   or folders are retrieved regardless if vbHidden is used or not
'*******************************************************************************
Private Sub AddFoldersTo(ByVal collTarget As Collection _
                       , ByVal folderPath As String _
                       , ByVal includeSubFolders As Boolean _
                       , ByVal fAttribute As VbFileAttribute _
)
    Const currentFolder As String = "."
    Const parentFolder As String = ".."
    Dim fixedPath As String
    Dim folderName As String
    Dim fullPath As String
    Dim collFolders As Collection
    '
    If includeSubFolders Then
        Set collFolders = New Collection 'Temp collection to be iterated later
    Else
        Set collFolders = collTarget 'No recusion so we add directly to target
    End If
    fixedPath = BuildPath(folderPath, vbNullString)
    folderName = Dir(fixedPath, fAttribute)
    Do While folderName <> vbNullString
        If folderName <> currentFolder And folderName <> parentFolder Then
            fullPath = fixedPath & folderName
            If GetAttr(fullPath) And vbDirectory Then collFolders.Add fullPath
        End If
        folderName = Dir
    Loop
    If includeSubFolders Then
        Dim subFolder As Variant
        '
        For Each subFolder In collFolders
            collTarget.Add subFolder
            AddFoldersTo collTarget, subFolder, True, fAttribute
        Next subFolder
    End If
End Sub

'*******************************************************************************
'Returns the local drive path for a given path or null string if path not local
'Note that the input path does not need to be an existing file/folder
'*******************************************************************************
#If Mac Then
#Else
Public Function GetLocalPath(ByVal fullPath As String) As String
    With GetDriveInfo(fullPath)
        If .driveLetter = vbNullString Then
            GetLocalPath = GetOneDriveLocalPath(fullPath, rebuildCache:=False)
        Else
            GetLocalPath = FixPathSeparators(Replace(fullPath _
            , .driveName, .driveLetter & ":", 1, 1, vbTextCompare))
            Exit Function
        End If
    End With
End Function
#End If

'*******************************************************************************
'Returns the path separator character. Encapsulates Application.PathSeparator
'*******************************************************************************
Public Function GetPathSeparator() As String
    Static pSeparator As String
    If pSeparator = vbNullString Then pSeparator = Application.PathSeparator
    GetPathSeparator = pSeparator
End Function

'*******************************************************************************
'Returns the UNC path for a given path or null string if path is not remote
'Note that the input path does not need to be an existing file/folder
'*******************************************************************************
#If Mac Then
#Else
Public Function GetUNCPath(ByVal fullPath As String) As String
    With GetDriveInfo(fullPath)
        If .shareName = vbNullString Then Exit Function 'Not UNC
        GetUNCPath = FixPathSeparators(Replace(fullPath _
        , .driveName, .shareName, 1, 1, vbTextCompare))
    End With
End Function
#End If

'*******************************************************************************
'Returns basic drive information about a full path
'*******************************************************************************
#If Mac Then
#Else
Private Function GetDriveInfo(ByVal fullPath As String) As DRIVE_INFO
    Dim fso As Object: Set fso = GetFileSystemObject()
    If fso Is Nothing Then Exit Function
    '
    Dim driveName As String: driveName = fso.GetDriveName(fullPath)
    If driveName = vbNullString Then Exit Function
    '
    Dim fsDrive As Object
    On Error Resume Next
    Set fsDrive = fso.GetDrive(driveName)
    On Error GoTo 0
    If fsDrive Is Nothing Then Exit Function
    '
    If fsDrive.driveLetter = vbNullString Then
        Dim sn As Long: sn = fsDrive.SerialNumber
        Dim tempDrive As Object
        Dim isFound As Boolean
        '
        For Each tempDrive In fso.Drives
            If tempDrive.SerialNumber = sn Then
                Set fsDrive = tempDrive
                isFound = True
                Exit For
            End If
        Next tempDrive
        If Not isFound Then Exit Function
    End If
    '
    With GetDriveInfo
        .driveName = driveName
        .driveLetter = fsDrive.driveLetter
        .fileSystem = fsDrive.fileSystem
        .shareName = fsDrive.shareName
        If .shareName <> vbNullString Then
            .driveName = AlignDriveNameIfNeeded(.driveName, .shareName)
        End If
    End With
End Function
#End If

'*******************************************************************************
'Late-bounded file system for Windows
'*******************************************************************************
#If Mac Then
#Else
Private Function GetFileSystemObject() As Object
    Static fso As Object
    '
    If fso Is Nothing Then
        On Error Resume Next
        Set fso = CreateObject("Scripting.FileSystemObject")
        On Error GoTo 0
    End If
    Set GetFileSystemObject = fso
End Function
#End If

'*******************************************************************************
'Aligns a wrong drive name with the share name
'Example: \\emea\ to \\emea.companyName.net\
'*******************************************************************************
#If Mac Then
#Else
Private Function AlignDriveNameIfNeeded(ByVal driveName As String _
                                      , ByVal shareName As String _
) As String
    Dim sepIndex As Long
    '
    sepIndex = VBA.InStr(3, driveName, GetPathSeparator())
    If sepIndex > 0 Then
        Dim newName As String: newName = VBA.Left$(driveName, sepIndex - 1)
        sepIndex = VBA.InStr(3, shareName, GetPathSeparator())
        newName = newName & Right$(shareName, Len(shareName) - sepIndex + 1)
        AlignDriveNameIfNeeded = newName
    Else
        AlignDriveNameIfNeeded = driveName
    End If
End Function
#End If

'*******************************************************************************
'Returns the local path for a OneDrive web path
'Returns null string if the path provided is not a valid OneDrive web path
'*******************************************************************************
#If Mac Then
#Else
Private Function GetOneDriveLocalPath(ByVal odWebPath As String _
                                    , ByVal rebuildCache As Boolean _
) As String
    If InStr(1, odWebPath, "https://", vbTextCompare) = 0 Then Exit Function
    '
    Dim odInfo As ONEDRIVE_INFO
    Dim odAccount As ONEDRIVE_ACCOUNT_INFO
    Dim i As Long
    Dim root As String
    Dim rPath As String
    '
    odInfo = GetOneDriveInfo(rebuildCache)
    If odInfo.accountsCount = 0 Then Exit Function
    '
    For i = LBound(odInfo.accounts) To UBound(odInfo.accounts)
        odAccount = odInfo.accounts(i)
        If odAccount.isBusiness Then
            root = odAccount.endpointUri
        Else
            root = "https://d.docs.live.net/" & odAccount.cid
        End If
        If StrComp(Left$(odWebPath, Len(root)), root, vbTextCompare) = 0 Then
            rPath = Right$(odWebPath, Len(odWebPath) - Len(root)) 'Remove root
            If odAccount.isBusiness Then 'Trim "/Documents"
                rPath = Right$(rPath, Len(rPath) - InStr(2, rPath, "/") + 1)
            End If
            rPath = BuildPath(odAccount.userFolder, rPath)
            Exit For
        End If
    Next i
    GetOneDriveLocalPath = rPath
End Function
#End If

'*******************************************************************************
'Returns info about valid OneDrive accounts associated to the current user
'https://docs.microsoft.com/en-us/windows/win32/wmisdk/obtaining-registry-data
'*******************************************************************************
#If Mac Then
#Else
Private Function GetOneDriveInfo(ByVal rebuildCache As Boolean) As ONEDRIVE_INFO
    Static odInfo As ONEDRIVE_INFO
    '
    If odInfo.isInitialized And Not rebuildCache Then
        GetOneDriveInfo = odInfo
        Exit Function
    ElseIf odInfo.isInitialized Then
        Dim tempAccounts As ONEDRIVE_INFO
        odInfo = tempAccounts 'Reset
    End If
    '
    Const HKCU = &H80000001 'HKEY_CURRENT_USER
    Const odAccountsKey As String = "Software\Microsoft\OneDrive\Accounts\"
    Const computerName As String = "."
    Dim oReg As Object
    Dim arrKeys() As Variant
    '
    'Read sub keys from OneDrive\Accounts\
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" _
        & computerName & "\root\default:StdRegProv")
    If oReg Is Nothing Then Exit Function
    oReg.EnumKey HKCU, odAccountsKey, arrKeys
    '
    On Error Resume Next
    odInfo.accountsCount = UBound(arrKeys) - LBound(arrKeys) + 1
    On Error GoTo 0
    If odInfo.accountsCount = 0 Then Exit Function
    '
    'Keep valid accounts only
    Dim subKey As Variant
    Dim fullKey As String
    Dim collValidKeys As New Collection
    Dim cid As String
    '
    For Each subKey In arrKeys
        fullKey = odAccountsKey & subKey
        oReg.GetStringValue HKCU, fullKey, "cid", cid
        If cid <> vbNullString Then collValidKeys.Add fullKey
    Next subKey
    If collValidKeys.Count = 0 Then Exit Function
    '
    Dim i As Long: i = 0
    Dim tempDWord As Long
    '
    'Read accounts info and write each to a UDT
    ReDim odInfo.accounts(0 To collValidKeys.Count - 1)
    For Each subKey In collValidKeys
        With odInfo.accounts(i)
            oReg.GetStringValue HKCU, subKey, "cid", .cid
            oReg.GetStringValue HKCU, subKey, "UserFolder", .userFolder
            oReg.GetStringValue HKCU, subKey, "ServiceEndpointUri", .endpointUri
            oReg.GetDWORDValue HKCU, subKey, "Business", tempDWord
            '
            .isBusiness = (tempDWord = 1)
            If .isBusiness Then
                .endpointUri = Replace(.endpointUri, "/_api", vbNullString)
            End If
        End With
        i = i + 1
    Next subKey
    '
    odInfo.accountsCount = collValidKeys.Count
    odInfo.isInitialized = True
    '
    GetOneDriveInfo = odInfo
End Function
#End If

'*******************************************************************************
'Checks if a path indicates a file path
'Note that if C:\Test\1.txt is valid then C:\Test\\///1.txt will also be valid
'Most VBA methods consider valid any path separators with multiple characters
'*******************************************************************************
Public Function IsFile(ByVal filePath As String) As Boolean
    On Error Resume Next
    IsFile = ((GetAttr(filePath) And vbDirectory) <> vbDirectory)
    On Error GoTo 0
End Function

'*******************************************************************************
'Checks if a path indicates a folder path
'Note that if C:\Test\Demo is valid then C:\Test\\///Demo will also be valid
'Most VBA methods consider valid any path separators with multiple characters
'*******************************************************************************
Public Function IsFolder(ByVal folderPath As String) As Boolean
    On Error Resume Next
    IsFolder = ((GetAttr(folderPath) And vbDirectory) = vbDirectory)
    On Error GoTo 0
End Function

'*******************************************************************************
'Moves (or renames) a file
'*******************************************************************************
Public Function MoveFile(ByVal sourcePath As String _
                       , ByVal destinationPath As String _
) As Boolean
    If sourcePath = vbNullString Then Exit Function
    If destinationPath = vbNullString Then Exit Function
    If Not IsFile(sourcePath) Then Exit Function
    '
    On Error Resume Next
    #If Mac Then
        Dim fAttr As VbFileAttribute: fAttr = GetAttr(sourcePath)
        If fAttr <> vbNormal Then SetAttr sourcePath, vbNormal
        Err.Clear
    #End If
    '
    Name sourcePath As destinationPath
    MoveFile = (Err.Number = 0)
    '
    #If Mac Then
        If fAttr <> vbNormal Then 'Restore attribute
            If MoveFile Then
                SetAttr destinationPath, fAttr
            Else
                SetAttr sourcePath, fAttr
            End If
        End If
    #End If
    On Error GoTo 0
End Function

'*******************************************************************************
'Moves (or renames) a folder
'*******************************************************************************
Public Function MoveFolder(ByVal sourcePath As String _
                         , ByVal destinationPath As String _
) As Boolean
    If sourcePath = vbNullString Then Exit Function
    If destinationPath = vbNullString Then Exit Function
    If Not IsFolder(sourcePath) Then Exit Function
    If IsFolder(destinationPath) Then Exit Function
    '
    'The 'Name' statement can move a file across drives, but it can only rename
    '   a directory or folder within the same drive. Try 'Name' first
    On Error Resume Next
    Name sourcePath As destinationPath
    If Err.Number = 0 Then
        MoveFolder = True
        Exit Function
    End If
    On Error GoTo 0
    '
    'Try FSO if available
    #If Mac Then
    #Else
        On Error Resume Next
        GetFileSystemObject().MoveFolder sourcePath, destinationPath
        If Err.Number = 0 Then
            MoveFolder = True
            Exit Function
        End If
        On Error GoTo 0
    #End If
    '
    'If all else failed, first make a copy and then delete the source
    If Not CopyFolder(sourcePath, destinationPath, True) Then 'Revert
        DeleteFolder destinationPath, True
        Exit Function
    ElseIf Not DeleteFolder(sourcePath, True) Then 'Files might be open. Revert
        CopyFolder destinationPath, sourcePath, ignoreFailedFiles:=True
        DeleteFolder destinationPath, True
        Exit Function
    End If
    '
    MoveFolder = True
End Function
