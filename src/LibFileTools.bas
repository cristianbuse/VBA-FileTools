Attribute VB_Name = "LibFileTools"
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
''    - GetUNCPath        (Windows only)
''    - GetWebPath        (Windows only)
''    - IsFile
''    - IsFolder
''    - IsFolderEditable
''    - MoveFile
''    - MoveFolder
''    - ReadBytes
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

Private Type ONEDRIVE_PROVIDER
    webPath As String
    mountPoint As String
    isBusiness As Boolean
    isMain As Boolean
    accountIndex As Long
End Type
Private Type ONEDRIVE_PROVIDERS
    arr() As ONEDRIVE_PROVIDER
    pCount As Long
    isSet As Boolean
End Type

#If Mac Then
    Public Const PATH_SEPARATOR = "/"
#Else
    Public Const PATH_SEPARATOR = "\"
#End If

Private m_providers As ONEDRIVE_PROVIDERS

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
                             , Optional ByVal allowMultiFiles As Boolean = True) As Collection
    'In case reference to Microsoft Office X.XX Object Library is missing
    Const dialogTypeFilePicker As Long = 3 'msoFileDialogFilePicker
    Const actionButton As Long = -1
    '
    With Application.FileDialog(dialogTypeFilePicker)
        If LenB(dialogTitle) > 0 Then .Title = dialogTitle
        If LenB(initialPath) > 0 Then .InitialFileName = initialPath
        If LenB(.InitialFileName) = 0 Then
            Dim app As Object: Set app = Application 'Needs to be late-binded
            Select Case Application.Name
                Case "Microsoft Excel": .InitialFileName = GetLocalPath(app.ThisWorkbook.Path)
                Case "Microsoft Word":  .InitialFileName = GetLocalPath(app.ThisDocument.Path)
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
#If Mac = 0 Then
Public Function BrowseForFolder(Optional ByVal initialPath As String _
                              , Optional ByVal dialogTitle As String) As String
    'In case reference to Microsoft Office X.XX Object Library is missing
    Const dialogTypeFolderPicker As Long = 4 'msoFileDialogFolderPicker
    Const actionButton As Long = -1
    '
    With Application.FileDialog(dialogTypeFolderPicker)
        If LenB(dialogTitle) > 0 Then .Title = dialogTitle
        If LenB(initialPath) > 0 Then .InitialFileName = initialPath
        If LenB(.InitialFileName) = 0 Then
            Dim app As Object: Set app = Application 'Needs to be late-binded
            Select Case Application.Name
                Case "Microsoft Excel": .InitialFileName = GetLocalPath(app.ThisWorkbook.Path)
                Case "Microsoft Word":  .InitialFileName = GetLocalPath(app.ThisDocument.Path)
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
                        , ByVal fsName As String) As String
    Const parentFolder As String = ".." & PATH_SEPARATOR
    '
    fsName = FixPathSeparators(fsName)
    If Left$(fsName, 3) = parentFolder Then
        Dim sepIndex As Long
        '
        folderPath = FixPathSeparators(folderPath)
        Do
            sepIndex = InStrRev(folderPath, PATH_SEPARATOR, Len(folderPath) - 1)
            If sepIndex < 3 Then Exit Do
            '
            folderPath = Left$(folderPath, sepIndex)
            fsName = Right$(fsName, Len(fsName) - 3)
        Loop Until Left$(fsName, 3) <> parentFolder
    End If
    BuildPath = FixPathSeparators(folderPath & PATH_SEPARATOR & fsName)
End Function

'*******************************************************************************
'Copies a file. Overwrites existing files unless 'failIfExists' is set to True
'Note that VBA.FileCopy does not copy opened files on Windows but it does on Mac
'If the destination file already exists and 'failIfExists' is set to False
'   then this method must be able to overwrite the destination file. Rather than
'   failing and then trying again with attribute set to vbNormal this method
'   sets the attribute for the destination path to vbNormal before copying.
'   This is slightly slower than just copying directly but far outperforms two
'   copy operations in the case where the first one fails and the second one is
'   done after setting the destination file attribute to vbNormal.
'*******************************************************************************
Public Function CopyFile(ByVal sourcePath As String _
                       , ByVal destinationPath As String _
                       , Optional ByVal failIfExists As Boolean = False) As Boolean
    If LenB(sourcePath) = 0 Then Exit Function
    If LenB(destinationPath) = 0 Then Exit Function
    '
    #If Mac Then
        If failIfExists Then If IsFile(destinationPath) Then Exit Function
        '
        On Error Resume Next
        SetAttr destinationPath, vbNormal 'Too costly to do after Copy fails
        Err.Clear 'Ignore any errors raised by 'SetAttr'
        VBA.FileCopy sourcePath, destinationPath 'Copies opened files as well
        CopyFile = (Err.Number = 0)
        On Error GoTo 0
    #Else
        If Not failIfExists Then
            On Error Resume Next
            SetAttr destinationPath, vbNormal 'Costly to do after Copy fails
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
'   different disk drives. Use this parameter with care
'*******************************************************************************
Public Function CopyFolder(ByVal sourcePath As String _
                         , ByVal destinationPath As String _
                         , Optional ByVal includeSubFolders As Boolean = True _
                         , Optional ByVal failIfExists As Boolean = False _
                         , Optional ByVal ignoreFailedFiles As Boolean = False) As Boolean
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
                           , Optional ByVal failIfExists As Boolean = False) As Boolean
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
        sepIndex = InStrRev(fullPath, PATH_SEPARATOR)
        If sepIndex < 3 Then Exit Do
        '
        fullPath = Left$(fullPath, sepIndex - 1)
        If IsFolder(fullPath) Then Exit Do
    Loop
    On Error Resume Next
    For i = collFoldersToCreate.Count To 1 Step -1
        'MkDir does not support all Unicode characters, unlike FSO
        #If Mac Then
            MkDir collFoldersToCreate(i)
        #Else
            GetFileSystemObject.CreateFolder collFoldersToCreate(i)
        #End If
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
    If LenB(filePath) = 0 Then Exit Function
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
                           , Optional ByVal failIfMissing As Boolean = False) As Boolean
    If LenB(folderPath) = 0 Then Exit Function
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
                          , Optional ByVal isFATFileSystem As Boolean = False) As String
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
#If Mac = 0 Then
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
#If Mac = 0 Then
Private Function IsReservedName(ByVal nameToCheck As String) As Boolean
    Static collReservedNames As Collection
    '
    If collReservedNames Is Nothing Then
        Dim v As Variant
        '
        Set collReservedNames = New Collection
        For Each v In Split("com1,com2,com3,com4,com5,com6,com7,com8,com9," _
        & "lpt1,lpt2,lpt3,lpt4,lpt5,lpt6,lpt7,lpt8,lpt9,con,nul,prn,aux", ",")
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
    Const oneSeparator As String = PATH_SEPARATOR
    Const twoSeparators As String = PATH_SEPARATOR & PATH_SEPARATOR
    Dim resultPath As String: resultPath = pathToFix
    '
    #If Mac = 0 Then 'Replace forward slashes with back slashes for Windows
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
    #If Mac = 0 Then
        If isUNC Then resultPath = oneSeparator & resultPath
    #End If
    '
    FixPathSeparators = resultPath
End Function

'*******************************************************************************
'Retrieves the owner name for a file path
'*******************************************************************************
#If Mac = 0 Then
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
    #If VBA7 Then
        Dim pOwner As LongPtr
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
    GetFileOwner = owDomain & PATH_SEPARATOR & owName
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
                       , Optional ByVal includeSystem As Boolean = False) As Collection
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
                     , ByVal fAttribute As VbFileAttribute)
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
                         , Optional ByVal includeSystem As Boolean = False) As Collection
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
                       , ByVal fAttribute As VbFileAttribute)
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
            If IsFolder(fullPath) Then collFolders.Add fullPath
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
#If Mac = 0 Then
Public Function GetLocalPath(ByVal fullPath As String) As String
    With GetDriveInfo(fullPath)
        If LenB(.driveLetter) = 0 Then
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
'Returns the UNC path for a given path or null string if path is not remote
'Note that the input path does not need to be an existing file/folder
'*******************************************************************************
#If Mac = 0 Then
Public Function GetUNCPath(ByVal fullPath As String) As String
    With GetDriveInfo(fullPath)
        If LenB(.shareName) = 0 Then Exit Function  'Not UNC
        GetUNCPath = FixPathSeparators(Replace(fullPath _
        , .driveName, .shareName, 1, 1, vbTextCompare))
    End With
End Function
#End If

'*******************************************************************************
'Returns the web path for a OneDrive local path or null string if not OneDrive
'Note that the input path does not need to be an existing file/folder
'*******************************************************************************
#If Mac = 0 Then
Public Function GetWebPath(ByVal fullPath As String) As String
    GetWebPath = GetOneDriveWebPath(fullPath, rebuildCache:=False)
End Function
#End If

'*******************************************************************************
'Returns basic drive information about a full path
'*******************************************************************************
#If Mac = 0 Then
Private Function GetDriveInfo(ByVal fullPath As String) As DRIVE_INFO
    Dim fso As Object: Set fso = GetFileSystemObject()
    If fso Is Nothing Then Exit Function
    '
    Dim driveName As String: driveName = fso.GetDriveName(fullPath)
    If LenB(driveName) = 0 Then Exit Function
    '
    Dim fsDrive As Object
    On Error Resume Next
    Set fsDrive = fso.GetDrive(driveName)
    On Error GoTo 0
    If fsDrive Is Nothing Then Exit Function
    '
    If LenB(fsDrive.driveLetter) = 0 Then
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
        If LenB(.shareName) > 0 Then
            .driveName = AlignDriveNameIfNeeded(.driveName, .shareName)
        End If
    End With
End Function
#End If

'*******************************************************************************
'Late-bounded file system for Windows
'*******************************************************************************
#If Mac = 0 Then
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
#If Mac = 0 Then
Private Function AlignDriveNameIfNeeded(ByVal driveName As String _
                                      , ByVal shareName As String) As String
    Dim sepIndex As Long
    '
    sepIndex = VBA.InStr(3, driveName, PATH_SEPARATOR)
    If sepIndex > 0 Then
        Dim newName As String: newName = VBA.Left$(driveName, sepIndex - 1)
        sepIndex = VBA.InStr(3, shareName, PATH_SEPARATOR)
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
'
'With the help of: @guwidoe (https://github.com/guwidoe)
'See: https://github.com/cristianbuse/VBA-FileTools/issues/1
'*******************************************************************************
#If Mac = 0 Then
Public Function GetOneDriveLocalPath(ByVal odWebPath As String _
                                   , ByVal rebuildCache As Boolean) As String
    If InStr(1, odWebPath, "https://", vbTextCompare) <> 1 Then Exit Function
    '
    Dim collMatches As New Collection
    Dim bestMatch As Long
    Dim mainIndex As Long
    Dim i As Long
    '
    If rebuildCache Or Not m_providers.isSet Then ReadProviders
    For i = 1 To m_providers.pCount
        With m_providers.arr(i)
            If StrCompLeft(odWebPath, .webPath, vbTextCompare) = 0 Then
                collMatches.Add i
                If Not .isBusiness Then Exit For
                If .isMain Then mainIndex = .accountIndex
            End If
        End With
    Next i
    '
    Select Case collMatches.Count
    Case 0: Exit Function
    Case 1: bestMatch = collMatches(1)
    Case Else
        Dim pos As Long: pos = Len(odWebPath) + 1
        Dim tempPath As String
        Dim webPath As String
        Dim rPart As String
        Dim localPath As String
        Dim v As Variant
        Do
            pos = InStrRev(odWebPath, "/", pos - 1)
            tempPath = Left$(odWebPath, pos)
            For Each v In collMatches
                With m_providers.arr(v)
                    rPart = Mid$(tempPath, Len(.webPath) + 1)
                    localPath = BuildPath(.mountPoint, rPart)
                    If IsFolder(localPath) Then
                        If bestMatch = 0 Or .isMain Then
                            bestMatch = v
                        Else
                            If IsBetterMatch(m_providers.arr(bestMatch) _
                                           , m_providers.arr(v) _
                                           , mainIndex _
                                           , localPath) Then
                                bestMatch = v
                            End If
                        End If
                    End If
                End With
            Next v
        Loop Until bestMatch > 0
    End Select
    With m_providers.arr(bestMatch)
        rPart = Mid$(odWebPath, Len(.webPath) + 1)
        GetOneDriveLocalPath = BuildPath(.mountPoint, rPart)
    End With
End Function
#End If
Private Function StrCompLeft(ByRef s1 As String _
                           , ByRef s2 As String _
                           , ByVal compareMethod As VbCompareMethod) As Long
    If Len(s1) > Len(s2) Then
        StrCompLeft = StrComp(Left$(s1, Len(s2)), s2, compareMethod)
    Else
        StrCompLeft = StrComp(s1, Left$(s2, Len(s1)), compareMethod)
    End If
End Function
Private Function IsBetterMatch(ByRef lastProvider As ONEDRIVE_PROVIDER _
                             , ByRef currProvider As ONEDRIVE_PROVIDER _
                             , ByRef mainIndex As Long _
                             , ByRef localPath As String) As Boolean
    If lastProvider.isMain Then Exit Function
    '
    Dim isLastOnMain As Boolean: isLastOnMain = (lastProvider.accountIndex = mainIndex)
    Dim isCurrOnMain As Boolean: isCurrOnMain = (currProvider.accountIndex = mainIndex)
    '
    If isLastOnMain Xor isCurrOnMain Then
        IsBetterMatch = isCurrOnMain
    Else
        IsBetterMatch = IsFolderEditable(localPath)
    End If
End Function

'*******************************************************************************
'Returns the web path for a OneDrive local path
'Returns null string if the path provided is not a valid OneDrive local path
'*******************************************************************************
#If Mac = 0 Then
Private Function GetOneDriveWebPath(ByVal odLocalPath As String _
                                  , ByVal rebuildCache As Boolean) As String
    If InStr(1, odLocalPath, ":\", vbTextCompare) <> 2 Then Exit Function
    '
    Dim localPath As String
    Dim rPart As String
    Dim bestMatch As Long
    Dim i As Long
    '
    odLocalPath = FixPathSeparators(odLocalPath)
    If rebuildCache Or Not m_providers.isSet Then ReadProviders
    For i = 1 To m_providers.pCount
        localPath = m_providers.arr(i).mountPoint
        If StrCompLeft(odLocalPath, localPath, vbTextCompare) = 0 Then
            If bestMatch = 0 Then
                bestMatch = i
            ElseIf Len(localPath) > Len(m_providers.arr(bestMatch).mountPoint) _
            Then
                bestMatch = i
            End If
        End If
    Next i
    If bestMatch = 0 Then Exit Function
    '
    With m_providers.arr(bestMatch)
        rPart = Replace(Mid$(odLocalPath, Len(.mountPoint) + 1), "\", "/")
        GetOneDriveWebPath = .webPath & rPart
    End With
End Function
#End If

'*******************************************************************************
'Utility for 'GetOneDriveLocalPath' and 'GetOneDriveWebPath'
'*******************************************************************************
Private Sub ReadProviders()
    Dim folderPath As Variant
    Dim folderName As String
    Dim appPath As String
    '
    m_providers.pCount = 0
    '
    appPath = BuildPath(Environ$("LOCALAPPDATA"), "Microsoft\OneDrive\settings\")
    For Each folderPath In GetFolders(appPath)
        folderName = Right$(folderPath, Len(folderPath) - Len(appPath))
        If folderName Like "Business*" Then
            AddBusinessPaths folderPath
        ElseIf folderName = "Personal" Then
            AddPersonalPaths folderPath
        End If
    Next folderPath
    m_providers.isSet = True
End Sub
Private Function AddProvider() As Long
    If m_providers.pCount = 0 Then
        ReDim m_providers.arr(1 To 1)
    Else
        ReDim Preserve m_providers.arr(1 To m_providers.pCount + 1)
    End If
    m_providers.pCount = m_providers.pCount + 1
    AddProvider = m_providers.pCount
End Function
Private Sub AddBusinessPaths(ByVal folderPath As String)
    Const businessIniMask As String = "????????-????-????-????-????????????.ini"
    Dim iniName As String: iniName = Dir(BuildPath(folderPath, businessIniMask))
    If LenB(iniName) = 0 Then Exit Sub
    '
    Dim iniPath As String: iniPath = BuildPath(folderPath, iniName)
    Dim datPath As String: datPath = Replace(iniPath, ".ini", ".dat")
    Dim bytes() As Byte:   ReadBytes iniPath, bytes
    Dim lineText As Variant
    Dim temp() As String
    Dim tempMount As String
    Dim mainMount As String
    Dim accountIndex As Long
    Dim tempURL As String
    Dim cFolders As Collection
    Dim cParents As Collection
    Dim cPending As New Collection
    Dim canAdd As Boolean
    '
    accountIndex = CLng(Mid$(folderPath, InStrRev(folderPath, "Business") + 8))
    For Each lineText In Split(bytes, vbNewLine)
        Dim parts() As String: parts = Split(lineText, """")
        Select Case Left$(lineText, InStr(1, lineText, " "))
        Case "libraryScope "
            tempMount = parts(9)
            canAdd = (LenB(tempMount) > 0)
            If parts(3) = "ODB" Then
                mainMount = tempMount
                tempURL = GetUrlNamespace(folderPath)
            Else
                temp = Split(parts(8), " ")
                tempURL = GetUrlNamespace(folderPath, "_" & temp(3) & temp(1))
            End If
            If Not canAdd Then cPending.Add tempURL, Split(parts(0), " ")(2)
        Case "libraryFolder "
            If cFolders Is Nothing Then
                Set cFolders = GetODFolders(datPath, cParents)
            End If
            tempMount = parts(1)
            temp = Split(parts(0), " ")
            tempURL = cPending(temp(3))
            Dim tempID As String:     tempID = Split(temp(4), "+")(0)
            Dim tempFolder As String: tempFolder = vbNullString
            On Error Resume Next
            Do
                tempFolder = cFolders(tempID) & "/" & tempFolder
                tempID = cParents(tempID)
            Loop Until Err.Number <> 0
            On Error GoTo 0
            canAdd = (LenB(tempFolder) > 0)
            tempURL = tempURL & tempFolder
        Case "AddedScope "
            If cFolders Is Nothing Then
                Set cFolders = GetODFolders(datPath, cParents)
            End If
            tempID = Split(parts(0), " ")(3)
            tempFolder = vbNullString
            On Error Resume Next
            Do
                tempFolder = cFolders(tempID) & "\" & tempFolder
                tempID = cParents(tempID)
            Loop Until Err.Number <> 0
            On Error GoTo 0
            tempMount = mainMount & "\" & tempFolder
            tempURL = parts(5)
            If tempURL = " " Or LenB(tempURL) = 0 Then
                tempURL = vbNullString
            Else
                tempURL = tempURL & "/"
            End If
            temp = Split(parts(4), " ")
            tempURL = GetUrlNamespace(folderPath, "_" & temp(3) & temp(1) _
                                                      & temp(4)) & tempURL
            canAdd = True
        Case Else
            Exit For
        End Select
        If canAdd Then
            With m_providers.arr(AddProvider())
                .webPath = tempURL
                .mountPoint = BuildPath(tempMount, vbNullString)
                .isBusiness = True
                .isMain = (tempMount = mainMount)
                .accountIndex = accountIndex
            End With
        End If
    Next lineText
End Sub
Private Function GetUrlNamespace(ByVal folderPath As String _
                               , Optional ByVal cSignature As String) As String
    Const nTag As String = "DavUrlNamespace"
    Dim bytes() As Byte
    Dim lineText As Variant
    '
    ReadBytes BuildPath(folderPath, "ClientPolicy" & cSignature & ".ini"), bytes
    For Each lineText In Split(bytes, vbNewLine)
        If Left$(lineText, Len(nTag)) = nTag Then
            GetUrlNamespace = Mid$(lineText, InStr(Len(nTag), lineText, "https"))
            Exit Function
        End If
    Next lineText
End Function
Private Sub AddPersonalPaths(ByVal folderPath As String)
    Dim datName As String: datName = Dir(BuildPath(folderPath, "*.dat"))
    Dim iniPath As String
    Do
        If LenB(datName) = 0 Then Exit Sub
        iniPath = BuildPath(folderPath, Replace(datName, ".dat", ".ini"))
        If IsFile(iniPath) Then Exit Do
        datName = Dir
    Loop
    Dim mainURL As String: mainURL = GetUrlNamespace(folderPath) & "/"
    Dim cid As String:     cid = Replace(datName, ".dat", vbNullString)
    Dim tempURL As String: tempURL = mainURL & cid & "/"
    Dim bytes() As Byte:   ReadBytes iniPath, bytes
    Dim lineText As Variant
    Dim mainMount As String
    '
    For Each lineText In Split(bytes, vbNewLine)
        If Left$(lineText, InStr(1, lineText, " ")) = "library " Then
            mainMount = Split(lineText, """")(3) & "\"
            Exit For
        End If
    Next lineText
    With m_providers.arr(AddProvider())
        .webPath = tempURL
        .mountPoint = mainMount
    End With
    '
    Dim groupPath As String
    Dim i As Long
    Dim relPath As String
    Dim folderID As String
    Dim cFolders As Collection
    Dim tempMount As String
    '
    groupPath = BuildPath(folderPath, "GroupFolders.ini")
    If Not IsFile(groupPath) Then Exit Sub
    ReadBytes groupPath, bytes
    '
    For Each lineText In Split(bytes, vbNewLine)
        If InStr(1, lineText, "_BaseUri", vbTextCompare) > 0 Then
            cid = LCase$(Mid$(lineText, InStrRev(lineText, "/") + 1))
            i = InStr(1, cid, "!")
            If i > 0 Then cid = Left$(cid, i - 1)
        Else
            i = InStr(1, lineText, "_Path", vbTextCompare)
            If i > 0 Then
                relPath = Mid$(lineText, InStr(lineText, " = ") + 3)
                folderID = Left$(lineText, i - 1)
                If cFolders Is Nothing Then
                    Set cFolders = GetODFolders(Replace(iniPath, ".ini", ".dat"))
                End If
                If cFolders.Count > 0 Then
                    With m_providers.arr(AddProvider())
                        .webPath = mainURL & cid & "/" & relPath & "/"
                        .mountPoint = mainMount & cFolders(folderID) & "\"
                    End With
                End If
            End If
        End If
    Next lineText
End Sub

'*******************************************************************************
'Utility - Retrieves all folders from an OneDrive user .dat file
'*******************************************************************************
Private Function GetODFolders(ByVal filePath As String _
                            , Optional ByRef outParents As Collection) As Collection
    If Not IsFile(filePath) Then Exit Function
    '
    Dim fileNumber As Long: fileNumber = FreeFile()
    '
    Open filePath For Binary Access Read As #fileNumber
    Dim size As Long: size = LOF(fileNumber)
    If size = 0 Then GoTo CloseFile
    '
    Const hCheckSize As Long = 8
    Const idSize As Long = 39
    Const fNameOffset As Long = 121
    Const checkToName As Long = hCheckSize + idSize + fNameOffset + fNameOffset
    '
    Const chunkSize As Long = &H100000 '1MB
    Dim b(1 To chunkSize) As Byte
    Dim s As String
    Dim lastRecord As Long
    Dim i As Long
    Dim cFolders As Collection
    Dim stepSize As Long
    Dim bytes As Long
    Dim folderID As String
    Dim parentID As String
    Dim folderName As String
    Dim lastFileChange As Date
    Dim currFileChange As Date
    Dim vbNullByte As String:   vbNullByte = MidB$(vbNullChar, 1, 1)
    Dim hFolder As String:      hFolder = StrConv(Chr$(&H2), vbFromUnicode) 'x02
    Dim hCheck As String:       hCheck = ChrW$(&H1) & String(3, vbNullChar) 'x01..
    '
    For stepSize = 16 To 8 Step -8
        lastFileChange = 0
        Do
            i = 0
            currFileChange = FileDateTime(filePath)
            If currFileChange > lastFileChange Then
                Set cFolders = New Collection
                Set outParents = New Collection
                lastFileChange = currFileChange
                lastRecord = 1
            End If
            Get fileNumber, lastRecord, b
            s = b
            i = InStrB(stepSize, s, hCheck)
            Do While i > stepSize And i < chunkSize - checkToName
                If MidB$(s, i - stepSize, 1) = hFolder Then
                    i = i + hCheckSize
                    bytes = Clamp(InStrB(i, s, vbNullByte) - i, 0, idSize)
                    folderID = StrConv(MidB$(s, i, bytes), vbUnicode)
                    i = i + idSize
                    bytes = Clamp(InStrB(i, s, vbNullByte) - i, 0, idSize)
                    parentID = StrConv(MidB$(s, i, bytes), vbUnicode)
                    i = i + fNameOffset
                    bytes = InStr(-Int(-(i - 1) / 2) + 1, s, vbNullChar) * 2 - i - 1
                    bytes = bytes + bytes Mod 2
                    If bytes < 0 Or i + bytes - 1 > chunkSize Then 'Read next chunk
                        i = i - checkToName
                        Exit Do
                    End If
                    folderName = MidB$(s, i, bytes)
                    If LenB(folderID) > 0 And LenB(parentID) > 0 Then
                        cFolders.Add folderName, folderID
                        outParents.Add parentID, folderID
                    End If
                End If
                i = InStrB(i + 1, s, hCheck)
            Loop
            lastRecord = lastRecord + chunkSize - stepSize
            If i > stepSize Then lastRecord = lastRecord - chunkSize + i - 1
        Loop Until lastRecord > size
        If cFolders.Count > 0 Then Exit For
    Next stepSize
    Set GetODFolders = cFolders
CloseFile:
    Close #fileNumber
End Function
Private Function Clamp(ByVal v As Long, ByVal lowB As Long, uppB As Long) As Long
    If v < lowB Then
        Clamp = lowB
    ElseIf v > uppB Then
        Clamp = uppB
    Else
        Clamp = v
    End If
End Function

'*******************************************************************************
'Checks if a path indicates a file path
'Note that if C:\Test\1.txt is valid then C:\Test\\///1.txt will also be valid
'Most VBA methods consider valid any path separators with multiple characters
'*******************************************************************************
Public Function IsFile(ByVal filePath As String) As Boolean
    Const errBadFileNameOrNumber As Long = 52
    Dim fAttr As VbFileAttribute
    '
    On Error Resume Next
    fAttr = GetAttr(filePath)
    If Err.Number = errBadFileNameOrNumber Then 'Unicode characters
        #If Mac Then
            
        #Else
            IsFile = GetFileSystemObject().FileExists(filePath)
        #End If
    ElseIf Err.Number = 0 Then
        IsFile = Not CBool(fAttr And vbDirectory)
    End If
    On Error GoTo 0
End Function
'*******************************************************************************
'Checks if a path indicates a folder path
'Note that if C:\Test\Demo is valid then C:\Test\\///Demo will also be valid
'Most VBA methods consider valid any path separators with multiple characters
'*******************************************************************************
Public Function IsFolder(ByVal folderPath As String) As Boolean
    Const errBadFileNameOrNumber As Long = 52
    Dim fAttr As VbFileAttribute
    '
    On Error Resume Next
    fAttr = GetAttr(folderPath)
    If Err.Number = errBadFileNameOrNumber Then 'Unicode characters
        #If Mac Then
            
        #Else
            IsFolder = GetFileSystemObject().FolderExists(folderPath)
        #End If
    Else
        IsFolder = CBool(fAttr And vbDirectory)
    End If
    On Error GoTo 0
End Function

'*******************************************************************************
'Checks if the contents of a folder can be edited
'*******************************************************************************
Public Function IsFolderEditable(ByVal folderPath As String) As Boolean
    Dim tempFolder As String
    '
    folderPath = BuildPath(folderPath, vbNullString)
    Do
        tempFolder = folderPath & Rnd()
    Loop Until Not IsFolder(tempFolder)
    '
    On Error Resume Next
    MkDir tempFolder
    If Err.Number = 0 Then RmDir tempFolder
    IsFolderEditable = (Err.Number = 0)
    On Error GoTo 0
End Function

'*******************************************************************************
'Moves (or renames) a file
'*******************************************************************************
Public Function MoveFile(ByVal sourcePath As String _
                       , ByVal destinationPath As String) As Boolean
    If LenB(sourcePath) = 0 Then Exit Function
    If LenB(destinationPath) = 0 Then Exit Function
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
                         , ByVal destinationPath As String) As Boolean
    If LenB(sourcePath) = 0 Then Exit Function
    If LenB(destinationPath) = 0 Then Exit Function
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
    #If Mac = 0 Then
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

'*******************************************************************************
'Reads a file into an array of Bytes
'*******************************************************************************
Public Sub ReadBytes(ByVal filePath As String, ByRef result() As Byte)
    If Not IsFile(filePath) Then
        Erase result
        Exit Sub
    End If
    '
    Dim fileNumber As Long: fileNumber = FreeFile()
    '
    Open filePath For Binary Access Read As #fileNumber
    Dim size As Long: size = LOF(fileNumber)
    If size > 0 Then
        ReDim result(0 To size - 1)
        Get fileNumber, , result
    Else
        Erase result
    End If
    Close #fileNumber
End Sub
