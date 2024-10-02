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
''    - BrowseForFiles           (Windows only)
''    - BrowseForFolder          (Windows only)
''    - BuildPath
''    - ConvertText
''    - CopyFile
''    - CopyFolder
''    - CreateFolder
''    - DecodeURL
''    - DeleteFile
''    - DeleteFolder
''    - FixFileName
''    - FixPathSeparators
''    - GetFileOwner             (Windows only)
''    - GetFiles
''    - GetFolders
''    - GetKnownFolderCLSID      (Windows only)
''    - GetKnownFolderPath       (Windows only)
''    - GetLocalPath
''    - GetRelativePath
''    - GetRemotePath
''    - GetSpecialFolderConstant (Mac only)
''    - GetSpecialFolderDomain   (Mac only)
''    - GetSpecialFolderPath     (Mac only)
''    - IsFile
''    - IsFolder
''    - IsFolderEditable
''    - MoveFile
''    - MoveFolder
''    - ParentFolder
''    - ReadBytes
'*******************************************************************************

Option Explicit
Option Private Module

#Const Windows = (Mac = 0)

#If Mac Then
    #If VBA7 Then 'https://developer.apple.com/library/archive/documentation/System/Conceptual/ManPages_iPhoneOS/man3/iconv.3.html
        Private Declare PtrSafe Function iconv Lib "/usr/lib/libiconv.dylib" (ByVal cd As LongPtr, ByRef inBuf As LongPtr, ByRef inBytesLeft As LongPtr, ByRef outBuf As LongPtr, ByRef outBytesLeft As LongPtr) As LongPtr
        Private Declare PtrSafe Function iconv_open Lib "/usr/lib/libiconv.dylib" (ByVal toCode As LongPtr, ByVal fromCode As LongPtr) As LongPtr
        Private Declare PtrSafe Function iconv_close Lib "/usr/lib/libiconv.dylib" (ByVal cd As LongPtr) As Long
    #Else
        Private Declare Function iconv Lib "/usr/lib/libiconv.dylib" (ByVal cd As Long, ByRef inBuf As Long, ByRef inBytesLeft As Long, ByRef outBuf As Long, ByRef outBytesLeft As Long) As Long
        Private Declare Function iconv_open Lib "/usr/lib/libiconv.dylib" (ByVal toCode As Long, ByVal fromCode As Long) As Long
        Private Declare Function iconv_close Lib "/usr/lib/libiconv.dylib" (ByVal cd As Long) As Long
    #End If
#Else
    #If VBA7 Then
        Private Declare PtrSafe Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
        Private Declare PtrSafe Function GetOpenFileNameW Lib "comdlg32.dll" (pOpenfilename As OPENFILENAME) As Long
        Private Declare PtrSafe Function CopyFileW Lib "kernel32" (ByVal lpExistingFileName As LongPtr, ByVal lpNewFileName As LongPtr, ByVal bFailIfExists As Long) As Long
        Private Declare PtrSafe Function DeleteFileW Lib "kernel32" (ByVal lpFileName As LongPtr) As Long
        Private Declare PtrSafe Function RemoveDirectoryW Lib "kernel32" (ByVal lpPathName As LongPtr) As Long
        Private Declare PtrSafe Function GetFileSecurity Lib "advapi32.dll" Alias "GetFileSecurityA" (ByVal lpFileName As String, ByVal RequestedInformation As Long, pSecurityDescriptor As Byte, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
        Private Declare PtrSafe Function GetSecurityDescriptorOwner Lib "advapi32.dll" (pSecurityDescriptor As Byte, pOwner As LongPtr, lpbOwnerDefaulted As LongPtr) As Long
        Private Declare PtrSafe Function LookupAccountSid Lib "advapi32.dll" Alias "LookupAccountSidA" (ByVal lpSystemName As String, ByVal Sid As LongPtr, ByVal Name As String, cbName As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As LongPtr) As Long
        Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal codePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
        Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal codePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
        Private Declare PtrSafe Function SHGetKnownFolderPath Lib "shell32" (ByRef rfID As GUID, ByVal dwFlags As Long, ByVal hToken As Long, ByRef pszPath As LongPtr) As Long
        Private Declare PtrSafe Function CLSIDFromString Lib "ole32" (ByVal lpszGuid As LongPtr, ByRef pGuid As GUID) As Long
        Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
        Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32" (ByVal hMem As LongPtr)
        Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As LongPtr)
    #Else
        Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
        Private Declare Function GetOpenFileNameW Lib "comdlg32.dll" (pOpenfilename As OPENFILENAME) As Long
        Private Declare Function CopyFileW Lib "kernel32" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long, ByVal bFailIfExists As Long) As Long
        Private Declare Function DeleteFileW Lib "kernel32" (ByVal lpFileName As Long) As Long
        Private Declare Function RemoveDirectoryW Lib "kernel32" (ByVal lpPathName As Long) As Long
        Private Declare Function GetFileSecurity Lib "advapi32.dll" Alias "GetFileSecurityA" (ByVal lpFileName As String, ByVal RequestedInformation As Long, pSecurityDescriptor As Byte, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
        Private Declare Function GetSecurityDescriptorOwner Lib "advapi32.dll" (pSecurityDescriptor As Byte, pOwner As Long, lpbOwnerDefaulted As Long) As Long
        Private Declare Function LookupAccountSid Lib "advapi32.dll" Alias "LookupAccountSidA" (ByVal lpSystemName As String, ByVal Sid As Long, ByVal Name As String, cbName As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Long) As Long
        Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal codePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
        Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal codePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
        Private Declare Function SHGetKnownFolderPath Lib "shell32" (rfID As Any, ByVal dwFlags As Long, ByVal hToken As Long, ppszPath As Long) As Long
        Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszGuid As Long, pGuid As Any) As Long
        Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
        Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
        Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
    #End If
#End If

#If VBA7 = 0 Then
    Private Enum LongPtr
        [_]
    End Enum
#End If

Public Enum PageCode
    [_pcCount] = 5
    codeUTF8 = 65001
    codeUTF16LE = 1200
    codeUTF16BE = 1201
#If Mac Then
    codeUTF32LE = 12000
    codeUTF32BE = 12001
#End If
End Enum

#If Mac Then
    Public Enum SpecialFolderConstant 'See 'GetSpecialFolderConstant'
        sfc_ApplicationSupport
        [_minSFC] = sfc_ApplicationSupport
        sfc_ApplicationsFolder
        sfc_Desktop
        sfc_DesktopPicturesFolder
        sfc_DocumentsFolder
        sfc_DownloadsFolder
        sfc_FavoritesFolder
        sfc_FolderActionScripts
        sfc_Fonts
        sfc_Help
        sfc_HomeFolder
        sfc_InternetPlugins
        sfc_KeychainFolder
        sfc_LibraryFolder
        sfc_ModemScripts
        sfc_MoviesFolder
        sfc_MusicFolder
        sfc_PicturesFolder
        sfc_Preferences
        sfc_PrinterDescriptions
        sfc_PublicFolder
        sfc_ScriptingAdditions
        sfc_ScriptsFolder
        sfc_ServicesFolder
        sfc_SharedDocuments
        sfc_SharedLibraries
        sfc_SitesFolder
        sfc_StartupDisk
        sfc_StartupItems
        sfc_SystemFolder
        sfc_SystemPreferences
        sfc_TemporaryItems
        sfc_Trash
        sfc_UsersFolder
        sfc_UtilitiesFolder
        sfc_WorkflowsFolder
        '
        'Classic domain only
        sfc_AppleMenu
        sfc_ControlPanels
        sfc_ControlStripModules
        sfc_Extensions
        sfc_LauncherItemsFolder
        sfc_PrinterDrivers
        sfc_Printmonitor
        sfc_ShutdownFolder
        sfc_SpeakableItems
        sfc_Stationery
        sfc_Voices
        [_maxSFC] = sfc_Voices
    End Enum
    Public Enum SpecialFolderDomain 'See 'GetSpecialFolderDomain
        [_sfdNone] = 0
        [_minSFD] = [_sfdNone]
        sfd_System
        sfd_Local
        sfd_Network
        sfd_User
        sfd_Classic
        [_maxSFD] = sfd_Classic
    End Enum
#Else
    Public Enum KnownFolderID 'See 'GetKnownFolderCLSID' method
        kfID_AccountPictures = 0
        [_minKfID] = kfID_AccountPictures
        kfID_AddNewPrograms
        kfID_AdminTools
        kfID_AllAppMods
        kfID_AppCaptures
        kfID_AppDataDesktop
        kfID_AppDataDocuments
        kfID_AppDataFavorites
        kfID_AppDataProgramData
        kfID_ApplicationShortcuts
        kfID_AppsFolder
        kfID_AppUpdates
        kfID_CameraRoll
        kfID_CameraRollLibrary
        kfID_CDBurning
        kfID_ChangeRemovePrograms
        kfID_CommonAdminTools
        kfID_CommonOEMLinks
        kfID_CommonPrograms
        kfID_CommonStartMenu
        kfID_CommonStartMenuPlaces
        kfID_CommonStartup
        kfID_CommonTemplates
        kfID_ComputerFolder
        kfID_ConflictFolder
        kfID_ConnectionsFolder
        kfID_Contacts
        kfID_ControlPanelFolder
        kfID_Cookies
        kfID_CurrentAppMods
        kfID_Desktop
        kfID_DevelopmentFiles
        kfID_Device
        kfID_DeviceMetadataStore
        kfID_Documents
        kfID_DocumentsLibrary
        kfID_Downloads
        kfID_Favorites
        kfID_Fonts
        kfID_Games
        kfID_GameTasks
        kfID_History
        kfID_HomeGroup
        kfID_HomeGroupCurrentUser
        kfID_ImplicitAppShortcuts
        kfID_InternetCache
        kfID_InternetFolder
        kfID_Libraries
        kfID_Links
        kfID_LocalAppData
        kfID_LocalAppDataLow
        kfID_LocalDocuments
        kfID_LocalDownloads
        kfID_LocalizedResourcesDir
        kfID_LocalMusic
        kfID_LocalPictures
        kfID_LocalStorage
        kfID_LocalVideos
        kfID_Music
        kfID_MusicLibrary
        kfID_NetHood
        kfID_NetworkFolder
        kfID_Objects3D
        kfID_OneDrive
        kfID_OriginalImages
        kfID_PhotoAlbums
        kfID_Pictures
        kfID_PicturesLibrary
        kfID_Playlists
        kfID_PrintersFolder
        kfID_PrintHood
        kfID_Profile
        kfID_ProgramData
        kfID_ProgramFiles
        kfID_ProgramFilesCommon
        kfID_ProgramFilesCommonX64
        kfID_ProgramFilesCommonX86
        kfID_ProgramFilesX64
        kfID_ProgramFilesX86
        kfID_Programs
        kfID_Public
        kfID_PublicDesktop
        kfID_PublicDocuments
        kfID_PublicDownloads
        kfID_PublicGameTasks
        kfID_PublicLibraries
        kfID_PublicMusic
        kfID_PublicPictures
        kfID_PublicRingtones
        kfID_PublicUserTiles
        kfID_PublicVideos
        kfID_QuickLaunch
        kfID_Recent
        kfID_RecordedCalls
        kfID_RecordedTVLibrary
        kfID_RecycleBinFolder
        kfID_ResourceDir
        kfID_RetailDemo
        kfID_Ringtones
        kfID_RoamedTileImages
        kfID_RoamingAppData
        kfID_RoamingTiles
        kfID_SampleMusic
        kfID_SamplePictures
        kfID_SamplePlaylists
        kfID_SampleVideos
        kfID_SavedGames
        kfID_SavedPictures
        kfID_SavedPicturesLibrary
        kfID_SavedSearches
        kfID_Screenshots
        kfID_SEARCH_CSC
        kfID_SEARCH_MAPI
        kfID_SearchHistory
        kfID_SearchHome
        kfID_SearchTemplates
        kfID_SendTo
        kfID_SidebarDefaultParts
        kfID_SidebarParts
        kfID_SkyDrive
        kfID_SkyDriveCameraRoll
        kfID_SkyDriveDocuments
        kfID_SkyDriveMusic
        kfID_SkyDrivePictures
        kfID_StartMenu
        kfID_StartMenuAllPrograms
        kfID_Startup
        kfID_SyncManagerFolder
        kfID_SyncResultsFolder
        kfID_SyncSetupFolder
        kfID_System
        kfID_SystemX86
        kfID_Templates
        kfID_UserPinned
        kfID_UserProfiles
        kfID_UserProgramFiles
        kfID_UserProgramFilesCommon
        kfID_UsersFiles
        kfID_UsersLibraries
        kfID_Videos
        kfID_VideosLibrary
        kfID_Windows
        [_maxKfID] = kfID_Windows
    End Enum
#End If

Private Type DRIVE_INFO
    driveName As String
    driveLetter As String
    fileSystem As String
    shareName As String
End Type

#If Windows Then
    Private Type GUID
        data1 As Long
        data2 As Integer
        data3 As Integer
        data4(0 To 7) As Byte
    End Type
    '
    'https://docs.microsoft.com/en-gb/windows/win32/api/commdlg/ns-commdlg-openfilenamea
    Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As LongPtr
        hInstance As LongPtr
        lpstrFilter As LongPtr
        lpstrCustomFilter As LongPtr
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As LongPtr
        nMaxFile As Long
        lpstrFileTitle As LongPtr
        nMaxFileTitle As Long
        lpstrInitialDir As LongPtr
        lpstrTitle As LongPtr
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As LongPtr
        lCustData As LongPtr
        lpfnHook As LongPtr
        lpTemplateName As LongPtr
        pvReserved As LongPtr
        dwReserved As Long
        flagsEx As Long
    End Type
#End If

Private Type ONEDRIVE_PROVIDER
    webPath As String
    mountPoint As String
    isBusiness As Boolean
    isMain As Boolean
    accountIndex As Long
    baseMount As String
    syncID As String
    #If Mac Then
        syncDir As String
    #End If
End Type
Private Type ONEDRIVE_PROVIDERS
    arr() As ONEDRIVE_PROVIDER
    pCount As Long
    isSet As Boolean
End Type

Private Type ONEDRIVE_ACCOUNT_INFO
    accountIndex As Long
    accountName As String
    cID As String
    clientPath As String
    datPath As String
    dbPath As String
    folderPath As String
    globalPath As String
    groupPath As String
    iniDateTime As Date
    iniPath As String
    isPersonal As Boolean
    isValid As Boolean
    hasDatFile As Boolean
End Type
Private Type ONEDRIVE_ACCOUNTS_INFO
    arr() As ONEDRIVE_ACCOUNT_INFO
    pCount As Long
    isSet As Boolean
End Type

Private Type DirInfo
    dirID As String
    parentID As String
    dirName As String
    isNameASCII As Boolean
End Type
Private Type DirsInfo
    idToIndex As Collection
    arrDirs() As DirInfo
    dirCount As Long
    dirUBound As Long
End Type

#If Mac Then
    Public Const PATH_SEPARATOR = "/"
#Else
    Public Const PATH_SEPARATOR = "\"
#End If

Private Const vbErrInvalidProcedureCall        As Long = 5
Private Const vbErrInternalError               As Long = 51
Private Const vbErrPathFileAccessError         As Long = 75
Private Const vbErrPathNotFound                As Long = 76
Private Const vbErrInvalidFormatInResourceFile As Long = 325
Private Const vbErrComponentNotRegistered      As Long = 336

Private m_providers As ONEDRIVE_PROVIDERS
#If Mac Then
    Private m_conversionDescriptors As New Collection
#End If

'*******************************************************************************
'Returns a Collection of file paths by using a FilePicker FileDialog
'Always returns an instantiated Collection
'
'More than one file extension may be specified in the 'filterExtensions' param
'   and each must be separated by a semi-colon. For example: "*.txt;*.csv".
'   Spaces will be ignored
'*******************************************************************************
Public Function BrowseForFiles(Optional ByRef initialPath As String _
                             , Optional ByRef dialogTitle As String _
                             , Optional ByRef filterDesc As String _
                             , Optional ByRef filterExtensions As String _
                             , Optional ByVal allowMultiFiles As Boolean = True) As Collection
    'msoFileDialogFilePicker = 3 - only available for some Microsoft apps
    Const dialogTypeFilePicker As Long = 3
    Const actionButton As Long = -1
    Dim filePicker As Object
    Dim app As Object: Set app = Application 'Late-binded for compatibility
    '
    On Error Resume Next
    Set filePicker = app.FileDialog(dialogTypeFilePicker)
    On Error GoTo 0
    '
    If filePicker Is Nothing Then
    #If Mac Then
        'Not implemented
        'Seems achievable via script:
        '   - https://stackoverflow.com/a/15546518/8488913
        '   - https://stackoverflow.com/a/37411960/8488913
    #Else
        Set BrowseForFiles = BrowseFilesAPI(initialPath, dialogTitle, filterDesc _
                                          , filterExtensions, allowMultiFiles)
    #End If
        Exit Function
    End If
    '
    With filePicker
        If LenB(dialogTitle) > 0 Then .Title = dialogTitle
        If LenB(initialPath) > 0 Then .InitialFileName = initialPath
        .allowMultiSelect = allowMultiFiles
        .filters.Clear
        If LenB(filterExtensions) > 0 Then
            On Error Resume Next
            .filters.Add filterDesc, filterExtensions
            On Error GoTo 0
        End If
        If .filters.Count = 0 Then .filters.Add "All Files", "*.*"
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

'*******************************************************************************
'Returns a Collection of file paths by creating an Open dialog box that lets the
'   user specify the drive, directory, and the name of the file(s)
'*******************************************************************************
#If Windows Then
Private Function BrowseFilesAPI(ByRef initialPath As String _
                              , ByRef dialogTitle As String _
                              , ByRef filterDesc As String _
                              , ByRef filterExtensions As String _
                              , ByVal allowMultiFiles As Boolean) As Collection
    Dim ofName As OPENFILENAME
    Dim resultPaths As New Collection
    Dim buffFiles As String
    Dim buffFilter As String
    Dim temp As String
    '
    With ofName
        On Error Resume Next
        Dim app As Object: Set app = Application
        .hwndOwner = app.Hwnd
        On Error GoTo 0
        '
        .lStructSize = LenB(ofName)
        If LenB(filterExtensions) = 0 Then
            buffFilter = "All Files (*.*)" & vbNullChar & "*.*"
        Else
            temp = Replace(filterExtensions, ",", ";")
            buffFilter = filterDesc & " (" & temp & ")" & vbNullChar & temp
        End If
        buffFilter = buffFilter & vbNullChar & vbNullChar
        .lpstrFilter = StrPtr(buffFilter)
        '
        .nMaxFile = &H100000
        buffFiles = VBA.Space$(.nMaxFile)
        .lpstrFile = StrPtr(buffFiles)
        .lpstrInitialDir = StrPtr(initialPath)
        .lpstrTitle = StrPtr(dialogTitle)
        '
        Const OFN_HIDEREADONLY As Long = &H4&
        Const OFN_ALLOWMULTISELECT As Long = &H200&
        Const OFN_PATHMUSTEXIST As Long = &H800&
        Const OFN_FILEMUSTEXIST As Long = &H1000&
        Const OFN_EXPLORER As Long = &H80000
        '
        .flags = OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST
        If allowMultiFiles Then
            .flags = .flags Or OFN_ALLOWMULTISELECT Or OFN_EXPLORER
        End If
    End With
    '
    Do
        Const FNERR_BUFFERTOOSMALL As Long = &H3003&
        Dim mustRetry As Boolean: mustRetry = False
        Dim i As Long
        Dim j As Long
        '
        If GetOpenFileNameW(ofName) Then
            i = InStr(1, buffFiles, vbNullChar)
            temp = Left$(buffFiles, i - 1)
            '
            If allowMultiFiles Then
                j = InStr(i + 1, buffFiles, vbNullChar)
                If j = i + 1 Then 'Single file selected
                    resultPaths.Add temp
                Else
                    temp = BuildPath(temp, vbNullString) 'Parent folder
                    Do
                        resultPaths.Add temp & Mid$(buffFiles, i + 1, j - i)
                        i = j
                        j = InStr(i + 1, buffFiles, vbNullChar)
                    Loop Until j = i + 1
                End If
            Else
                resultPaths.Add temp
            End If
        ElseIf CommDlgExtendedError() = FNERR_BUFFERTOOSMALL Then
            Dim b() As Byte: b = LeftB$(buffFiles, 4)
            '
            If b(3) And &H80 Then
                mustRetry = (MsgBox("Try selecting fewer files" _
                                  , vbExclamation + vbRetryCancel _
                                  , "Insufficient memory") = vbRetry)
            Else
                With ofName
                    .nMaxFile = b(3)
                    For i = 2 To 0 Step -1
                        .nMaxFile = .nMaxFile * &H100& + b(i)
                    Next i
                    buffFiles = VBA.Space$(.nMaxFile)
                    .lpstrFile = StrPtr(buffFiles)
                End With
                MsgBox "Did not expect so many files. Please select again!" _
                     , vbInformation, "Repeat selection"
                mustRetry = True
            End If
        End If
    Loop Until Not mustRetry
    Set BrowseFilesAPI = resultPaths
End Function
#End If

'*******************************************************************************
'Returns a folder path by using a FolderPicker FileDialog
'*******************************************************************************
Public Function BrowseForFolder(Optional ByRef initialPath As String _
                              , Optional ByRef dialogTitle As String) As String
#If Mac Then
    'If user has not accesss [initialPath] previously, will be prompted by
    'Mac OS to Grant permission to directory
    If LenB(initialPath) > 0 Then
        If Not Right(initialPath, 1) = Application.PathSeparator Then
            initialPath = initialPath & Application.PathSeparator
        End If
        Dir initialPath, Attributes:=vbDirectory
    End If
    Dim retPath
    If LenB(dialogTitle) = 0 Then dialogTitle = "Choose Foldler"
    retPath = MacScript("choose folder with prompt """ & dialogTitle & """ as string")
    If Len(retPath) > 0 Then
        retPath = MacScript("POSIX path of """ & retPath & """")
        If LenB(retPath) > 0 Then
            BrowseForFolder = retPath
        End If
    End If
#ElseIf Windows Then
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
                Case "Microsoft Excel": .InitialFileName = GetLocalPath(app.ThisWorkbook.Path, , True)
                Case "Microsoft Word":  .InitialFileName = GetLocalPath(app.ThisDocument.Path, , True)
            End Select
        End If
        If .Show = actionButton Then
            .InitialFileName = .SelectedItems.Item(1)
            BrowseForFolder = .InitialFileName
        End If
    End With
#End If
End Function

'*******************************************************************************
'Combines a folder path with a file/folder name or an incomplete path (ex. \a\b)
'*******************************************************************************
Public Function BuildPath(ParamArray PathComponents() As Variant) As String
    BuildPath = FixPathSeparators(Join(PathComponents, PATH_SEPARATOR))
End Function

'*******************************************************************************
'Converts a text between 2 page codes
'*******************************************************************************
#If Mac Then
Public Function ConvertText(ByRef textToConvert As String _
                          , ByVal toCode As PageCode _
                          , ByVal fromCode As PageCode _
                          , Optional ByVal persistDescriptor As Boolean = False) As String
#Else
Public Function ConvertText(ByRef textToConvert As String _
                          , ByVal toCode As PageCode _
                          , ByVal fromCode As PageCode) As String
#End If
    If toCode = fromCode Then
        ConvertText = textToConvert
        Exit Function
    End If
    #If Mac Then
        Dim inBytesLeft As LongPtr:  inBytesLeft = LenB(textToConvert)
        Dim outBytesLeft As LongPtr: outBytesLeft = inBytesLeft * 4
        Dim buffer As String:        buffer = Space$(CLng(inBytesLeft) * 2)
        Dim inBuf As LongPtr:        inBuf = StrPtr(textToConvert)
        Dim outBuf As LongPtr:       outBuf = StrPtr(buffer)
        Dim cd As LongPtr
        Dim cdKey As String:         cdKey = fromCode & "_" & toCode
        Dim cdFound As Boolean
        '
        On Error Resume Next
        cd = m_conversionDescriptors(cdKey)
        cdFound = (cd <> 0)
        On Error GoTo 0
        If Not cdFound Then
            cd = iconv_open(StrPtr(PageCodeToText(toCode)) _
                          , StrPtr(PageCodeToText(fromCode)))
            If persistDescriptor Then m_conversionDescriptors.Add cd, cdKey
        End If
        If iconv(cd, inBuf, inBytesLeft, outBuf, outBytesLeft) <> -1 Then
            ConvertText = LeftB$(buffer, LenB(buffer) - CLng(outBytesLeft))
        End If
        If Not (cdFound Or persistDescriptor) Then iconv_close cd
    #Else
        If toCode = codeUTF16LE Then
            ConvertText = EncodeToUTF16LE(textToConvert, fromCode)
        ElseIf fromCode = codeUTF16LE Then
            ConvertText = EncodeFromUTF16LE(textToConvert, toCode)
        Else
            ConvertText = EncodeFromUTF16LE( _
                          EncodeToUTF16LE(textToConvert, fromCode), toCode)
        End If
    #End If
End Function
#If Mac Then
Public Sub ClearConversionDescriptors()
    If m_conversionDescriptors.Count = 0 Then Exit Sub
    Dim v As Variant
    '
    For Each v In m_conversionDescriptors
        iconv_close v
    Next v
    Set m_conversionDescriptors = Nothing
End Sub
Private Function PageCodeToText(ByVal pc As PageCode) As String
    Dim result As String
    Select Case pc
        Case codeUTF8:    result = "UTF-8"
        Case codeUTF16LE: result = "UTF-16LE"
        Case codeUTF16BE: result = "UTF-16BE"
        Case codeUTF32LE: result = "UTF-32LE"
        Case codeUTF32BE: result = "UTF-32BE"
    End Select
    PageCodeToText = StrConv(result, vbFromUnicode)
End Function
#Else
Private Function EncodeToUTF16LE(ByRef textToConvert As String _
                               , ByVal fromCode As PageCode) As String
    Dim charCount As Long
    charCount = MultiByteToWideChar(fromCode, 0, StrPtr(textToConvert) _
                                  , LenB(textToConvert), 0, 0)
    If charCount = 0 Then Exit Function

    EncodeToUTF16LE = Space$(charCount)
    MultiByteToWideChar fromCode, 0, StrPtr(textToConvert) _
                      , LenB(textToConvert), StrPtr(EncodeToUTF16LE), charCount
End Function
Private Function EncodeFromUTF16LE(ByRef textToConvert As String _
                                 , ByVal toCode As PageCode) As String
    Dim byteCount As Long
    byteCount = WideCharToMultiByte(toCode, 0, StrPtr(textToConvert) _
                                  , Len(textToConvert), 0, 0, 0, 0)
    If byteCount = 0 Then Exit Function
    '
    EncodeFromUTF16LE = Space$((byteCount + 1) \ 2)
    If byteCount Mod 2 = 1 Then
        EncodeFromUTF16LE = LeftB$(EncodeFromUTF16LE, byteCount)
    End If
    WideCharToMultiByte toCode, 0, StrPtr(textToConvert), Len(textToConvert) _
                      , StrPtr(EncodeFromUTF16LE), byteCount, 0, 0
End Function
#End If

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
Public Function CopyFile(ByRef sourcePath As String _
                       , ByRef destinationPath As String _
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
        FileCopy sourcePath, destinationPath 'Copies opened files as well
        CopyFile = (Err.Number = 0)
        On Error GoTo 0
    #Else
        If Not failIfExists Then
            On Error Resume Next
            SetAttr destinationPath, vbNormal 'Costly to do after Copy fails
            On Error GoTo 0
        End If
        CopyFile = CopyFileW(StrPtr(sourcePath), StrPtr(destinationPath), failIfExists)
    #End If
End Function

'*******************************************************************************
'Copies a folder. Ability to copy all subfolders
'If 'failIfExists' is set to True then this method will fail if any file or
'   subFolder already exists (including the main 'destinationPath')
'If 'ignoreFailedChildren' is set to True then the method continues to copy the
'   remaining files and subfolders. This is useful when reverting a 'MoveFolder'
'   call across different disk drives. Use this parameter with care
'*******************************************************************************
Public Function CopyFolder(ByRef sourcePath As String _
                         , ByRef destinationPath As String _
                         , Optional ByVal includeSubFolders As Boolean = True _
                         , Optional ByVal failIfExists As Boolean = False _
                         , Optional ByVal ignoreFailedChildren As Boolean = False) As Boolean
    If Not IsFolder(sourcePath) Then Exit Function
    If Not CreateFolder(destinationPath, failIfExists) Then Exit Function
    '
    Dim fixedSrc As String: fixedSrc = BuildPath(sourcePath, vbNullString)
    Dim fixedDst As String: fixedDst = BuildPath(destinationPath, vbNullString)
    '
    If includeSubFolders Then
        Dim subFolderPath As Variant
        Dim newFolderPath As String
        '
        For Each subFolderPath In GetFolders(fixedSrc, True, True, True)
            newFolderPath = Replace(subFolderPath, fixedSrc, fixedDst)
            If Not CreateFolder(newFolderPath, failIfExists) Then
                If Not ignoreFailedChildren Then Exit Function
            End If
        Next subFolderPath
    End If
    '
    Dim filePath As Variant
    Dim newFilePath As String
    '
    For Each filePath In GetFiles(fixedSrc, includeSubFolders, True, True)
        newFilePath = Replace(filePath, fixedSrc, fixedDst)
        If Not CopyFile(CStr(filePath), newFilePath, failIfExists) Then
            If Not ignoreFailedChildren Then Exit Function
        End If
    Next filePath
    '
    CopyFolder = True
End Function

'*******************************************************************************
'Creates a folder including parent folders if needed
'*******************************************************************************
Public Function CreateFolder(ByRef folderPath As String _
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
    Dim v As Variant
    '
    'Note that the same outcome could be achieved via recursivity but this
    '   approach avoids adding extra stack frames to the call stack
    collFoldersToCreate.Add fullPath
    Do
        sepIndex = InStrRev(fullPath, PATH_SEPARATOR)
        If sepIndex < 3 Then Exit Do
        '
        fullPath = Left$(fullPath, sepIndex - 1)
        If IsFolder(fullPath) Then Exit Do
        collFoldersToCreate.Add fullPath, Before:=1
    Loop
    On Error Resume Next
    For Each v In collFoldersToCreate
        'MkDir does not support all Unicode characters, unlike FSO
        #If Mac Then
            MkDir v
        #Else
            GetFSO.CreateFolder v
        #End If
        If Err.Number <> 0 Then Exit For
    Next v
    CreateFolder = (Err.Number = 0)
    On Error GoTo 0
End Function

'*******************************************************************************
'Deletes a file only. Does not support wildcards * ?
'*******************************************************************************
Public Function DeleteFile(ByRef filePath As String) As Boolean
    If LenB(filePath) = 0 Then Exit Function
    If Not IsFile(filePath) Then Exit Function 'Avoid 'Kill' on folder
    '
    On Error Resume Next
    #If Windows Then
        GetFSO.DeleteFile filePath, True
        DeleteFile = (Err.Number = 0)
        If DeleteFile Then Exit Function
        Err.Clear
    #End If
    SetAttr filePath, vbNormal 'Too costly to do after failing Kill
    Err.Clear
    Kill filePath
    DeleteFile = (Err.Number = 0)
    On Error GoTo 0
    '
    #If Windows Then
        If Not DeleteFile Then DeleteFile = CBool(DeleteFileW(StrPtr(filePath)))
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
Public Function DeleteFolder(ByRef folderPath As String _
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
    DeleteFolder = (Err.Number = 0)
    If DeleteFolder Then Exit Function
    '
    #If Windows Then
        Err.Clear
        GetFSO.DeleteFolder folderPath, True
        DeleteFolder = (Err.Number = 0)
        If DeleteFolder Then Exit Function
    #End If
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
Private Function DeleteBottomMostFolder(ByRef folderPath As String) As Boolean
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
        If Not DeleteFile(CStr(filePath)) Then Exit Function
    Next filePath
    '
    On Error Resume Next
    RmDir fixedPath
    DeleteBottomMostFolder = (Err.Number = 0)
    On Error GoTo 0
    '
    #If Windows Then
        If Not DeleteBottomMostFolder Then
            DeleteBottomMostFolder = CBool(RemoveDirectoryW(StrPtr(fixedPath)))
        End If
    #End If
End Function

'*******************************************************************************
'Fixes a file or folder name, NOT a path
'Before creating a file/folder it's useful to fix the name so that the creation
'   does not fail because of forbidden characters, reserved names or other rules
'*******************************************************************************
#If Mac Then
Public Function FixFileName(ByRef nameToFix As String) As String
    Dim resultName As String
    Dim i As Long: i = 1
    '
    resultName = Replace(nameToFix, ":", vbNullString)
    resultName = Replace(resultName, "/", vbNullString)
    '
    'Names cannot start with a space character
    Do While Mid$(resultName, i, 1) = "."
        i = i + 1
    Loop
    If i > 1 Then resultName = Mid$(resultName, i)
    '
    FixFileName = resultName
End Function
#Else
Public Function FixFileName(ByRef nameToFix As String _
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
    Dim i As Long:       i = nameLen
    '
    If i = 0 Then Exit Function
    Do While InStr(1, dotSpace, Mid$(resultName, i, 1)) > 0
        i = i - 1
        If i = 0 Then Exit Function
    Loop
    If i < nameLen Then resultName = Left$(resultName, i)
    If IsReservedName(resultName) Then Exit Function
    '
    FixFileName = resultName
End Function
#End If

'*******************************************************************************
'Returns a collection of forbidden characters for a file/folder name
'Ability to add the caret ^ char - forbidden on FAT file systems but not on NTFS
'*******************************************************************************
#If Windows Then
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
            collForbiddenChars.Add Chr$(i)
        Next i
    End If
    If hasCaret And Not addCaret Then
        collForbiddenChars.Remove 1
    ElseIf Not hasCaret And addCaret Then
        collForbiddenChars.Add Item:="^", Before:=1
    End If
    hasCaret = addCaret
    '
    Set ForbiddenNameChars = collForbiddenChars
End Function
#End If

'*******************************************************************************
'Windows file/folder reserved names: com1 to com9, lpt1 to lpt9, con, nul, prn
'*******************************************************************************
#If Windows Then
Private Function IsReservedName(ByRef nameToCheck As String) As Boolean
    Static collReservedNames As Collection
    Dim v As Variant
    '
    If collReservedNames Is Nothing Then
        Set collReservedNames = New Collection
        For Each v In Split("com1,com2,com3,com4,com5,com6,com7,com8,com9," _
                          & "lpt1,lpt2,lpt3,lpt4,lpt5,lpt6,lpt7,lpt8,lpt9," _
                          & "con,nul,prn,aux", ",")
            collReservedNames.Add Empty, v
        Next v
    End If
    On Error Resume Next
    collReservedNames.Item nameToCheck
    IsReservedName = (Err.Number = 0)
    On Error GoTo 0
End Function
#End If

'*******************************************************************************
'Fixes path separators for a path
'Windows example: replace \\, \\\, \\\\, \\//, \/\/\, /, // etc. with a single \
'Note that on a Mac, the network paths (smb:// or afp://) need to be mounted and
'   are only valid via the mounted volumes: /volumes/VolumeName/... unlike on a
'   PC where \\share\data\... is a valid file/folder path (UNC)
'Trims any current paths \. as well as any parent folder pairs \{parentName}\..
'*******************************************************************************
Public Function FixPathSeparators(ByRef pathToFix As String) As String
    Const ps As String = PATH_SEPARATOR
    Dim resultPath As String
    Dim isUNC As Boolean
    '
    If LenB(pathToFix) = 0 Then Exit Function
    #If Mac Then
        resultPath = Replace(pathToFix, "\", ps)
    #Else
        resultPath = Replace(pathToFix, "/", ps)
        If Left$(resultPath, 4) = "\\?\" Then
            If Mid$(resultPath, 5, 4) = "UNC\" Then
                Mid$(resultPath, 7, 1) = "\"
                resultPath = Mid$(resultPath, 7)
            Else
                resultPath = Mid$(resultPath, 5)
            End If
        End If
        isUNC = (Left$(resultPath, 2) = "\\")
    #End If
    '
    Const fCurrent As String = ps & "." & ps
    Const fParent As String = ps & ".." & ps
    Dim sepIndex As Long
    Dim i As Long: i = 0
    '
    'Remove any current folder references
    Do
        i = InStr(i + 1, resultPath, fCurrent)
        If i = 0 Then i = InStr(Len(resultPath) - 1, resultPath, ps & ".")
        If i > 0 Then Mid$(resultPath, i + 1, 1) = ps
    Loop Until i = 0
    '
    FixPathSeparators = RemoveDuplicatePS(resultPath, isUNC)
    '
    'Remove any parent folder references
    i = 1
    Do
        i = InStr(i, FixPathSeparators, fParent)
        If i = 0 And Len(FixPathSeparators) > 2 Then
            i = InStr(Len(FixPathSeparators) - 2, FixPathSeparators, ps & "..")
        End If
        If i > 1 Then
            sepIndex = InStrRev(FixPathSeparators, ps, i - 1)
            If sepIndex < 3 Then sepIndex = i
            FixPathSeparators = Left$(FixPathSeparators, sepIndex) _
                              & Mid$(FixPathSeparators, i + 4)
            If sepIndex < i Then i = i - sepIndex
        End If
    Loop Until i = 0
End Function

'*******************************************************************************
'Utility for 'FixPathSeparators'. Removes any duplicate path separators
'*******************************************************************************
Private Function RemoveDuplicatePS(ByRef pathToFix As String _
                                 , ByVal isUNC As Boolean) As String
    Const ps As String = PATH_SEPARATOR
    Dim startPos As Long
    Dim currPos As Long
    Dim prevPos As Long
    Dim diff As Long
    Dim i As Long
    '
    If isUNC Then currPos = 2 'Skip the leading UNC prefix: \\
    RemoveDuplicatePS = pathToFix
    Do
        prevPos = currPos
        currPos = InStr(currPos + 1, pathToFix, ps)
        If startPos = 0 Then startPos = prevPos + 1
        If currPos - prevPos <= 1 Then
            diff = currPos - startPos
            If currPos = 0 Then diff = diff + Len(pathToFix) + 1
            If startPos * Sgn(i * diff) > 1 Then
                Mid$(RemoveDuplicatePS, i) = Mid$(pathToFix, startPos, diff)
                i = i + diff
            End If
            If i = 0 Then i = (startPos + diff) * Sgn(prevPos)
            startPos = 0
        End If
    Loop Until currPos = 0
    If i > 1 Then RemoveDuplicatePS = Left$(RemoveDuplicatePS, i - 1)
End Function

'*******************************************************************************
'Retrieves the owner name for a file path
'*******************************************************************************
#If Windows Then
Public Function GetFileOwner(ByRef filePath As String) As String
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
    If GetSecurityDescriptorOwner(sd(0), pOwner, 0&) = 0 Then Exit Function
    '
    'Get name and domain length
    Dim nameLen As Long, domainLen As Long
    LookupAccountSid vbNullString, pOwner, vbNullString _
                   , nameLen, vbNullString, domainLen, 0&
    If nameLen = 0 Then Exit Function
    '
    'Get name and domain
    Dim owName As String: owName = Space$(nameLen - 1) '-1 less Null Char
    Dim owDomain As String: owDomain = Space$(domainLen - 1)
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
Public Function GetFiles(ByRef folderPath As String _
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
        Dim subFolderPath As Variant
        For Each subFolderPath In GetFolders(folderPath, True, True, True)
            AddFilesTo collFiles, CStr(subFolderPath), fAttribute
        Next subFolderPath
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
                     , ByRef folderPath As String _
                     , ByVal fAttribute As VbFileAttribute)
    #If Mac Then
        Const maxDirLen As Long = 247 'To be updated
    #Else
        Const maxDirLen As Long = 247
    #End If
    Const errBadFileNameOrNumber As Long = 52
    Dim fileName As String
    Dim fullPath As String
    Dim collTemp As New Collection
    Dim dirFailed As Boolean
    Dim v As Variant
    Dim fixedPath As String: fixedPath = BuildPath(folderPath, vbNullString)
    '
    On Error Resume Next
    fileName = Dir(fixedPath, fAttribute)
    dirFailed = (Err.Number = errBadFileNameOrNumber) 'Unsupported Unicode
    On Error GoTo 0
    '
    Do While LenB(fileName) > 0
        collTemp.Add fileName
        If InStr(1, fileName, "?") > 0 Then 'Unsupported Unicode
            Set collTemp = New Collection
            dirFailed = True
            Exit Do
        End If
        fileName = Dir
    Loop
    If dirFailed Or Len(fixedPath) > maxDirLen Then
        #If Mac Then

        #Else
            Dim fsoFile As Object
            Dim fsoFolder As Object: Set fsoFolder = GetFSOFolder(fixedPath)
            '
            If Not fsoFolder Is Nothing Then
                With fsoFolder
                    For Each fsoFile In .Files
                        collTemp.Add fsoFile.Name
                    Next fsoFile
                End With
            End If
        #End If
    End If
    For Each v In collTemp
        collTarget.Add fixedPath & v
    Next v
End Sub

'*******************************************************************************
'For long paths FSO fails in either retrieving the folder or it retrieves the
'   folder but the SubFolders or Files collections are not correctly populated
'*******************************************************************************
#If Windows Then
Private Function GetFSOFolder(ByRef folderPath As String) As Object
    If Not IsFolder(folderPath) Then Exit Function
    '
    Dim fso As Object: Set fso = GetFSO()
    Dim fsoFolder As Object
    Dim tempFolder As Object
    '
    On Error Resume Next
    Set fsoFolder = fso.GetFolder(folderPath)
    If Err.Number <> 0 Then
        Const ps As String = PATH_SEPARATOR
        Dim collNames As New Collection
        Dim i As Long
        Dim parentPath As String: parentPath = folderPath
        Dim folderName As String
        '
        If Right$(parentPath, 1) = ps Then
            parentPath = Left$(parentPath, Len(parentPath) - 1)
        End If
        Do
            i = InStrRev(parentPath, ps)
            folderName = Mid$(parentPath, i + 1)
            parentPath = Left$(parentPath, i - 1)
            '
            If collNames.Count = 0 Then
                collNames.Add folderName
            Else
                collNames.Add folderName, Before:=1
            End If
            Err.Clear
            Set fsoFolder = fso.GetFolder(parentPath)
        Loop Until Err.Number = 0
        Do
            Set fsoFolder = fso.GetFolder(fsoFolder.ShortPath) 'Fix .SubFolders
            Set fsoFolder = fsoFolder.SubFolders(collNames(1))
            collNames.Remove 1
        Loop Until collNames.Count = 0
    End If
    On Error GoTo 0
    Set GetFSOFolder = fso.GetFolder(fsoFolder.ShortPath) 'Fix .Files Bug
End Function
#End If

'*******************************************************************************
'Returns the FOLDERID of a 'known folder' on Windows
'Returns a null string if 'kfID' is not a valid enum value
'Source: KnownFolders.h (Windows 11 SDK 10.0.22621.0) (sorted alphabetically)
'Note: Most of the FOLDERIDs that are available on a specific device seem to
'      be registered in the windows registry under
'      HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions
'      However, it seems that sometimes the SHGetKnownFolderPath function can
'      process a FOLDERID even if not present in said registry location.
'*******************************************************************************
#If Windows Then
Public Function GetKnownFolderCLSID(ByVal kfID As KnownFolderID) As String
    Static cids([_minKfID] To [_maxKfID]) As String
    '
    If kfID < [_minKfID] Or kfID > [_maxKfID] Then Exit Function
    If LenB(cids([_minKfID])) = 0 Then
        cids(kfID_AccountPictures) = "{008ca0b1-55b4-4c56-b8a8-4de4b299d3be}"
        cids(kfID_AddNewPrograms) = "{de61d971-5ebc-4f02-a3a9-6c82895e5c04}"
        cids(kfID_AdminTools) = "{724EF170-A42D-4FEF-9F26-B60E846FBA4F}"
        cids(kfID_AllAppMods) = "{7ad67899-66af-43ba-9156-6aad42e6c596}"
        cids(kfID_AppCaptures) = "{EDC0FE71-98D8-4F4A-B920-C8DC133CB165}"
        cids(kfID_AppDataDesktop) = "{B2C5E279-7ADD-439F-B28C-C41FE1BBF672}"
        cids(kfID_AppDataDocuments) = "{7BE16610-1F7F-44AC-BFF0-83E15F2FFCA1}"
        cids(kfID_AppDataFavorites) = "{7CFBEFBC-DE1F-45AA-B843-A542AC536CC9}"
        cids(kfID_AppDataProgramData) = "{559D40A3-A036-40FA-AF61-84CB430A4D34}"
        cids(kfID_ApplicationShortcuts) = "{A3918781-E5F2-4890-B3D9-A7E54332328C}"
        cids(kfID_AppsFolder) = "{1e87508d-89c2-42f0-8a7e-645a0f50ca58}"
        cids(kfID_AppUpdates) = "{a305ce99-f527-492b-8b1a-7e76fa98d6e4}"
        cids(kfID_CameraRoll) = "{AB5FB87B-7CE2-4F83-915D-550846C9537B}"
        cids(kfID_CameraRollLibrary) = "{2B20DF75-1EDA-4039-8097-38798227D5B7}"
        cids(kfID_CDBurning) = "{9E52AB10-F80D-49DF-ACB8-4330F5687855}"
        cids(kfID_ChangeRemovePrograms) = "{df7266ac-9274-4867-8d55-3bd661de872d}"
        cids(kfID_CommonAdminTools) = "{D0384E7D-BAC3-4797-8F14-CBA229B392B5}"
        cids(kfID_CommonOEMLinks) = "{C1BAE2D0-10DF-4334-BEDD-7AA20B227A9D}"
        cids(kfID_CommonPrograms) = "{0139D44E-6AFE-49F2-8690-3DAFCAE6FFB8}"
        cids(kfID_CommonStartMenu) = "{A4115719-D62E-491D-AA7C-E74B8BE3B067}"
        cids(kfID_CommonStartMenuPlaces) = "{A440879F-87A0-4F7D-B700-0207B966194A}"
        cids(kfID_CommonStartup) = "{82A5EA35-D9CD-47C5-9629-E15D2F714E6E}"
        cids(kfID_CommonTemplates) = "{B94237E7-57AC-4347-9151-B08C6C32D1F7}"
        cids(kfID_ComputerFolder) = "{0AC0837C-BBF8-452A-850D-79D08E667CA7}"
        cids(kfID_ConflictFolder) = "{4bfefb45-347d-4006-a5be-ac0cb0567192}"
        cids(kfID_ConnectionsFolder) = "{6F0CD92B-2E97-45D1-88FF-B0D186B8DEDD}"
        cids(kfID_Contacts) = "{56784854-C6CB-462b-8169-88E350ACB882}"
        cids(kfID_ControlPanelFolder) = "{82A74AEB-AEB4-465C-A014-D097EE346D63}"
        cids(kfID_Cookies) = "{2B0F765D-C0E9-4171-908E-08A611B84FF6}"
        cids(kfID_CurrentAppMods) = "{3db40b20-2a30-4dbe-917e-771dd21dd099}"
        cids(kfID_Desktop) = "{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}"
        cids(kfID_DevelopmentFiles) = "{DBE8E08E-3053-4BBC-B183-2A7B2B191E59}"
        cids(kfID_Device) = "{1C2AC1DC-4358-4B6C-9733-AF21156576F0}"
        cids(kfID_DeviceMetadataStore) = "{5CE4A5E9-E4EB-479D-B89F-130C02886155}"
        cids(kfID_Documents) = "{FDD39AD0-238F-46AF-ADB4-6C85480369C7}"
        cids(kfID_DocumentsLibrary) = "{7b0db17d-9cd2-4a93-9733-46cc89022e7c}"
        cids(kfID_Downloads) = "{374DE290-123F-4565-9164-39C4925E467B}"
        cids(kfID_Favorites) = "{1777F761-68AD-4D8A-87BD-30B759FA33DD}"
        cids(kfID_Fonts) = "{FD228CB7-AE11-4AE3-864C-16F3910AB8FE}"
        cids(kfID_Games) = "{CAC52C1A-B53D-4edc-92D7-6B2E8AC19434}"
        cids(kfID_GameTasks) = "{054FAE61-4DD8-4787-80B6-090220C4B700}"
        cids(kfID_History) = "{D9DC8A3B-B784-432E-A781-5A1130A75963}"
        cids(kfID_HomeGroup) = "{52528A6B-B9E3-4add-B60D-588C2DBA842D}"
        cids(kfID_HomeGroupCurrentUser) = "{9B74B6A3-0DFD-4f11-9E78-5F7800F2E772}"
        cids(kfID_ImplicitAppShortcuts) = "{bcb5256f-79f6-4cee-b725-dc34e402fd46}"
        cids(kfID_InternetCache) = "{352481E8-33BE-4251-BA85-6007CAEDCF9D}"
        cids(kfID_InternetFolder) = "{4D9F7874-4E0C-4904-967B-40B0D20C3E4B}"
        cids(kfID_Libraries) = "{1B3EA5DC-B587-4786-B4EF-BD1DC332AEAE}"
        cids(kfID_Links) = "{bfb9d5e0-c6a9-404c-b2b2-ae6db6af4968}"
        cids(kfID_LocalAppData) = "{F1B32785-6FBA-4FCF-9D55-7B8E7F157091}"
        cids(kfID_LocalAppDataLow) = "{A520A1A4-1780-4FF6-BD18-167343C5AF16}"
        cids(kfID_LocalDocuments) = "{f42ee2d3-909f-4907-8871-4c22fc0bf756}"
        cids(kfID_LocalDownloads) = "{7d83ee9b-2244-4e70-b1f5-5393042af1e4}"
        cids(kfID_LocalizedResourcesDir) = "{2A00375E-224C-49DE-B8D1-440DF7EF3DDC}"
        cids(kfID_LocalMusic) = "{a0c69a99-21c8-4671-8703-7934162fcf1d}"
        cids(kfID_LocalPictures) = "{0ddd015d-b06c-45d5-8c4c-f59713854639}"
        cids(kfID_LocalStorage) = "{B3EB08D3-A1F3-496B-865A-42B536CDA0EC}"
        cids(kfID_LocalVideos) = "{35286a68-3c57-41a1-bbb1-0eae73d76c95}"
        cids(kfID_Music) = "{4BD8D571-6D19-48D3-BE97-422220080E43}"
        cids(kfID_MusicLibrary) = "{2112AB0A-C86A-4ffe-A368-0DE96E47012E}"
        cids(kfID_NetHood) = "{C5ABBF53-E17F-4121-8900-86626FC2C973}"
        cids(kfID_NetworkFolder) = "{D20BEEC4-5CA8-4905-AE3B-BF251EA09B53}"
        cids(kfID_Objects3D) = "{31C0DD25-9439-4F12-BF41-7FF4EDA38722}"
        cids(kfID_OneDrive) = "{A52BBA46-E9E1-435f-B3D9-28DAA648C0F6}"
        cids(kfID_OriginalImages) = "{2C36C0AA-5812-4b87-BFD0-4CD0DFB19B39}"
        cids(kfID_PhotoAlbums) = "{69D2CF90-FC33-4FB7-9A0C-EBB0F0FCB43C}"
        cids(kfID_Pictures) = "{33E28130-4E1E-4676-835A-98395C3BC3BB}"
        cids(kfID_PicturesLibrary) = "{A990AE9F-A03B-4e80-94BC-9912D7504104}"
        cids(kfID_Playlists) = "{DE92C1C7-837F-4F69-A3BB-86E631204A23}"
        cids(kfID_PrintersFolder) = "{76FC4E2D-D6AD-4519-A663-37BD56068185}"
        cids(kfID_PrintHood) = "{9274BD8D-CFD1-41C3-B35E-B13F55A758F4}"
        cids(kfID_Profile) = "{5E6C858F-0E22-4760-9AFE-EA3317B67173}"
        cids(kfID_ProgramData) = "{62AB5D82-FDC1-4DC3-A9DD-070D1D495D97}"
        cids(kfID_ProgramFiles) = "{905e63b6-c1bf-494e-b29c-65b732d3d21a}"
        cids(kfID_ProgramFilesCommon) = "{F7F1ED05-9F6D-47A2-AAAE-29D317C6F066}"
        cids(kfID_ProgramFilesCommonX64) = "{6365D5A7-0F0D-45e5-87F6-0DA56B6A4F7D}"
        cids(kfID_ProgramFilesCommonX86) = "{DE974D24-D9C6-4D3E-BF91-F4455120B917}"
        cids(kfID_ProgramFilesX64) = "{6D809377-6AF0-444b-8957-A3773F02200E}"
        cids(kfID_ProgramFilesX86) = "{7C5A40EF-A0FB-4BFC-874A-C0F2E0B9FA8E}"
        cids(kfID_Programs) = "{A77F5D77-2E2B-44C3-A6A2-ABA601054A51}"
        cids(kfID_Public) = "{DFDF76A2-C82A-4D63-906A-5644AC457385}"
        cids(kfID_PublicDesktop) = "{C4AA340D-F20F-4863-AFEF-F87EF2E6BA25}"
        cids(kfID_PublicDocuments) = "{ED4824AF-DCE4-45A8-81E2-FC7965083634}"
        cids(kfID_PublicDownloads) = "{3D644C9B-1FB8-4f30-9B45-F670235F79C0}"
        cids(kfID_PublicGameTasks) = "{DEBF2536-E1A8-4c59-B6A2-414586476AEA}"
        cids(kfID_PublicLibraries) = "{48daf80b-e6cf-4f4e-b800-0e69d84ee384}"
        cids(kfID_PublicMusic) = "{3214FAB5-9757-4298-BB61-92A9DEAA44FF}"
        cids(kfID_PublicPictures) = "{B6EBFB86-6907-413C-9AF7-4FC2ABF07CC5}"
        cids(kfID_PublicRingtones) = "{E555AB60-153B-4D17-9F04-A5FE99FC15EC}"
        cids(kfID_PublicUserTiles) = "{0482af6c-08f1-4c34-8c90-e17ec98b1e17}"
        cids(kfID_PublicVideos) = "{2400183A-6185-49FB-A2D8-4A392A602BA3}"
        cids(kfID_QuickLaunch) = "{52a4f021-7b75-48a9-9f6b-4b87a210bc8f}"
        cids(kfID_Recent) = "{AE50C081-EBD2-438A-8655-8A092E34987A}"
        cids(kfID_RecordedCalls) = "{2f8b40c2-83ed-48ee-b383-a1f157ec6f9a}"
        cids(kfID_RecordedTVLibrary) = "{1A6FDBA2-F42D-4358-A798-B74D745926C5}"
        cids(kfID_RecycleBinFolder) = "{B7534046-3ECB-4C18-BE4E-64CD4CB7D6AC}"
        cids(kfID_ResourceDir) = "{8AD10C31-2ADB-4296-A8F7-E4701232C972}"
        cids(kfID_RetailDemo) = "{12D4C69E-24AD-4923-BE19-31321C43A767}"
        cids(kfID_Ringtones) = "{C870044B-F49E-4126-A9C3-B52A1FF411E8}"
        cids(kfID_RoamedTileImages) = "{AAA8D5A5-F1D6-4259-BAA8-78E7EF60835E}"
        cids(kfID_RoamingAppData) = "{3EB685DB-65F9-4CF6-A03A-E3EF65729F3D}"
        cids(kfID_RoamingTiles) = "{00BCFC5A-ED94-4e48-96A1-3F6217F21990}"
        cids(kfID_SampleMusic) = "{B250C668-F57D-4EE1-A63C-290EE7D1AA1F}"
        cids(kfID_SamplePictures) = "{C4900540-2379-4C75-844B-64E6FAF8716B}"
        cids(kfID_SamplePlaylists) = "{15CA69B3-30EE-49C1-ACE1-6B5EC372AFB5}"
        cids(kfID_SampleVideos) = "{859EAD94-2E85-48AD-A71A-0969CB56A6CD}"
        cids(kfID_SavedGames) = "{4C5C32FF-BB9D-43b0-B5B4-2D72E54EAAA4}"
        cids(kfID_SavedPictures) = "{3B193882-D3AD-4eab-965A-69829D1FB59F}"
        cids(kfID_SavedPicturesLibrary) = "{E25B5812-BE88-4bd9-94B0-29233477B6C3}"
        cids(kfID_SavedSearches) = "{7d1d3a04-debb-4115-95cf-2f29da2920da}"
        cids(kfID_Screenshots) = "{b7bede81-df94-4682-a7d8-57a52620b86f}"
        cids(kfID_SEARCH_CSC) = "{ee32e446-31ca-4aba-814f-a5ebd2fd6d5e}"
        cids(kfID_SEARCH_MAPI) = "{98ec0e18-2098-4d44-8644-66979315a281}"
        cids(kfID_SearchHistory) = "{0D4C3DB6-03A3-462F-A0E6-08924C41B5D4}"
        cids(kfID_SearchHome) = "{190337d1-b8ca-4121-a639-6d472d16972a}"
        cids(kfID_SearchTemplates) = "{7E636BFE-DFA9-4D5E-B456-D7B39851D8A9}"
        cids(kfID_SendTo) = "{8983036C-27C0-404B-8F08-102D10DCFD74}"
        cids(kfID_SidebarDefaultParts) = "{7B396E54-9EC5-4300-BE0A-2482EBAE1A26}"
        cids(kfID_SidebarParts) = "{A75D362E-50FC-4fb7-AC2C-A8BEAA314493}"
        cids(kfID_SkyDrive) = "{A52BBA46-E9E1-435f-B3D9-28DAA648C0F6}"
        cids(kfID_SkyDriveCameraRoll) = "{767E6811-49CB-4273-87C2-20F355E1085B}"
        cids(kfID_SkyDriveDocuments) = "{24D89E24-2F19-4534-9DDE-6A6671FBB8FE}"
        cids(kfID_SkyDriveMusic) = "{C3F2459E-80D6-45DC-BFEF-1F769F2BE730}"
        cids(kfID_SkyDrivePictures) = "{339719B5-8C47-4894-94C2-D8F77ADD44A6}"
        cids(kfID_StartMenu) = "{625B53C3-AB48-4EC1-BA1F-A1EF4146FC19}"
        cids(kfID_StartMenuAllPrograms) = "{F26305EF-6948-40B9-B255-81453D09C785}"
        cids(kfID_Startup) = "{B97D20BB-F46A-4C97-BA10-5E3608430854}"
        cids(kfID_SyncManagerFolder) = "{43668BF8-C14E-49B2-97C9-747784D784B7}"
        cids(kfID_SyncResultsFolder) = "{289a9a43-be44-4057-a41b-587a76d7e7f9}"
        cids(kfID_SyncSetupFolder) = "{0F214138-B1D3-4a90-BBA9-27CBC0C5389A}"
        cids(kfID_System) = "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}"
        cids(kfID_SystemX86) = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}"
        cids(kfID_Templates) = "{A63293E8-664E-48DB-A079-DF759E0509F7}"
        cids(kfID_UserPinned) = "{9e3995ab-1f9c-4f13-b827-48b24b6c7174}"
        cids(kfID_UserProfiles) = "{0762D272-C50A-4BB0-A382-697DCD729B80}"
        cids(kfID_UserProgramFiles) = "{5cd7aee2-2219-4a67-b85d-6c9ce15660cb}"
        cids(kfID_UserProgramFilesCommon) = "{bcbd3057-ca5c-4622-b42d-bc56db0ae516}"
        cids(kfID_UsersFiles) = "{f3ce0f7c-4901-4acc-8648-d5d44b04ef8f}"
        cids(kfID_UsersLibraries) = "{A302545D-DEFF-464b-ABE8-61C8648D939B}"
        cids(kfID_Videos) = "{18989B1D-99B5-455B-841C-AB7C74E4DDFC}"
        cids(kfID_VideosLibrary) = "{491E922F-5643-4af4-A7EB-4E7A138D8174}"
        cids(kfID_Windows) = "{F38BF404-1D43-42F2-9305-67DE0B28FC23}"
    End If
    GetKnownFolderCLSID = cids(kfID)
End Function
#End If

'*******************************************************************************
'Returns the path of a 'known folder' on Windows
'If 'createIfMissing' is set to True, the windows API function will be called
'   with flags 'KF_FLAG_CREATE' and 'KF_FLAG_INIT' and will create the folder
'   if it does not currently exist on the system.
'The function can raise the following errors:
'   -   5: (Invalid procedure call) if 'kfID' is not valid
'   -  76: (Path not found) if 'createIfMissing' = False AND path not found
'   -  75: (Path/File access error) if path not found because either:
'          * the specified folder ID is for a known virtual folder
'          * there are insufficient permissions to create the folder
'   - 336: (Component not correctly registered) if the path, or the known
'          folder ID itself are not registered in the windows registry
'   -  51: (Internal error) if an unexpected error occurs
'*******************************************************************************
#If Windows Then
Public Function GetKnownFolderPath(ByVal kfID As KnownFolderID _
                                 , Optional ByVal createIfMissing As Boolean = False) As String
    Const methodName As String = "GetKnownFolderPath"
    Const NOERROR As Long = 0
    Static guids([_minKfID] To [_maxKfID]) As GUID
    '
    If kfID < [_minKfID] Or kfID > [_maxKfID] Then
        Err.Raise vbErrInvalidProcedureCall, methodName, "Invalid Folder ID"
    ElseIf guids(kfID).data1 = 0 Then
        If CLSIDFromString(StrPtr(GetKnownFolderCLSID(kfID)), guids(kfID)) <> NOERROR Then
            Err.Raise vbErrInvalidProcedureCall, methodName, "Invalid CLSID"
        End If
    End If
    '
    Const KF_FLAG_CREATE As Long = &H8000&  'Other flags not relevant
    Const KF_FLAG_INIT   As Long = &H800&
    Const flagCreateInit As Long = KF_FLAG_CREATE Or KF_FLAG_INIT
    Dim dwFlags As Long: If createIfMissing Then dwFlags = flagCreateInit
    '
    Const S_OK As Long = 0
    Dim ppszPath As LongPtr
    Dim hRes As Long: hRes = SHGetKnownFolderPath(guids(kfID), dwFlags, 0, ppszPath)
    '
    If hRes = S_OK Then
        GetKnownFolderPath = Space$(lstrlenW(ppszPath))
        CopyMemory StrPtr(GetKnownFolderPath), ppszPath, LenB(GetKnownFolderPath)
    End If
    CoTaskMemFree ppszPath 'Memory must be freed, even on fail
    If hRes = S_OK Then Exit Function
    '
    Const E_FAIL                        As Long = &H80004005
    Const E_INVALIDARG                  As Long = &H80070057
    Const HRESULT_ERROR_FILE_NOT_FOUND  As Long = &H80070002
    Const HRESULT_ERROR_PATH_NOT_FOUND  As Long = &H80070003
    Const HRESULT_ERROR_ACCESS_DENIED   As Long = &H80070005
    Const HRESULT_ERROR_NOT_FOUND       As Long = &H80070490
    '
    Select Case hRes
    Case E_FAIL
        Err.Raise vbErrPathFileAccessError, methodName, "Known folder might " _
                & "be marked 'KF_CATEGORY_VIRTUAL' which does not have a path"
    Case E_INVALIDARG
        Err.Raise vbErrInvalidProcedureCall, methodName _
                , "Known folder not present on system"
    Case HRESULT_ERROR_FILE_NOT_FOUND
        Err.Raise vbErrPathNotFound, methodName, "KnownFolderID might not exist"
    Case HRESULT_ERROR_PATH_NOT_FOUND, HRESULT_ERROR_NOT_FOUND
        Err.Raise vbErrComponentNotRegistered, methodName, "KnownFolderID " _
                & "might be registered, but no path registered for it"
    Case HRESULT_ERROR_ACCESS_DENIED
        Err.Raise vbErrPathFileAccessError, methodName, "Access denied"
    Case Else
        Err.Raise vbErrInternalError, methodName, "Unexpected error code"
    End Select
End Function
#End If

'*******************************************************************************
'Returns a Collection of all the subfolders (paths) in a specified folder
'Warning! On Mac the 'Dir' method only accepts the vbHidden and the vbDirectory
'   attributes. However the vbHidden attribute does not work - no hidden files
'   or folders are retrieved regardless if vbHidden is used or not
'On Windows, the vbHidden, and vbSystem attributes work fine with 'Dir'
'*******************************************************************************
Public Function GetFolders(ByRef folderPath As String _
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
                       , ByRef folderPath As String _
                       , ByVal includeSubFolders As Boolean _
                       , ByVal fAttribute As VbFileAttribute)
    #If Mac Then
        Const maxDirLen As Long = 247 'To be updated
    #Else
        Const maxDirLen As Long = 247
    #End If
    Const errBadFileNameOrNumber As Long = 52
    Const currentDir As String = "."
    Const parentDir As String = ".."
    Dim folderName As String
    Dim fullPath As String
    Dim collFolders As Collection
    Dim collTemp As New Collection
    Dim dirFailed As Boolean
    Dim v As Variant
    Dim fixedPath As String: fixedPath = BuildPath(folderPath, vbNullString)
    '
    If includeSubFolders Then
        Set collFolders = New Collection 'Temp collection to be iterated later
    Else
        Set collFolders = collTarget 'No recusion so we add directly to target
    End If
    '
    On Error Resume Next
    folderName = Dir(fixedPath, fAttribute)
    dirFailed = (Err.Number = errBadFileNameOrNumber) 'Unsupported Unicode
    On Error GoTo 0
    '
    Do While LenB(folderName) > 0
        If folderName <> currentDir And folderName <> parentDir Then
            collTemp.Add folderName
            If InStr(1, folderName, "?") > 0 Then 'Unsupported Unicode
                Set collTemp = New Collection
                dirFailed = True
                Exit Do
            End If
        End If
        folderName = Dir
    Loop
    If dirFailed Or Len(fixedPath) > maxDirLen Then
        #If Mac Then

        #Else
            Dim fsoDir As Object
            Dim fsoFolder As Object: Set fsoFolder = GetFSOFolder(fixedPath)
            '
            If Not fsoFolder Is Nothing Then
                On Error Resume Next
                For Each fsoDir In fsoFolder.SubFolders
                    collFolders.Add fixedPath & fsoDir.Name
                Next fsoDir
                On Error GoTo 0
            End If
        #End If
    End If
    For Each v In collTemp
        fullPath = fixedPath & v
        If IsFolder(fullPath) Then collFolders.Add fullPath
    Next v
    If includeSubFolders Then
        Dim subFolderPath As Variant
        '
        For Each subFolderPath In collFolders
            collTarget.Add subFolderPath
            AddFoldersTo collTarget, CStr(subFolderPath), True, fAttribute
        Next subFolderPath
    End If
End Sub

'*******************************************************************************
'Returns the local drive path for a given path or null string if path not local
'Note that the input path does not need to be an existing file/folder
'Works with both UNC paths (Win) and OneDrive/SharePoint synchronized paths
'
'Important!
'The expectation is that 'fullPath' is NOT URL encoded. If you have an encoded
'   path (e.g. in Word, ActiveDocument.Path returns an encoded URL) then use
'   GetLocalPath(DecodeURL(fullPath)...
'*******************************************************************************
Public Function GetLocalPath(ByRef fullPath As String _
                           , Optional ByVal rebuildCache As Boolean = False _
                           , Optional ByVal returnInputOnFail As Boolean = False) As String
    #If Windows Then
        If InStr(1, fullPath, "https://", vbTextCompare) <> 1 Then
            Dim tempPath As String: tempPath = FixPathSeparators(fullPath)
            With GetDriveInfo(tempPath)
                If LenB(.driveLetter) > 0 Then
                    GetLocalPath = Replace(tempPath, .driveName _
                                , .driveLetter & ":", 1, 1, vbTextCompare)
                    Exit Function
                End If
            End With
        End If
    #End If
    GetLocalPath = GetOneDriveLocalPath(fullPath, rebuildCache)
    If LenB(GetLocalPath) = 0 And returnInputOnFail Then
        GetLocalPath = fullPath
    End If
End Function

'*******************************************************************************
'Returns the UNC path for a given path or null string if path is not remote
'Note that the input path does not need to be an existing file/folder
'*******************************************************************************
#If Windows Then
Private Function GetUNCPath(ByRef fullPath As String) As String
    With GetDriveInfo(fullPath)
        If LenB(.shareName) = 0 Then Exit Function  'Not UNC
        GetUNCPath = FixPathSeparators(Replace(fullPath, .driveName, .shareName _
                                             , 1, 1, vbTextCompare))
    End With
End Function
#End If

'*******************************************************************************
'Returns the relative path for a given 'fullPath' based on another full path
'   or a Null String if the two paths do not have a common root
'*******************************************************************************
Public Function GetRelativePath(ByRef fullPath As String _
                              , ByRef relativeToFullPath As String) As String
    Dim fPath As String
    Dim rPath As String
    '
    fPath = GetLocalPath(fullPath, , True)
    rPath = GetLocalPath(relativeToFullPath, , True)
    '
    Const ps As String = PATH_SEPARATOR
    Dim fParent As String
    Dim rParent As String
    Dim prevPos As Long
    Dim currPos As Long
    Dim rcurrPos As Long
    Dim diff As Long
    Dim isRFile As Boolean
    '
    Do
        prevPos = currPos
        currPos = InStr(currPos + 1, fPath, ps)
        rcurrPos = InStr(prevPos + 1, rPath, ps)
	    If currPos <> rcurrPos Then Exit Do
        diff = currPos - prevPos - 1
        If diff > 0 Then
            fParent = Mid$(fPath, prevPos + 1, diff)
            rParent = Mid$(rPath, prevPos + 1, diff)
            If StrComp(fParent, rParent, vbTextCompare) <> 0 Then Exit Do
        End If
    Loop Until currPos = 0
    If prevPos = 0 Then Exit Function
    '
    fPath = Mid$(fPath, prevPos + 1)
    isRFile = IsFile(rPath)
    rPath = Mid$(rPath, prevPos + 1)
    '
    If LenB(rPath) > 0 Then
        Dim psCount As Long
        currPos = 0
        Do
            currPos = InStr(currPos + 1, rPath, ps)
            psCount = psCount + 1
        Loop Until currPos = 0
        If Right$(rPath, 1) = ps Or isRFile Then psCount = psCount - 1
    End If
    If psCount > 0 Then
        GetRelativePath = Replace(Space$(psCount), " ", ".." & ps) & fPath
    Else
        GetRelativePath = "." & ps & fPath
    End If
End Function

'*******************************************************************************
'Returns the web path for a OneDrive local path or null string if not OneDrive
'Note that the input path does not need to be an existing file/folder
'*******************************************************************************
Public Function GetRemotePath(ByRef fullPath As String _
                            , Optional ByVal rebuildCache As Boolean = False _
                            , Optional ByVal returnInputOnFail As Boolean = False) As String
    Dim tempPath As String: tempPath = FixPathSeparators(fullPath)
    #If Windows Then
        GetRemotePath = GetUNCPath(tempPath)
        If LenB(GetRemotePath) > 0 Then Exit Function
    #End If
    GetRemotePath = GetOneDriveWebPath(tempPath, rebuildCache)
    If LenB(GetRemotePath) = 0 And returnInputOnFail Then
        GetRemotePath = fullPath
    End If
End Function

'*******************************************************************************
'Source: https://developer.apple.com/library/archive/documentation/AppleScript/Conceptual/AppleScriptLangGuide/reference/ASLR_cmds.html
'Returns a special folder constant on Mac based on the corresponding enum value
'*******************************************************************************
#If Mac Then
Public Function GetSpecialFolderConstant(ByVal sfc As SpecialFolderConstant) As String
    Static sfcs([_minSFC] To [_maxSFC]) As String
    '
    If sfc < [_minSFC] Or sfc > [_maxSFC] Then Exit Function
    If LenB(sfcs([_minSFC])) = 0 Then
        sfcs(sfc_ApplicationSupport) = "application support"
        sfcs(sfc_ApplicationsFolder) = "applications folder"
        sfcs(sfc_Desktop) = "desktop"
        sfcs(sfc_DesktopPicturesFolder) = "desktop pictures folder"
        sfcs(sfc_DocumentsFolder) = "documents folder"
        sfcs(sfc_DownloadsFolder) = "downloads folder"
        sfcs(sfc_FavoritesFolder) = "favorites folder"
        sfcs(sfc_FolderActionScripts) = "Folder Action scripts"
        sfcs(sfc_Fonts) = "fonts"
        sfcs(sfc_Help) = "help"
        sfcs(sfc_HomeFolder) = "home folder"
        sfcs(sfc_InternetPlugins) = "internet plugins"
        sfcs(sfc_KeychainFolder) = "keychain folder"
        sfcs(sfc_LibraryFolder) = "library folder"
        sfcs(sfc_ModemScripts) = "modem scripts"
        sfcs(sfc_MoviesFolder) = "movies folder"
        sfcs(sfc_MusicFolder) = "music folder"
        sfcs(sfc_PicturesFolder) = "pictures folder"
        sfcs(sfc_Preferences) = "preferences"
        sfcs(sfc_PrinterDescriptions) = "printer descriptions"
        sfcs(sfc_PublicFolder) = "public folder"
        sfcs(sfc_ScriptingAdditions) = "scripting additions"
        sfcs(sfc_ScriptsFolder) = "scripts folder"
        sfcs(sfc_ServicesFolder) = "services folder"
        sfcs(sfc_SharedDocuments) = "shared documents"
        sfcs(sfc_SharedLibraries) = "shared libraries"
        sfcs(sfc_SitesFolder) = "sites folder"
        sfcs(sfc_StartupDisk) = "startup disk"
        sfcs(sfc_StartupItems) = "startup items"
        sfcs(sfc_SystemFolder) = "system folder"
        sfcs(sfc_SystemPreferences) = "system preferences"
        sfcs(sfc_TemporaryItems) = "temporary items"
        sfcs(sfc_Trash) = "trash"
        sfcs(sfc_UsersFolder) = "users folder"
        sfcs(sfc_UtilitiesFolder) = "utilities folder"
        sfcs(sfc_WorkflowsFolder) = "workflows folder"
        '
        'Classic domain only
        sfcs(sfc_AppleMenu) = "apple menu"
        sfcs(sfc_ControlPanels) = "control panels"
        sfcs(sfc_ControlStripModules) = "control strip modules"
        sfcs(sfc_Extensions) = "extensions"
        sfcs(sfc_LauncherItemsFolder) = "launcher items folder"
        sfcs(sfc_PrinterDrivers) = "printer drivers"
        sfcs(sfc_Printmonitor) = "printmonitor"
        sfcs(sfc_ShutdownFolder) = "shutdown folder"
        sfcs(sfc_SpeakableItems) = "speakable items"
        sfcs(sfc_Stationery) = "stationery"
        sfcs(sfc_Voices) = "voices"
    End If
    GetSpecialFolderConstant = sfcs(sfc)
End Function
#End If

'*******************************************************************************
'Returns a special folder domain on Mac based on the corresponding enum value
'*******************************************************************************
#If Mac Then
Public Function GetSpecialFolderDomain(ByVal sfd As SpecialFolderDomain) As String
    Static sfds([_minSFD] To [_maxSFD]) As String
    '
    If sfd < [_minSFD] Or sfd > [_maxSFD] Then Exit Function
    If LenB(sfds([_maxSFD])) = 0 Then
        sfds(sfd_System) = "system"
        sfds(sfd_Local) = "local"
        sfds(sfd_Network) = "network"
        sfds(sfd_User) = "user"
        sfds(sfd_Classic) = "classic"
    End If
    GetSpecialFolderDomain = sfds(sfd)
End Function
#End If

'*******************************************************************************
'Returns the path of a 'special folder' on Mac
'If 'createIfMissing' is set to True, the function will try to create the folder
'   if it does not currently exist on the system. Note that this argument
'   ignores the 'forceNonSandboxedPath' option, and it can happen that the
'   folder gets created in the sandboxed location and the function returns a non
'   sandboxed path. This behavior can not be avoided without creating access
'   requests, therefore it should be taken into account by the user
'The function can raise the following errors:
'   -  5: (Invalid procedure call) if 'sfc' or 'sfd' is invalid
'   - 76: (Path not found) if 'createIfMissing' = False AND path not found
'   - 75: (Path/File access error) if 'createIfMissing'= True AND path not found
'*******************************************************************************
#If Mac Then
Public Function GetSpecialFolderPath(ByVal sfc As SpecialFolderConstant _
                                   , Optional ByVal sfd As SpecialFolderDomain = [_sfdNone] _
                                   , Optional ByVal forceNonSandboxedPath As Boolean = True _
                                   , Optional ByVal createIfMissing As Boolean = False) As String
    Const methodName As String = "GetSpecialFolderPath"
    '
    If sfc < [_minSFC] Or sfc > [_maxSFC] _
    Or sfd < [_minSFD] Or sfd > [_maxSFD] Then
        Err.Raise vbErrInvalidProcedureCall, methodName, "Invalid constant/domain"
    End If
    '
    Dim cmd As String: cmd = GetSpecialFolderConstant(sfc)
    '
    If sfd <> [_sfdNone] Then cmd = cmd & " from " & GetSpecialFolderDomain(sfd) & " domain"
    cmd = cmd & IIf(createIfMissing, " with", " without") & " folder creation"
    cmd = "return POSIX path of (path to " & cmd & ") as string"
    '
    On Error Resume Next
    Dim app As Object:        Set app = Application
    Dim inExcel As Boolean:   inExcel = (app.Name = "Microsoft Excel")
    Dim appVersion As Double: appVersion = Val(app.Version)
    On Error GoTo 0
    '
    If inExcel And appVersion < 15 Then 'Old excel version
        cmd = Replace(cmd, "POSIX path of ", vbNullString, , 1)
    End If
    '
    On Error GoTo PathDoesNotExist
    GetSpecialFolderPath = MacScript(cmd)
    On Error GoTo 0
    '
    If forceNonSandboxedPath Then
        Dim sboxPath As String:    sboxPath = Environ$("HOME")
        Dim i As Long:             i = InStrRev(sboxPath, "/Library/Containers/")
        Dim sboxRelPath As String: If i > 0 Then sboxRelPath = Mid$(sboxPath, i)
        GetSpecialFolderPath = Replace(GetSpecialFolderPath, sboxRelPath _
                                     , vbNullString, , 1, vbTextCompare)
    End If
    If LenB(GetSpecialFolderPath) > 0 Then Exit Function
PathDoesNotExist:
    Const errMsg As String = "Not available or needs specific domain"
    If createIfMissing Then
        Err.Raise vbErrPathFileAccessError, methodName, errMsg
    Else
        Err.Raise vbErrPathNotFound, methodName, errMsg
    End If
End Function
#End If

'*******************************************************************************
'Returns basic drive information about a full path
'*******************************************************************************
#If Windows Then
Private Function GetDriveInfo(ByRef fullPath As String) As DRIVE_INFO
    Dim fso As Object: Set fso = GetFSO()
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
        Dim tempSN As Long
        Dim isFound As Boolean
        '
        On Error Resume Next 'In case Drive is not connected
        For Each tempDrive In fso.Drives
            tempSN = tempDrive.SerialNumber
            If tempSN = sn Then
                Set fsDrive = tempDrive
                isFound = True
                Exit For
            End If
        Next tempDrive
        On Error GoTo 0
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
#If Windows Then
Private Function GetFSO() As Object
    Static fso As Object
    '
    If fso Is Nothing Then
        On Error Resume Next
        Set fso = CreateObject("Scripting.FileSystemObject")
        On Error GoTo 0
    End If
    Set GetFSO = fso
End Function
#End If

'*******************************************************************************
'Aligns a wrong drive name with the share name
'Example: \\emea\ to \\emea.companyName.net\
'*******************************************************************************
#If Windows Then
Private Function AlignDriveNameIfNeeded(ByRef driveName As String _
                                      , ByRef shareName As String) As String
    Dim sepIndex As Long
    '
    sepIndex = InStr(3, driveName, PATH_SEPARATOR)
    If sepIndex > 0 Then
        Dim newName As String: newName = Left$(driveName, sepIndex - 1)
        sepIndex = InStr(3, shareName, PATH_SEPARATOR)
        newName = newName & Right$(shareName, Len(shareName) - sepIndex + 1)
        AlignDriveNameIfNeeded = newName
    Else
        AlignDriveNameIfNeeded = driveName
    End If
End Function
#End If

Public Function DecodeURL(ByRef odWebPath As String) As String
    Static nibbleMap(0 To 255) As Long 'Nibble: 0 to F. Byte: 00 to FF
    Static charMap(0 To 255) As String
    Dim i As Long
    '
    If nibbleMap(0) = 0 Then
        For i = 0 To 255
            nibbleMap(i) = -256 'To force invalid character code
            charMap(i) = ChrW$(i)
        Next i
        For i = 0 To 9
            nibbleMap(Asc(CStr(i))) = i
        Next i
        For i = 10 To 15
            nibbleMap(i + 55) = i 'Asc("A") to Asc("F")
            nibbleMap(i + 87) = i 'Asc("a") to Asc("f")
        Next i
    End If
    '
    DecodeURL = odWebPath 'Buffer
    '
    Dim b() As Byte:     b = odWebPath
    Dim pathLen As Long: pathLen = Len(odWebPath)
    Dim maxFind As Long: maxFind = pathLen * 2 - 4
    Dim codeW As Integer
    Dim j As Long
    Dim diff As Long
    Dim chunkLen As Long
    '
    i = InStrB(1, odWebPath, "%")
    Do While i > 0 And i < maxFind
        codeW = nibbleMap(b(i + 1)) * &H10& + nibbleMap(b(i + 3))
        If codeW > 0 And b(i + 2) = 0 And b(i + 4) = 0 Then
            If j > 0 Then
                chunkLen = i - j
                If chunkLen > 0 Then
                    MidB$(DecodeURL, j - diff) = MidB$(odWebPath, j, chunkLen)
                End If
            End If
            MidB$(DecodeURL, i - diff) = charMap(codeW)
            i = i + 4
            j = i + 2
            diff = diff + 4
        End If
        i = InStrB(i + 2, odWebPath, "%")
    Loop
    If diff > 0 Then
        chunkLen = pathLen * 2 + 1 - j
        If chunkLen > 0 Then
            MidB$(DecodeURL, j - diff) = MidB$(odWebPath, j, chunkLen)
        End If
        DecodeURL = Left$(DecodeURL, pathLen - diff / 2)
    End If
End Function

'*******************************************************************************
'Returns the local path for a OneDrive web path
'Returns null string if the path provided is not a valid OneDrive web path
'
'With the help of: @guwidoe (https://github.com/guwidoe)
'See: https://github.com/cristianbuse/VBA-FileTools/issues/1
'*******************************************************************************
Private Function GetOneDriveLocalPath(ByVal odWebPath As String _
                                    , ByVal rebuildCache As Boolean) As String
    If InStr(1, odWebPath, "https://", vbTextCompare) <> 1 Then Exit Function
    '
    Dim collMatches As New Collection
    Dim bestMatch As Long
    Dim mainIndex As Long
    Dim i As Long
    '
    If rebuildCache Or Not m_providers.isSet Then ReadODProviders
    For i = 1 To m_providers.pCount
        If StrCompLeft(odWebPath, m_providers.arr(i).webPath, vbTextCompare) = 0 Then
            collMatches.Add i
            If Not m_providers.arr(i).isBusiness Then Exit For
            If m_providers.arr(i).isMain Then
                mainIndex = m_providers.arr(i).accountIndex
            End If
        End If
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
Private Function GetOneDriveWebPath(ByRef odLocalPath As String _
                                  , ByVal rebuildCache As Boolean) As String
    Dim localPath As String
    Dim rPart As String
    Dim bestMatch As Long
    Dim i As Long
    Dim fixedPath As String: fixedPath = FixPathSeparators(odLocalPath)
    '
    If rebuildCache Or Not m_providers.isSet Then ReadODProviders
    For i = 1 To m_providers.pCount
        localPath = m_providers.arr(i).mountPoint
        If StrCompLeft(fixedPath, localPath, vbTextCompare) = 0 Then
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
        rPart = Replace(Mid$(fixedPath, Len(.mountPoint) + 1), "\", "/")
        GetOneDriveWebPath = .webPath & rPart
    End With
End Function

'*******************************************************************************
'Populates the OneDrive providers in the 'm_providers' structure
'Utility for 'GetOneDriveLocalPath' and 'GetOneDriveWebPath'
'*******************************************************************************
Private Sub ReadODProviders()
    Dim i As Long
    Dim accountsInfo As ONEDRIVE_ACCOUNTS_INFO
    '
    m_providers.pCount = 0
    m_providers.isSet = False
    '
    ReadODAccountsInfo accountsInfo
    If Not accountsInfo.isSet Then Exit Sub
    '
    #If Mac Then 'Grant access to all needed files/folders, in batch
        Dim collFiles As New Collection
        Dim fileName As String
        '
        For i = 1 To accountsInfo.pCount
            With accountsInfo.arr(i)
                collFiles.Add .iniPath
                collFiles.Add .datPath
                collFiles.Add .dbPath
                collFiles.Add .clientPath
                collFiles.Add .globalPath
                If .isPersonal Then
                    collFiles.Add .groupPath
                Else
                    fileName = Dir(Replace(.clientPath, ".ini", "_*.ini"))
                    Do While LenB(fileName) > 0
                        collFiles.Add .folderPath & "/" & fileName
                        fileName = Dir
                    Loop
                End If
            End With
        Next i
        '
        Const syncIDFileName As String = ".849C9593-D756-4E56-8D6E-42412F2A707B"
        Dim collCloudDirs As Collection: Set collCloudDirs = GetODCloudDirs()
        Dim odCloudDir As Variant
        Dim arrPaths() As String
        Dim syncID As String
        Dim folderPath As String
        Dim collSyncIDToDir As New Collection
        '
        For Each odCloudDir In collCloudDirs
            collFiles.Add odCloudDir
            collFiles.Add odCloudDir & "/" & syncIDFileName
        Next odCloudDir
        arrPaths = CollectionToStrings(collFiles)
        If Not GrantAccessToMultipleFiles(arrPaths) Then Exit Sub
        '
        Set collFiles = New Collection
        For Each odCloudDir In collCloudDirs
            syncID = ReadSyncID(odCloudDir & "/" & syncIDFileName)
            If LenB(syncID) > 0 Then
                collSyncIDToDir.Add odCloudDir, syncID
            Else
                fileName = Dir(odCloudDir & "/", vbDirectory)
                Do While LenB(fileName) > 0
                    folderPath = odCloudDir & "/" & fileName
                    collFiles.Add folderPath
                    collFiles.Add folderPath & "/" & syncIDFileName
                    fileName = Dir
                Loop
            End If
        Next odCloudDir
        If collFiles.Count > 0 Then
            arrPaths = CollectionToStrings(collFiles)
            If Not GrantAccessToMultipleFiles(arrPaths) Then Exit Sub
            '
            For i = LBound(arrPaths) To UBound(arrPaths) Step 2
                syncID = ReadSyncID(arrPaths(i + 1))
                If LenB(syncID) > 0 Then collSyncIDToDir.Add arrPaths(i), syncID
            Next i
        End If
    #End If
    For i = 1 To accountsInfo.pCount 'Check for unsynchronized accounts
        Dim j As Long
        For j = i + 1 To accountsInfo.pCount
            ValidateAccounts accountsInfo.arr(i), accountsInfo.arr(j)
        Next j
    Next i
    For i = 1 To accountsInfo.pCount
        If accountsInfo.arr(i).isValid Then
            If accountsInfo.arr(i).isPersonal Then
                AddPersonalProviders accountsInfo.arr(i)
            Else
                AddBusinessProviders accountsInfo.arr(i)
            End If
        End If
    Next i
    #If Mac Then
        If collSyncIDToDir.Count > 0 Then 'Replace sandbox paths
            For i = 1 To m_providers.pCount
                With m_providers.arr(i)
                    On Error Resume Next
                    .syncDir = collSyncIDToDir(.syncID)
                    .mountPoint = Replace(.mountPoint, .baseMount, .syncDir)
                    On Error GoTo 0
                End With
            Next i
        End If
    #End If
    m_providers.isSet = True
#If Mac Then
    ClearConversionDescriptors
#End If
End Sub

'*******************************************************************************
'Mac utilities for reading OneDrive providers
'*******************************************************************************
#If Mac Then
Private Function GetODCloudDirs() As Collection
    Dim coll As New Collection
    Dim cloudPath As String:  cloudPath = GetCloudPath()
    Dim folderName As String: folderName = Dir(cloudPath, vbDirectory)
    '
    Do While LenB(folderName) > 0
        If folderName Like "OneDrive*" Then
            coll.Add BuildPath(cloudPath, folderName)
        End If
        folderName = Dir
    Loop
    Set GetODCloudDirs = coll
End Function
Private Function GetCloudPath() As String
    GetCloudPath = GetUserPath() & "Library/CloudStorage/"
End Function
Private Function GetUserPath() As String
    GetUserPath = "/Users/" & Environ$("USER") & "/"
End Function
Private Function CollectionToStrings(ByVal coll As Collection) As String()
    If coll.Count = 0 Then
        CollectionToStrings = Split(vbNullString)
        Exit Function
    End If
    '
    Dim res() As String: ReDim res(0 To coll.Count - 1)
    Dim i As Long
    Dim v As Variant
    '
    For Each v In coll
        res(i) = v
        i = i + 1
    Next v
    CollectionToStrings = res
End Function
Private Function ReadSyncID(ByRef syncFilePath As String) As String
    Dim b() As Byte:       ReadBytes syncFilePath, b
    Dim parts() As String: parts = Split(StrConv(b, vbUnicode), """guid"" : """)
    '
    If UBound(parts) < 1 Then Exit Function
    ReadSyncID = Left$(parts(1), InStr(1, parts(1), """") - 1)
End Function
#End If
Private Sub ValidateAccounts(ByRef a1 As ONEDRIVE_ACCOUNT_INFO _
                           , ByRef a2 As ONEDRIVE_ACCOUNT_INFO)
    If a1.accountName <> a2.accountName Then Exit Sub
    If Not (a1.isValid And a2.isValid) Then Exit Sub
    '
    If a1.iniDateTime = 0 Then a1.iniDateTime = FileDateTime(a1.iniPath)
    If a2.iniDateTime = 0 Then a2.iniDateTime = FileDateTime(a2.iniPath)
    '
    a1.isValid = (a1.iniDateTime > a2.iniDateTime)
    a2.isValid = Not a1.isValid
End Sub

'*******************************************************************************
'Utility for reading folder information for all the OneDrive accounts
'*******************************************************************************
Private Sub ReadODAccountsInfo(ByRef accountsInfo As ONEDRIVE_ACCOUNTS_INFO)
    Const ps As String = PATH_SEPARATOR
    Dim folderPath As Variant
    Dim i As Long
    Dim hasIniFile As Boolean
    Dim collFolders As Collection: Set collFolders = GetODAccountDirs()
    '
    accountsInfo.pCount = 0
    accountsInfo.isSet = False
    '
    If collFolders Is Nothing Then Exit Sub
    If collFolders.Count > 0 Then ReDim accountsInfo.arr(1 To collFolders.Count)
    '
    For Each folderPath In collFolders
        i = i + 1
        With accountsInfo.arr(i)
            .folderPath = folderPath
            .accountName = Mid$(folderPath, InStrRev(folderPath, ps) + 1)
            .isPersonal = (.accountName = "Personal")
            If Not .isPersonal Then
                .accountIndex = CLng(Right$(.accountName, 1))
            End If
            .globalPath = .folderPath & ps & "global.ini"
            .cID = GetTagValue(.globalPath, "cid = ")
            .iniPath = .folderPath & ps & .cID & ".ini"
            #If Mac Then 'Avoid Mac File Access Request
                hasIniFile = (Dir(.iniPath & "*") = .cID & ".ini")
            #Else
                hasIniFile = IsFile(.iniPath)
            #End If
            If hasIniFile Then
                .datPath = .folderPath & ps & .cID & ".dat"
                .dbPath = .folderPath & ps & "SyncEngineDatabase.db"
                .groupPath = .folderPath & ps & "GroupFolders.ini"
                .clientPath = .folderPath & ps & "ClientPolicy.ini"
                #If Mac Then 'Avoid Mac File Access Request
                    .hasDatFile = (Dir(.datPath & "*") = .cID & ".dat")
                #Else
                    .hasDatFile = IsFile(.datPath)
                #End If
                .isValid = True
            End If
            If Not .isValid Then i = i - 1
        End With
    Next folderPath
    With accountsInfo
        If i > 0 And i < collFolders.Count Then ReDim Preserve .arr(1 To i)
        .pCount = i
        .isSet = True
    End With
End Sub

'*******************************************************************************
'Utility for reading all OneDrive account folder paths within OneDrive Settings
'*******************************************************************************
Private Function GetODAccountDirs() As Collection
    Dim collSettings As Collection: Set collSettings = GetODSettingsDirs()
    Dim settingsPath As Variant
    '
    #If Mac Then 'Grant access, if needed, to all possbile folders, in batch
        Dim arrDirs() As Variant: ReDim arrDirs(0 To collSettings.Count * 11)
        Dim i As Long
        '
        arrDirs(i) = GetCloudPath()
        For Each settingsPath In collSettings
            For i = i + 1 To i + 9
                arrDirs(i) = settingsPath & "Business" & i Mod 11
            Next i
            arrDirs(i) = settingsPath
            i = i + 1
            arrDirs(i) = settingsPath & "Personal"
        Next settingsPath
        If Not GrantAccessToMultipleFiles(arrDirs) Then Exit Function
    #End If
    '
    Dim folderPath As Variant
    Dim folderName As String
    Dim collFolders As New Collection
    '
    For Each settingsPath In collSettings
        folderName = Dir(settingsPath, vbDirectory)
        Do While LenB(folderName) > 0
            If folderName Like "Business#" Or folderName = "Personal" Then
                folderPath = settingsPath & folderName
                If IsFolder(CStr(folderPath)) Then collFolders.Add folderPath
            End If
            folderName = Dir
        Loop
    Next settingsPath
    Set GetODAccountDirs = collFolders
End Function

'*******************************************************************************
'Utility returning all possible OneDrive Settings folders
'*******************************************************************************
Private Function GetODSettingsDirs() As Collection
    Set GetODSettingsDirs = New Collection
    With GetODSettingsDirs
    #If Mac Then
        Const settingsPath = "Library/Application Support/OneDrive/settings/"
        Const dataPath = "Library/Containers/com.microsoft.OneDrive-mac/Data/"
        .Add GetUserPath() & settingsPath
        .Add GetUserPath() & dataPath & settingsPath
    #Else
        .Add BuildPath(Environ$("LOCALAPPDATA"), "Microsoft\OneDrive\settings\")
    #End If
    End With
End Function

'*******************************************************************************
'Returns the index of the newly added OneDrive provider struct
'*******************************************************************************
Private Function AddProvider() As Long
    With m_providers
        If .pCount = 0 Then Erase .arr
        .pCount = .pCount + 1
        ReDim Preserve .arr(1 To .pCount)
        AddProvider = .pCount
    End With
End Function

'*******************************************************************************
'Adds all providers for a Business OneDrive account
'*******************************************************************************
Private Sub AddBusinessProviders(ByRef aInfo As ONEDRIVE_ACCOUNT_INFO)
    Dim bytes() As Byte:   ReadBytes aInfo.iniPath, bytes
    Dim iniText As String: iniText = bytes
    Dim lineText As Variant
    Dim tempMount As String
    Dim mainMount As String
    Dim syncID As String
    Dim mainSyncID As String
    Dim tempURL As String
    Dim cSignature As String
    Dim oDirs As DirsInfo
    Dim cParents As Collection
    Dim cPending As New Collection
    Dim canAdd As Boolean
    Dim collTags As New Collection
    Dim arrTags() As Variant
    Dim vTag As Variant
    Dim tempColl As Collection
    Dim collSortedLines As New Collection
    Dim i As Long, j As Long
    Dim targetCount As Long
    '
    #If Mac Then
        iniText = ConvertText(iniText, codeUTF16LE, codeUTF8, True)
    #End If
    arrTags = Array("libraryScope", "libraryFolder", "AddedScope")
    For Each vTag In arrTags
        collTags.Add New Collection, vTag
    Next vTag
    For Each lineText In Split(iniText, vbNewLine)
        i = InStr(1, lineText, " = ", vbBinaryCompare)
        If i > 0 Then
            vTag = Left$(lineText, i - 1)
            Select Case vTag
            Case arrTags(0), arrTags(1), arrTags(2)
                i = i + 3
                j = InStr(i, lineText, " ", vbBinaryCompare)
                collTags(vTag).Add lineText, Mid$(lineText, i, j - i)
            End Select
        End If
    Next lineText
    On Error Resume Next
    For Each tempColl In collTags
        i = 0
        targetCount = collSortedLines.Count + tempColl.Count
        Do
            collSortedLines.Add tempColl(CStr(i))
            i = i + 1
        Loop Until collSortedLines.Count = targetCount
    Next tempColl
    On Error GoTo 0
    For Each lineText In collSortedLines
        Dim parts() As String: parts = SplitIniLine(lineText)
        Select Case parts(0)
        Case "libraryScope"
            tempMount = parts(14)
            syncID = parts(16)
            canAdd = (LenB(tempMount) > 0)
            If parts(2) = "0" Then
                mainMount = tempMount
                mainSyncID = syncID
                tempURL = GetUrlNamespace(aInfo.clientPath)
            Else
                cSignature = "_" & parts(12) & parts(10)
                tempURL = GetUrlNamespace(aInfo.clientPath, cSignature)
            End If
            cPending.Add tempURL, parts(2)
        Case "libraryFolder"
            If oDirs.dirCount = 0 Then ReadODDirs aInfo, oDirs
            tempMount = parts(6)
            tempURL = cPending(parts(3))
            syncID = parts(9)
            Dim tempID As String:     tempID = parts(4)
            Dim tempFolder As String: tempFolder = vbNullString
            If aInfo.hasDatFile Then tempID = Split(tempID, "+")(0)
            On Error Resume Next
            Do
                i = oDirs.idToIndex(tempID)
                If Err.Number <> 0 Then Exit Do
                With oDirs.arrDirs(i)
                    tempFolder = .dirName & "/" & tempFolder
                    tempID = .parentID
                End With
            Loop
            On Error GoTo 0
            canAdd = (LenB(tempFolder) > 0)
            tempURL = tempURL & tempFolder
        Case "AddedScope"
            If LenB(mainMount) = 0 Then Err.Raise vbErrInvalidFormatInResourceFile
            If oDirs.dirCount = 0 Then ReadODDirs aInfo, oDirs
            tempID = parts(3)
            tempFolder = vbNullString
            On Error Resume Next
            Do
                i = oDirs.idToIndex(tempID)
                If Err.Number <> 0 Then Exit Do
                With oDirs.arrDirs(i)
                    tempFolder = .dirName & PATH_SEPARATOR & tempFolder
                    tempID = .parentID
                End With
            Loop
            On Error GoTo 0
            tempMount = mainMount & PATH_SEPARATOR & tempFolder
            syncID = mainSyncID
            tempURL = parts(11)
            If tempURL = " " Or LenB(tempURL) = 0 Then
                tempURL = vbNullString
            Else
                tempURL = tempURL & "/"
            End If
            cSignature = "_" & parts(9) & parts(7) & parts(10)
            tempURL = GetUrlNamespace(aInfo.clientPath, cSignature) & tempURL
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
                .accountIndex = aInfo.accountIndex
                If syncID = mainSyncID Then
                    .baseMount = mainMount
                Else
                    .baseMount = tempMount
                End If
                .syncID = syncID
            End With
        End If
    Next lineText
End Sub

'*******************************************************************************
'Splits a cid.ini file into space delimited parts
'*******************************************************************************
Private Function SplitIniLine(ByVal lineText As String) As String()
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim res() As String: ReDim res(0 To 20)
    Dim v As Variant
    Dim s As String
    Dim c As Long: c = Len(lineText)
    '
    i = InStr(1, lineText, " ")
    res(0) = Left$(lineText, i - 1)
    Do
        Do
            i = i + 1
            s = Mid$(lineText, i, 1)
        Loop Until s <> " "
        If i > c Then Exit Do
        If s = """" Then
            i = i + 1
            j = InStr(i, lineText, """")
        Else
            j = InStr(i + 1, lineText, " ")
        End If
        If j = 0 Then j = c + 1
        k = k + 1
        If k > UBound(res) Then ReDim Preserve res(0 To k)
        res(k) = Mid$(lineText, i, j - i)
        i = j
    Loop Until j > c
    ReDim Preserve res(0 To k)
    SplitIniLine = res
End Function

'*******************************************************************************
'Returns the URLNamespace from a provider's ClientPolicy*.ini file
'*******************************************************************************
Private Function GetUrlNamespace(ByRef clientPath As String _
                               , Optional ByVal cSignature As String) As String
    Dim cPath As String
    '
    cPath = Left$(clientPath, Len(clientPath) - 4) & cSignature & ".ini"
    GetUrlNamespace = GetTagValue(cPath, "DavUrlNamespace = ")
End Function

'*******************************************************************************
'Returns the required value from an ini file text line based on given tag
'*******************************************************************************
Private Function GetTagValue(ByRef filePath As String _
                           , ByRef vTag As String) As String
    Dim bytes() As Byte
    Dim fText As String
    Dim i As Long
    Dim j As Long
    '
    On Error Resume Next
    ReadBytes filePath, bytes
    fText = bytes
    #If Mac Then
        If Err.Number = 0 Then
            fText = ConvertText(fText, codeUTF16LE, codeUTF8, True)
        Else 'Open failed, try AppleScript with no text conversion needed
            Dim tempPath As String
            tempPath = MacScript("return path to startup disk as string") _
                     & Replace(Mid$(filePath, 2), PATH_SEPARATOR, ":")
            fText = MacScript("return read file """ & tempPath & """ as string")
        End If
    #End If
    On Error GoTo 0
    '
    If Len(fText) = 0 Then Exit Function
    i = InStr(1, fText, vTag)
    If i = 0 Then Exit Function
    i = i + Len(vTag)
    '
    j = InStr(i + 1, fText, vbNewLine)
    If j = 0 Then
        GetTagValue = Mid$(fText, i)
    Else
        GetTagValue = Mid$(fText, i, j - i)
    End If
End Function

'*******************************************************************************
'Adds all providers for a Personal OneDrive account
'*******************************************************************************
Private Sub AddPersonalProviders(ByRef aInfo As ONEDRIVE_ACCOUNT_INFO)
    Dim mainURL As String
    Dim libText As String
    Dim libParts() As String
    Dim mainMount As String
    Dim bytes() As Byte
    Dim groupText As String
    Dim syncID As String
    Dim lineText As Variant
    Dim cID As String
    Dim i As Long
    Dim relPath As String
    Dim folderID As String
    Dim oDirs As DirsInfo
    '
    ReadBytes aInfo.groupPath, bytes
    groupText = bytes
    '
    mainURL = GetUrlNamespace(aInfo.clientPath) & "/"
    libText = GetTagValue(aInfo.iniPath, "library = ")
    If LenB(libText) > 0 Then
        libParts = SplitIniLine(libText)
        mainMount = libParts(7)
        syncID = libParts(9)
    Else
        libText = GetTagValue(aInfo.iniPath, "libraryScope = ")
        libParts = SplitIniLine(libText)
        mainMount = libParts(12)
        syncID = libParts(7)
    End If
    '
    With m_providers.arr(AddProvider())
        .webPath = mainURL & aInfo.cID & "/"
        .mountPoint = mainMount & PATH_SEPARATOR
        .baseMount = mainMount
        .syncID = syncID
    End With
    #If Mac Then
        groupText = ConvertText(groupText, codeUTF16LE, codeUTF8, True)
    #End If
    For Each lineText In Split(groupText, vbNewLine)
        If InStr(1, lineText, "_BaseUri", vbTextCompare) > 0 Then
            cID = LCase$(Mid$(lineText, InStrRev(lineText, "/") + 1))
            i = InStr(1, cID, "!")
            If i > 0 Then cID = Left$(cID, i - 1)
        Else
            i = InStr(1, lineText, "_Path", vbTextCompare)
            If i > 0 Then
                relPath = Mid$(lineText, i + 8)
                folderID = Left$(lineText, i - 1)
                If oDirs.dirCount = 0 Then ReadODDirs aInfo, oDirs
                With m_providers.arr(AddProvider())
                    .webPath = mainURL & cID & "/" & relPath & "/"
                    relPath = oDirs.arrDirs(oDirs.idToIndex(folderID)).dirName
                    .mountPoint = BuildPath(mainMount, relPath & "/")
                    .baseMount = mainMount
                    .syncID = syncID
                End With
            End If
        End If
    Next lineText
End Sub

'*******************************************************************************
'Utility - Retrieves all folders from an OneDrive account
'*******************************************************************************
Private Sub ReadODDirs(ByRef aInfo As ONEDRIVE_ACCOUNT_INFO _
                     , ByRef outdirs As DirsInfo)
    If aInfo.hasDatFile Then
        ReadDirsFromDat aInfo.datPath, outdirs
    End If
    If outdirs.dirCount = 0 Then
        ReadDirsFromDB aInfo.dbPath, aInfo.isPersonal, outdirs
    End If
End Sub

'*******************************************************************************
'Utility - Retrieves all folders from an OneDrive user dat file
'*******************************************************************************
Private Sub ReadDirsFromDat(ByRef filePath As String, ByRef outdirs As DirsInfo)
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
    Const chunkSize As Long = &H100000 '1MB
    Const maxDirName As Long = 255
    #If Mac Then
        Const nameEnd As String = vbNullChar & vbNullChar
    #Else
        Const nameEnd As String = vbNullChar
    #End If
    '
    Dim b(1 To chunkSize) As Byte
    Dim s As String
    Dim lastRecord As Long
    Dim i As Long
    Dim lastFileChange As Date
    Dim currFileChange As Date
    Dim stepSize As Long
    Dim bytes As Long
    Dim dirID As String
    Dim parentID As String
    Dim dirName As String
    Dim idPattern As String
    Dim vbNullByte As String: vbNullByte = ChrB$(0)
    Dim hFolder As String:    hFolder = ChrB$(2) 'x02
    Dim hCheck As String * 4: MidB$(hCheck, 1) = ChrB$(1) 'x01000000
    '
    idPattern = Replace(Space$(12), " ", "[a-fA-F0-9]") & "*"
    For stepSize = 16 To 8 Step -8
        lastFileChange = 0
        Do
            i = 0
            currFileChange = FileDateTime(filePath)
            If currFileChange > lastFileChange Then
                With outdirs
                    Set .idToIndex = New Collection
                    .dirCount = 0
                    .dirUBound = 256
                    ReDim .arrDirs(1 To .dirUBound)
                End With
                lastFileChange = currFileChange
                lastRecord = 1
            End If
            Get fileNumber, lastRecord, b
            s = b
            i = InStrB(stepSize + 1, s, hCheck)
            Do While i > 0 And i < chunkSize - checkToName
                If MidB$(s, i - stepSize, 1) = hFolder Then
                    i = i + hCheckSize
                    bytes = Clamp(InStrB(i, s, vbNullByte) - i, 0, idSize)
                    dirID = StrConv(MidB$(s, i, bytes), vbUnicode)
                    '
                    i = i + idSize
                    bytes = Clamp(InStrB(i, s, vbNullByte) - i, 0, idSize)
                    parentID = StrConv(MidB$(s, i, bytes), vbUnicode)
                    '
                    i = i + fNameOffset
                    If dirID Like idPattern And parentID Like idPattern Then
                        bytes = InStr((i + 1) \ 2, s, nameEnd) * 2 - i - 1
                        #If Mac Then
                            Do While bytes Mod 4 > 0 And bytes > 0
                                If bytes > maxDirName * 4 Then
                                    bytes = maxDirName * 4
                                    Exit Do
                                End If
                                bytes = InStr((i + bytes + 1) \ 2 + 1, s, nameEnd) _
                                      * 2 - i - 1
                            Loop
                        #Else
                            If bytes > maxDirName * 2 Then bytes = maxDirName * 2
                        #End If
                        If bytes < 0 Or i + bytes - 1 > chunkSize Then 'Next chunk
                            i = i - checkToName
                            Exit Do
                        End If
                        dirName = MidB$(s, i, bytes)
                        #If Mac Then
                            dirName = ConvertText(dirName, codeUTF16LE _
                                                   , codeUTF32LE, True)
                        #End If
                        With outdirs
                            .dirCount = .dirCount + 1
                            If .dirCount > .dirUBound Then
                                .dirUBound = .dirUBound * 2
                                ReDim Preserve .arrDirs(1 To .dirUBound)
                            End If
                            .idToIndex.Add .dirCount, dirID
                            With outdirs.arrDirs(.dirCount)
                                .dirID = dirID
                                .dirName = dirName
                                .parentID = parentID
                            End With
                        End With
                    End If
                End If
                i = InStrB(i + 1, s, hCheck)
            Loop
            lastRecord = lastRecord + chunkSize - stepSize
            If i > stepSize Then
                lastRecord = lastRecord - chunkSize + (i \ 2) * 2
            End If
        Loop Until lastRecord > size
        If outdirs.dirCount > 0 Then Exit For
    Next stepSize
    If outdirs.dirCount > 0 Then
        ReDim Preserve outdirs.arrDirs(1 To outdirs.dirCount)
    End If
CloseFile:
    Close #fileNumber
End Sub
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
'Utility - Retrieves all folders from an OneDrive user database file
'*******************************************************************************
Private Sub ReadDirsFromDB(ByRef filePath As String _
                         , ByVal isPersonal As Boolean _
                         , ByRef outdirs As DirsInfo)
    If Not IsFile(filePath) Then Exit Sub
    Dim fileNumber As Long: fileNumber = FreeFile()
    '
    Open filePath For Binary Access Read As #fileNumber
    Dim size As Long: size = LOF(fileNumber)
    If size = 0 Then GoTo CloseFile
    '
    Const chunkSize As Long = &H100000 '1MB
    Const minName As Long = 15
    Const maxSigByte As Byte = 9
    Const maxHeader As Long = 21
    Const minIDSize As Long = 12
    Const maxIDSize As Long = 48
    Const minThreeIDSizes As Long = minIDSize * 3
    Const maxThreeIDSizes As Long = maxIDSize * 3
    Const leadingBuff As Long = maxHeader + maxThreeIDSizes
    Const headBytesOffset As Long = 15
    Const bangCode As Long = 33 'Asc("!")
    Dim curlyStart As String: curlyStart = ChrW$(&H7B22) '"{
    Dim quoteB As String:     quoteB = ChrB$(&H22)       '"
    Dim bangB As String:      bangB = ChrB$(bangCode)    '!
    Dim sig As String
    Dim b(1 To chunkSize) As Byte
    Dim s As String
    Dim lastRecord As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim idSize(1 To 4) As Long
    Dim nameSize As Long
    Dim dirID As String
    Dim parentID As String
    Dim dirName As String
    Dim nameEnd As Long
    Dim nameStart As Long
    Dim isASCII As Boolean
    Dim mustAdd As Boolean
    Dim idPattern As String
    '
    idPattern = Replace(Space$(12), " ", "[a-fA-F0-9]")
    If isPersonal Then
        sig = bangB
        idPattern = "*" & idPattern & "![a-fA-F0-9]*"
    Else
        sig = curlyStart
        idPattern = idPattern & "*"
    End If
    Do
        Dim currFileChange As Date: currFileChange = FileDateTime(filePath)
        Dim lastFileChange As Date
        '
        i = 0
        If currFileChange > lastFileChange Then
            With outdirs
                Set .idToIndex = New Collection
                .dirCount = 0
                .dirUBound = 256
                ReDim .arrDirs(1 To .dirUBound)
            End With
            lastFileChange = currFileChange
            lastRecord = 1
        End If
        Get fileNumber, lastRecord, b
        s = b
        i = InStrB(1, s, sig)
        Do While i > 0
            If isPersonal Then
                For j = i - 1 To i - maxIDSize Step -1
                    If j = 0 Then GoTo NextSig
                    If b(j) < bangCode Then Exit For
                Next j
                If (j < maxHeader) Or (i - j < minIDSize) Then GoTo NextSig
            Else
                j = InStrB(i + 2, s, quoteB)
                If j = 0 Then Exit Do 'Next chunk
                idSize(4) = j - i + 1
                If idSize(4) > maxIDSize Then GoTo NextSig
                For j = i - 1 To i - maxThreeIDSizes Step -1
                    If j = 0 Then GoTo NextSig
                    If b(j) < bangCode Then Exit For
                Next j
                If j < maxHeader Then GoTo NextSig
                idSize(1) = i - j - 1 'ID 1+2+3
                If idSize(1) < minThreeIDSizes Then GoTo NextSig
            End If
            '
            k = j + 1 'ID1 Start
            For j = j To j - headBytesOffset + 1 Step -1
                If b(j) > maxSigByte Then GoTo NextSig
            Next j
            If (b(j) <= maxSigByte) And (b(j - 1) < &H80) Then j = j - 1
            If b(j) < minName Then j = j - 1
            '
            nameSize = b(j)
            If nameSize Mod 2 = 0 Then GoTo NextSig
            nameSize = (nameSize - 13) / 2
            If b(j - 1) > &H7F Then
                nameSize = (b(j - 1) - &H80) * &H40 + nameSize
                j = j - 1
            End If
            If j < 5 Then GoTo NextSig
            If (nameSize < 1) Or (b(j - 4) = 0) Then GoTo NextSig
            '
            If isPersonal Then
                idSize(4) = (b(j - 1) - 13) / 2
                idSize(3) = (b(j - 2) - 13) / 2
                idSize(2) = (b(j - 3) - 13) / 2
                idSize(1) = (b(j - 4) - 13) / 2
                nameStart = k + idSize(1) + idSize(2) + idSize(3) + idSize(4)
            Else
                If b(j - 1) <> idSize(4) * 2 + 13 Then GoTo NextSig
                idSize(3) = (b(j - 2) - 13) / 2
                idSize(2) = (b(j - 3) - 13) / 2
                idSize(1) = idSize(1) - idSize(2) - idSize(3)
                nameStart = i + idSize(4)
            End If
            For j = 1 To 4
                If (idSize(j) < minIDSize) _
                Or (idSize(j) > maxIDSize) Then GoTo NextSig
            Next j
            '
            nameEnd = nameStart + nameSize - 1
            If nameEnd > chunkSize Then Exit Do 'Next chunk
            '
            dirID = StrConv(MidB$(s, k, idSize(1)), vbUnicode)
            If Not dirID Like idPattern Then GoTo NextSig
            '
            k = k + idSize(1)
            parentID = StrConv(MidB$(s, k, idSize(2)), vbUnicode)
            If Not parentID Like idPattern Then GoTo NextSig
            '
            If isPersonal Then
                k = k + idSize(2)
                If Not StrConv(MidB$(s, k, idSize(3)), vbUnicode) _
                       Like idPattern Then GoTo NextSig
                If Not StrConv(MidB$(s, k + idSize(3), idSize(4)), vbUnicode) _
                       Like idPattern Then GoTo NextSig
            End If
            '
            On Error Resume Next
            j = outdirs.idToIndex(dirID)
            mustAdd = (Err.Number <> 0)
            On Error GoTo 0
            '
            If mustAdd Then
                With outdirs
                    .dirCount = .dirCount + 1
                    If .dirCount > .dirUBound Then
                        .dirUBound = .dirUBound * 2
                        ReDim Preserve .arrDirs(1 To .dirUBound)
                    End If
                    .idToIndex.Add .dirCount, dirID
                    j = .dirCount
                End With
                With outdirs.arrDirs(j)
                    .dirName = MidB$(s, nameStart, nameSize)
                    .isNameASCII = True
                    For k = nameStart To nameEnd
                        If b(k) > &H7F Then
                            .isNameASCII = False
                            Exit For
                        End If
                    Next k
                    If .isNameASCII Then
                        .dirName = StrConv(.dirName, vbUnicode)
                    Else
                        .dirName = ConvertText(.dirName, codeUTF16LE, codeUTF8)
                    End If
                    .dirID = dirID
                    .parentID = parentID
                End With
            Else
                With outdirs.arrDirs(j)
                    If (Not .isNameASCII) Or (Len(.dirName) < nameSize) Then
                        dirName = MidB$(s, nameStart, nameSize)
                        isASCII = True
                        For k = nameStart To nameEnd
                            If b(k) > &H7F Then
                                isASCII = False
                                Exit For
                            End If
                        Next k
                        If isASCII Then
                            .dirName = StrConv(dirName, vbUnicode)
                        Else
                            .dirName = ConvertText(dirName, codeUTF16LE, codeUTF8)
                        End If
                        .isNameASCII = isASCII
                    End If
                End With
            End If
            i = nameEnd
NextSig:
            i = InStrB(i + 1, s, sig)
        Loop
        If i = 0 Then
            lastRecord = lastRecord + chunkSize - leadingBuff
        ElseIf i > leadingBuff Then
            lastRecord = lastRecord + i - leadingBuff
        Else
            lastRecord = lastRecord + i
        End If
    Loop Until lastRecord > size
    ReDim Preserve outdirs.arrDirs(1 To outdirs.dirCount)
CloseFile:
    Close #fileNumber
End Sub

'*******************************************************************************
'Checks if a path indicates a file path
'Note that if C:\Test\1.txt is valid then C:\Test\\///1.txt will also be valid
'Most VBA methods consider valid any path separators with multiple characters
'*******************************************************************************
Public Function IsFile(ByRef filePath As String) As Boolean
    #If Mac Then
        Const maxFileLen As Long = 259 'To be updated
    #Else
        Const maxFileLen As Long = 259
    #End If
    Const errBadFileNameOrNumber As Long = 52
    Dim fAttr As VbFileAttribute
    '
    On Error Resume Next
    fAttr = GetAttr(filePath)
    If Err.Number = errBadFileNameOrNumber Then 'Unicode characters
        #If Mac Then
            
        #Else
            IsFile = GetFSO().FileExists(filePath)
        #End If
    ElseIf Err.Number = 0 Then
        IsFile = Not CBool(fAttr And vbDirectory)
    ElseIf Len(filePath) > maxFileLen Then
        #If Mac Then

        #Else
            If Left$(filePath, 4) = "\\?\" Then
                IsFile = GetFSO().FileExists(filePath)
            ElseIf Left$(filePath, 2) = "\\" Then
                IsFile = GetFSO().FileExists("\\?\UNC" & Mid$(filePath, 2))
            Else
                IsFile = GetFSO().FileExists("\\?\" & filePath)
            End If
        #End If
    End If
    On Error GoTo 0
End Function
'*******************************************************************************
'Checks if a path indicates a folder path
'Note that if C:\Test\Demo is valid then C:\Test\\///Demo will also be valid
'Most VBA methods consider valid any path separators with multiple characters
'*******************************************************************************
Public Function IsFolder(ByRef folderPath As String) As Boolean
    #If Mac Then
        Const maxDirLen As Long = 247 'To be updated
    #Else
        Const maxDirLen As Long = 247
    #End If
    Const errBadFileNameOrNumber As Long = 52
    Dim fAttr As VbFileAttribute
    '
    On Error Resume Next
    fAttr = GetAttr(folderPath)
    If Err.Number = errBadFileNameOrNumber Then 'Unicode characters
        #If Mac Then

        #Else
            IsFolder = GetFSO().FolderExists(folderPath)
        #End If
    ElseIf Err.Number = 0 Then
        IsFolder = CBool(fAttr And vbDirectory)
    ElseIf Len(folderPath) > maxDirLen Then
        #If Mac Then

        #Else
            If Left$(folderPath, 4) = "\\?\" Then
                IsFolder = GetFSO().FolderExists(folderPath)
            ElseIf Left$(folderPath, 2) = "\\" Then
                IsFolder = GetFSO().FolderExists("\\?\UNC" & Mid$(folderPath, 2))
            Else
                IsFolder = GetFSO().FolderExists("\\?\" & folderPath)
            End If
        #End If
    End If
    On Error GoTo 0
End Function

'*******************************************************************************
'Checks if the contents of a folder can be edited
'*******************************************************************************
Public Function IsFolderEditable(ByRef folderPath As String) As Boolean
    Dim tempFolder As String
    Dim fixedPath As String: fixedPath = BuildPath(folderPath, vbNullString)
    '
    Do
        tempFolder = fixedPath & Rnd()
    Loop Until Not IsFolder(tempFolder)
    '
    On Error Resume Next
    MkDir tempFolder
    IsFolderEditable = (Err.Number = 0)
    If IsFolderEditable Then RmDir tempFolder
    On Error GoTo 0
End Function

'*******************************************************************************
'Moves (or renames) a file
'*******************************************************************************
Public Function MoveFile(ByRef sourcePath As String _
                       , ByRef destinationPath As String) As Boolean
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
Public Function MoveFolder(ByRef sourcePath As String _
                         , ByRef destinationPath As String) As Boolean
    If LenB(sourcePath) = 0 Then Exit Function
    If LenB(destinationPath) = 0 Then Exit Function
    If Not IsFolder(sourcePath) Then Exit Function
    If IsFolder(destinationPath) Then Exit Function
    '
    'The 'Name' statement can move a file across drives, but it can only rename
    '   a directory or folder within the same drive. Try 'Name' first
    On Error Resume Next
    Name sourcePath As destinationPath
    MoveFolder = (Err.Number = 0)
    If MoveFolder Then Exit Function
    On Error GoTo 0
    '
    'Try FSO if available
    #If Windows Then
        On Error Resume Next
        GetFSO().MoveFolder sourcePath, destinationPath
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
        CopyFolder destinationPath, sourcePath, ignoreFailedChildren:=True
        DeleteFolder destinationPath, True
        Exit Function
    End If
    '
    MoveFolder = True
End Function

'*******************************************************************************
'Returns the parent folder path for a given file or folder local path
'*******************************************************************************
Public Function ParentFolder(ByRef localPath As String) As String
    Const ps As String = PATH_SEPARATOR
    Dim fixedPath As String: fixedPath = FixPathSeparators(localPath)
    Dim i As Long
    '
    If Len(fixedPath) < 3 Then Exit Function
    i = InStrRev(fixedPath, ps, Len(fixedPath) - 1)
    If i < 2 Then Exit Function
    '
    If Mid$(fixedPath, i - 1, 1) <> ps Then
        ParentFolder = Left$(fixedPath, i - 1)
    End If
End Function

'*******************************************************************************
'Reads a file into an array of Bytes
'*******************************************************************************
Public Sub ReadBytes(ByRef filePath As String, ByRef result() As Byte)
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
        Get fileNumber, 1, result
    Else
        Erase result
    End If
    Close #fileNumber
End Sub

'*******************************************************************************
'Creates a text file used for diagnosing OneDrive logic issues
'*******************************************************************************
Private Sub CreateODDiagnosticsFile()
    Dim folderPath As String
    Do
        folderPath = BrowseForFolder(, "Choose target folder for diagnostics")
        If LenB(folderPath) = 0 Then Exit Sub
        If IsFolderEditable(folderPath) Then Exit Do
        MsgBox "Please choose a folder with write access"
    Loop
    '
    Const vbTwoNewLines As String = vbNewLine & vbNewLine
    Const fileName As String = "DiagnosticsOD.txt"
    Dim accountsInfo As ONEDRIVE_ACCOUNTS_INFO
    Dim fileNumber As Long: fileNumber = FreeFile()
    Dim filePath As String: filePath = BuildPath(folderPath, fileName)
    Dim res As String
    Dim i As Long
    Dim temp(0 To 2) As String
    '
    #If Mac Then
        temp(0) = "Mac"
    #Else
        temp(0) = "Win"
    #End If
    #If VBA7 Then
        temp(1) = "VBA7"
    #Else
        temp(1) = "VBA6"
    #End If
    #If Win64 Then
        temp(2) = "x64"
    #Else
        temp(2) = "x32"
    #End If
    res = Join(temp, " ") & vbTwoNewLines & String$(80, "-") & vbTwoNewLines
    '
    ReadODAccountsInfo accountsInfo
    For i = 1 To accountsInfo.pCount 'Check for unsynchronized accounts
        Dim j As Long
        For j = i + 1 To accountsInfo.pCount
            ValidateAccounts accountsInfo.arr(i), accountsInfo.arr(j)
        Next j
    Next i
    res = res & "Accounts found: " & accountsInfo.pCount & vbTwoNewLines
    '
    For i = 1 To accountsInfo.pCount
        With accountsInfo.arr(i)
            res = res & "Name: " & .accountName & vbNewLine
            res = res & "ID: " & .cID & vbNewLine
            res = res & "Has DAT: " & .hasDatFile & vbNewLine
            res = res & "Is Valid: " & .isValid & vbNewLine
        End With
        res = res & vbNewLine
    Next i
    res = res & String$(80, "-")
    res = res & vbTwoNewLines
    '
    ReadODProviders
    res = res & "Providers found: " & m_providers.pCount & vbTwoNewLines
    For i = 1 To m_providers.pCount
        With m_providers.arr(i)
            res = res & "Base Mount: " & .baseMount & vbNewLine
            res = res & "Is Business: " & .isBusiness & vbNewLine
            res = res & "Is Main: " & .isMain & vbNewLine
            res = res & "Mount Point: " & .mountPoint & vbNewLine
            res = res & "Sync ID: " & .syncID & vbNewLine
            res = res & "Web Path: " & .webPath & vbNewLine
            #If Mac Then
                res = res & "Sync Dir: " & .syncDir & vbNewLine
            #End If
        End With
        res = res & vbNewLine
    Next i
    '
    Open filePath For Output As #fileNumber
    Print #fileNumber, res
    Close #fileNumber
    '
    MsgBox "Created [" & fileName & "] diagnostics file", vbInformation
End Sub
