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
''    - BrowseForFiles      (Windows only)
''    - BrowseForFolder     (Windows only)
''    - BuildPath
''    - ConvertText
''    - CopyFile
''    - CopyFolder
''    - CreateFolder
''    - DeleteFile
''    - DeleteFolder
''    - FixFileName
''    - FixPathSeparators
''    - GetFileOwner        (Windows only)
''    - GetFiles
''    - GetFolders
''    - GetKnownFolderWin   (Windows only)
''    - GetLocalPath
''    - GetRelativePath
''    - GetRemotePath
''    - GetSpecialFolderMac (Mac only)
''    - IsFile
''    - IsFolder
''    - IsFolderEditable
''    - MoveFile
''    - MoveFolder
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
    Public Enum LongPtr
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

#If Mac Then 'Special folder constants for Mac
    'Source: https://developer.apple.com/library/archive/documentation/AppleScript/Conceptual/AppleScriptLangGuide/reference/ASLR_cmds.html
    Public Const SFC_ApplicationSupport    As String = "application support"
    Public Const SFC_ApplicationsFolder    As String = "applications folder"
    Public Const SFC_Desktop               As String = "desktop"
    Public Const SFC_DesktopPicturesFolder As String = "desktop pictures folder"
    Public Const SFC_DocumentsFolder       As String = "documents folder"
    Public Const SFC_DownloadsFolder       As String = "downloads folder"
    Public Const SFC_FavoritesFolder       As String = "favorites folder"
    Public Const SFC_FolderActionScripts   As String = "Folder Action scripts"
    Public Const SFC_Fonts                 As String = "fonts"
    Public Const SFC_Help                  As String = "help"
    Public Const SFC_HomeFolder            As String = "home folder"
    Public Const SFC_InternetPlugins       As String = "internet plugins"
    Public Const SFC_KeychainFolder        As String = "keychain folder"
    Public Const SFC_LibraryFolder         As String = "library folder"
    Public Const SFC_ModemScripts          As String = "modem scripts"
    Public Const SFC_MoviesFolder          As String = "movies folder"
    Public Const SFC_MusicFolder           As String = "music folder"
    Public Const SFC_PicturesFolder        As String = "pictures folder"
    Public Const SFC_Preferences           As String = "preferences"
    Public Const SFC_PrinterDescriptions   As String = "printer descriptions"
    Public Const SFC_PublicFolder          As String = "public folder"
    Public Const SFC_ScriptingAdditions    As String = "scripting additions"
    Public Const SFC_ScriptsFolder         As String = "scripts folder"
    Public Const SFC_ServicesFolder        As String = "services folder"
    Public Const SFC_SharedDocuments       As String = "shared documents"
    Public Const SFC_SharedLibraries       As String = "shared libraries"
    Public Const SFC_SitesFolder           As String = "sites folder"
    Public Const SFC_StartupDisk           As String = "startup disk"
    Public Const SFC_StartupItems          As String = "startup items"
    Public Const SFC_SystemFolder          As String = "system folder"
    Public Const SFC_SystemPreferences     As String = "system preferences"
    Public Const SFC_TemporaryItems        As String = "temporary items"
    Public Const SFC_Trash                 As String = "trash"
    Public Const SFC_UsersFolder           As String = "users folder"
    Public Const SFC_UtilitiesFolder       As String = "utilities folder"
    Public Const SFC_WorkflowsFolder       As String = "workflows folder"
                                      
    'Classic domain only
    Public Const SFC_AppleMenu             As String = "apple menu"
    Public Const SFC_ControlPanels         As String = "control panels"
    Public Const SFC_ControlStripModules   As String = "control strip modules"
    Public Const SFC_Extensions            As String = "extensions"
    Public Const SFC_LauncherItemsFolder   As String = "launcher items folder"
    Public Const SFC_PrinterDrivers        As String = "printer drivers"
    Public Const SFC_Printmonitor          As String = "printmonitor"
    Public Const SFC_ShutdownFolder        As String = "shutdown folder"
    Public Const SFC_SpeakableItems        As String = "speakable items"
    Public Const SFC_Stationery            As String = "stationery"
    Public Const SFC_Voices                As String = "voices"

    'The following domain names are valid:
    Public Const DOMAIN_System  As String = "system"
    Public Const DOMAIN_Local   As String = "local"
    Public Const DOMAIN_Network As String = "network"
    Public Const DOMAIN_User    As String = "user"
    Public Const DOMAIN_Classic As String = "classic"
#Else
    'List of standard KnownFolderIDs, declarations for VBA
    'Source: KnownFolders.h (Windows 11 SDK 10.0.22621.0) (sorted alphabetically)
    'Note: Most of the FOLDERIDs that are available on a specific device seem to
    '      be registered in the windows registry under
    '      HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions
    '      However, it seems that sometimes the SHGetKnownFolderPath function can
    '      process a FOLDERID even if not present in said registry location.
    Public Const FOLDERID_AccountPictures        As String = "{008ca0b1-55b4-4c56-b8a8-4de4b299d3be}"
    Public Const FOLDERID_AddNewPrograms         As String = "{de61d971-5ebc-4f02-a3a9-6c82895e5c04}"
    Public Const FOLDERID_AdminTools             As String = "{724EF170-A42D-4FEF-9F26-B60E846FBA4F}"
    Public Const FOLDERID_AllAppMods             As String = "{7ad67899-66af-43ba-9156-6aad42e6c596}"
    Public Const FOLDERID_AppCaptures            As String = "{EDC0FE71-98D8-4F4A-B920-C8DC133CB165}"
    Public Const FOLDERID_AppDataDesktop         As String = "{B2C5E279-7ADD-439F-B28C-C41FE1BBF672}"
    Public Const FOLDERID_AppDataDocuments       As String = "{7BE16610-1F7F-44AC-BFF0-83E15F2FFCA1}"
    Public Const FOLDERID_AppDataFavorites       As String = "{7CFBEFBC-DE1F-45AA-B843-A542AC536CC9}"
    Public Const FOLDERID_AppDataProgramData     As String = "{559D40A3-A036-40FA-AF61-84CB430A4D34}"
    Public Const FOLDERID_ApplicationShortcuts   As String = "{A3918781-E5F2-4890-B3D9-A7E54332328C}"
    Public Const FOLDERID_AppsFolder             As String = "{1e87508d-89c2-42f0-8a7e-645a0f50ca58}"
    Public Const FOLDERID_AppUpdates             As String = "{a305ce99-f527-492b-8b1a-7e76fa98d6e4}"
    Public Const FOLDERID_CameraRoll             As String = "{AB5FB87B-7CE2-4F83-915D-550846C9537B}"
    Public Const FOLDERID_CameraRollLibrary      As String = "{2B20DF75-1EDA-4039-8097-38798227D5B7}"
    Public Const FOLDERID_CDBurning              As String = "{9E52AB10-F80D-49DF-ACB8-4330F5687855}"
    Public Const FOLDERID_ChangeRemovePrograms   As String = "{df7266ac-9274-4867-8d55-3bd661de872d}"
    Public Const FOLDERID_CommonAdminTools       As String = "{D0384E7D-BAC3-4797-8F14-CBA229B392B5}"
    Public Const FOLDERID_CommonOEMLinks         As String = "{C1BAE2D0-10DF-4334-BEDD-7AA20B227A9D}"
    Public Const FOLDERID_CommonPrograms         As String = "{0139D44E-6AFE-49F2-8690-3DAFCAE6FFB8}"
    Public Const FOLDERID_CommonStartMenu        As String = "{A4115719-D62E-491D-AA7C-E74B8BE3B067}"
    Public Const FOLDERID_CommonStartMenuPlaces  As String = "{A440879F-87A0-4F7D-B700-0207B966194A}"
    Public Const FOLDERID_CommonStartup          As String = "{82A5EA35-D9CD-47C5-9629-E15D2F714E6E}"
    Public Const FOLDERID_CommonTemplates        As String = "{B94237E7-57AC-4347-9151-B08C6C32D1F7}"
    Public Const FOLDERID_ComputerFolder         As String = "{0AC0837C-BBF8-452A-850D-79D08E667CA7}"
    Public Const FOLDERID_ConflictFolder         As String = "{4bfefb45-347d-4006-a5be-ac0cb0567192}"
    Public Const FOLDERID_ConnectionsFolder      As String = "{6F0CD92B-2E97-45D1-88FF-B0D186B8DEDD}"
    Public Const FOLDERID_Contacts               As String = "{56784854-C6CB-462b-8169-88E350ACB882}"
    Public Const FOLDERID_ControlPanelFolder     As String = "{82A74AEB-AEB4-465C-A014-D097EE346D63}"
    Public Const FOLDERID_Cookies                As String = "{2B0F765D-C0E9-4171-908E-08A611B84FF6}"
    Public Const FOLDERID_CurrentAppMods         As String = "{3db40b20-2a30-4dbe-917e-771dd21dd099}"
    Public Const FOLDERID_Desktop                As String = "{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}"
    Public Const FOLDERID_DevelopmentFiles       As String = "{DBE8E08E-3053-4BBC-B183-2A7B2B191E59}"
    Public Const FOLDERID_Device                 As String = "{1C2AC1DC-4358-4B6C-9733-AF21156576F0}"
    Public Const FOLDERID_DeviceMetadataStore    As String = "{5CE4A5E9-E4EB-479D-B89F-130C02886155}"
    Public Const FOLDERID_Documents              As String = "{FDD39AD0-238F-46AF-ADB4-6C85480369C7}"
    Public Const FOLDERID_DocumentsLibrary       As String = "{7b0db17d-9cd2-4a93-9733-46cc89022e7c}"
    Public Const FOLDERID_Downloads              As String = "{374DE290-123F-4565-9164-39C4925E467B}"
    Public Const FOLDERID_Favorites              As String = "{1777F761-68AD-4D8A-87BD-30B759FA33DD}"
    Public Const FOLDERID_Fonts                  As String = "{FD228CB7-AE11-4AE3-864C-16F3910AB8FE}"
    Public Const FOLDERID_Games                  As String = "{CAC52C1A-B53D-4edc-92D7-6B2E8AC19434}"
    Public Const FOLDERID_GameTasks              As String = "{054FAE61-4DD8-4787-80B6-090220C4B700}"
    Public Const FOLDERID_History                As String = "{D9DC8A3B-B784-432E-A781-5A1130A75963}"
    Public Const FOLDERID_HomeGroup              As String = "{52528A6B-B9E3-4add-B60D-588C2DBA842D}"
    Public Const FOLDERID_HomeGroupCurrentUser   As String = "{9B74B6A3-0DFD-4f11-9E78-5F7800F2E772}"
    Public Const FOLDERID_ImplicitAppShortcuts   As String = "{bcb5256f-79f6-4cee-b725-dc34e402fd46}"
    Public Const FOLDERID_InternetCache          As String = "{352481E8-33BE-4251-BA85-6007CAEDCF9D}"
    Public Const FOLDERID_InternetFolder         As String = "{4D9F7874-4E0C-4904-967B-40B0D20C3E4B}"
    Public Const FOLDERID_Libraries              As String = "{1B3EA5DC-B587-4786-B4EF-BD1DC332AEAE}"
    Public Const FOLDERID_Links                  As String = "{bfb9d5e0-c6a9-404c-b2b2-ae6db6af4968}"
    Public Const FOLDERID_LocalAppData           As String = "{F1B32785-6FBA-4FCF-9D55-7B8E7F157091}"
    Public Const FOLDERID_LocalAppDataLow        As String = "{A520A1A4-1780-4FF6-BD18-167343C5AF16}"
    Public Const FOLDERID_LocalDocuments         As String = "{f42ee2d3-909f-4907-8871-4c22fc0bf756}"
    Public Const FOLDERID_LocalDownloads         As String = "{7d83ee9b-2244-4e70-b1f5-5393042af1e4}"
    Public Const FOLDERID_LocalizedResourcesDir  As String = "{2A00375E-224C-49DE-B8D1-440DF7EF3DDC}"
    Public Const FOLDERID_LocalMusic             As String = "{a0c69a99-21c8-4671-8703-7934162fcf1d}"
    Public Const FOLDERID_LocalPictures          As String = "{0ddd015d-b06c-45d5-8c4c-f59713854639}"
    Public Const FOLDERID_LocalStorage           As String = "{B3EB08D3-A1F3-496B-865A-42B536CDA0EC}"
    Public Const FOLDERID_LocalVideos            As String = "{35286a68-3c57-41a1-bbb1-0eae73d76c95}"
    Public Const FOLDERID_Music                  As String = "{4BD8D571-6D19-48D3-BE97-422220080E43}"
    Public Const FOLDERID_MusicLibrary           As String = "{2112AB0A-C86A-4ffe-A368-0DE96E47012E}"
    Public Const FOLDERID_NetHood                As String = "{C5ABBF53-E17F-4121-8900-86626FC2C973}"
    Public Const FOLDERID_NetworkFolder          As String = "{D20BEEC4-5CA8-4905-AE3B-BF251EA09B53}"
    Public Const FOLDERID_Objects3D              As String = "{31C0DD25-9439-4F12-BF41-7FF4EDA38722}"
    Public Const FOLDERID_OneDrive               As String = "{A52BBA46-E9E1-435f-B3D9-28DAA648C0F6}"
    Public Const FOLDERID_OriginalImages         As String = "{2C36C0AA-5812-4b87-BFD0-4CD0DFB19B39}"
    Public Const FOLDERID_PhotoAlbums            As String = "{69D2CF90-FC33-4FB7-9A0C-EBB0F0FCB43C}"
    Public Const FOLDERID_Pictures               As String = "{33E28130-4E1E-4676-835A-98395C3BC3BB}"
    Public Const FOLDERID_PicturesLibrary        As String = "{A990AE9F-A03B-4e80-94BC-9912D7504104}"
    Public Const FOLDERID_Playlists              As String = "{DE92C1C7-837F-4F69-A3BB-86E631204A23}"
    Public Const FOLDERID_PrintersFolder         As String = "{76FC4E2D-D6AD-4519-A663-37BD56068185}"
    Public Const FOLDERID_PrintHood              As String = "{9274BD8D-CFD1-41C3-B35E-B13F55A758F4}"
    Public Const FOLDERID_Profile                As String = "{5E6C858F-0E22-4760-9AFE-EA3317B67173}"
    Public Const FOLDERID_ProgramData            As String = "{62AB5D82-FDC1-4DC3-A9DD-070D1D495D97}"
    Public Const FOLDERID_ProgramFiles           As String = "{905e63b6-c1bf-494e-b29c-65b732d3d21a}"
    Public Const FOLDERID_ProgramFilesCommon     As String = "{F7F1ED05-9F6D-47A2-AAAE-29D317C6F066}"
    Public Const FOLDERID_ProgramFilesCommonX64  As String = "{6365D5A7-0F0D-45e5-87F6-0DA56B6A4F7D}"
    Public Const FOLDERID_ProgramFilesCommonX86  As String = "{DE974D24-D9C6-4D3E-BF91-F4455120B917}"
    Public Const FOLDERID_ProgramFilesX64        As String = "{6D809377-6AF0-444b-8957-A3773F02200E}"
    Public Const FOLDERID_ProgramFilesX86        As String = "{7C5A40EF-A0FB-4BFC-874A-C0F2E0B9FA8E}"
    Public Const FOLDERID_Programs               As String = "{A77F5D77-2E2B-44C3-A6A2-ABA601054A51}"
    Public Const FOLDERID_Public                 As String = "{DFDF76A2-C82A-4D63-906A-5644AC457385}"
    Public Const FOLDERID_PublicDesktop          As String = "{C4AA340D-F20F-4863-AFEF-F87EF2E6BA25}"
    Public Const FOLDERID_PublicDocuments        As String = "{ED4824AF-DCE4-45A8-81E2-FC7965083634}"
    Public Const FOLDERID_PublicDownloads        As String = "{3D644C9B-1FB8-4f30-9B45-F670235F79C0}"
    Public Const FOLDERID_PublicGameTasks        As String = "{DEBF2536-E1A8-4c59-B6A2-414586476AEA}"
    Public Const FOLDERID_PublicLibraries        As String = "{48daf80b-e6cf-4f4e-b800-0e69d84ee384}"
    Public Const FOLDERID_PublicMusic            As String = "{3214FAB5-9757-4298-BB61-92A9DEAA44FF}"
    Public Const FOLDERID_PublicPictures         As String = "{B6EBFB86-6907-413C-9AF7-4FC2ABF07CC5}"
    Public Const FOLDERID_PublicRingtones        As String = "{E555AB60-153B-4D17-9F04-A5FE99FC15EC}"
    Public Const FOLDERID_PublicUserTiles        As String = "{0482af6c-08f1-4c34-8c90-e17ec98b1e17}"
    Public Const FOLDERID_PublicVideos           As String = "{2400183A-6185-49FB-A2D8-4A392A602BA3}"
    Public Const FOLDERID_QuickLaunch            As String = "{52a4f021-7b75-48a9-9f6b-4b87a210bc8f}"
    Public Const FOLDERID_Recent                 As String = "{AE50C081-EBD2-438A-8655-8A092E34987A}"
    Public Const FOLDERID_RecordedCalls          As String = "{2f8b40c2-83ed-48ee-b383-a1f157ec6f9a}"
    Public Const FOLDERID_RecordedTVLibrary      As String = "{1A6FDBA2-F42D-4358-A798-B74D745926C5}"
    Public Const FOLDERID_RecycleBinFolder       As String = "{B7534046-3ECB-4C18-BE4E-64CD4CB7D6AC}"
    Public Const FOLDERID_ResourceDir            As String = "{8AD10C31-2ADB-4296-A8F7-E4701232C972}"
    Public Const FOLDERID_RetailDemo             As String = "{12D4C69E-24AD-4923-BE19-31321C43A767}"
    Public Const FOLDERID_Ringtones              As String = "{C870044B-F49E-4126-A9C3-B52A1FF411E8}"
    Public Const FOLDERID_RoamedTileImages       As String = "{AAA8D5A5-F1D6-4259-BAA8-78E7EF60835E}"
    Public Const FOLDERID_RoamingAppData         As String = "{3EB685DB-65F9-4CF6-A03A-E3EF65729F3D}"
    Public Const FOLDERID_RoamingTiles           As String = "{00BCFC5A-ED94-4e48-96A1-3F6217F21990}"
    Public Const FOLDERID_SampleMusic            As String = "{B250C668-F57D-4EE1-A63C-290EE7D1AA1F}"
    Public Const FOLDERID_SamplePictures         As String = "{C4900540-2379-4C75-844B-64E6FAF8716B}"
    Public Const FOLDERID_SamplePlaylists        As String = "{15CA69B3-30EE-49C1-ACE1-6B5EC372AFB5}"
    Public Const FOLDERID_SampleVideos           As String = "{859EAD94-2E85-48AD-A71A-0969CB56A6CD}"
    Public Const FOLDERID_SavedGames             As String = "{4C5C32FF-BB9D-43b0-B5B4-2D72E54EAAA4}"
    Public Const FOLDERID_SavedPictures          As String = "{3B193882-D3AD-4eab-965A-69829D1FB59F}"
    Public Const FOLDERID_SavedPicturesLibrary   As String = "{E25B5812-BE88-4bd9-94B0-29233477B6C3}"
    Public Const FOLDERID_SavedSearches          As String = "{7d1d3a04-debb-4115-95cf-2f29da2920da}"
    Public Const FOLDERID_Screenshots            As String = "{b7bede81-df94-4682-a7d8-57a52620b86f}"
    Public Const FOLDERID_SEARCH_CSC             As String = "{ee32e446-31ca-4aba-814f-a5ebd2fd6d5e}"
    Public Const FOLDERID_SEARCH_MAPI            As String = "{98ec0e18-2098-4d44-8644-66979315a281}"
    Public Const FOLDERID_SearchHistory          As String = "{0D4C3DB6-03A3-462F-A0E6-08924C41B5D4}"
    Public Const FOLDERID_SearchHome             As String = "{190337d1-b8ca-4121-a639-6d472d16972a}"
    Public Const FOLDERID_SearchTemplates        As String = "{7E636BFE-DFA9-4D5E-B456-D7B39851D8A9}"
    Public Const FOLDERID_SendTo                 As String = "{8983036C-27C0-404B-8F08-102D10DCFD74}"
    Public Const FOLDERID_SidebarDefaultParts    As String = "{7B396E54-9EC5-4300-BE0A-2482EBAE1A26}"
    Public Const FOLDERID_SidebarParts           As String = "{A75D362E-50FC-4fb7-AC2C-A8BEAA314493}"
    Public Const FOLDERID_SkyDrive               As String = "{A52BBA46-E9E1-435f-B3D9-28DAA648C0F6}"
    Public Const FOLDERID_SkyDriveCameraRoll     As String = "{767E6811-49CB-4273-87C2-20F355E1085B}"
    Public Const FOLDERID_SkyDriveDocuments      As String = "{24D89E24-2F19-4534-9DDE-6A6671FBB8FE}"
    Public Const FOLDERID_SkyDriveMusic          As String = "{C3F2459E-80D6-45DC-BFEF-1F769F2BE730}"
    Public Const FOLDERID_SkyDrivePictures       As String = "{339719B5-8C47-4894-94C2-D8F77ADD44A6}"
    Public Const FOLDERID_StartMenu              As String = "{625B53C3-AB48-4EC1-BA1F-A1EF4146FC19}"
    Public Const FOLDERID_StartMenuAllPrograms   As String = "{F26305EF-6948-40B9-B255-81453D09C785}"
    Public Const FOLDERID_Startup                As String = "{B97D20BB-F46A-4C97-BA10-5E3608430854}"
    Public Const FOLDERID_SyncManagerFolder      As String = "{43668BF8-C14E-49B2-97C9-747784D784B7}"
    Public Const FOLDERID_SyncResultsFolder      As String = "{289a9a43-be44-4057-a41b-587a76d7e7f9}"
    Public Const FOLDERID_SyncSetupFolder        As String = "{0F214138-B1D3-4a90-BBA9-27CBC0C5389A}"
    Public Const FOLDERID_System                 As String = "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}"
    Public Const FOLDERID_SystemX86              As String = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}"
    Public Const FOLDERID_Templates              As String = "{A63293E8-664E-48DB-A079-DF759E0509F7}"
    Public Const FOLDERID_UserPinned             As String = "{9e3995ab-1f9c-4f13-b827-48b24b6c7174}"
    Public Const FOLDERID_UserProfiles           As String = "{0762D272-C50A-4BB0-A382-697DCD729B80}"
    Public Const FOLDERID_UserProgramFiles       As String = "{5cd7aee2-2219-4a67-b85d-6c9ce15660cb}"
    Public Const FOLDERID_UserProgramFilesCommon As String = "{bcbd3057-ca5c-4622-b42d-bc56db0ae516}"
    Public Const FOLDERID_UsersFiles             As String = "{f3ce0f7c-4901-4acc-8648-d5d44b04ef8f}"
    Public Const FOLDERID_UsersLibraries         As String = "{A302545D-DEFF-464b-ABE8-61C8648D939B}"
    Public Const FOLDERID_Videos                 As String = "{18989B1D-99B5-455B-841C-AB7C74E4DDFC}"
    Public Const FOLDERID_VideosLibrary          As String = "{491E922F-5643-4af4-A7EB-4E7A138D8174}"
    Public Const FOLDERID_Windows                As String = "{F38BF404-1D43-42F2-9305-67DE0B28FC23}"
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
#End If

Private Type ONEDRIVE_PROVIDER
    webPath As String
    mountPoint As String
    isBusiness As Boolean
    isMain As Boolean
    accountIndex As Long
    baseMount As String
    syncID As String
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

#If Mac Then
    Public Const PATH_SEPARATOR = "/"
#Else
    Public Const PATH_SEPARATOR = "\"
#End If

Private Const vbErrInvalidProcedureCall   As Long = 5
Private Const vbErrInternalError          As Long = 51
Private Const vbErrPathFileAccessError    As Long = 75
Private Const vbErrPathNotFound           As Long = 76
Private Const vbErrComponentNotRegistered As Long = 336

Private m_providers As ONEDRIVE_PROVIDERS
#If Mac Then
    Private m_conversionDescriptors As New Collection
#End If

'*******************************************************************************
'Returns a Collection of file paths by using a FilePicker FileDialog
'*******************************************************************************
#If Mac Then
    'Not implemented
    'Seems achievable via script:
    '   - https://stackoverflow.com/a/15546518/8488913
    '   - https://stackoverflow.com/a/37411960/8488913
#Else
Public Function BrowseForFiles(Optional ByRef initialPath As String _
                             , Optional ByRef dialogTitle As String _
                             , Optional ByRef filterDesc As String _
                             , Optional ByRef filterList As String _
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
                Case "Microsoft Excel": .InitialFileName = GetLocalPath(app.ThisWorkbook.Path, , True)
                Case "Microsoft Word":  .InitialFileName = GetLocalPath(app.ThisDocument.Path, , True)
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
    SetAttr filePath, vbNormal 'Too costly to do after failing Delete
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
        If i = 0 Then
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
'Returns path of a 'known folder' using the respective 'FOLDERID' on Windows
'Use prefixed constants 'FOLDERID_' for the 'knownFolderID' argument
'If 'createIfMissing' is set to True, the windows API function will be called
'   with flags 'KF_FLAG_CREATE' and 'KF_FLAG_INIT' and will create the folder
'   if it does not currently exist on the system.
'The function can raise the following errors:
'   -   5: (Invalid procedure call) if 'knownFolderID' is not a valid CLSID
'   -  76: (Path not found) if 'createIfMissing' = False AND path not found
'   -  75: (Path/File access error) if path not found because either:
'          * the specified FOLDERID is for a known virtual folder
'          * there are insufficient permissions to create the folder
'   - 336: (Component not correctly registered) if the path, or the known
'          FOLDERID itself are not registered in the windows registry
'   -  51: (Internal error) if an unexpected error occurs
'*******************************************************************************
#If Windows Then
Public Function GetKnownFolderWin(ByRef knownFolderID As String, _
                         Optional ByVal createIfMissing As Boolean = False) As String
    Const methodName As String = "GetKnownFolderWin"
    Const NOERROR As Long = 0
    Dim rfID As GUID
    '
    If CLSIDFromString(StrPtr(knownFolderID), rfID) <> NOERROR Then
        Err.Raise vbErrInvalidProcedureCall, methodName, "Invalid CLSID"
    End If
    '
    Const KF_FLAG_CREATE As Long = &H8000&  'Other flags not relevant
    Const KF_FLAG_INIT   As Long = &H800&
    Const flagCreateInit As Long = KF_FLAG_CREATE Or KF_FLAG_INIT
    Dim dwFlags As Long: If createIfMissing Then dwFlags = flagCreateInit
    '
    Const S_OK As Long = 0
    Dim ppszPath As LongPtr
    Dim hRes As Long: hRes = SHGetKnownFolderPath(rfID, dwFlags, 0, ppszPath)
    '
    If hRes = S_OK Then
        GetKnownFolderWin = Space$(lstrlenW(ppszPath))
        CopyMemory StrPtr(GetKnownFolderWin), ppszPath, LenB(GetKnownFolderWin)
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
    Const currentFolder As String = "."
    Const parentFolder As String = ".."
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
        If folderName <> currentFolder And folderName <> parentFolder Then
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
    Dim diff As Long
    Dim isRFile As Boolean
    '
    Do
        prevPos = currPos
        currPos = InStr(currPos + 1, fPath, ps)
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

#If Mac Then
'*******************************************************************************
'Gets path of a 'special folder' using the respective 'folder constant' on Mac
'Use prefixed constants 'SFC_' and 'DOMAIN_' for the first 2 arguments
'If 'createIfMissing' is set to True, the function will try to create the folder
'   if it does not currently exist on the system. Note that this argument
'   ignores the 'forceNonSandboxedPath' option, and it can happen that the
'   folder gets created in the sandboxed location and the function returns a non
'   sandboxed path. This behavior can not be avoided without creating access
'   requests, therefore it should be taken into account by the user
'The function can raise the following errors:
'   -  5: (Invalid procedure call) if 'domainName' is invalid
'   - 76: (Path not found) if 'createIfMissing' = False AND path not found
'   - 75: (Path/File access error) if 'createIfMissing'= True AND path not found
'*******************************************************************************
Public Function GetSpecialFolderMac(ByRef specialFolderConstant As String _
                                  , Optional ByRef domainName As String = vbNullString _
                                  , Optional ByVal forceNonSandboxedPath As Boolean = True _
                                  , Optional ByVal createIfMissing As Boolean = False) As String
    Const methodName As String = "GetSpecialFolderMac"
    '
    Select Case LCase$(domainName)
    Case "system", "local", "network", "user", "classic", vbNullString
    Case Else
        Err.Raise vbErrInvalidProcedureCall, methodName, "Invalid domain name" _
                & ". Expected one of: system, local, network, user, classic"
    End Select
    '
    Dim cmd As String: cmd = specialFolderConstant
    '
    If LenB(domainName) > 0 Then cmd = cmd & " from " & domainName & " domain"
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
    GetSpecialFolderMac = MacScript(cmd)
    On Error GoTo 0
    '
    If forceNonSandboxedPath Then
        Dim sboxPath As String:    sboxPath = Environ$("HOME")
        Dim i As Long:             i = InStrRev(sboxPath, "/Library/Containers/")
        Dim sboxRelPath As String: If i > 0 Then sboxRelPath = Mid$(sboxPath, i)
        GetSpecialFolderMac = Replace(GetSpecialFolderMac, sboxRelPath _
                                    , vbNullString, , 1, vbTextCompare)
    End If
    If LenB(GetSpecialFolderMac) > 0 Then Exit Function
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

'*******************************************************************************
'Returns the local path for a OneDrive web path
'Returns null string if the path provided is not a valid OneDrive web path
'
'With the help of: @guwidoe (https://github.com/guwidoe)
'See: https://github.com/cristianbuse/VBA-FileTools/issues/1
'*******************************************************************************
Private Function GetOneDriveLocalPath(ByRef odWebPath As String _
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
                    Dim syncDir As String: syncDir = collSyncIDToDir(.syncID)
                    .mountPoint = Replace(.mountPoint, .baseMount, syncDir)
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
    Const businessMask As String = "????????-????-????-????-????????????"
    Const personalMask As String = "????????????????"
    Const ps As String = PATH_SEPARATOR
    Dim folderPath As Variant
    Dim i As Long
    Dim mask As String
    Dim iniName As String
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
            If .isPersonal Then
                mask = personalMask
            Else
                mask = businessMask
                .accountIndex = CLng(Right$(.accountName, 1))
            End If
            iniName = Dir(BuildPath(.folderPath, mask & ".ini"))
            If LenB(iniName) > 0 And iniName Like mask & ".ini" Then
                .cID = Left$(iniName, Len(iniName) - 4)
                .datPath = .folderPath & ps & .cID & ".dat"
                .dbPath = .folderPath & ps & "SyncEngineDatabase.db"
                .groupPath = .folderPath & ps & "GroupFolders.ini"
                .iniPath = .folderPath & ps & iniName
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
    Dim temp() As String
    Dim tempMount As String
    Dim mainMount As String
    Dim syncID As String
    Dim mainSyncID As String
    Dim tempURL As String
    Dim cSignature As String
    Dim cDirs As Collection
    Dim cParents As Collection
    Dim cPending As New Collection
    Dim canAdd As Boolean
    '
    #If Mac Then
        iniText = ConvertText(iniText, codeUTF16LE, codeUTF8, True)
    #End If
    For Each lineText In Split(iniText, vbNewLine)
        Dim parts() As String: parts = Split(lineText, """")
        Select Case Left$(lineText, InStr(1, lineText, " "))
        Case "libraryScope "
            tempMount = parts(9)
            syncID = Split(parts(10), " ")(2)
            canAdd = (LenB(tempMount) > 0)
            If parts(3) = "ODB" Then
                mainMount = tempMount
                mainSyncID = syncID
                tempURL = GetUrlNamespace(aInfo.clientPath)
            Else
                temp = Split(parts(8), " ")
                cSignature = "_" & temp(3) & temp(1)
                tempURL = GetUrlNamespace(aInfo.clientPath, cSignature)
            End If
            cPending.Add tempURL, Split(parts(0), " ")(2)
        Case "libraryFolder "
            If cDirs Is Nothing Then Set cDirs = GetODDirs(aInfo, cParents)
            tempMount = parts(1)
            temp = Split(parts(0), " ")
            tempURL = cPending(temp(3))
            syncID = Split(parts(4), " ")(1)
            Dim tempID As String:     tempID = temp(4)
            Dim tempFolder As String: tempFolder = vbNullString
            If aInfo.hasDatFile Then tempID = Split(tempID, "+")(0)
            On Error Resume Next
            Do
                tempFolder = cDirs(tempID) & "/" & tempFolder
                tempID = cParents(tempID)
            Loop Until Err.Number <> 0
            On Error GoTo 0
            canAdd = (LenB(tempFolder) > 0)
            tempURL = tempURL & tempFolder
        Case "AddedScope "
            If cDirs Is Nothing Then Set cDirs = GetODDirs(aInfo, cParents)
            tempID = Split(parts(0), " ")(3)
            tempFolder = vbNullString
            On Error Resume Next
            Do
                tempFolder = cDirs(tempID) & PATH_SEPARATOR & tempFolder
                tempID = cParents(tempID)
            Loop Until Err.Number <> 0
            On Error GoTo 0
            tempMount = mainMount & PATH_SEPARATOR & tempFolder
            syncID = mainSyncID
            tempURL = parts(5)
            If tempURL = " " Or LenB(tempURL) = 0 Then
                tempURL = vbNullString
            Else
                tempURL = tempURL & "/"
            End If
            temp = Split(parts(4), " ")
            cSignature = "_" & temp(3) & temp(1) & temp(4)
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
    i = InStr(1, fText, vTag) + Len(vTag)
    GetTagValue = Mid$(fText, i, InStr(i, fText, vbNewLine) - i)
End Function

'*******************************************************************************
'Adds all providers for a Personal OneDrive account
'*******************************************************************************
Private Sub AddPersonalProviders(ByRef aInfo As ONEDRIVE_ACCOUNT_INFO)
    Dim mainURL As String:    mainURL = GetUrlNamespace(aInfo.clientPath) & "/"
    Dim libText As String:    libText = GetTagValue(aInfo.iniPath, "library = ")
    Dim libParts() As String: libParts = Split(libText, """")
    Dim mainMount As String:  mainMount = libParts(3)
    Dim bytes() As Byte:      ReadBytes aInfo.groupPath, bytes
    Dim groupText As String:  groupText = bytes
    Dim syncID As String:     syncID = Split(libParts(4), " ")(2)
    Dim lineText As Variant
    Dim cID As String
    Dim i As Long
    Dim relPath As String
    Dim folderID As String
    Dim cDirs As Collection
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
                If cDirs Is Nothing Then Set cDirs = GetODDirs(aInfo)
                With m_providers.arr(AddProvider())
                    .webPath = mainURL & cID & "/" & relPath & "/"
                    .mountPoint = BuildPath(mainMount, cDirs(folderID) & "/")
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
Private Function GetODDirs(ByRef aInfo As ONEDRIVE_ACCOUNT_INFO _
                         , Optional ByRef outParents As Collection) As Collection
    If aInfo.hasDatFile Then
        Set GetODDirs = GetODDirsFromDat(aInfo.datPath, outParents)
    End If
    If GetODDirs Is Nothing Then
        Set GetODDirs = GetODDirsFromDB(aInfo.dbPath, outParents)
    End If
End Function

'*******************************************************************************
'Utility - Retrieves all folders from an OneDrive user dat file
'*******************************************************************************
Private Function GetODDirsFromDat(ByRef filePath As String _
                                , ByRef outParents As Collection) As Collection
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
    Dim cDirs As Collection
    Dim lastFileChange As Date
    Dim currFileChange As Date
    Dim stepSize As Long
    Dim bytes As Long
    Dim folderID As String
    Dim parentID As String
    Dim folderName As String
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
                Set cDirs = New Collection
                Set outParents = New Collection
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
                    folderID = StrConv(MidB$(s, i, bytes), vbUnicode)
                    '
                    i = i + idSize
                    bytes = Clamp(InStrB(i, s, vbNullByte) - i, 0, idSize)
                    parentID = StrConv(MidB$(s, i, bytes), vbUnicode)
                    '
                    i = i + fNameOffset
                    If folderID Like idPattern And parentID Like idPattern Then
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
                        folderName = MidB$(s, i, bytes)
                        #If Mac Then
                            folderName = ConvertText(folderName, codeUTF16LE _
                                                   , codeUTF32LE, True)
                        #End If
                        cDirs.Add folderName, folderID
                        outParents.Add parentID, folderID
                    End If
                End If
                i = InStrB(i + 1, s, hCheck)
            Loop
            lastRecord = lastRecord + chunkSize - stepSize
            If i > stepSize Then
                lastRecord = lastRecord - chunkSize + (i \ 2) * 2
            End If
        Loop Until lastRecord > size
        If cDirs.Count > 0 Then Exit For
    Next stepSize
    Set GetODDirsFromDat = cDirs
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
'Utility - Retrieves all folders from an OneDrive user database file
'*******************************************************************************
Private Function GetODDirsFromDB(ByRef filePath As String _
                               , ByRef outParents As Collection) As Collection
    If Not IsFile(filePath) Then Exit Function
    Dim fileNumber As Long: fileNumber = FreeFile()
    '
    Open filePath For Binary Access Read As #fileNumber
    Dim size As Long: size = LOF(fileNumber)
    If size = 0 Then GoTo CloseFile
    '                             __    ____
    'Signature bytes: 0b0b0b0b0b0b080b0b08080b0b0b0b where b>=0, b <= 9
    Dim sig88 As String: sig88 = ChrW$(&H808)
    Const sig8 As Long = 8
    Const sig8Offset As Long = -3
    Const maxSigByte As Byte = 9
    Const sig88ToDataOffset As Long = 6 'Data comes after the signature
    Const headBytes6 As Long = &H16
    Const headBytes5 As Long = &H15
    Const headBytes6Offset As Long = -16 'Header comes before the signature
    Const headBytes5Offset As Long = -15
    Const chunkSize As Long = &H100000 '1MB
    '
    Dim b(1 To chunkSize) As Byte
    Dim s As String
    Dim lastRecord As Long
    Dim idPattern As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim cDirs As Collection
    Dim bytes As Long
    Dim idSize(1 To 4) As Byte
    Dim nameSize As Long
    Dim folderID As String
    Dim lastFolderID As String
    Dim parentID As String
    Dim lastParentID As String
    Dim folderName As String
    Dim lastFolderName As String
    Dim currDataEnd As Long
    Dim lastDataEnd As Long
    Dim headByte As Byte
    Dim lastHeadByte
    Dim has5HeadBytes As Boolean
    Dim heads As New Collection
    '
    idPattern = Replace(Space$(12), " ", "[a-fA-F0-9]") & "*"
    Do
        Dim currFileChange As Date: currFileChange = FileDateTime(filePath)
        Dim lastFileChange As Date
        '
        i = 0
        If currFileChange > lastFileChange Then
            Set cDirs = New Collection
            Set outParents = New Collection
            lastFileChange = currFileChange
            lastRecord = 1
        End If
        Get fileNumber, lastRecord, b
        s = b
        i = InStrB(1 - headBytes6Offset, s, sig88)
        lastDataEnd = 0
        Do While i > 0
            If i + headBytes6Offset - 2 > lastDataEnd And LenB(lastFolderID) > 0 Then
                On Error Resume Next 'Ignore duplicates
                cDirs.Add lastFolderName, lastFolderID
                If Err.Number <> 0 Then
                    If cDirs(lastFolderID) <> lastFolderName _
                    Or outParents(lastFolderID) <> lastParentID Then
                        If heads(lastFolderID) < lastHeadByte Then
                            cDirs.Remove lastFolderID
                            outParents.Remove lastFolderID
                            heads.Remove lastFolderID
                            cDirs.Add lastFolderName, lastFolderID
                        End If
                    End If
                End If
                outParents.Add lastParentID, lastFolderID
                heads.Add lastHeadByte, lastFolderID
                On Error GoTo 0
                lastFolderID = vbNullString
            End If
            '
            If b(i + sig8Offset) <> sig8 Then GoTo NextSig
            has5HeadBytes = True
            If b(i + headBytes5Offset) = headBytes5 Then
                j = i + headBytes5Offset
            ElseIf b(i + headBytes6Offset) = headBytes6 Then
                j = i + headBytes6Offset
                has5HeadBytes = False 'Has 6 bytes header
            ElseIf b(i + headBytes5Offset) <= maxSigByte Then
                j = i + headBytes5Offset
            Else
                GoTo NextSig
            End If
            headByte = b(j)
            '
            bytes = sig88ToDataOffset
            For k = 1 To 4
                If k = 1 And headByte <= maxSigByte Then
                    idSize(k) = b(j + 2) 'Ignore first header byte
                Else
                    idSize(k) = b(j + k)
                End If
                If idSize(k) < 37 Or idSize(k) Mod 2 = 0 Then GoTo NextSig
                idSize(k) = (idSize(k) - 13) / 2
                bytes = bytes + idSize(k)
            Next k
            If has5HeadBytes Then
                nameSize = b(j + 5)
                If nameSize < 15 Or nameSize Mod 2 = 0 Then GoTo NextSig
                nameSize = (nameSize - 13) / 2
            Else
                nameSize = (b(j + 5) - 128) * 64 + (b(j + 6) - 13) / 2
                If nameSize < 1 Or b(j + 6) Mod 2 = 0 Then GoTo NextSig
            End If
            bytes = bytes + nameSize
            '
            currDataEnd = i + bytes - 1
            If currDataEnd > chunkSize Then 'Next chunk
                i = i - 1
                Exit Do
            End If
            j = i + sig88ToDataOffset
            folderID = StrConv(MidB$(s, j, idSize(1)), vbUnicode)
            j = j + idSize(1)
            parentID = StrConv(MidB$(s, j, idSize(2)), vbUnicode)
            '
            If folderID Like idPattern And parentID Like idPattern Then
                j = j + idSize(2) + idSize(3) + idSize(4)
                folderName = MidB$(s, j, nameSize)
                '
                Dim isASCII As Boolean: isASCII = True
                For j = j To j + nameSize - 1
                    If b(j) And &H80& Then
                        isASCII = False
                        Exit For
                    End If
                Next j
                If isASCII Then
                    folderName = StrConv(folderName, vbUnicode)
                Else
                    folderName = ConvertText(folderName, codeUTF16LE, codeUTF8)
                End If
                '
                lastFolderID = folderID
                lastParentID = parentID
                lastFolderName = folderName
                lastHeadByte = headByte
                lastDataEnd = currDataEnd
            End If
NextSig:
            i = InStrB(i + 1, s, sig88)
        Loop
        If i = 0 Then
            lastRecord = lastRecord + chunkSize + headBytes6Offset
        Else
            lastRecord = lastRecord + i + headBytes6Offset
        End If
    Loop Until lastRecord > size
    Set GetODDirsFromDB = cDirs
CloseFile:
    Close #fileNumber
End Function

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
    If Err.Number = 0 Then RmDir tempFolder
    IsFolderEditable = (Err.Number = 0)
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
    If Err.Number = 0 Then
        MoveFolder = True
        Exit Function
    End If
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
