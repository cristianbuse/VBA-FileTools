# VBA-FileTools

FileTools is a small VBA library that is useful for interacting with the file system.

Supports **OneDrive/SharePoint/UNC** path conversion to **local** drive path (```GetLocalPath```) and viceversa (```GetRemotePath```) written in collaboration with Guido ([@guwidoe](https://github.com/guwidoe) / [GWD](https://stackoverflow.com/users/12287457/gwd)). See relevant [SO Answer](https://stackoverflow.com/a/73577057/8488913) written by Guido. Many thanks to him!

## Installation

Just import the following code module(s) in your VBA Project:

* [**LibFileTools.bas**](src/LibFileTools.bas)
* [**UDF_FileTools.bas**](src/UDF_FileTools.bas) (optional - works in MS Excel interface only, with exposed User Defined Functions)

## Usage

A couple of demoes saved in the [Demo](src/Demo/DemoLibFileTools.bas) module.

Public/Exposed methods:
 - BrowseForFiles           (Windows only)
 - BrowseForFolder          (Windows only)
 - BuildPath
 - ConvertText
 - CopyFile
 - CopyFolder
 - CreateFolder
 - DecodeURL
 - DeleteFile
 - DeleteFolder
 - FixFileName
 - FixPathSeparators
 - GetFileOwner             (Windows only)
 - GetFiles
 - GetFolders
 - GetKnownFolderCLSID      (Windows only)
 - GetKnownFolderPath       (Windows only)
 - GetLocalPath             (covers UNC/OneDrive/SharePoint paths)
 - GetMainBusinessURLs
 - GetRelativePath
 - GetRemotePath            (covers UNC/OneDrive/SharePoint paths)
 - GetSpecialFolderConstant (Mac only)
 - GetSpecialFolderDomain   (Mac only)
 - GetSpecialFolderPath     (Mac only)
 - IsFile
 - IsFolder
 - IsFolderEditable
 - MoveFile
 - MoveFolder
 - ParentFolder
 - ReadBytes
 
 Please note that ```GetLocalPath(path)``` can handle only unencoded URL paths. For encoded paths use ```GetLocalPath(DecodeURL(path))```!

## Notes
* No extra library references are needed (e.g. Microsoft Scripting Runtime)
* Works in any host Application (Excel, Word, AutoCAD etc.)
* Works on both Windows and Mac. On Mac, 3 of the methods are not available 
* Works in both x32 and x64 application environments

## License
MIT License

Copyright (c) 2012 Ion Cristian Buse

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
