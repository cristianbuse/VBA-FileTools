# VBA-FileTools

FileTools is a small VBA library that is useful for interacting with the file system.

## Installation

Just import the following code module in your VBA Project:

* [**LibFileTools.bas**](src/LibFileTools.bas)

## Testing

Section will follow soon

## Usage

Section will follow soon

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
