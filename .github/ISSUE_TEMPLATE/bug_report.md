---
name: Bug report
about: Create a report to help us improve
title: Bug
labels: bug
assignees: ''

---

**Describe the bug**
A clear and concise description of what the bug is.

**To Reproduce**
Steps to reproduce the behavior:
1. Go to '...'
2. Click on '....'
3. Scroll down to '....'
4. ...
5. See error

**Expected behavior**
A clear and concise description of what you expected to happen.

**Screenshots**
If applicable, add screenshots to help explain your problem.

**Version (please complete the following information):**
 - OS: [e.g. Win, Mac]
 - VBA Version [e.g. VBA7]
 - Bitness [e.g x32, x64]

Run this code for quick version:
```VBA
Public Sub ShowVBInfo()
    Dim res(0 To 2) As String
    #If Mac Then
        res(0) = "Mac"
    #Else
        res(0) = "Win"
    #End If
    #If VBA7 Then
        res(1) = "VBA7"
    #Else
        res(1) = "VBA6"
    #End If
    #If Win64 Then
        res(2) = "x64"
    #Else
        res(2) = "x32"
    #End If
   MsgBox Join(res, " ")
End Sub
```

**Additional context**
Add any other context about the problem here.
