Attribute VB_Name = "ModSelfDelete"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32.dll" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Private Function FixPath(lpPath As String) As String
    'Fix a given path adding \ backslash when required
    If Right(lpPath, 1) = "\" Then FixPath = lpPath: Exit Function Else FixPath = lpPath & "\"
End Function

Private Function GetFileName(lpPath As String) As String
Dim x As Integer, e_pos As Integer
    'Used to Return a filename from a path
    'eg GetFileName(C:\Windows\system32\this.exe) returns this.exe
    
    For x = 1 To Len(lpPath) 'Loop the string
        If Mid$(lpPath, x, 1) = "\" Then e_pos = x ' if we find \ in the string store it's poistion
    Next x
    
    x = 0 'Clear var
    GetFileName = LTrim$(Mid$(lpPath, e_pos + 1, Len(lpPath))) 'Extract the filename
    
End Function

Private Function GetSortFileName(lpFile As String) As String
Dim iRet As Long
Dim sPath As String
    'Function used to return the sort path of a long path
    sPath = Space$(260) 'Make some room to store the path
    iRet = GetShortPathName(lpFile, sPath, 260)
    GetSortFileName = Left(sPath, iRet) 'Return path and remove trailing spaces
    sPath = "" 'Clear var
End Function

Public Sub SelfDelete()
Dim absFile As String
Dim sGetFile As String
Dim fp As Long

    'This is the main self delete sub
    
    absFile = GetSortFileName(FixPath(App.Path) & App.EXEName & ".exe") 'Get the correct path and filename
    sGetFile = GetFileName(absFile) 'Extract the filename from string above
    
    fp = FreeFile 'Pointer to file
    
    Open FixPath(App.Path) & "del.bat" For Output As #fp
        Print #fp, "@ECHO OFF"
        Print #fp, "attrib -s -h -r -a" ' Remove the file's attribites
        Print #fp, "try:" ' Start
        Print #fp, "del " & sGetFile ' Delete the file
        Print #fp, "if exist " & sGetFile & " goto Try" ' While file is still here jump to start
        Print #fp, "del del.bat" ' delete this batch file
        Print #fp, "exit" ' Exit to OS
    Close #fp
    
    ShellExecute hwnd, "open", FixPath(App.Path) & "del.bat", vbNullString, vbNullString, 0 'used to execure the batch file
    sGetFile = ""
    apFile = ""
    
End Sub

