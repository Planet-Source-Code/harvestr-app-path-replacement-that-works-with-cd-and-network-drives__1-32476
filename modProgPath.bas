Attribute VB_Name = "modProgPath"
' API and Variables declaration

Declare Function GetModuleHandle Lib "kernel32" _
    Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
    
Declare Function GetModuleFileName Lib "kernel32" _
    Alias "GetModuleFileNameA" (ByVal hModule As Long, _
    ByVal lpFileName As String, ByVal nSize As Long) As Long

Const MAX_PATH = 260      ' Max Path size used in API

Global ProgPath As String ' Stores the Program path and make it available
                          ' in all modules of our Application

Function GetProgPath()
Dim lHandle As Long, lresult As Long, sBuffer As String

' Uses the API to get the program path
sBuffer = Space$(MAX_PATH)
lHandle = GetModuleHandle(App.EXEName)
lresult = GetModuleFileName(lHandle, sBuffer, MAX_PATH)
tmpPath = sBuffer

' Clean the String the API returned
For I = Len(tmpPath) To 1 Step -1
If Mid(tmpPath, I, 1) = "\" Then
    ProgPath = Left(tmpPath, I)
    Exit For
End If
Next I

' Add a terminal \ to the path if there's not one
' (if the path is a Drive root, as C:, then the \ is automatically
'  added, it's not in all other cases)
If Right(ProgPath, 1) <> "\" Then ProgPath = ProgPath + "\"

' Now the tricky part !
' (1) The problem with this API based ProgPath is, when you run it from
'     VB Editor (IDE), it will return the VB6.EXE path, not the real program
'     path. So, we test if our EXE is existing in the ProgPath. If it doesn't, then
'     we now we're in the IDE, and then we can use the classical App.Path
'
' (2) We use App.EXEName, so if our program is renamed, then ProgPath will
'     still return a valid value.

If Not FileExists(ProgPath + App.EXEName + ".exe") Then
    ProgPath = App.Path
    ' Add a terminal \ to the path if there's not one
    If Right(ProgPath, 1) <> "\" Then ProgPath = ProgPath + "\"
End If

End Function

Function FileExists%(FileName$)
Dim f%

' Here is a good way to check if a file exists. Sometimes, the Dir() command
' fails for some reasons, so this way is more accurate IMHO.
On Error Resume Next
f% = FreeFile
Open FileName$ For Input As #f%
Close #f%
FileExists% = Not (Err <> 0)

End Function

' This function is not used for ProgPath, but I think you may want to have
' a look at it :-)

Function DirExists%(FileName$)
Dim tmpDir As String

' Here is a good way to check if a folder exists. This one works with all
' Drives (Local, Network, CD...)
On Error Resume Next
tmpDir = CurDir
ChDir FileName
ChDir CurDir
DirExists% = Not (Err <> 0)

End Function

