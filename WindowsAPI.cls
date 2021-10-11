VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WindowsAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
(ByVal lpClassname As String, ByVal lpWindowName As String) As Long '------------Import FindWindow function for locating an open file.

Private Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" ( _
    ByVal HWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long '-----------------------------------------------------Import PostMessage function for sending a message to the Windows System.

Private Declare PtrSafe Function GetForegroundWindow Lib "user32" _
    () As Long '-----------------------------------------------------------------Import function for getting the handle of the active window.
    
Private Declare PtrSafe Function GetWindowText Lib "user32" _
    Alias "GetWindowTextA" (ByVal HWnd As Long, _
    ByVal lpString As String, ByVal cch As Long) As Long '-----------------------Import function for getting the name of a window given its handle.

Private Const WM_CLOSE = &H10 '--------------------------------------------------Define Windows command WM_CLOSE as hexadecimal 16.


Dim FolderPath              As String '------------------------------------------The absolute path of the specified folder.

Dim FilePath                As String '------------------------------------------The absolute path of the specified file.
Dim FileName                As String

Dim AppPath                 As String '------------------------------------------The absolute path of the specified application.
Dim AppName                 As String '------------------------------------------The *.exe name of the specified application.
Dim AppPID                  As Integer '-----------------------------------------The process ID of the specified application.


Public Sub InitializeFile( _
    ByVal ThisFile As String, _
    ByVal ThisApp As String)
    'This sub must be run before interacting with a file.

    AppPID = 0 '-----------------------------------------------------------------Should be 0 by default, but want ensure that we don't tamper with the wrong process.

    FilePath = ThisFile
    AppPath = ThisApp
    
    ExtractFileName
    ExtractAppName
    
End Sub

Public Sub InitializeFolder(ByVal ThisFolder As String)
    'This sub must be run before interacting with a folder.

    FolderPath = ThisFolder

End Sub


'************************************Begin File Interaction Functions*****************************************

Private Sub ExtractFileName()
    'Extract the filename from a filepath, if the file exists.
    
    FileName = Dir(FilePath) '---------------------------------------------------Extract the filename from the path.
    
End Sub

Private Sub ExtractAppName()
    'Extract the application name from an application path, if it exists.

    AppName = Dir(AppPath)

End Sub

Public Sub OpenFile(Optional ByVal WindowStyle As Integer = 3)
    'Access the shell directly and input command line prompts to open a file.
    
    GetPID '---------------------------------------------------------------------Get app PID if the app is open.
    CloseApp '-------------------------------------------------------------------Close the app if it's open.

    AppPID = Shell(AppPath & " " & FilePath, WindowStyle) '----------------------Use command line to open file and also save the PID.

    WaitUntilFileOpen '----------------------------------------------------------Don't proceed until the file has fully loaded.
    
    
End Sub

Public Sub CloseFile()
                    
    Dim HWnd                As Long
                    
    HWnd = FindWindow(vbNullString, FileName)
    PostMessage HWnd, WM_CLOSE, 0, ByVal 0&
    
    WaitUntilFileClosed
                      
End Sub

Private Function FileStatus()
    'Return True if the file is open, otherwise return False.
    
    Dim FileNum             As Integer
    Dim ErrNum              As Integer
 
    On Error Resume Next '-------------------------------------------------------Turn error checking off.
    
    FileNum = FreeFile() '-------------------------------------------------------Get the next available file number.
    
    Open FilePath For Input Lock Read As #FileNum '------------------------------Attempt to open the file and lock it.
    Close FileNum '--------------------------------------------------------------Close the file.
    ErrNum = Err '---------------------------------------------------------------Save the error number that occurred.
    
    On Error GoTo 0 '------------------------------------------------------------Turn error checking back on.
 
    Select Case ErrNum '---------------------------------------------------------Select the appropriate error.
        Case 0
            FileStatus = False '-------------------------------------------------No error. The file is not open.
        Case 70
            FileStatus = True '--------------------------------------------------Permission denied. The file is open.
        Case Else
            Error ErrNum '-------------------------------------------------------Some other error occurred, but is not relevant for this context.
    End Select
    
End Function

Private Sub WaitUntilFileOpen()
    'Repeatedly check the status of the file until it's open.
    
    Do Until FileStatus() = True
    Loop
    
    Application.Wait Now + #12:00:03 AM# '---------------------------------------Even though the file is open, it takes time to finish loading.
    
End Sub

Private Sub WaitUntilFileClosed()
    'Repeatedly check the status of the file until it's closed.

    Do Until FileStatus() = False
    Loop
    
    Application.Wait Now + #12:00:01 AM#

End Sub

Public Function ActiveWindow()
    'Return the name of the active window (the window that the OS is focusing on).

    Dim WinText As String
    Dim HWnd As Long
    Dim L As Long
    
    HWnd = GetForegroundWindow() '-----------------------------------------------Get the window handle.
    WinText = String(255, vbNullChar) '------------------------------------------Create a null string.
    
    L = GetWindowText(HWnd, WinText, 255) '--------------------------------------Get the length of the active window and set the beginning of WinText to window name.
    WinText = Left(WinText, InStr(1, WinText, vbNullChar) - 1) '-----------------Remove excess null characters from WinText.
    
    ActiveWindow = WinText

End Function

Public Sub CloseApp()
    'Close the initialized application if it's open.

    If AppPID <> 0 Then '--------------------------------------------------------If the app has been opened.
        Shell "TaskKill /F /PID " & CStr(AppPID), vbHide
    End If

End Sub

Public Sub GetPID()
    'Set AppPID to the PID of the initialized application, if it's already open.
    
    Dim CmdOutput()        As String
    
    Dim n As Integer
    
    CmdOutput = ShellRun( _
        "tasklist /FI " & Chr(34) & _
        "ImageName eq " & AppName & Chr(34) & _
        " /FI ""Status eq Running"" /FO LIST")
    
    For n = 1 To UBound(CmdOutput)
        If Left(CmdOutput(n), 4) = "PID:" Then
            CmdOutput(n) = Replace(CmdOutput(n), " ", "") '----------------------Remove white space.
            CmdOutput(n) = Split(CmdOutput(n), ":")(1)
            AppPID = CInt(CmdOutput(n))
        End If
    Next n

End Sub

Public Function ShellRun(sCmd As String) As Variant
    'Return an array of line items from a command prompt output.

    Dim ReturnArray() As String
    Dim n As Integer
    
    Dim ShellObj As Object
    Set ShellObj = CreateObject("WScript.Shell")

    'run command
    Dim oExec As Object
    Dim oOutput As Object
    Set oExec = ShellObj.Exec(sCmd)
    Do While oExec.Status = 0
        Application.Wait Now + #12:00:01 AM#
    Loop
    
    Set oOutput = oExec.StdOut

    'handle the results as they are written to and read from the StdOut object
    Dim s As String
    Dim sLine As String
    n = 1
    While Not oOutput.AtEndOfStream
        sLine = oOutput.ReadLine
        
        If sLine <> "" Then
            ReDim Preserve ReturnArray(n)
            ReturnArray(n) = sLine
            n = n + 1
        End If
    Wend

    ShellRun = ReturnArray

End Function

'************************************Begin Folder Interaction Functions*****************************************

Public Function GetFileList()
    'Return a list of filenames in a specified folder path.
    
    'Argument(s):
    '   FolderPath As String - Absolute path of desired folder.
    'Return:
    '   List of file names As String.

    'Define file system objects.
    Dim Filesys             As Object
    Dim Folder              As Object
    Dim File                As Object
    Dim Files               As Object
    
    Dim s()                 As String '------------------------------------------Define an array for storing file names.
    
    Dim i                   As Integer '-----------------------------------------Define a generic iterator.
    
    Set Filesys = CreateObject("Scripting.FileSystemObject") '-------------------Set filesys as a Windows file system object for navigation.
    Set Folder = Filesys.GetFolder(FolderPath) '---------------------------------Set folder as a folder object for the specified path.
    Set Files = Folder.Files '---------------------------------------------------Set files as the set of file objects within the folder.
    
    'Append each file name to S().
    i = 1
    For Each File In Files
        If Right(File.Name, 4) = ".pdf" Then '-----------------------------------This is necessary to avoid including desktop.ini or any other unwanted file.
            ReDim Preserve s(i)
            s(i) = File.Name '---------------------------------------------------Append the file name to S at the ith position.
            i = i + 1
        End If
    Next
    
    GetFileList = s
    
End Function
