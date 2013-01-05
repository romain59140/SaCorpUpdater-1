Attribute VB_Name = "modMain"
Option Explicit

'|--------------------------------|
'| Eclipse Origins - Autoupdater  |
'| Created by: Robin Perris       |
'| Website: freemmorpgmaker.com   |
'|--------------------------------|

' file host
Private UpdateURL As String

' stores the variables for the version downloaders
Private VersionCount As Long

Sub Main()
    If Not FileExist(App.Path & "\Data Files\updaterInfo.ini") Then DestroyUpdater
    UpdateURL = GetVar(App.Path & "\Data Files\updaterInfo.ini", "UPDATER", "updateURL")
    Load frmMain
End Sub

Public Sub DestroyUpdater()
    ' kill temp files
    If FileExist(App.Path & "\tmpUpdate.ini") Then Kill App.Path & "\tmpUpdate.ini"
    ' end updater
    Unload frmMain
    End
End Sub

Public Sub Update()
Dim CurVersion As Long
Dim Filename As String
Dim i As Long

    AddProgress "Connecting to server..."

    ' get the file which contains the info of updated files
    DownloadFile UpdateURL & "/update.ini", App.Path & "\tmpUpdate.ini"
    
    AddProgress "Connected to server!"
    AddProgress "Retrieving version information."
    
    ' read the version count
    VersionCount = GetVar(App.Path & "\tmpUpdate.ini", "FILES", "Versions")
    
    ' check if we've got a current client version saved
    If FileExist(App.Path & "\Data Files\version.ini") Then
        CurVersion = GetVar(App.Path & "\Data Files\version.ini", "UPDATER", "CurVersion")
    Else
        CurVersion = 0
    End If
    
    ' are we up to date?
    If CurVersion < VersionCount Then
        ' make sure it's not 0!
        If CurVersion = 0 Then CurVersion = 1
        ' loop around, download and unrar each update
        For i = CurVersion To VersionCount
            ' let them know!
            AddProgress "Downloading version " & i & "."
            Filename = "version" & i & ".rar"
            ' set the download going through inet
            DownloadFile UpdateURL & "/" & Filename, App.Path & "\" & Filename
            ' us the unrar.dll to extract data
            RARExecute OP_EXTRACT, Filename
            ' kill the temp update file
            Kill App.Path & "\" & Filename
            ' update the current version
            PutVar App.Path & "\Data Files\version.ini", "UPDATER", "CurVersion", Str(i)
            ' let them know!
            AddProgress "Version " & i & " installed."
        Next
        ' let them know the update has finished
        AddProgress ""
        AddProgress "Update Complete!"
        AddProgress "You can now exit the updater.", False
    Else
        ' they're at the correct version, or perhaps higher!
        AddProgress ""
        AddProgress "You are completely up to date!"
        AddProgress "You can now exit the updater.", False
    End If
End Sub

Public Sub AddProgress(ByVal sProgress As String, Optional ByVal newline As Boolean = True)
    ' add a string to the textbox on the form
    frmMain.txtProgress.Text = frmMain.txtProgress.Text & sProgress
    If newline = True Then frmMain.txtProgress.Text = frmMain.txtProgress.Text & vbNewLine
End Sub

Private Sub DownloadFile(ByVal URL As String, ByVal Filename As String)
    Dim fileBytes() As Byte
    Dim fileNum As Integer
    
    On Error GoTo DownloadError
    
    ' download data to byte array
    fileBytes() = frmMain.inetDownload.OpenURL(URL, icByteArray)
    
    fileNum = FreeFile
    Open Filename For Binary Access Write As #fileNum
        ' dump the byte array as binary
        Put #fileNum, , fileBytes()
    Close #fileNum
    
    Exit Sub
    
DownloadError:
    MsgBox Err.Description
End Sub
