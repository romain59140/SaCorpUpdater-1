Attribute VB_Name = "modMain"
Option Explicit

'-------------------
'| SaCorpUpdater   |
'-------------------
'| By Herasus      |
'-------------------

Private UpdateURL As String
Dim appname As String
Private VersionCount As Long

Sub Main()
    If Not FileExist(App.Path & "\Config\Launcher.ini") Then DestroyUpdater
    UpdateURL = GetVar(App.Path & "\Config\Launcher.ini", "SERVER", "updateurl")
    appname = GetVar(App.Path & "\Config\Launcher.ini", "GENERAL", "gamename")
    frmMain.lblTitle.Caption = appname
    Load frmMain
End Sub

Public Sub DestroyUpdater()
    'Supprimer fichier temp
    If FileExist(App.Path & "\tmpUpdate.ini") Then Kill App.Path & "\tmpUpdate.ini"
    'Quitter
    Unload frmMain
    End
End Sub

Public Sub Update()
Dim CurVersion As Long
Dim Filename As String
Dim I As Long

    AddProgress "Connexion au serveur de " & appname & " en cours..."

    'Télécharge le fichier de mise à jour
    DownloadFile UpdateURL & "/update.ini", App.Path & "\tmpUpdate.ini"
    
    'A t-on téléchargé le fichier ?
    
    If Not FileExist(App.Path & "\tmpUpdate.ini") Then
    GoTo downfail
    End If
    
    AddProgress "Connexion au serveur réussie, en attente d'informations..."
    
    'quelle version le client a ?
    VersionCount = GetVar(App.Path & "\tmpUpdate.ini", "FILES", "versions")
    
    'voir si une version du client existe
    If FileExist(App.Path & "\Config\version.ini") Then
        CurVersion = GetVar(App.Path & "\Config\version.ini", "UPDATER", "CurVersion")
    Else
        CurVersion = 1
    End If
    
    'Sommes-nous à jour ?
    If CurVersion < VersionCount Then
        If CurVersion = 0 Then CurVersion = 1
        For I = CurVersion To VersionCount
            AddProgress "Téléchargement de la version " & I & "."
            Filename = "version" & I & ".rar"
            DownloadFile UpdateURL & "/" & Filename, App.Path & "\" & Filename
            RARExecute OP_EXTRACT, Filename
            Kill App.Path & "\" & Filename
            PutVar App.Path & "\Config\version.ini", "UPDATER", "CurVersion", Str(I)
            AddProgress "Version " & I & " installé."
        Next
        'Fin de la mise à jour
        AddProgress "Les mises à jour sont terminés, vous pouvez dorénavant lancer " & appname & "."
    Else
        'Déjà à jour
        AddProgress "Vous êtes à jour, vous pouvez dorénavant lancer " & appname & "."
    End If
    frmMain.cmdLaunch.Enabled = True
    
=downfail
    
    AddProgress "Connexion au serveur de mise à jour impossible."
    frmMain.cmdLaunch.Enabled = True
End Sub


Public Sub AddProgress(ByVal sProgress As String, Optional ByVal newline As Boolean = True)
    frmMain.lblprogress.Caption = sProgress
End Sub

Sub DownloadProgress(intPercent As String)
    frmMain.ctlProgressBar1.value = intPercent
End Sub


Public Sub DownloadFile(strURL As String, strDestination As String)
'Code inspiré de http://www.codeitbetter.com/download-file-inet-control-progress/
    Const CHUNK_SIZE As Long = 1024
    Dim iFile As Integer
    Dim lBytesReceived As Long
    Dim lFileLength As Long
    Dim strHeader As String
    Dim b() As Byte
    Dim I As Integer
    DoEvents
    With frmMain.inetDownload
        .URL = strURL
        .Execute , "GET", , "Range: bytes=" & CStr(lBytesReceived) & "-" & vbCrLf
        While .StillExecuting
            DoEvents
        Wend
        strHeader = .GetHeader
    End With
    strHeader = frmMain.inetDownload.GetHeader("Content-Length")
    lFileLength = Val(strHeader)
    DoEvents
    lBytesReceived = 0
    iFile = FreeFile()
    Open strDestination For Binary Access Write As #iFile
    Do
        b = frmMain.inetDownload.GetChunk(CHUNK_SIZE, icByteArray)
        Put #iFile, , b
        lBytesReceived = lBytesReceived + UBound(b, 1) + 1
        DownloadProgress (Round((lBytesReceived / lFileLength) * 100))
        DoEvents
    Loop While UBound(b, 1) > 0
    Close #iFile
End Sub
