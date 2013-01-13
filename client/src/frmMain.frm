VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SaCorp Updater"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8775
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   376
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   585
   StartUpPosition =   2  'CenterScreen
   Begin SaCorpUpdater.ctlProgressBar ctlProgressBar1 
      Height          =   375
      Left            =   120
      Top             =   4080
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   661
      Appearance      =   1
   End
   Begin VB.CommandButton cmdLaunch 
      Caption         =   "Démarrer"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Quitter"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   5160
      Width           =   1575
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   8535
      ExtentX         =   15055
      ExtentY         =   5741
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin InetCtlsObjects.Inet inetDownload 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label apropos 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "A propos"
      Height          =   255
      Left            =   7920
      TabIndex        =   5
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label lblprogress 
      BackStyle       =   0  'Transparent
      Caption         =   "Initialisation... Merci de patienter."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Width           =   8535
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SaCorp Updater"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub apropos_Click()
    frmAbout.Show
End Sub

Private Sub cmdLaunch_Click()
Dim appfile As String
appfile = GetVar(App.Path & "\Config\Launcher.ini", "GENERAL", "exename")
If FileExist(App.Path & "\" & appfile) Then
    Shell appfile, vbNormalFocus
Else
    MsgBox "Erreur : fichier introuvable"
End If
End Sub

Private Sub Form_Load()
Dim Filename As String
Dim appname As String
    WebBrowser1.Navigate GetVar(App.Path & "\Config\Launcher.ini", "GENERAL", "newsurl")
    
    'fond de l'updater
    Filename = App.Path & "\Config\sacorpupdater.jpg"
    
    If FileExist(Filename) Then
        Me.Picture = LoadPicture(Filename)
    End If
    appname = GetVar(App.Path & "\Config\Launcher.ini", "GENERAL", "gamename")
    Me.Caption = appname
    cmdLaunch.Caption = GetVar(App.Path & "\Config\Launcher.ini", "TEXT", "launch")
    cmdExit.Caption = GetVar(App.Path & "\Config\Launcher.ini", "TEXT", "quit")
    
    'on désactive le bouton lancer, pour le réactiver après la mise à jour
    cmdLaunch.Enabled = False
    
    Me.Show
    
    Update
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DestroyUpdater
End Sub

Private Sub cmdExit_Click()
    ' call the game destroy sub
    DestroyUpdater
End Sub
