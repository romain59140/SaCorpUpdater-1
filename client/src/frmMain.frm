VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autoupdater"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4455
   ControlBox      =   0   'False
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
   MinButton       =   0   'False
   ScaleHeight     =   322
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   297
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtProgress 
      Appearance      =   0  'Flat
      Height          =   3495
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   720
      Width           =   4215
   End
   Begin InetCtlsObjects.Inet inetDownload 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label lblConnect 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Autoupdater"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim Filename As String

    ' if a GUI image exists, then load it onto the form
    ' and make the labels invisible.
    Filename = App.Path & "\Data Files\graphics\autoupdater.jpg"
    
    If FileExist(Filename) Then
        Me.Picture = LoadPicture(Filename)
        lblTitle.Visible = False
        lblConnect.Caption = vbNullString
        lblExit.Caption = vbNullString
    End If
    
    Me.Show
    
    AddProgress "Welcome to the Eclipse Origins autoupdater."
    AddProgress "Press 'Connect' to update your client."
    AddProgress ""
End Sub

Private Sub lblConnect_Click()
    ' call the update sub
    Update
End Sub

Private Sub lblExit_Click()
    ' call the game destroy sub
    DestroyUpdater
End Sub
