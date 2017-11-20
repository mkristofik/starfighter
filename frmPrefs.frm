VERSION 5.00
Begin VB.Form frmPrefs 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit File Preferences"
   ClientHeight    =   3030
   ClientLeft      =   1470
   ClientTop       =   3075
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1935
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Default Starfighter Folder:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmPrefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()

    Unload Me

End Sub

Private Sub Dir1_Change()

    Drive1.Drive = Dir1.Path

End Sub

Private Sub Drive1_Change()

    Dir1.Path = Drive1.Drive

End Sub

Private Sub Form_Load()

    Dim strFile As String
    
    On Error GoTo Handle
    Open App.Path + "\prefs.txt" For Input As #1
    Input #1, strFile
    
    Dir1.Path = strFile
    Close #1
    
Handle:
    Dir1.Path = App.Path
    Close #1

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then CancelButton_Click

End Sub

Private Sub OKButton_Click()

    Open App.Path + "\prefs.txt" For Output As #1
    Write #1, Dir1.Path
    Close #1
    
    frmMain.dlgOpen.InitDir = Dir1.Path
    Unload Me

End Sub
