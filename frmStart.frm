VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create New"
   ClientHeight    =   2556
   ClientLeft      =   3072
   ClientTop       =   3516
   ClientWidth     =   5004
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2556
   ScaleWidth      =   5004
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open &Existing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   720
      Width           =   1575
   End
   Begin VB.CheckBox chkIsometal 
      BackColor       =   &H80000009&
      Caption         =   "Isometal Internal Structure"
      Height          =   495
      Left            =   3360
      TabIndex        =   7
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame fraWings 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Fuselage Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   2655
      Begin VB.OptionButton optTriWing 
         BackColor       =   &H80000009&
         Caption         =   "Tri-Wing (i.e. TIE Defender)"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   2295
      End
      Begin VB.OptionButton optWingless 
         BackColor       =   &H80000009&
         Caption         =   "No Wings (i.e. A-Wing)"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton optStandard 
         BackColor       =   &H80000009&
         Caption         =   "Standard"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.ComboBox cboTechBase 
      Height          =   315
      ItemData        =   "frmStart.frx":0000
      Left            =   2040
      List            =   "frmStart.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label lblTechBase 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Technology Base:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   1575
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    
    Dim strFile As String
    Open App.Path + "\prefs.txt" For Input As #1
    Input #1, strFile
    Close #1
 
    Craft.TechBase = left(cboTechBase.Text, 1)

    If optStandard.Value Then Craft.Wings = 2
    If optWingless.Value Then Craft.Wings = 0
    If optTriWing.Value Then Craft.Wings = 3
    
    If chkIsometal.Value Then
        Internal.Iso = True
    Else
        Internal.Iso = False
    End If
    
    Unload Me
    frmMain.Show
    frmMain.dlgOpen.InitDir = strFile

End Sub

Private Sub cmdOpen_Click()
    
    cmdOK_Click
    frmMain.mnuFileOpen_Click
    
End Sub

Private Sub Form_Load()
    cboTechBase.ListIndex = 0
    
    On Error GoTo FileCreate
    Open App.Path + "\prefs.txt" For Input As #1
    Close #1
    Exit Sub
        
FileCreate:
    Open App.Path + "\prefs.txt" For Output As #1
    Write #1, App.Path
    Close #1
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then cmdOK_Click
End Sub
