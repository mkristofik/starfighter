VERSION 5.00
Begin VB.Form frmEquipInfo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Isometal Internal Structure"
   ClientHeight    =   2310
   ClientLeft      =   1785
   ClientTop       =   2805
   ClientWidth     =   4590
   ForeColor       =   &H80000015&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   Tag             =   "0"
   Begin VB.CommandButton cmdOK 
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
      Left            =   3720
      TabIndex        =   10
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblMax 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "1"
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
      Height          =   195
      Left            =   3240
      TabIndex        =   18
      Top             =   1920
      Width           =   120
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Max #:"
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
      Left            =   2520
      TabIndex        =   17
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Allowed Locations"
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
      Left            =   2640
      TabIndex        =   16
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblLoc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   15
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label lblLoc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   2
      Left            =   3120
      TabIndex        =   14
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label lblLoc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "LW, RW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   3
      Left            =   3600
      TabIndex        =   13
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Type:"
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
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblType 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Internal Structure"
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
      Height          =   315
      Left            =   1560
      TabIndex        =   11
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblSpace 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "0.00"
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
      Height          =   195
      Left            =   1560
      TabIndex        =   9
      Top             =   480
      Width           =   390
   End
   Begin VB.Label lblCriticals 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "4"
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
      Height          =   195
      Left            =   1560
      TabIndex        =   8
      Top             =   1200
      Width           =   120
   End
   Begin VB.Label lblDamage 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "---"
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
      Height          =   195
      Left            =   1560
      TabIndex        =   7
      Top             =   840
      Width           =   195
   End
   Begin VB.Label lblRange 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "---"
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
      Height          =   195
      Left            =   1560
      TabIndex        =   6
      Top             =   1560
      Width           =   195
   End
   Begin VB.Label lblTohit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "---"
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
      Height          =   195
      Left            =   1560
      TabIndex        =   5
      Top             =   1920
      Width           =   195
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Damage:"
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
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Range:"
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
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "To-hit:"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Criticals:"
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
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Space:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "frmEquipInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Activate()

    If Me.Tag > 0 Then
        Call ShowInfo
    ElseIf Me.Tag < 0 Then
        Call ShowEngine
    Else
        lblSpace.Caption = FormatNumber(RoundUp(frmMain.cboTotSpc.Text / 20), 2)
    End If

End Sub

Private Sub ShowInfo()

    Dim i As Integer
    
    Get #1, Me.Tag, Weapon
    Me.Caption = RTrim(Weapon.WeapName)
    lblSpace.Caption = FormatNumber(Weapon.WeapSpace, 2)
    lblCriticals.Caption = Weapon.Criticals
    lblTohit.Caption = Weapon.Tohit
    lblDamage.Caption = Weapon.Damage
    lblRange.Caption = Weapon.Range
    
    If InStr(Weapon.Options, 1) Then
        lblType.Caption = "Warhead Launcher"
    ElseIf InStr(Weapon.Options, 2) Then
        lblType.Caption = "Weapon"
    ElseIf InStr(Weapon.WeapName, "Sensors") Then
        lblType.Caption = "Sensors"
    Else
        lblType.Caption = "Miscellaneous"
    End If
    
    If Weapon.MaxNum Then
        lblMax.Caption = Weapon.MaxNum
    Else
        lblMax.Caption = "---"
    End If
    
    For i = 1 To 3
        If InStr(Weapon.Locations, i) Then
            lblLoc(i).ForeColor = vbGreen
        Else
            lblLoc(i).ForeColor = vbRed
        End If
    Next i
    
End Sub

Private Sub ShowEngine()

    Get #2, -CInt(Me.Tag), Engine
    Me.Caption = RTrim(Engine.EngName)
    lblCriticals.Caption = Engine.Criticals
    
    lblLoc(1).ForeColor = vbRed
    lblLoc(2).ForeColor = vbGreen
    lblLoc(3).ForeColor = vbGreen
    
    If Engine.EngType = 0 Then
        lblType.Caption = "Engine"
        lblSpace.Caption = "varies"
        lblMax.Caption = 1
    ElseIf Engine.EngType = 1 Then
        lblType.Caption = "Engine w/ Hyperdrive"
        lblSpace.Caption = "varies"
        lblMax.Caption = 1
    ElseIf Engine.EngType = 2 Then
        lblType.Caption = "Hyperdrive"
        lblCriticals.Caption = RoundUp(frmMain.cboTotSpc.Text / Engine.Criticals)
        lblSpace.Caption = FormatNumber(lblCriticals.Caption, 2)
        lblMax.Caption = 1
        lblLoc(3).ForeColor = vbRed
    Else
        lblType.Caption = "After Burner"
        lblLoc(3).ForeColor = vbRed
        If Me.Caption = "SLAM System" Then
            lblSpace.Caption = "2.00"
        Else
            lblSpace.Caption = "1.00"
        End If
        lblMax.Caption = "---"
    End If
    
    lblTohit.Caption = "---"
    lblDamage.Caption = "---"
    lblRange.Caption = "---"

End Sub
