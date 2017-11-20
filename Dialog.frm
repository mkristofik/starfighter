VERSION 5.00
Begin VB.Form frmChgSpd 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Speed"
   ClientHeight    =   1290
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.HScrollBar hsbSpeed 
      Height          =   255
      Left            =   120
      Max             =   10
      TabIndex        =   2
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   2760
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
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
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblMaxSpeed 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Max: 10"
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
      Left            =   960
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblCurSpeed 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "3"
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
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "frmChgSpd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
' Click the Cancel button.
    frmChgSpd.Visible = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
' Click the OK button.
    frmChgSpd.Visible = False
    Engine.Speed = hsbSpeed.value
    CalcEngine
    Unload Me
    frmEquipment.UpdateTotalCrits
End Sub

Private Sub Form_Load()
' Initialize variables.
    hsbSpeed.value = Engine.Speed
    Dim MaxSpeed As Integer
    
    Open App.Path & "\engine.dat" For Input As #1
    Dim EngineName, TechBase As String
    Dim ManBase, Criticals As Integer
    Dim Modifier As Single
    
    Do Until EOF(1)
        Input #1, EngineName, Modifier, ManBase, Criticals, TechBase
        If EngineName = RTrim(Engine.EngType) Then
            Close #1
            Exit Do
        End If
    Loop
    
    If Modifier = 0.5 Then
        MaxSpeed = Int(400 / Craft.TotalSpace) * 2
    Else
        MaxSpeed = Int(400 / Craft.TotalSpace) - Modifier
    End If
    
    If MaxSpeed > 10 Then MaxSpeed = 10
    lblMaxSpeed.Caption = "Max: " & MaxSpeed
    hsbSpeed.max = MaxSpeed
End Sub

Private Sub hsbSpeed_Change()
' Move the Speed scrollbar.
    lblCurSpeed.Caption = hsbSpeed.value
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        cmdOK_Click
    End If
End Sub
