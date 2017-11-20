VERSION 5.00
Begin VB.Form frmEdit 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Equipment Editor"
   ClientHeight    =   5655
   ClientLeft      =   1440
   ClientTop       =   2130
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   6510
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
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
      Left            =   5160
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
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
      Left            =   5160
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Frame fraInfo 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   3615
      Left            =   120
      TabIndex        =   15
      Top             =   1920
      Width           =   6255
      Begin VB.ListBox lstOptions 
         Height          =   510
         ItemData        =   "frmEdit.frx":0000
         Left            =   4200
         List            =   "frmEdit.frx":000A
         Style           =   1  'Checkbox
         TabIndex        =   13
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox txtCrits 
         Height          =   285
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   8
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtMaxNum 
         Height          =   285
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   11
         Top             =   2760
         Width           =   735
      End
      Begin VB.ListBox lstTech 
         Height          =   1185
         ItemData        =   "frmEdit.frx":0028
         Left            =   4200
         List            =   "frmEdit.frx":003B
         Style           =   1  'Checkbox
         TabIndex        =   14
         Top             =   2280
         Width           =   1935
      End
      Begin VB.ListBox lstLoc 
         Height          =   735
         ItemData        =   "frmEdit.frx":006F
         Left            =   4200
         List            =   "frmEdit.frx":007C
         Style           =   1  'Checkbox
         TabIndex        =   12
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtWeapName 
         Height          =   285
         Left            =   1080
         MaxLength       =   25
         TabIndex        =   5
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtDamage 
         Height          =   285
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   6
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtSpace 
         Height          =   285
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   7
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtRange 
         Height          =   285
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   9
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox txtTohit 
         Height          =   285
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   10
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Allowed Locations:"
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
         Height          =   495
         Left            =   3120
         TabIndex        =   28
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label10 
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
         Left            =   120
         TabIndex        =   27
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Max Qty (0 = Unlimited):"
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
         TabIndex        =   26
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tech-Base:"
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
         Left            =   3120
         TabIndex        =   25
         Top             =   2280
         Width           =   975
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Options:"
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
         Left            =   3360
         TabIndex        =   24
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Name:"
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
         TabIndex        =   23
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label2 
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
         Left            =   120
         TabIndex        =   22
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label3 
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
         Left            =   120
         TabIndex        =   21
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label4 
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
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   615
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
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblInstr 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Click to Add/Edit"
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
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblNum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   255
         Left            =   5880
         TabIndex        =   17
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Record #"
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
         Left            =   5040
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add New"
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
      Left            =   5160
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Data"
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
      Left            =   5160
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ListBox lstEquipment 
      Height          =   1815
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim numRecs As Integer, curRec As Integer, IsDirty As Boolean, flag As Boolean

Private Sub cmdAdd_Click()

    Dim i As Integer
    
    lblInstr.Caption = "Add New:"
    txtWeapName.Text = ""
    txtDamage.Text = ""
    txtSpace.Text = ""
    txtCrits.Text = ""
    txtRange.Text = ""
    txtTohit.Text = ""
    txtMaxNum.Text = ""
    curRec = numRecs + 1
    lblNum.Caption = curRec
    IsDirty = True
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdSave.Enabled = True
    fraInfo.Enabled = True
    Weapon.Deleted = False
    
    For i = 0 To 2
        lstLoc.Selected(i) = True
        If i < 2 Then lstOptions.Selected(i) = False
    Next i
    lstTech.Selected(0) = True
    txtWeapName.SetFocus
    
End Sub

Private Sub cmdDelete_Click()

    Dim m As Integer
    m = MsgBox("Warning!  Item will be permantently deleted.  Continue?", vbYesNo Or _
        vbExclamation)
    If m = vbYes Then
        Weapon.Deleted = True
        Put #1, curRec, Weapon
        ReloadList
    End If
        
    cmdDelete.Enabled = False

End Sub

Private Sub cmdEdit_Click()

    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    lblInstr.Caption = "Modify Existing:"
    fraInfo.Enabled = True
    cmdSave.Enabled = True
    
End Sub

Private Sub cmdSave_Click()

    Dim i As Integer, loc As String, opt As String, tech As String
        
    If BadData Then Exit Sub
    Weapon.WeapName = txtWeapName.Text
    Weapon.Damage = txtDamage.Text
    Weapon.WeapSpace = Val(txtSpace.Text)
    Weapon.Range = txtRange.Text
    Weapon.Tohit = txtTohit.Text
    Weapon.Criticals = Val(txtCrits.Text)
    Weapon.MaxNum = Val(txtMaxNum.Text)
    
    For i = 0 To 4
        If i <= 1 Then
            If lstLoc.Selected(i) Then loc = loc & CStr(i + 1)
            If lstOptions.Selected(i) Then opt = opt & CStr(i + 1)
            If lstTech.Selected(i) Then tech = tech & CStr(i)
        ElseIf i = 2 Then
            If lstLoc.Selected(i) Then loc = loc & "34"
            If lstTech.Selected(i) Then tech = tech & CStr(i)
        Else
            If lstTech.Selected(i) Then tech = tech & CStr(i)
        End If
    Next i
    
    Weapon.Locations = Val(loc)
    Weapon.Options = Val(opt)
    Weapon.TechBase = Val(tech)
    
    fraInfo.Enabled = False
    Put #1, curRec, Weapon
    MsgBox "Data saved."
    ReloadList
    IsDirty = False

End Sub

Function BadData() As Boolean
  
    Dim ret As Boolean
    
    If txtWeapName.Text = "" Then
        MsgBox "Must enter a name.", vbCritical
        ret = True
    End If
    If Val(txtCrits.Text) < 1 Then
        MsgBox "All equipment needs at least one critical slot.", vbCritical
        ret = True
    End If
    If lstLoc.SelCount = 0 Then
        MsgBox "At least one location must be available.", vbCritical
        ret = True
    End If
    If Val(txtSpace.Text) > 3 And lstLoc.Selected(0) Then
        MsgBox "3.0 space limit on cockpit.", vbCritical
        ret = True
    End If
    
    BadData = ret

End Function

Private Sub Form_Load()

    numRecs = GetNumRecords()
    curRec = 1
    ReloadList
    IsDirty = False
    Call SetTabs(lstEquipment)
    
End Sub

Private Sub ReloadList()

    Dim i As Integer
    lstEquipment.Clear
    numRecs = GetNumRecords()
    
    For i = 1 To numRecs
        Get #1, i, Weapon
        If Not Weapon.Deleted Then
            lstEquipment.AddItem (Weapon.WeapName) & Chr$(9) & TechString(Weapon.TechBase)
            lstEquipment.ItemData(lstEquipment.NewIndex) = i
        End If
    Next i
    cmdEdit.Enabled = False
    cmdSave.Enabled = False
    
End Sub

Private Sub lstEquipment_Click()
    
    Dim i As Integer
    
    ' Error trap for when user clicks cancel in code block below.
    If flag Then
        flag = False
        Exit Sub
    End If
    
    ' Check to change weapons.
    If IsDirty Then
        Dim m As Integer
        m = MsgBox("This action will reset any changes made to current weapon." _
            + vbCrLf + "Do you wish to continue?", vbYesNo)
        If m = 7 Then
            flag = True
            lstEquipment.ListIndex = curRec
            Exit Sub
        End If
    End If
    
    curRec = lstEquipment.ItemData(lstEquipment.ListIndex)
    Get #1, curRec, Weapon
    lblNum.Caption = curRec
    
    txtWeapName.Text = RTrim(Weapon.WeapName)
    txtSpace.Text = Weapon.WeapSpace
    txtDamage.Text = RTrim(Weapon.Damage)
    txtRange.Text = RTrim(Weapon.Range)
    txtTohit.Text = RTrim(Weapon.Tohit)
    txtCrits.Text = Weapon.Criticals
    txtMaxNum.Text = Weapon.MaxNum
    
    ' Fill in the checkboxes
    For i = 0 To 4
        If i < 2 Then
            lstOptions.Selected(i) = False
            If InStr(CStr(Weapon.Options), i + 1) Then lstOptions.Selected(i) = True
        End If
        
        If i < 3 Then
            lstLoc.Selected(i) = False
            If InStr(CStr(Weapon.Locations), i + 1) Then lstLoc.Selected(i) = True
        End If
        
        lstTech.Selected(i) = False
        If InStr(CStr(Weapon.TechBase), i) Then lstTech.Selected(i) = True
    Next i
    
    lblInstr.Caption = "View Existing:"
    cmdDelete.Enabled = True
    cmdSave.Enabled = False
    fraInfo.Enabled = False
    cmdEdit.Enabled = True
    IsDirty = False
    
End Sub

Private Sub mnuClose_Click()
    Unload Me
End Sub

Private Sub txtDamage_Change()
    IsDirty = True
End Sub

Private Sub txtRange_Change()
    IsDirty = True
End Sub

Private Sub txtTohit_Change()
    IsDirty = True
End Sub

Private Sub txtWeapName_Change()
    IsDirty = True
End Sub

Private Sub lstLoc_Click()
    IsDirty = True
End Sub

Private Sub lstOptions_Click()
    IsDirty = True
    If lstOptions.Selected(0) Then lstOptions.Selected(1) = True
End Sub

Private Sub lstTech_Click()
    
    Dim i As Integer
    IsDirty = True
    If lstTech.Selected(0) Then
        For i = 1 To 4
            lstTech.Selected(i) = False
        Next i
    End If

End Sub

Private Function GetNumRecords()
    GetNumRecords = LOF(1) / Len(Weapon)
End Function
