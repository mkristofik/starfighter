VERSION 5.00
Begin VB.Form frmArmor 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modify Armor"
   ClientHeight    =   4050
   ClientLeft      =   975
   ClientTop       =   1365
   ClientWidth     =   7260
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   7260
   Begin VB.CheckBox chkUpdate 
      Caption         =   "Update"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   32
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.HScrollBar hsbShields 
      Height          =   255
      LargeChange     =   20
      Left            =   4680
      Max             =   200
      SmallChange     =   10
      TabIndex        =   29
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
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
      Left            =   6120
      TabIndex        =   24
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton cmdMaxArmor 
      Caption         =   "&Maximize Armor/Shields"
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
      Left            =   4560
      TabIndex        =   23
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CheckBox chkBalance 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Balance Armor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4920
      TabIndex        =   22
      Top             =   2040
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      ItemData        =   "frmArmor.frx":0000
      Left            =   240
      List            =   "frmArmor.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   480
      Width           =   1575
   End
   Begin VB.VScrollBar vsbArmor 
      Height          =   855
      Index           =   4
      LargeChange     =   2
      Left            =   3600
      Max             =   0
      Min             =   30
      TabIndex        =   3
      Top             =   2280
      Width           =   375
   End
   Begin VB.VScrollBar vsbArmor 
      Height          =   855
      Index           =   2
      LargeChange     =   2
      Left            =   2400
      Max             =   0
      Min             =   40
      TabIndex        =   2
      Top             =   2280
      Width           =   375
   End
   Begin VB.VScrollBar vsbArmor 
      Height          =   855
      Index           =   3
      LargeChange     =   2
      Left            =   1200
      Max             =   0
      Min             =   30
      TabIndex        =   1
      Top             =   2280
      Width           =   375
   End
   Begin VB.VScrollBar vsbArmor 
      Height          =   855
      Index           =   1
      LargeChange     =   5
      Left            =   2400
      Max             =   0
      Min             =   5
      TabIndex        =   0
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblCurShields 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Shields: 0"
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
      TabIndex        =   31
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lblMaxShields 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Max: 200"
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
      Left            =   5280
      TabIndex        =   30
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblShields 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0"
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
      Left            =   5280
      TabIndex        =   28
      Top             =   960
      Width           =   375
   End
   Begin VB.Label lblShieldSpc 
      Appearance      =   0  'Flat
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
      Height          =   255
      Left            =   5280
      TabIndex        =   27
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblShdSpc 
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
      Left            =   3840
      TabIndex        =   26
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblShd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Shields:"
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
      Left            =   3840
      TabIndex        =   25
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblArmorType 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Armor Type"
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
      Left            =   480
      TabIndex        =   21
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblSpaceLeft 
      Appearance      =   0  'Flat
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
      Height          =   255
      Left            =   5280
      TabIndex        =   19
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblArmorLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0"
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
      Left            =   5280
      TabIndex        =   18
      Top             =   720
      Width           =   375
   End
   Begin VB.Label lblArmorSpace 
      Appearance      =   0  'Flat
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
      Height          =   255
      Left            =   5280
      TabIndex        =   17
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblArmorFactor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
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
      Left            =   5280
      TabIndex        =   16
      Top             =   240
      Width           =   375
   End
   Begin VB.Label lblSpaceRemaining 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Space Left:"
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
      Left            =   3840
      TabIndex        =   15
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblUnused 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Unused Armor:"
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
      Left            =   3840
      TabIndex        =   14
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblSpace 
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
      Left            =   3840
      TabIndex        =   13
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Armor Factor:"
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
      Left            =   3840
      TabIndex        =   12
      Top             =   240
      Width           =   1215
   End
   Begin VB.Shape shpBox 
      Height          =   1695
      Left            =   3720
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label lblArmor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
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
      Index           =   4
      Left            =   3600
      TabIndex        =   11
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lblArmor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
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
      Index           =   2
      Left            =   2400
      TabIndex        =   10
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lblArmor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
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
      Index           =   3
      Left            =   1200
      TabIndex        =   9
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lblArmor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
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
      Index           =   1
      Left            =   2400
      TabIndex        =   8
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblRightWing 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Right Wing"
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
      Left            =   3240
      TabIndex        =   7
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblLeftWing 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Left Wing"
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
      Left            =   840
      TabIndex        =   6
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblFuselage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Fuselage"
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
      Left            =   2040
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblCockpit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cockpit"
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
      Left            =   2040
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmArmor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboType_Click()
    
    If cboType.ListIndex = 4 Then
        If Craft.TechBase = 3 Then
            Call UpdateArmor
        Else
            MsgBox "Clearplast armor only available to Herald craft.", vbExclamation
            cboType.ListIndex = 0
        End If
    Else
        Call UpdateArmor
    End If
    Craft.ArmType = cboType.ListIndex
    
End Sub

Private Sub chkBalance_Click()
    vsbArmor(4).Value = vsbArmor(3).Value
End Sub

Private Sub chkUpdate_Click()
    If chkUpdate.Value <> 1 Then Exit Sub
    Call cboType_Click
    Call SetMaxima
    chkUpdate.Value = 0
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdMaxArmor_Click()

    Dim i As Integer
    For i = 1 To 4
        vsbArmor(i).Value = vsbArmor(i).Min
    Next i
    hsbShields.Value = hsbShields.max
    
End Sub

Private Sub Form_Activate()
    frmMain.Enabled = False
End Sub

Private Sub Form_Load()
    cboType.ListIndex = 0
    Call SetMaxima
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMain.Enabled = True
    Cancel = NotNewShip
    Me.Hide
End Sub

Private Sub UpdateArmor()
' Update armor totals on both the Armor form and Main form.
        
    Dim tot As Integer, i As Integer
    For i = 1 To 4
        tot = tot + Val(lblArmor(i).Caption)
    Next i
        
    If tot Mod 5 Then
        lblArmorLeft.Caption = 5 - tot Mod 5
        tot = tot + 5 - tot Mod 5 ' Round up to nearest multiple of 5
        lblArmorLeft.ForeColor = vbRed
    Else
        lblArmorLeft.ForeColor = vbBlack
        lblArmorLeft.Caption = 0
    End If
    lblArmorFactor.Caption = tot
    
    lblArmorSpace.Caption = FormatNumber(tot * cboType.ItemData(cboType.ListIndex) / 10, 2)
    
    ' Update the main form
    With frmMain
        .lblArmor(0).Caption = lblArmorFactor.Caption
        .lblArmorType.Caption = cboType.Text
        For i = 1 To 4
            .lblArmor(i).Caption = lblArmor(i).Caption
        Next i
        .lblSpace(4).Caption = lblArmorSpace.Caption
    End With
        
End Sub

Private Sub SetMaxima()
    
    Dim a As Integer
    
    ' Armor
    Dim i As Integer, max As Integer, max2 As Integer
    If RTrim(frmEquipment.lstCrits(1).List(4)) = "Co-Pilot" Or RTrim(frmEquipment.lstCrits(1).List(5)) = "Co-Pilot" Then
        vsbArmor(1).Min = 6
    Else
        vsbArmor(1).Min = 5
    End If
    
    For i = 2 To 4
        a = frmMain.lblInternal(i).Caption * 2
        If Craft.TechBase >= 3 Then a = a * 2
        vsbArmor(i).Min = a
    Next i
    
    ' Shields
    max = (CInt(hsbShields.Tag) + CInt(frmEquipment.lblCritsLeft(2).Caption)) * 50
    max2 = frmMain.cboTotSpc.Text * 2
    If Craft.TechBase >= 3 Then max2 = max2 * 2
    
    If max < max2 Then ' Take lower of the two values
        hsbShields.max = max
    Else
        hsbShields.max = max2
    End If
    
    lblMaxShields.Caption = "Max: " & hsbShields.max
    
End Sub

Private Sub hsbShields_Change()

    hsbShields.Enabled = False
    If hsbShields.Value Mod 10 Then hsbShields.Value = hsbShields.Value + 10 - _
        hsbShields.Value Mod 10
    
    lblCurShields.Caption = "Shields: " & hsbShields.Value
    lblShields.Caption = hsbShields.Value
    lblShieldSpc.Caption = FormatNumber(hsbShields.Value / 20)
    
    ' Update the main form
    With frmMain
        .lblShields.Caption = hsbShields.Value
        .lblSpace(3).Caption = lblShieldSpc.Caption
    End With
    
    Craft.Shields = hsbShields.Value
    frmEquipment.chkUpdate.Value = 1
    hsbShields.Enabled = True

End Sub

Private Sub lblSpaceLeft_Change()
    Call FormatLabel(lblSpaceLeft)
End Sub

Private Sub vsbArmor_Change(Index As Integer)
    
    If chkBalance.Value Then
        If Index = 3 Then vsbArmor(4).Value = vsbArmor(3).Value
        If Index = 4 Then vsbArmor(3).Value = vsbArmor(4).Value
    End If
    
    lblArmor(Index).Caption = vsbArmor(Index).Value
    Craft.Armor(Index) = vsbArmor(Index).Value
    Call UpdateArmor
    
End Sub
