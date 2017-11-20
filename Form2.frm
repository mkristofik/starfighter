VERSION 5.00
Begin VB.Form frmArmor 
   BackColor       =   &H80000009&
   Caption         =   "Modify Armor"
   ClientHeight    =   3840
   ClientLeft      =   936
   ClientTop       =   1368
   ClientWidth     =   7296
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   7296
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdShields 
      Caption         =   "Modify &Shields"
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
      Left            =   4920
      TabIndex        =   29
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
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
      Left            =   6120
      TabIndex        =   24
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton cmdMaxArmor 
      Caption         =   "&Maximize Armor"
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
      Left            =   4920
      TabIndex        =   23
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CheckBox chkBalance 
      BackColor       =   &H80000009&
      Caption         =   "&Balance Armor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
   Begin VB.ComboBox cboArmorType 
      Height          =   315
      ItemData        =   "Form2.frx":0000
      Left            =   240
      List            =   "Form2.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   480
      Width           =   1575
   End
   Begin VB.VScrollBar vsbRWArmor 
      Height          =   855
      LargeChange     =   2
      Left            =   3600
      Max             =   0
      Min             =   30
      TabIndex        =   3
      Top             =   2280
      Width           =   375
   End
   Begin VB.VScrollBar vsbFArmor 
      Height          =   855
      LargeChange     =   2
      Left            =   2400
      Max             =   0
      Min             =   40
      TabIndex        =   2
      Top             =   2280
      Width           =   375
   End
   Begin VB.VScrollBar vsbLWArmor 
      Height          =   855
      LargeChange     =   2
      Left            =   1200
      Max             =   0
      Min             =   30
      TabIndex        =   1
      Top             =   2280
      Width           =   375
   End
   Begin VB.VScrollBar vsbCArmor 
      Height          =   855
      LargeChange     =   5
      Left            =   2400
      Max             =   0
      Min             =   5
      TabIndex        =   0
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblShields 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0"
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
         Size            =   7.8
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
         Size            =   7.8
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
         Size            =   7.8
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
         Size            =   7.8
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
         Size            =   7.8
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
         Size            =   7.8
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
         Size            =   7.8
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
         Size            =   7.8
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
         Size            =   7.8
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
         Size            =   7.8
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
         Size            =   7.8
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
   Begin VB.Label lblArmor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Armor Factor:"
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
   Begin VB.Label lblRWArmor 
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
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3600
      TabIndex        =   11
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lblFArmor 
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
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lblLWArmor 
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
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lblCArmor 
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
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
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
         Size            =   7.8
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
         Size            =   7.8
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
         Size            =   7.8
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
         Size            =   7.8
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


Private Sub cboArmorType_Click()
' Choose Armor Type.
    UpdateArmor
End Sub

Private Sub chkBalance_Click()
' Check/uncheck the Balance Armor checkbox.
    If chkBalance.Value Then vsbLWArmor.Value = vsbRWArmor.Value
    UpdateArmor
End Sub

Public Sub cmdClose_Click()
' Click on the Close button. (called also from main form)
    frmArmor.Visible = False
    Armor.C = vsbCArmor.Value
    Armor.F = vsbFArmor.Value
    Armor.LW = vsbLWArmor.Value
    Armor.RW = vsbRWArmor.Value
    Armor.Total = lblArmorFactor.Caption
    Armor.Size = lblArmorSpace.Caption
    
    With frmMain
        .lblArmor = Armor.Total
        .lblCArmor = Armor.C
        .lblFArmor = Armor.F
        .lblLWArmor = Armor.LW
        .lblRWArmor = Armor.RW
    End With
    
    frmMain.lblArmorType = Armor.ArmType
    frmMain.lblArmorSpc = FormatNumber(Armor.Size, 2)
    
    Unload Me
    CalcEngine
End Sub

Private Sub cmdMaxArmor_Click()
' Click on the Maximize Armor button.
    vsbCArmor.Value = vsbCArmor.Min
    vsbFArmor.Value = vsbFArmor.Min
    
    If Craft.Wings Then
        vsbLWArmor.Value = vsbLWArmor.Min
        vsbRWArmor.Value = vsbRWArmor.Min
    End If
    
    UpdateArmor
End Sub

Private Sub cmdShields_Click()
' Click on the Modify Shields button.
    frmChgShd.Visible = True
End Sub
Private Sub Form_Load()
    Dim AType As Integer
        
    lblShields.Caption = Craft.Shields
    lblShieldSpc.Caption = FormatNumber(Craft.Shields / 20, 2)
    
    If cboArmorType.ListCount = 4 And (Craft.TechBase = "H") Then
        cboArmorType.AddItem ("Clear Plast")
    End If
            
    If RTrim(Armor.ArmType) = "Standard" Then AType = 0
    If RTrim(Armor.ArmType) = "Didrate" Then AType = 1
    If RTrim(Armor.ArmType) = "Trinnium" Then AType = 2
    If RTrim(Armor.ArmType) = "Tri-Di Composite" Then AType = 3
    If RTrim(Armor.ArmType) = "Clear Plast" And Craft.TechBase = "H" _
        Then AType = 4
            
    cboArmorType.ListIndex = AType
    
    SetMaxArmor
    
    With Armor
        If .C > vsbCArmor.Min Then
            vsbCArmor.Value = vsbCArmor.Min
        Else
            vsbCArmor.Value = .C
        End If
        
        If .F > vsbFArmor.Min Then
            vsbFArmor.Value = vsbFArmor.Min
        Else
            vsbFArmor.Value = .F
        End If
        
        If .LW > vsbLWArmor.Min Then
            vsbLWArmor.Value = vsbLWArmor.Min
        Else
            vsbLWArmor.Value = .LW
        End If
        
        If .RW > vsbRWArmor.Min Then
            vsbRWArmor.Value = vsbRWArmor.Min
        Else
            vsbRWArmor.Value = .RW
        End If
    End With
    
    ' Hide the wings if fighter has no wings.
    If Craft.Wings = 0 Then
        lblLeftWing.Enabled = False
        lblLWArmor.Enabled = False
        vsbLWArmor.Enabled = False
        lblRightWing.Enabled = False
        lblRWArmor.Enabled = False
        vsbRWArmor.Enabled = False
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        cmdClose_Click
    End If
End Sub

Private Sub vsbCArmor_Change()
' Change the Cockpit armor.
    lblCArmor.Caption = vsbCArmor.Value
    UpdateArmor
End Sub

Private Sub vsbFArmor_Change()
' Change the Fuselage armor.
    lblFArmor.Caption = vsbFArmor.Value
    UpdateArmor
End Sub

Private Sub vsbLWArmor_Change()
' Change the Left Wing armor.
    If chkBalance.Value Then vsbRWArmor.Value = vsbLWArmor.Value
    lblLWArmor.Caption = vsbLWArmor.Value
    UpdateArmor
End Sub

Private Sub vsbRWArmor_Change()
' Change the Right Wing armor.
    If chkBalance.Value Then vsbLWArmor.Value = vsbRWArmor.Value
    lblRWArmor.Caption = vsbRWArmor.Value
    UpdateArmor
End Sub

Public Sub UpdateArmor()
' Update armor totals on both the Armor form and Main form.
    Dim Total As Integer
    Dim Used As Integer
    Dim left As Integer
    Dim ArmorType As Integer
    Dim C As Integer
    Dim F As Integer
    Dim LW As Integer
    Dim RW As Integer
    
    C = vsbCArmor.Value
    F = vsbFArmor.Value
    LW = vsbLWArmor.Value
    RW = vsbRWArmor.Value
    
    Used = C + F + LW + RW
    Total = Used
    Do While Total Mod 5 <> 0
        Total = Total + 1
    Loop
    
    left = Total - Used
    
    lblArmorFactor.Caption = Total
    lblArmorLeft.Caption = left
    If left <> 0 Then
        lblArmorLeft.ForeColor = vbRed
    Else
        lblArmorLeft.ForeColor = vbBlack
    End If
    
    ArmorType = cboArmorType.ListIndex
    
    Select Case ArmorType
        Case 0
            Armor.ArmType = "Standard"
            lblArmorSpace.Caption = FormatNumber(Total * 0.6, 2)
        Case 1
            Armor.ArmType = "Didrate"
            lblArmorSpace.Caption = FormatNumber(Total * 0.4, 2)
        Case 2
            Armor.ArmType = "Trinnium"
            lblArmorSpace.Caption = FormatNumber(Total * 0.4, 2)
        Case 3
            Armor.ArmType = "Tri-Di Composite"
            lblArmorSpace.Caption = FormatNumber(Total * 0.2, 2)
        Case 4
            Armor.ArmType = "Clear Plast"
            lblArmorSpace.Caption = FormatNumber(Total * 0.1, 2)
    End Select
    
    Armor.Size = lblArmorSpace.Caption
    
    ' Update total space information.
    lblSpaceLeft.Caption = FormatNumber(SpaceLeft, 2)
    If lblSpaceLeft.Caption >= 0 Then
        lblSpaceLeft.ForeColor = vbBlack
    Else
        lblSpaceLeft.ForeColor = vbRed
    End If
      
End Sub

Private Sub SetMaxArmor()
    Dim Mult As Integer
    If Craft.TechBase = "N" Or Craft.TechBase = "I" Then
        Mult = 2
    Else
        Mult = 4
    End If
    
    vsbCArmor.Min = 5
    If RTrim(Criticals(1, 5).WeapName) = "Co-pilot" Or RTrim(Criticals(1, 6).WeapName) = "Co-pilot" Then _
    vsbCArmor.Min = 6
        
    vsbFArmor.Min = Internal.F * Mult
    If Craft.Wings = 3 Then vsbFArmor.Min = vsbFArmor.Min * 1.5
    
    vsbLWArmor.Min = Internal.LW * Mult
    vsbRWArmor.Min = Internal.RW * Mult
End Sub
