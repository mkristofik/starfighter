VERSION 5.00
Begin VB.Form frmChgShd 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modify Shields"
   ClientHeight    =   1245
   ClientLeft      =   1005
   ClientTop       =   2430
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.HScrollBar hsbShields 
      Height          =   255
      LargeChange     =   20
      Left            =   120
      Max             =   200
      SmallChange     =   10
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
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
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   1215
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
      Left            =   720
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblCurShields 
      Alignment       =   2  'Center
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
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmChgShd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim current, trap As Integer

Private Sub cmdCancel_Click()
' Click the Cancel button.
    frmChgShd.Visible = False
    Unload Me
End Sub

Public Sub cmdOK_Click()
' Click the OK button.
    frmChgShd.Visible = False
    Craft.Shields = hsbShields.Value
    With frmArmor
        .lblShields = Craft.Shields
        .lblShieldSpc = FormatNumber(Craft.Shields / 20, 2)
    End With
    
    With frmMain
        .lblShields = Craft.Shields
        .lblShieldSpc = FormatNumber(Craft.Shields / 20, 2)
    End With
    
    
    If trap = 0 Then ShieldCriticals
    ' trap fixes a bug that caused the shield criticals to be counted twice
                                    
    frmArmor.UpdateArmor
    frmArmor.Visible = True     ' Prevents a bug that causes the armor form
                                ' to disappear.
    Unload Me
End Sub

Private Sub Form_Load()
' Initialize Variables.
    trap = 0
    current = Int((Craft.Shields - 1) / 50) + 1
    
    Dim MaxShields, RealMax As Integer
    MaxShields = Craft.TotalSpace * 2
    RealMax = (12 - TotalCrits.F + current) * 50
    If RealMax < MaxShields Then MaxShields = RealMax
    
    If Craft.Techbase = "H" Or Craft.Techbase = "P" Then MaxShields = MaxShields * 2
    lblMaxShields = "Max: " & MaxShields
    hsbShields.Max = MaxShields
    
    If Craft.Shields > MaxShields Then
        hsbShields.Value = MaxShields
    Else
        hsbShields.Value = Craft.Shields
    End If
End Sub

Private Sub hsbShields_Change()
' Move the Shield scrollbar, must be a value divisible by 10.
    Do While hsbShields.Value Mod 10 <> 0
        hsbShields.Value = hsbShields.Value - 1
    Loop
    
    lblCurShields.Caption = hsbShields.Value
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        cmdOK_Click
    End If
End Sub

Private Sub ShieldCriticals()
    Dim newCrit, i, j, k As Integer
    newCrit = Int((Craft.Shields - 1) / 50) + 1
    
    For i = 0 To 11
        If RTrim(frmEquipment.lstFuselage.List(i)) = "Shields" Then Exit For
    Next i
    
    ' Remove the Shield criticals.
    If i <> 12 Then
        TotalCrits.F = TotalCrits.F - current
        With frmEquipment
            .lstFuselage.RemoveItem (i)
            .lstFCriticals.RemoveItem (i)
        End With
    
        For j = i + 1 To 11
            k = j + 1
            With Criticals(2, j)
               .WeapName = Criticals(2, k).WeapName
               .NumCrits = Criticals(2, k).NumCrits
               .WeapSpace = Criticals(2, k).WeapSpace
            End With
        Next j
    
        With Criticals(2, k)
            .NumCrits = 0
            .WeapName = ""
            .WeapSpace = 0
        End With
    End If
    
    ' Add back the new amount.
    If newCrit Then
        With frmEquipment
            .lstFuselage.AddItem ("Shields")
            .lstFCriticals.AddItem (newCrit)
        End With
        
        k = frmEquipment.lstFuselage.ListCount
        
        With Criticals(2, k)
            .NumCrits = newCrit
            .WeapName = "Shields"
            .WeapSpace = Craft.Shields / 20
        End With
        
        TotalCrits.F = TotalCrits.F + newCrit
    End If
    
    Unload frmEquipment
    trap = 1
End Sub
