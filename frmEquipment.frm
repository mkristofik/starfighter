VERSION 5.00
Begin VB.Form frmEquipment 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modify Equipment"
   ClientHeight    =   5175
   ClientLeft      =   3870
   ClientTop       =   3300
   ClientWidth     =   7590
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmEquipment.frx":0000
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   7590
   Begin VB.ListBox lstPrint 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      ItemData        =   "frmEquipment.frx":0442
      Left            =   5160
      List            =   "frmEquipment.frx":0444
      Sorted          =   -1  'True
      TabIndex        =   32
      Top             =   3720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chkSave 
      Caption         =   "S"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4560
      TabIndex        =   31
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chkLoad 
      Caption         =   "L"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5040
      TabIndex        =   30
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox lstNum 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      ItemData        =   "frmEquipment.frx":0446
      Left            =   6720
      List            =   "frmEquipment.frx":0448
      Sorted          =   -1  'True
      TabIndex        =   29
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.HScrollBar hsbSpeed 
      Height          =   255
      LargeChange     =   2
      Left            =   5040
      Max             =   10
      TabIndex        =   8
      Tag             =   "0"
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CheckBox chkUpdate 
      Caption         =   "Update"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6360
      TabIndex        =   26
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Ca&ncel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   1320
      Width           =   855
   End
   Begin VB.ListBox lstCrits 
      Appearance      =   0  'Flat
      Height          =   2370
      Index           =   2
      ItemData        =   "frmEquipment.frx":044A
      Left            =   2640
      List            =   "frmEquipment.frx":044C
      MultiSelect     =   2  'Extended
      TabIndex        =   5
      Top             =   2280
      Width           =   2175
   End
   Begin VB.ListBox lstCrits 
      Appearance      =   0  'Flat
      Height          =   1200
      Index           =   4
      ItemData        =   "frmEquipment.frx":044E
      Left            =   5160
      List            =   "frmEquipment.frx":0450
      MultiSelect     =   2  'Extended
      TabIndex        =   7
      Top             =   2280
      Width           =   2175
   End
   Begin VB.ListBox lstCrits 
      Appearance      =   0  'Flat
      Height          =   1200
      Index           =   3
      ItemData        =   "frmEquipment.frx":0452
      Left            =   120
      List            =   "frmEquipment.frx":0454
      MultiSelect     =   2  'Extended
      TabIndex        =   6
      Top             =   2280
      Width           =   2175
   End
   Begin VB.ListBox lstCrits 
      Appearance      =   0  'Flat
      Height          =   1200
      Index           =   1
      ItemData        =   "frmEquipment.frx":0456
      Left            =   4080
      List            =   "frmEquipment.frx":0458
      MouseIcon       =   "frmEquipment.frx":045A
      MultiSelect     =   2  'Extended
      TabIndex        =   4
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6600
      TabIndex        =   9
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.ListBox lstEquipment 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      ItemData        =   "frmEquipment.frx":089C
      Left            =   120
      List            =   "frmEquipment.frx":089E
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblCurSpeed 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Speed: 0"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5760
      TabIndex        =   28
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblMaxSpeed 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Max: 10"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5880
      TabIndex        =   27
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label lblCritsLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "12"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   25
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label lblCritsLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "6"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   1680
      TabIndex        =   24
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label lblCritsLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "6"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   1680
      TabIndex        =   23
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label lblCritsLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "6"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   22
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label lblRW 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "RW:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   21
      Top             =   4800
      Width           =   495
   End
   Begin VB.Label lblLW 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "LW:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   20
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label lblF 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "F:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label lblC 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "C:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label lblTotCrits 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "30"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   17
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label lblCrits 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Criticals Remaining:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label lblSpaceLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   15
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label lblSpace 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Space Left:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label lblRightWing 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Right Wing"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5520
      TabIndex        =   13
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lblFuselage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Fuselage"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      TabIndex        =   12
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblLeftWing 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Left Wing"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblCockpit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cockpit"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmEquipment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim counter As Integer, isAdd As Boolean, curRec As Integer, Warheads As Integer

Private Sub chkLoad_Click()

    Dim i As Integer, j As Integer, count As Integer, high As Integer
    If chkLoad.Value = 0 Then Exit Sub
    
    counter = 4 ' Set initial counter so cockpit standard items aren't touched
    Do While Craft.Criticals(count).Location
       For i = 0 To lstEquipment.ListCount - 1
            If lstEquipment.ItemData(i) = Craft.Criticals(count).recNum Then
                lstEquipment.ListIndex = i
                Call cmdAdd_Click
                'counter = Abs(Craft.Criticals(count).idNum)
                If Abs(counter) + 1 > high Then high = Abs(counter) + 1
                Call lstCrits_MouseUp(Craft.Criticals(count).Location, 1, 0, 1, 1)
                Call DeselectAll
                Exit For
            End If
        Next i
        
        count = count + 1
    Loop
    counter = high
    chkLoad.Value = 0

End Sub

Private Sub chkSave_Click()

    Dim i As Integer, j As Integer, count As Integer
    
    If chkSave.Value = 0 Then Exit Sub
    
    lstNum.Clear
    count = 0
    
    For i = 1 To 4
        For j = 0 To lstCrits(i).ListCount - 1
            If i <> 1 Or j >= 4 Then
                Call GetWeaponInfo(lstCrits(i).List(j))
                If curRec <> 900 Or InStr(lstCrits(i).List(j), "Isometal") Then
                    If NotUsed(lstCrits(i).ItemData(j)) Then
                        Craft.Criticals(count).Location = i
                        Craft.Criticals(count).idNum = lstCrits(i).ItemData(j)
                        lstNum.AddItem lstCrits(i).ItemData(j)
                        If curRec = 900 Then curRec = 0 ' Isometal
                        Craft.Criticals(count).recNum = curRec
                        count = count + 1
                    End If
                End If
            End If
        Next j
    Next i
    
    For i = count To 27
        Craft.Criticals(i).Location = 0
    Next i
    chkSave.Value = 0

End Sub

Private Function NotUsed(num As Integer) As Boolean

    Dim i As Integer, retval As Boolean
    
    retval = True
    For i = 0 To lstNum.ListCount - 1
        If num = lstNum.List(i) Then
            retval = False
            Exit For
        End If
    Next i
    NotUsed = retval

End Function

Private Sub chkUpdate_Click()
    
    Dim i As Integer, j As Integer, num As Integer, temp As Integer
    
    If chkUpdate.Value = 0 Then Exit Sub
    If lstEquipment.ListCount = 0 Then Call ReloadList ' Reload the equipment list
    
    ' Shields
    For i = 0 To lstCrits(2).ListCount - 1
        If lstCrits(2).ItemData(i) = 900 Then
            lstCrits(2).Selected(i) = True
            Call lstCrits_MouseUp(2, 1, 0, 1, 1) ' Select all shield criticals
            Exit For
        End If
    Next i
    
    num = RoundUp(frmArmor.hsbShields.Value / 50)
    If num <> lstCrits(2).SelCount Then
        Weapon.WeapSpace = 0
        Call cmdRemove_Click ' Remove all the shield criticals
         
        Weapon.MaxNum = 0 ' Add the new shield criticals
        Weapon.Criticals = num
        Weapon.WeapName = "Shields"
        temp = counter
        counter = 900
        Call AddStuff(2)
        counter = temp
        cmdRemove.Enabled = False
    End If
    
    Call DeselectAll
    frmArmor.hsbShields.Tag = num ' frmArmor needs to know current # of shield criticals
    Call SetMaxSpeed
    
    ' ********************
    ' Hyperdrive
    For i = 0 To lstCrits(2).ListCount - 1
        If lstCrits(2).List(i) = "Hyperdrive" Then
            lstCrits(2).Selected(i) = True
            Call lstCrits_MouseUp(2, 1, 0, 1, 1) ' Select all hyperdrive criticals
            Exit For
        End If
    Next i
    
    If lstCrits(2).SelCount Then
        num = RoundUp(frmMain.cboTotSpc.Text / 20)
        If num <> lstCrits(2).SelCount Then
            Call cmdRemove_Click ' Remove all hyperdrive criticals
            If num > lblCritsLeft(2).Caption Then
                MsgBox "Number of fuselage criticals exceeded." & vbCrLf & _
                    "Removing hyperdrive.", vbInformation
            Else
                Call GetWeaponInfo("Hyperdrive")
                Call AddEngine(2)
                cmdRemove.Enabled = False
            End If
        End If
    End If
    
    Call DeselectAll
    
    ' ********************
    ' Techbase stuff (if an item doesn't belong, remove it)
    For i = 1 To 4
        For j = 0 To lstCrits(i).ListCount - 1
            If i = 1 And j = 0 Then j = 4
            If j > lstCrits(i).ListCount - 1 Then Exit For
            
            lstCrits(i).Selected(j) = True
            Call lstCrits_MouseUp(i, 1, 0, 1, 1)
            If curRec = 0 Then
                Call cmdRemove_Click
                j = j - 1
            End If
        Next j
    Next i
    
    Call DeselectAll
    
    ' ********************
    ' Wing change stuff (if an item is on a wing, remove it)
    If Craft.Wings = 0 Then
        For i = 3 To 4
            Do While lstCrits(i).ListCount > 0
                If lstCrits(i).ItemData(0) <> 0 Then  ' Don't accidentally remove Isometal.
                    lstCrits(i).Selected(0) = True
                    Call lstCrits_MouseUp(i, 1, 0, 1, 1)
                    Call cmdRemove_Click
                ' If Isometal is present, remove the next item in the list (if any).
                ElseIf lstCrits(i).ListCount > 1 Then
                    lstCrits(i).Selected(1) = True
                    Call lstCrits_MouseUp(i, 1, 0, 1, 1)
                    Call cmdRemove_Click
                Else
                    Exit Do
                End If
            Loop
        Next i
    End If
    
    ' If going from tri-wing to standard or wingless...
    If Craft.Wings <> 3 Then
        Do While lstCrits(2).ListCount > 12
            lstCrits(2).Selected(lstCrits(2).ListCount - 1) = True
            Call lstCrits_MouseUp(2, 1, 0, 1, 1)
            Call cmdRemove_Click
        Loop
        
        If lstCrits(2).Height > 2700 Then
            lstCrits(2).Height = 2370
            lblCritsLeft(2).Caption = lblCritsLeft(2).Caption - 2
        End If
    ' If switching to tri-wing...
    Else
        If lstCrits(2).Height < 2400 Then
            lstCrits(2).Height = 2760
            lblCritsLeft(2).Caption = lblCritsLeft(2).Caption + 2
        End If
    End If
    
    Call DeselectAll
    cmdRemove.Enabled = False
    chkUpdate.Value = 0
    
End Sub

Private Sub SetMaxSpeed()

    Dim totSpc As Integer
    
    ' Set the maximum speed
    If hsbSpeed.Tag = 0 Then
        hsbSpeed.max = 0
    Else
        Get #2, CInt(hsbSpeed.Tag), Engine
        totSpc = frmMain.cboTotSpc.Text
        hsbSpeed.max = Int((400 - totSpc * Engine.SpeedMod) / totSpc)
        If hsbSpeed.max > 10 Then hsbSpeed.max = 10
    End If
    lblMaxSpeed.Caption = "Max: " & hsbSpeed.max

End Sub

Private Sub cmdAdd_Click()

    Dim i As Integer
    
    Call lstEquipment_Click
    If cmdAdd.Tag = "Y" Then
        Call AddIso
    Else
        ' Set up the different pointers for valid and invalid locations
        For i = 1 To 4
            If InStr(Weapon.Locations, CStr(i)) Then
                lstCrits(i).MousePointer = 2
            Else
                lstCrits(i).MousePointer = 12
            End If
        Next i
        isAdd = True
        cmdCancel.Enabled = True
        cmdRemove.Enabled = False
    End If

End Sub

Private Sub AddIso()

    Dim i As Integer, flag As Boolean
    
    If frmMain.lblInternalType.Caption = "Isometal" Then
        MsgBox "This craft already uses Isometal internal structure.", vbExclamation
        Exit Sub
    End If
    
    ' Check for at least one critical left in each location.
    For i = 1 To 4
        If lblCritsLeft(i).Caption = 0 Then
            flag = True
            MsgBox "Number of critical slots exceeded.", vbCritical
            Exit For
        End If
    Next i
        
    If Not flag Then
        Call DeselectAll
        For i = 1 To 4
            lstCrits(i).AddItem "Isometal Internal Structure"
            lstCrits(i).Selected(lstCrits(i).NewIndex) = True
            lblCritsLeft(i).Caption = CInt(lblCritsLeft(i).Caption) - 1
        Next i
        frmMain.lblInternalType.Caption = "Isometal"
        Call cmdCancel_Click
        cmdRemove.Enabled = True
    End If

End Sub

Private Sub cmdCancel_Click()

    ' Restore the settings to normal
    Dim i As Integer
    
    isAdd = False
    For i = 1 To 4
        lstCrits(i).MousePointer = 0
    Next i
    cmdCancel.Enabled = False
    cmdRemove.Enabled = False

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRemove_Click()

    Dim i As Integer, j As Integer, C As Integer, flag As Boolean
    C = 32767
    
    For i = 1 To 4
        j = 0
        Do While j <= lstCrits(i).ListCount - 1
            If lstCrits(i).Selected(j) Then
                ' Grab the ID before the item gets removed
                If Not flag Then C = lstCrits(i).ItemData(j)
                flag = True
                
                lstCrits(i).RemoveItem (j)
                lblCritsLeft(i).Caption = CInt(lblCritsLeft(i).Caption) + 1
            Else
                j = j + 1
            End If
        Loop
    Next i
    
    
    If C = 0 Then
        frmMain.lblInternalType = "Standard" ' Removing Isometal
    ElseIf C = 900 Then ' Removing Shields
        Weapon.WeapSpace = 0
        If chkUpdate.Value Then
            frmArmor.chkUpdate.Value = 1 ' If this is an update
        Else
            frmArmor.hsbShields.Value = 0 ' Otherwise the shields are being removed
            Weapon.WeapSpace = 0
        End If
    ElseIf C < 0 Then ' Removing an engine component
        Dim s As Integer
        s = Engine.EngType
        Call RemoveEngine
    End If
    
    frmMain.lblSpace(5).Caption = CDbl(frmMain.lblSpace(5).Caption) - Weapon.WeapSpace
    Call cmdCancel_Click ' Return settings to normal
    cmdAdd.Enabled = False ' Force the user to choose from the equipment box again
    
    ' If removing co-pilot, reset max cockpit armor to 5
    If RTrim(Weapon.WeapName) = "Co-Pilot" Then frmArmor.vsbArmor(1).Min = 5
    
    ' Warheads
    If InStr(Weapon.Options, 1) Then Warheads = Warheads - 1

End Sub

Private Sub RemoveEngine()

    Dim temp As Integer
    
    Select Case Engine.EngType
        Case 0 ' Standard Engine
            hsbSpeed.Tag = 0
            Call SetMaxSpeed
            frmMain.lblEngineType.Caption = "None"
        Case 1 ' Goofy Engine
            hsbSpeed.Tag = 0
            Call SetMaxSpeed
            frmMain.lblEngineType.Caption = "None"
            frmMain.lblHDriveType.Caption = "None"
            frmMain.lblSpace(2).Caption = 0
        Case 2 ' Hyperdrive
            frmMain.lblHDriveType.Caption = "None"
            frmMain.lblSpace(2).Caption = 0
        Case 3 ' After Burner
            If RTrim(Engine.EngName) = "SLAM System" Then
                Weapon.WeapSpace = 2
            ElseIf RTrim(Engine.EngName) = "After Burner (10)" Then
                Weapon.WeapSpace = 1
            End If
            frmMain.lblEngineType.Tag = 0
    End Select

End Sub

Private Sub Form_Activate()
    frmMain.Enabled = False
End Sub

Private Sub ReloadList()

    Dim numRecs As Integer, i As Integer, j As Integer
    
    j = lstEquipment.ListIndex
    lstEquipment.Clear
    
    ' Add weapons and equipment to the list
    numRecs = LOF(1) / Len(Weapon)
    For i = 5 To numRecs ' Don't show the first four (required cockpit stuff)
        Get #1, i, Weapon
        If Not Weapon.Deleted Then
            If Weapon.TechBase = 0 Or InStr(Weapon.TechBase, Craft.TechBase) Then
                lstEquipment.AddItem RTrim(Weapon.WeapName)
                lstEquipment.ItemData(lstEquipment.NewIndex) = i
            End If
        End If
    Next i
    
    ' Add engine stuff to the list
    numRecs = LOF(2) / Len(Engine)
    For i = 1 To numRecs
        Get #2, i, Engine
        If Not Engine.Deleted Then
            If Engine.TechBase = 0 Or InStr(Engine.TechBase, Craft.TechBase) Then
                lstEquipment.AddItem RTrim(Engine.EngName)
                lstEquipment.ItemData(lstEquipment.NewIndex) = -i ' Negative #'s are engines
            End If
        End If
    Next i
    
    lstEquipment.AddItem "Isometal Internal Structure"
    lstEquipment.ListIndex = j

End Sub

Private Sub Form_Load()

    Dim i As Integer
    
    Open App.Path & "\weapons.db" For Random As #1 Len = Len(Weapon)
    Open App.Path & "\engines.db" For Random As #2 Len = Len(Engine)

    counter = 0
    Warheads = 0
    
    For i = 1 To 4
        Get #1, i, Weapon
        Call AddStuff(1)
    Next i
    Call DeselectAll
    cmdRemove.Enabled = False

End Sub

Private Function GetNumRecords()
    GetNumRecords = LOF(1) / Len(Weapon)
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Call DeselectAll
    cmdAdd.Enabled = False
    cmdRemove.Enabled = False
    cmdCancel.Enabled = False
    frmMain.Enabled = True
    frmArmor.chkUpdate.Value = 1
    Cancel = NotNewShip
    Me.Hide
    
End Sub

Private Sub hsbSpeed_Change()

    Craft.Speed = hsbSpeed.Value
    lblCurSpeed.Caption = "Speed: " & hsbSpeed.Value
    frmMain.lblSpeed.Caption = hsbSpeed.Value

End Sub

Private Sub lblCritsLeft_Change(Index As Integer)
   
    ' Update the criticals total
    Dim i As Integer
    
    lblTotCrits.Caption = 0
    For i = 1 To 4
        If lblCritsLeft(i).Visible Then lblTotCrits.Caption = CInt(lblTotCrits.Caption) + _
            CInt(lblCritsLeft(i).Caption)
    Next i

End Sub

Private Sub lblSpaceLeft_Change()
    Call FormatLabel(lblSpaceLeft)
End Sub

Private Sub lstCrits_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X _
    As Single, Y As Single)

    Dim i As Integer, C As Integer, loc As Integer
    
    If Button <> 1 Then Call cmdCancel_Click ' If user clicks right button, cancel operation
        
    If isAdd Then
        If lstCrits(Index).MousePointer = 12 Then
            MsgBox "Equipment not allowed in this location.", vbCritical
            Call DeselectAll
            Exit Sub
        ElseIf curRec < 0 Then
            Call AddEngine(Index)
        Else
            If RTrim(Weapon.WeapName) = "S Foil" Then
                Call AddStuff(3, 4)
            Else
                Call AddStuff(Index)
            End If
        End If
    Else
        If lstCrits(Index).SelCount Then
            C = lstCrits(Index).ItemData(lstCrits(Index).ListIndex)
            
            If C = 900 Or C <= 0 Then
                cmdRemove.Enabled = True
                Weapon.WeapSpace = 0 ' Isometal and shield space accounted for elsewhere
                'curRec = c
                'If c = 0 Then curRec = 900
                Call GetWeaponInfo(lstCrits(Index).List(lstCrits(Index).ListIndex))
            ElseIf C > 4 Then
                'curRec = c
                cmdRemove.Enabled = True
                Call GetWeaponInfo(lstCrits(Index).List(lstCrits(Index).ListIndex))
            Else
                cmdRemove.Enabled = False ' First four cockpit items cannot be removed
            End If
            
            ' Deselect everything before selecting the current item's criticals
            Call DeselectAll
            For loc = 1 To 4
                For i = 0 To lstCrits(loc).ListCount - 1
                    If lstCrits(loc).ItemData(i) = C Then lstCrits(loc).Selected(i) = True
                Next i
            Next loc
        End If
    End If

End Sub

Private Sub DeselectAll()

    Dim i As Integer, j As Integer
    For i = 1 To 4
        For j = 0 To lstCrits(i).ListCount - 1
            lstCrits(i).Selected(j) = False
        Next j
    Next i

End Sub

Private Sub AddStuff(ByVal loc As Integer, Optional ByVal loc2 As Integer = 0)
   
    Dim i As Integer, Crits As Integer
    
    Crits = Weapon.Criticals
    
    If OverMax Then Exit Sub
    
    If loc2 Then
        Crits = RoundUp(Weapon.Criticals / 2) ' Split # criticals in half (round up)
        If Crits > lblCritsLeft(loc2).Caption Then
            MsgBox "Number of critical slots exceeded.", vbCritical
            Call cmdCancel_Click
            Exit Sub
        End If
    End If
    
    Dim a As Integer
    a = lblCritsLeft(loc).Caption
    If Crits > lblCritsLeft(loc).Caption Then
        MsgBox "Number of critical slots exceeded.", vbCritical
        Call cmdCancel_Click
        Exit Sub
    End If
   
    If counter <> 900 Then
        counter = counter + 1
        If counter = 900 Then counter = 901 ' 900 is a reserved ID for shields
    End If
    Call DeselectAll
    
    Do While loc
        For i = 1 To Crits
            lstCrits(loc).AddItem RTrim(Weapon.WeapName)
            lstCrits(loc).ItemData(lstCrits(loc).NewIndex) = counter
            lstCrits(loc).Selected(lstCrits(loc).NewIndex) = True
        Next i
    
        lblCritsLeft(loc).Caption = CInt(lblCritsLeft(loc).Caption) - Crits
                
        loc = loc2 ' Repeat for the other split location if necessary
        loc2 = 0
    Loop
    
    ' Update the space total on the main form
    frmMain.lblSpace(5).Caption = CDbl(frmMain.lblSpace(5).Caption) + Weapon.WeapSpace
    Call cmdCancel_Click 'Return the settings to normal
    cmdRemove.Enabled = True
    
    ' If adding co-pilot, set max cockpit armor to 6
    If RTrim(Weapon.WeapName) = "Co-Pilot" Then frmArmor.vsbArmor(1).Min = 6
    
    ' Warheads
    If InStr(Weapon.Options, 1) Then Warheads = Warheads + 1
    
End Sub

Private Function OverMax() As Boolean

    Dim i As Integer, j As Integer, C As Integer
    For i = 1 To 4
        For j = 0 To lstCrits(i).ListCount - 1
            If lstCrits(i).List(j) = RTrim(Weapon.WeapName) Then C = C + 1
        Next j
    Next i
    
    If (C = (Weapon.MaxNum + 1) * Weapon.Criticals And Weapon.MaxNum) Or (InStr(Weapon.Options, 1) _
        And Warheads = 4) Then
            MsgBox "Maximum quanitity exceeded for this equipment.", vbExclamation
            OverMax = True
    Else
        OverMax = False
    End If

End Function

Private Sub AddEngine(ByVal loc As Integer)

    Dim i As Integer, Crits As Integer, C As Integer, first As Boolean
    
    If Invalid Then
        Call cmdCancel_Click
        Exit Sub
    End If
    
    Crits = Engine.Criticals
    If Engine.EngType = 2 Then Crits = RoundUp(frmMain.cboTotSpc.Text / Crits) ' Hyperdrive
    first = True
    
    If loc <> 2 Then
        Crits = RoundUp(Crits / 2) ' Split # criticals in half (round up)
        If Crits > lblCritsLeft(3).Caption Or Crits > lblCritsLeft(4).Caption Then
            MsgBox "Number of critical slots exceeded.", vbCritical
            Call cmdCancel_Click
            Exit Sub
        End If
    Else
        If Crits > lblCritsLeft(2).Caption Then
            MsgBox "Number of critical slots exceeded.", vbCritical
            Call cmdCancel_Click
            Exit Sub
        End If
    End If
   
    C = -counter
    counter = counter + 1
    Call DeselectAll
    
    Do While loc
        For i = 1 To Crits
            lstCrits(loc).AddItem RTrim(Engine.EngName)
            lstCrits(loc).ItemData(lstCrits(loc).NewIndex) = C
            lstCrits(loc).Selected(lstCrits(loc).NewIndex) = True
        Next i
    
        lblCritsLeft(loc).Caption = CInt(lblCritsLeft(loc).Caption) - Crits
                
        If loc = 3 And first Then
            loc = 4 ' Repeat for the other wing if necessary
            first = False
        ElseIf loc = 4 And first Then
            loc = 3 ' Repeat for the other wing if necessary
            first = False
        Else
            loc = 0
        End If
    Loop
    
    If Engine.EngType <= 1 Then
        ' Tag the speed scrollbar with the engine's record number
        hsbSpeed.Tag = -curRec
        Call SetMaxSpeed
        
        ' Update the main form
        frmMain.lblEngineType.Caption = RTrim(Engine.EngName) ' Engine Name
        If Engine.EngType Then frmMain.lblHDriveType.Caption = RTrim(Engine.EngName)
    
    ElseIf Engine.EngType = 2 Then
        If RTrim(Engine.EngName) = "Hyperdrive" Then
            frmMain.lblHDriveType.Caption = "Standard Hyperdrive"
        Else
            frmMain.lblHDriveType.Caption = RTrim(Engine.EngName)
        End If
        
        frmMain.lblSpace(2).Caption = Crits
    Else
        If RTrim(Engine.EngName) = "SLAM System" Then
            frmMain.lblSpace(5).Caption = CDbl(frmMain.lblSpace(5).Caption) + 2
        Else
            frmMain.lblSpace(5).Caption = CDbl(frmMain.lblSpace(5).Caption) + 1
        End If
        frmMain.lblEngineType.Tag = -curRec  ' Tag the engine type label w/ burner's rec #
        frmMain.lblSpeed.Caption = frmMain.lblSpeed.Caption & " "
    End If
        
    Call cmdCancel_Click 'Return the settings to normal
    cmdRemove.Enabled = True

End Sub

Private Sub lstEquipment_Click()
          
    cmdAdd.Enabled = True
    curRec = lstEquipment.ItemData(lstEquipment.ListIndex)
    
    If curRec > 0 Then
        Get #1, curRec, Weapon
        cmdAdd.Tag = "N"
    ElseIf curRec < 0 Then
        Get #2, -curRec, Engine
        If Engine.EngType <= 1 Then
            Weapon.Locations = "234" ' Engines aren't allowed in the cockpit
        Else
            Weapon.Locations = "2" ' Accessories are fuselage only
        End If
        cmdAdd.Tag = "N"
    Else
        cmdAdd.Tag = "Y" ' Error trap - Isometal has no record in file
    End If

End Sub

Private Sub GetWeaponInfo(ByVal searchStr As String)
' Search the list for the selected weapon

    If searchStr = "Shields" Or searchStr = "Isometal Internal Structure" Then
        curRec = 900
        Exit Sub
    End If
    
    Dim first As Integer, last As Integer, mid As Integer, isDone As Boolean
    first = 0
    last = lstEquipment.ListCount - 1
    
    Do While last >= first And isDone = False
        mid = Int((first + last) / 2)
        
        If LCase$(searchStr) < LCase$(lstEquipment.List(mid)) Then
            last = mid - 1
        Else
            If LCase$(searchStr) > LCase$(lstEquipment.List(mid)) Then
                first = mid + 1
            Else
                isDone = True
            End If
        End If
    Loop
    
    If isDone Then
        curRec = lstEquipment.ItemData(mid)
        If curRec < 0 Then
            Get #2, -curRec, Engine
            Weapon.Options = 0
        ElseIf curRec > 0 Then
            Get #1, curRec, Weapon
        End If
    Else
        curRec = 0
        ' Still need to get info even if not in the list
        If FileSearch(searchStr, 1) = 0 Then
            Dim a As Integer
            a = FileSearch(searchStr, 2)
        End If
    End If
        
End Sub

Private Function FileSearch(searchStr As String, fileNum As Integer) As Integer

    Dim myStr As String, i As Integer
    
    i = 1
    Do Until EOF(fileNum) Or myStr = searchStr
        If fileNum = 1 Then
            Get #1, i, Weapon
        Else
            Get #2, i, Engine
        End If
        i = i + 1
    Loop
    
    If myStr = searchStr Then
        FileSearch = 1
    Else
        FileSearch = 0
    End If
    
End Function

Private Function Invalid() As Boolean

    If frmMain.lblEngineType.Caption <> "None" And Engine.EngType <= 1 Then
        MsgBox "Must remove current engine first.", vbInformation
        Invalid = True
    ElseIf frmMain.lblHDriveType.Caption <> "None" And Engine.EngType = 2 Then
        MsgBox "Hyperdrive already present.", vbInformation
        Invalid = True
    ElseIf frmMain.lblEngineType.Tag = 2 And RTrim(Engine.EngName) = "SLAM System" Then
        MsgBox "SLAM System already present.", vbInformation
        Invalid = True
    Else
        Invalid = False
    End If

End Function

Private Sub lstEquipment_DblClick()
       
    frmEquipInfo.Tag = lstEquipment.ItemData(lstEquipment.ListIndex)
    frmEquipInfo.Show
    
End Sub
