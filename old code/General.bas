Attribute VB_Name = "General"
Type EngineInfo
    Rating As Integer
    EngType As String * 20
    Speed As Integer
    Maneuver As Integer
    Size As Single
    Criticals As Integer
End Type

Type ArmorInfo
    ArmType As String * 16
    C As Integer
    F As Integer
    LW As Integer
    RW As Integer
    Total As Integer
    Size As Single
End Type

Type GeneralInfo
    CraftName As String * 25
    Abbr As String * 6
    TotalSpace As Integer
    Shields As Integer
    HDrive As Integer
    TechBase As String * 1
    Wings As Integer
    Slam As Boolean
    Tag As Boolean
    Warheads As Integer
End Type

Type InternalInfo
    Size As Single
    F As Integer
    LW As Integer
    RW As Integer
    Iso As Boolean
End Type

Type CriticalInfo
    WeapName As String * 21
    NumCrits As Integer
    WeapSpace As Single
End Type

Type TotalCriticalData
    C As Integer
    F As Integer
    LW As Integer
    RW As Integer
End Type

Public Engine As EngineInfo
Public Armor As ArmorInfo
Public Craft As GeneralInfo
Public Internal As InternalInfo
Public Criticals(1 To 4, 1 To 12) As CriticalInfo
Public TotalCrits As TotalCriticalData
Private myMan As Single

Public Sub CalcEngine()
' Calculate relevant engine info.
    Dim ER, i As Integer
    ER = GetEngRating
    If ER > 400 Then
        MsgBox "Maximum engine rating exceeded.  Modifying Speed and Total Space to compensate" _
            , vbOKOnly, "Engine Overflow"
        Do While ER > 400
            Engine.Speed = Engine.Speed - 1
            If Engine.Speed < 3 Then
                Engine.Speed = 3
                Craft.TotalSpace = Craft.TotalSpace - 5
            End If
            ER = GetEngRating
        Loop
    End If
    
    frmMain.cboTotSpc.ListIndex = Craft.TotalSpace / 5 - 2
    frmMain.lblEngRating.Caption = ER
    Engine.Rating = ER
    Engine.Size = ER / 100
    frmMain.lblEngineSpc.Caption = FormatNumber(Engine.Size, 2)
    
    ' Special formatting for the speed.
    frmMain.lblSpeed.Caption = Engine.Speed
    
    Dim count, fastSpeed As Integer
    For count = 1 To 12
        If RTrim(Criticals(2, count).WeapName) = "SLAM System" Or RTrim(Criticals(2, count).WeapName) _
        = "After Burner (10)" Then Exit For
    Next count
        
    If count <> 13 Then
        frmMain.lblSpeed.Caption = CStr(Engine.Speed) + " / " + CStr(Engine.Speed * 2)
        
        If RTrim(Engine.EngType) = "Tzo Converter" Then
            fastSpeed = Int(Engine.Speed * -1.5) * -1
            frmMain.lblSpeed.Caption = frmMain.lblSpeed.Caption + " (" + CStr(fastSpeed) + _
            " / " + CStr(fastSpeed * 2) + ")"
        End If
    
    Else
    
        If RTrim(Engine.EngType) = "Tzo Converter" Then
            fastSpeed = Int(Engine.Speed * -1.5) * -1
            frmMain.lblSpeed.Caption = CStr(Engine.Speed) + " (" + CStr(fastSpeed) + ")"
        End If
    End If
    ' End special formatting section.
        
    frmMain.lblEngineType = Engine.EngType
    
    frmMain.lblManeuver.Caption = CalcManeuver
    Engine.Maneuver = frmMain.lblManeuver.Caption
    
    ' Find where the engine is located and update the space.
    For i = 1 To 12
        If InStr(Criticals(2, i).WeapName, "Engine") Or InStr(Criticals(2, i).WeapName, "Converter") _
        Then Exit For
    Next i
    
    Criticals(2, i).WeapSpace = Engine.Size
    
End Sub
    
Private Function CalcManeuver()
    Dim Man, left
    Dim strMan As String
    Man = myMan
    Man = Man - Engine.Speed
    Man = Man + Int(Craft.TotalSpace / -20) * -1
    
    left = SpaceLeft
    If left < 0 Then left = 0
    Man = Man - Int(left / (Craft.TotalSpace / -20)) * -1
    
    If Man >= 0 Then
        strMan = "+" & Man
    Else
        strMan = Man
    End If
    
    CalcManeuver = strMan
End Function

Private Function GetEngRating()
' Calculate the engine rating with this engine.
    Dim ER As Single
    
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
        ER = CInt(Craft.TotalSpace * Engine.Speed / 2)
        Do While ER Mod 5 <> 0
            ER = ER + 1
        Loop
    Else
        ER = Craft.TotalSpace * (Engine.Speed + Modifier)
    End If
    
    myMan = ManBase
    GetEngRating = ER
End Function

Public Function SpaceLeft()
    Dim SpcTotal As Single
    Dim i, j As Integer
    
    SpcTotal = Armor.Size + Internal.Size
    ' Only count armor and internal structure (rest included in criticals)
    
    For i = 1 To 4
        For j = 1 To 12
            SpcTotal = SpcTotal + Criticals(i, j).WeapSpace
        Next j
    Next i
    
    SpaceLeft = Craft.TotalSpace - SpcTotal
End Function
