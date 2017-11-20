Attribute VB_Name = "ModConvert"
'Old Format:

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

Public Type FileInfo
    Eng As EngineInfo
    Arm As ArmorInfo
    Ship As GeneralInfo
    IntStr As InternalInfo
    Crits(1 To 4, 1 To 12) As CriticalInfo
    TotCrits As TotalCriticalData
End Type


