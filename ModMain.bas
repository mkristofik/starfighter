Attribute VB_Name = "ModMain"
Option Explicit

Type WeapInfo
    WeapName As String * 25
    Damage As String * 6
    WeapSpace As Double
    Criticals As Integer
    Range As String * 15
    Tohit As String * 6
    MaxNum As Integer
    Techbase As Integer
    Locations As Integer
    Options As Integer
    Deleted As Boolean
End Type

Type EngineInfo
    EngName As String * 25
    Criticals As Integer
    EngType As Integer
    Techbase As Integer
    ManBase As Integer
    SpeedMod As Integer
    SpeedMult As Integer
    Deleted As Boolean
End Type

Type CriticalType
    Location As Integer
    recNum As Integer
    idNum As Integer
End Type

Type CraftInfo
    CraftName As String * 25
    Abbr As String * 6
    TotSpace As Integer
    Criticals(27) As CriticalType
    Armor(1 To 4) As Integer
    ArmType As Integer
    Shields As Integer
    Speed As Integer
    Techbase As Integer
    Wings As Integer
End Type

Public Weapon As WeapInfo, Engine As EngineInfo, Craft As CraftInfo, NotNewShip As Boolean

' Declarations required to place tab stops in a list box.
Const LB_SETTABSTOPS = &H192
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Sub SetTabs(lst As ListBox)
' Calls Windows DLL file to set tab stops in a list box
    ReDim tabs(0) As Long           ' # of tab stops
    Dim returnVal As Long           ' DLL function returns a long
    tabs(0) = 150                   ' Twips measurement for the tab stops
    returnVal = SendMessage(lst.hwnd, LB_SETTABSTOPS, 1, tabs(0)) ' Call the DLL

End Sub

Public Function TechString(t As Integer) As String
' Returns the string to identify tech bases

    Dim str As String
    
    If t = 0 Then
        str = "C"
    Else
        If InStr(t, "1") Then str = "NR "
        If InStr(t, "2") Then str = str & "I "
        If InStr(t, "3") Then str = str & "H "
        If InStr(t, "4") Then str = str & "P "
    End If
    
    TechString = str
        
End Function

Public Sub FormatLabel(lbl As Label)

    ' Format red for negative numbers
    If Val(lbl.Caption) < 0 Then
        lbl.ForeColor = vbRed
    Else
        lbl.ForeColor = vbBlack
    End If

End Sub

Public Function RoundUp(ByVal stuff As Double) As Integer
    RoundUp = Int(-stuff) * -1
End Function

Sub Main()
    NotNewShip = True
    frmMain.Show
End Sub

Public Sub NewShip()
    
    Dim i As Integer
    
    NotNewShip = False
    Unload frmArmor
    Unload frmEquipment
    Unload frmMain
    
    ' Reset craft defaults
    With Craft
        .Abbr = ""
        .ArmType = 0
        .CraftName = ""
        .Shields = 0
        .Speed = 0
        .Techbase = 0 ' ok to set to 0
        .TotSpace = 0
        .Wings = 0 ' ok to set to 0
        For i = 1 To 4
            .Armor(i) = 0
        Next i
        
        For i = 0 To 27
            .Criticals(i).idNum = 0
            .Criticals(i).Location = 0
            .Criticals(i).recNum = 0
        Next i
    End With
        
    NotNewShip = True
    frmMain.Show
End Sub
