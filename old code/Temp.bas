Attribute VB_Name = "Temp"
Private Sub mnuTechH_Click()
    mnuTechH.Checked = True
    Craft.TechBase = "H"
    mnuTechI.Checked = False
    mnuTechNR.Checked = False
    mnuTechP.Checked = False
End Sub

Private Sub mnuTechI_Click()
    mnuTechI.Checked = True
    Craft.TechBase = "I"
    mnuTechNR.Checked = False
    mnuTechH.Checked = False
    mnuTechP.Checked = False
    WrongArmor
End Sub

Private Sub mnuTechNR_Click()
    mnuTechNR.Checked = True
    Craft.TechBase = "N"
    mnuTechI.Checked = False
    mnuTechH.Checked = False
    mnuTechP.Checked = False
    WrongArmor
End Sub

Private Sub mnuTechP_Click()
    mnuTechP.Checked = True
    Craft.TechBase = "P"
    mnuTechI.Checked = False
    mnuTechH.Checked = False
    mnuTechNR.Checked = False
    WrongArmor
End Sub
Private Sub WrongArmor()
' Check for valid armor type.  Only Herald may use Clear Plast.
    If RTrim(Armor.ArmType) = "Clear Plast" Then
        Dim strMessage As String
        strMessage = "Invalid Armor Type.  Changing to Standard."
        MsgBox strMessage, vbExclamation, "Armor Error"
        Armor.ArmType = "Standard"
        
        Armor.Size = Armor.Total * 0.6
        lblArmorType.Caption = "Standard"
        lblArmorSpc.Caption = FormatNumber(Armor.Size, 2)
                
    End If
    If frmArmor.cboArmorType.ListCount = 5 Then frmArmor.cboArmorType.RemoveItem (4)
       
End Sub

