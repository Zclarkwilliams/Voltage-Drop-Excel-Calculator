'
' Author: Zachary Clark_Williams
' Last Edited: 08-31-2017
'
' Excel Voltage Drop Calculator
' 
' This code is to the main calculator userform used in Main_Rev1.vba.
'



Public Amperes, CableLen, PwrFctr, VoltSupply, PhaseNum As Double
Public DevDesc, ConductType, ConduitType, WireGauge As String

Private Sub Amps_AfterUpdate()
' Check to if Number Entered, Not Char or Other or Nothing
If CheckIfNum(Amps.Value) = True Then   ' Numeric Value Entered!
    Amperes = CDbl(Amps.Value)  ' Number Entered Set Global to Value
    Amps.BackColor = vbWhite    ' Make sure backroung is set to white
Else    ' Non-Numeric Entered! ERROR!
    Amps.BackColor = RGB(220, 10, 10)   ' Set red Background for ERROR
    Amps.Value = "(A)"  ' Reset box value to Default
    Exit Sub            ' Exit the function due to ERROR
End If
End Sub

Private Sub CableLength_AfterUpdate()
' Check to if Number Entered, Not Char or Other or Nothing
If CheckIfNum(CableLength.Value) = True Then   ' Numeric Value Entered!
    CableLen = CDbl(CableLength.Value)  ' Number Entered Set Global to Value
    CableLength.BackColor = vbWhite     ' Make sure backroung is set to white
Else    ' Non-Numeric Entered! ERROR!
    CableLength.BackColor = RGB(220, 10, 10)   ' Set red Background for ERROR
    CableLength.Value = "(ft)"  ' Reset box value to Default
    Exit Sub                    ' Exit the function due to ERROR
End If
End Sub

Private Sub Calculate_Click()
Dim AddDesc As Integer
    
    '****** Check If Device Description Entered or Not
    If DeviceDesc.Value = vbNullString Or _
       DeviceDesc.Value = "Insert Device Description Here." Then   ' Nothing entered in Device Description Textbox
        AddDesc = MsgBox("No Desctription Entered. Would you like to add one?", vbYesNo)
        If AddDesc = vbYes Then ' User wants to enter a device description
            Exit Sub    ' So exit this process so user can enter
        Else    ' User doesn't want to enter device description
            DevDesc = vbNullString  ' Set it to Null String " "
        End If
    Else    ' Description Entered Pass to Global
        DevDesc = DeviceDesc.Value  ' Pass to global
    End If

    '****** Check if Conductor Material was selected or not
    If ConductorMtrl.Value = vbNullString Then    ' No Conductor Selected
        MsgBox ("**Err: Conductor Material Not Selected. Please Fix and Re-Try.")
        Exit Sub    ' This must be done so Exit till this is entered
    Else    ' Conductor type selected pass to Global
        ConductType = ConductorMtrl.Value
    End If
    
    '****** Check if Conduit Material was selected or not
    If ConduitMtrl.Value = vbNullString Then    ' No Conduit Selected
        MsgBox ("**Err: Conduit Material Not Selected. Please Fix and Re-Try.")
        Exit Sub    ' This must be done so Exit till this is entered
    Else    ' Conduit type selected pass to Global
        ConduitType = ConduitMtrl.Value
    End If
    
    '****** Check if Wire Gauge was selected or not
    If GaugeSize.Value = vbNullString Then    ' No Wire Gauge Selected
        MsgBox ("**Err: Wire Gauge Not Selected. Please Fix and Re-Try.")
        Exit Sub    ' This must be done so Exit till this is entered
    Else    ' Wire Gauge type selected pass to Global
        WireGauge = GaugeSize.Value
    End If
    
    '****** Check if Supply Voltage Entered
    If VoltSupp.Value = "(V)" Then    ' No Supply Voltage Entered
        MsgBox ("**Err: Supply Voltage Not Entered. Please Fix and Re-Try.")
        Exit Sub    ' This must be done so Exit till this is entered
    End If
    
    '****** Check if Current Entered
    If Amps.Value = "(A)" Then    ' No Current Entered
        MsgBox ("**Err: Current Not Entered. Please Fix and Re-Try.")
        Exit Sub    ' This must be done so Exit till this is entered
    End If
    
    '****** Check if Power Factor Entered
    If PF.Value = "PF" Then    ' No PF Entered
        MsgBox ("**Err: Power Factor Not Entered. Please Fix and Re-Try.")
        Exit Sub    ' This must be done so Exit till this is entered
    End If
    
    '****** Check if Est. Cable Length Entered
    If CableLength.Value = "(ft)" Then    ' No CableLength Entered
        MsgBox ("**Err: Est. Cable Length Not Entered. Please Fix and Re-Try.")
        Exit Sub    ' This must be done so Exit till this is entered
    End If
    
    '****** Check if Number of Phases Selected
    If SinglePhase.Value = Flase And ThreePhase.Value = False Then  ' No Phase Selected
        MsgBox ("**Err: Number Of Phases Not Selected. Please Fix and Re-Try.")
        Exit Sub    ' This must be done so Exit till this is entered
    End If
    
    Hide
End Sub

Private Sub PF_AfterUpdate()
' Check to if Number Entered, Not Char or Other or Nothing
If CheckIfNum(PF.Value) = True Then   ' Numeric Value Entered!
    If PF.Value < 1 And PF.Value > 0 Then ' Value entered > 0 && < 1
        PwrFctr = CDbl(PF.Value)  ' Number Entered Set Global to Value
        PF.BackColor = vbWhite    ' Make sure backroung is set to white
    Else    ' Value entered greater than 1 or less than 0 !
        MsgBox ("**Err: Power Factor InValid. Please Fix and Re-Try.")
        PF.BackColor = RGB(220, 10, 10)   ' Set red Background for ERROR
        PF.Value = "PF"     ' Reset box value to Default
        Exit Sub            ' Exit the function due to ERROR
    End If
Else    ' Non-Numeric Entered! ERROR!
    PF.BackColor = RGB(220, 10, 10)   ' Set red Background for ERROR
    PF.Value = "PF"     ' Reset box value to Default
    Exit Sub            ' Exit the function due to ERROR
End If
End Sub

Private Sub SinglePhase_Click()
    ThreePhase.Value = False    ' Single Phase Set ThreePhase to Invalid
    PhaseNum = 1    ' Pass Phase Number to Global
End Sub

Private Sub ThreePhase_Click()
    SinglePhase.Value = False   ' Single Phase Set ThreePhase to Invalid
    PhaseNum = 3    ' Pass Phase Number to Global
End Sub

Private Sub VoltSupp_AfterUpdate()
' Check to if Number Entered, Not Char or Other or Nothing
If CheckIfNum(VoltSupp.Value) = True Then   ' Numeric Value Entered!
    VoltSupply = CDbl(VoltSupp.Value)   ' Number Entered Set Global to Value
    VoltSupp.BackColor = vbWhite        ' Make sure backroung is set to white
Else    ' Non-Numeric Entered! ERROR!
    VoltSupp.BackColor = RGB(220, 10, 10)   ' Set red Background for ERROR
    VoltSupp.Value = "(V)"  ' Reset box value to Default
    Exit Sub                ' Exit the function due to ERROR
End If

End Sub

Public Sub User_Info_Panel_Initialize()
    User_Info_Panel.Amps.Value = "(A)"
    User_Info_Panel.CableLength.Value = "(ft)"
    User_Info_Panel.ConductorMtrl.Value = ""
    User_Info_Panel.ConduitMtrl.Value = ""
    User_Info_Panel.DeviceDesc.Value = "Insert Device Description Here."
    User_Info_Panel.GaugeSize.Value = ""
    User_Info_Panel.PF.Value = "PF"
    User_Info_Panel.SinglePhase.Value = False
    User_Info_Panel.ThreePhase.Value = False
    User_Info_Panel.VoltSupp.Value = "(V)"
End Sub

Function CheckIfNum(ByVal NumCheck As Variant) As Boolean
    If IsNumeric(NumCheck) = False Or NumCheck = vbNullString Then 'Check if a number
        MsgBox ("**Err: Non-Numeric Value Entered. Please Re-Try.") 'Output error msg
        CheckIfNum = False  ' Set function to false / Activate Flag
        Exit Function       ' Kill Function!
    Else: CheckIfNum = True 'Is number Deactivate flag
    End If
End Function
