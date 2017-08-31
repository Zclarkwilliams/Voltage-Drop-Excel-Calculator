'
'   Author: Zachary Clark-Williams
'   Date Last Edited: 08/30/2017
'   
'   Voltage Drop Calculator
'   
'   ** To run properly you will need the userform Conduitconductor.vba
'


'-------------------------------------------------------------------------------------'
'------------Enter New Row Of Data And Calculations To Votlage Drop Table-------------'
'-------------------------------------------------------------------------------------'
Private Sub CalculateVoltageDrop_Click()

'*************************************************************************************'
'                            Variable Declaration                                     '
'*************************************************************************************'
Dim USER_KVA, USER_PF, USER_EstCableLen, USER_VoltSupply As Variant
Dim KW, KVA, PF, VoltDrop, VoltDropPerc, Zeff As Double
Dim Rev, MotorLeadSize, ConduitType As String
Dim Vcond, ThetaRad, ACRes, InductReact, Zcond, ThetaDeg As Double
Dim iRow, NumPhase, Flag As Integer
Dim ResRange, ReactRange As Range

'*************************************************************************************'
'                 Set the Column Headers Last to AutoFit All Data Entered             '
'*************************************************************************************'
Sheets("Voltage Drop Calculator").Range("A4:L600").NumberFormat = Text
SetTitleBlock

'*************************************************************************************'
'                                 Get User Information                                '
'*************************************************************************************'

DeviceDescription = Application.InputBox("Pleas Enter Device Description Info:", Type:=2)
Do  'Get the Current from user and check to make sure it's an integer
    USER_Amps = Application.InputBox("Please Enter Current (If continuous load please multiply by 125% per NEC requiremnets prior entering)(A):")
    If USER_Amps = False Then                             'User Hit Cancel so Exit Program
        Exit Sub
    End If
Loop While CheckIfNum(USER_Amps) = False 'Check Flag Status
    Amps = CDbl(USER_Amps)                'Convert Input To Double
Do  'Get the PF and check to make sure its valid i.e. less than 1
    Do
        USER_PF = Application.InputBox("Please Enter Power Factor:") ' Get PF from user
        If USER_PF = False Then            'User Hit Cancel so Exit Program
            Exit Sub
        End If
    Loop While CheckIfNum(USER_PF) = False 'Check Flag Status
    PF = CDbl(USER_PF)                'Convert Input To Double
    If USER_PF > 1 Then                    'Is the entered PF Less Than 1?
        MsgBox ("**Err: Power Factor Greater than Possible. Re-Try.") 'Output error msg
        Flag = 1                                'Not Less Than 1 Activate Flag
    Else: Flag = 0                              'Is Less Than 1 Deactivate Flag
    End If
Loop While Flag = 1                             'Check Flag Status
Do  'Get the Cable Length from user and check to make sure it's an integer
    USER_EstCableLen = Application.InputBox("Please Enter Est. Cable Length (ft):") 'Get Length from User
    If USER_EstCableLen = False Then            'User Hit Cancel so Exit Program
        Exit Sub
    End If
Loop While CheckIfNum(USER_EstCableLen) = False 'Check Flag Status
    EstCableLen = CDbl(USER_EstCableLen)        'Convert Input To Double
Do  'Get Motor Lead Gauge and Check if Valid Size
    MotorLeadSize = Application.InputBox("Please Enter Conduit Size (No # Req.):", Type:=2) 'Get Gauge from User
    If MotorLeadSize = "False" Then             'User Hit Cancel so Exit Program
        Exit Sub
    End If
    'Go check against Table 9 Gauge Column to check validity or user entered value
    If VBA.IsError(Application.Match(MotorLeadSize, Sheets("Table 9").Range("A7:A27").Value, 0)) Then
        Flag = 1    'Invalid Lead Size Activate Flag
        MsgBox ("**Err: Not a Valid AWG. Please Re-Try.")
    Else: Flag = 0  'Matching Lead Size Deactivate Flag
    End If
Loop While Flag = 1 'Check Flag Status
Do  'Get the Supply Voltage from user and check to make sure it's an integer
    USER_VoltSupply = Application.InputBox("Please Enter Supply Voltage:", Type:=1)
    If USER_VoltSupply = False Then             'User Hit Cancel so Exit Program
        Exit Sub
    End If
Loop While CheckIfNum(USER_VoltSupply) = False  'Check Flag Status
    VoltSupply = CDbl(USER_VoltSupply)          'Convert Input To Double
Do  'Get and Check the user number of phases to see if valid i.e. single or 3 phase
    Do  'Get the Cable Length from user and check to make sure it's an integer
        USER_NumPhase = Application.InputBox("Please Enter the Number of Phases:")
        If USER_NumPhase = False Then               'User Hit Cancel so Exit Program
            Exit Sub
        End If
    Loop While CheckIfNum(USER_NumPhase) = False    'Check Flag Status
        NumPhase = CInt(USER_NumPhase)              'Convert to int
    If NumPhase = 1 Or NumPhase = 3 Then            'Check if Phase Number is Valid
        Flag = 0                                    'Valid #Phase Deactivate Flag
    Else
        Flag = 1                                    'Not Valid #Phase Activate Flag
        MsgBox ("**Err: Please Use Single or 3 phase only! Re-Try.") 'Display error msg
    End If
Loop While Flag = 1                                 'Check Flag Status
Do
'****Get Conduit and Conductor from userform dropdown lists
ConduitConductor.ConduitConductor_Initialize
ConduitConductor.Show vbModal   'Open small conductor/conduit user choice gui
ConductorType = ConduitConductor.User_ConductorType 'Get User Conductor Type
ConduitType = ConduitConductor.User_ConduitType 'Get User Conduit Type
If ConduitType <> "PVC" And ConduitType <> "Aluminum" And ConduitType <> "Steel" Then
    MsgBox ("**Err: Invalid Conduit Type. Please Fix and Try Again.")
    Flag = 1
Else: Flag = 0
End If
Loop While Flag = 1
'****Find Out If the Motor in Caclulations is continuously Run or NOT!!
'ContRunning.ContRunning_Initialize  'Run Initialize to Unselect Radio Buttons
'ContRunning.Show vbModal            'Show the OptionButton Form for Cont. Run
'RunCont = ContRunning.MotorRunType  'Get the User entered Value as Bool T/F

'*************************************************************************************'
'                                Voltage Drop Equations                               '
'*************************************************************************************'
'****Get AC Resistance and Inductive Reactiance From NEC 2017 Table 9
Select Case ConduitType
    Case "PVC":         'Use PVC Column From Table 9
        Set ReactRange = Sheets("Table 9").Range("B7:B27")
        If ConductorType = "Copper" Then
            Set ResRange = Sheets("Table 9").Range("D7:D27")
        Else: Set ResRange = Sheets("Table 9").Range("G7:G27")
        End If
    Case "Aluminum":    'Use Aluminum Column From Table 9
        Set ReactRange = Sheets("Table 9").Range("B7:B27")
        If ConductorType = "Copper" Then
            Set ResRange = Sheets("Table 9").Range("E7:E27")
        Else: Set ResRange = Sheets("Table 9").Range("H7:H27")
        End If
    Case "Steel":       'Use Steel Column From Table 9
        Set ReactRange = Sheets("Table 9").Range("C7:C27")
        If ConductorType = "Copper" Then
            Set ResRange = Sheets("Table 9").Range("F7:F27")
        Else: Set ResRange = Sheets("Table 9").Range("I7:I27")
        End If
End Select
'Now Go To Table 9 and Get Resistance Value and Inductive Reactance Value
ACRes = Application.WorksheetFunction.Index(ResRange, _
        Application.WorksheetFunction.Match(MotorLeadSize, _
                                            Sheets("Table 9").Range("A7:A27").Value, 0))
InductReact = Application.WorksheetFunction.Index(ReactRange, _
              Application.WorksheetFunction.Match(MotorLeadSize, _
                                            Sheets("Table 9").Range("A7:A27").Value, 0))
ThetaRad = Application.Acos(PF)
ThetaDeg = (ThetaRad * 180) / WorksheetFunction.Pi
Zeff = ACRes * Cos(ThetaDeg) + InductReact * Sin(ThetaDeg)
Zcond = (EstCableLen / 1000) * Zeff
If NumPhase = 1 Then    '******* Single Phase
    KVA = Amps * VoltSupply / 1000
    VoltDrop = Amps * 2 * Zcond
Else                    '******* Three Phae
    KVA = Amps * VoltSupply * Sqr(3) / 1000
    VoltDrop = Amps * Sqr(3) * Zcond
End If
VoltDropPerc = VoltDrop / USER_VoltSupply * 100
KW = PF * KVA

'*************************************************************************************'
'                           Place Values in Appropriate Columns                       '
'*************************************************************************************'
'****Get Row to print data to
On Error Resume Next
LastRow = Sheets("Voltage Drop Calculator").Cells.Find(What:="Total", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).Row
If LastRow <> 0 Then
    Sheets("Voltage Drop Calculator").Range(Cells(LastRow - 6, 3), Cells(LastRow, 3)).ClearContents
    Sheets("Voltage Drop Calculator").Range(Cells(LastRow - 6, 3), Cells(LastRow, 3)).ClearFormats
End If
Sheets("Voltage Drop Calculator").Range("A4:L600").NumberFormat = Text
iRow = Sheets("Voltage Drop Calculator").Cells.Find(What:="*", _
                                            LookAt:=xlPart, _
                                            LookIn:=xlValues, _
                                            SearchOrder:=xlByRows, _
                                            SearchDirection:=xlPrevious, _
                                            MatchCase:=False).Row + 1
Sheets("Voltage Drop Calculator").Cells(iRow, 1) = DeviceDescription
Sheets("Voltage Drop Calculator").Cells(iRow, 2) = Amps
Sheets("Voltage Drop Calculator").Cells(iRow, 3) = Round(KVA, 3)
Sheets("Voltage Drop Calculator").Cells(iRow, 4) = Round(PF, 3)
Sheets("Voltage Drop Calculator").Cells(iRow, 5) = Round(KW, 3)
Sheets("Voltage Drop Calculator").Cells(iRow, 6) = "'" & MotorLeadSize
Sheets("Voltage Drop Calculator").Cells(iRow, 7) = NumPhase
Sheets("Voltage Drop Calculator").Cells(iRow, 8) = EstCableLen
Sheets("Voltage Drop Calculator").Cells(iRow, 9) = Round(Zeff, 5)
Sheets("Voltage Drop Calculator").Cells(iRow, 10) = Round(VoltDrop, 3)
Sheets("Voltage Drop Calculator").Cells(iRow, 11) = Round(VoltDropPerc, 3)
Sheets("Voltage Drop Calculator").Cells(iRow, 12) = VoltSupply
Sheets("Voltage Drop Calculator").Cells(iRow, 13) = ConduitType

'*************************************************************************************'
'  Professionalize That Row Look With Borders and Centering And Adjust Column Width   '
'*************************************************************************************'

'****Border That Row
For C = 1 To 14     'Set Left Side Boarders Row By Row, Column By Column
    Sheets("Voltage Drop Calculator").Cells(iRow, C).Borders(xlEdgeLeft). _
                                                            ColorIndex = xlAutomatic
    If C < 14 Then
        Sheets("Voltage Drop Calculator").Cells(iRow, C).Borders(xlEdgeBottom). _
                                                            ColorIndex = xlAutomatic
    End If
Next C
'****Set All Cells to Correct Column Width
Sheets("Voltage Drop Calculator").Range("A:M").EntireColumn.AutoFit
'****Set Text in Cells to Center
Sheets("Voltage Drop Calculator").Range("A1:M600").VerticalAlignment = xlCenter
Sheets("Voltage Drop Calculator").Range("A1:M600").HorizontalAlignment = xlCenter

TotalTableValues

End Sub

'-------------------------------------------------------------------------------------'
'---------Edit Existing Data And Re-Calculate Row Data In Votlage Drop Table----------'
'-------------------------------------------------------------------------------------'
Private Sub EditMadeReCalculate_Click()

Dim iRow, GETOUT, Flag, LastRow As Integer
Dim CblLen, VSupply, Amps As Variant
Dim WireGauge, ConduitType As String
Dim ACRes, InductReact, ThetaRad, ThetaDeg As Double
Dim ClearTots As Range
'Dim FLA As Variant, PF As Variant, ContRun As String, HP As Variant, RPM As Variant

Sheets("Voltage Drop Calculator").Range("A4:L600").NumberFormat = Text

'****Get Number of Rows With Valid Data Present
On Error Resume Next
LastRow = Sheets("Voltage Drop Calculator").Cells.Find(What:="Total", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).Row
If LastRow <> 0 Then
    Sheets("Voltage Drop Calculator").Range(Cells(LastRow - 6, 3), Cells(LastRow, 3)).ClearContents
    Sheets("Voltage Drop Calculator").Range(Cells(LastRow - 6, 3), Cells(LastRow, 3)).ClearFormats
End If
Sheets("Voltage Drop Calculator").Range("A4:L600").NumberFormat = Text
LastRow = Cells.Find(What:="*", _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row
iRow = 7
Do
    Amps = Sheets("Voltage Drop Calculator").Cells(iRow, 2).Value
    If CheckIfNum(Amps) = False Then
        Sheets("Voltage Drop Calculator").Cells(iRow, 2).Interior.Color = RGB(211, 100, 100)
        Exit Sub
    Else
        Amps = CDbl(Amps)
        Sheets("Voltage Drop Calculator").Cells(iRow, 2).Interior.Color = xlNone
    End If
    PF = Sheets("Voltage Drop Calculator").Cells(iRow, 4).Value
    If CheckIfNum(PF) = False Then
        Sheets("Voltage Drop Calculator").Cells(iRow, 4).Interior.Color = RGB(211, 100, 100)
        Exit Sub
    Else
        If PF > 1 Then
            MsgBox ("**Err: PF Value Valid, Greater Than 1. Please Fix.")
            Sheets("Voltage Drop Calculator").Cells(iRow, 4).Interior.Color = RGB(211, 100, 100)
            Exit Sub
        Else
            PF = CDbl(PF)
            Sheets("Voltage Drop Calculator").Cells(iRow, 4).Interior.Color = xlNone
        End If
    End If
    WireGauge = Sheets("Voltage Drop Calculator").Cells(iRow, 6).Text
    If VBA.IsError(Application.Match(WireGauge, Sheets("Table 9").Range("A7:A27").Value, 0)) Then
        Sheets("Voltage Drop Calculator").Cells(iRow, 6).Interior.Color = RGB(211, 100, 100)
        MsgBox ("**Err: Invalid Motor Lead Gauge Size. Please Fix And Try Again.")
        Exit Sub
    Else
        Sheets("Voltage Drop Calculator").Cells(iRow, 6).Interior.Color = xlNone
    End If
    NumPhase = Sheets("Voltage Drop Calculator").Cells(iRow, 7).Value
    If CheckIfNum(NumPhase) = False Then
        Sheets("Voltage Drop Calculator").Cells(iRow, 7).Interior.Color = RGB(211, 100, 100)
        Exit Sub
    Else
        NumPhase = CDbl(NumPhase)
        Sheets("Voltage Drop Calculator").Cells(iRow, 7).Interior.Color = xlNone
    End If
    CblLen = Sheets("Voltage Drop Calculator").Cells(iRow, 8).Value
    If CheckIfNum(CblLen) = False Then
        Sheets("Voltage Drop Calculator").Cells(iRow, 8).Interior.Color = RGB(211, 100, 100)
        Exit Sub
    Else
        CblLen = CDbl(CblLen)
        Sheets("Voltage Drop Calculator").Cells(iRow, 8).Interior.Color = xlNone
    End If
    VSupply = Sheets("Voltage Drop Calculator").Cells(iRow, 12).Value
    If CheckIfNum(VSupply) = False Then
        Sheets("Voltage Drop Calculator").Cells(iRow, 12).Interior.Color = RGB(211, 100, 100)
        Exit Sub
    Else
        VSupply = CDbl(VSupply)
        Sheets("Voltage Drop Calculator").Cells(iRow, 12).Interior.Color = xlNone
    End If
    ConduitType = Sheets("Voltage Drop Calculator").Cells(iRow, 13).Text
    If ConduitType <> "PVC" And ConduitType <> "Aluminum" And ConduitType <> "Steel" Then
        MsgBox ("**Err: Invalid Conduit Type. Please Fix and Try Again.")
        Sheets("Voltage Drop Calculator").Cells(iRow, 13).Interior.Color = RGB(211, 100, 100)
        Exit Sub
    Else
        Sheets("Voltage Drop Calculator").Cells(iRow, 13).Interior.Color = xlNone
    End If
    '*************************************************************************************'
    '                                Voltage Drop Equations                               '
    '*************************************************************************************'
    '****Get AC Resistance and Inductive Reactiance From NEC 2017 Table 9
    Select Case ConduitType
        Case "PVC":         'Use PVC Column From Table 9
            Set ResRange = Sheets("Table 9").Range("D7:D27")
            Set ReactRange = Sheets("Table 9").Range("B7:B27")
        Case "Aluminum":    'Use Aluminum Column From Table 9
            Set ResRange = Sheets("Table 9").Range("E7:E27")
            Set ReactRange = Sheets("Table 9").Range("B7:B27")
        Case "Steel":       'Use Steel Column From Table 9
            Set ResRange = Sheets("Table 9").Range("F7:F27")
            Set ReactRange = Sheets("Table 9").Range("C7:C27")
    End Select
    'Now Go To Table 9 and Get Resistance Value and Inductive Reactance Value
    ACRes = Application.WorksheetFunction.Index(ResRange, _
            Application.WorksheetFunction.Match(WireGauge, _
                                                Sheets("Table 9").Range("A7:A27").Value, 0))
    InductReact = Application.WorksheetFunction.Index(ReactRange, _
                  Application.WorksheetFunction.Match(WireGauge, _
                                                Sheets("Table 9").Range("A7:A27").Value, 0))
    ThetaRad = Application.Acos(PF)
    ThetaDeg = (ThetaRad * 180) / WorksheetFunction.Pi
    Zeff = ACRes * Cos(ThetaDeg) + InductReact * Sin(ThetaDeg)
    Zcond = (CblLen / 1000) * Zeff
    If NumPhase = 1 Then    '******* Single Phase
        KVA = Amps * VSupply / 1000
        VoltDrop = Amps * 2 * Zcond
    Else                    '******* Three Phae
        KVA = Amps * VSupply * Sqr(3) / 1000
        VoltDrop = Amps * Sqr(3) * Zcond
    End If
    VoltDropPerc = VoltDrop / VSupply * 100
    KW = PF * KVA
    
    '*************************************************************************************'
    '                           Place Values in Appropriate Columns                       '
    '*************************************************************************************'
    Sheets("Voltage Drop Calculator").Range("A4:L600").NumberFormat = Text
    Sheets("Voltage Drop Calculator").Cells(iRow, 2) = Amps
    Sheets("Voltage Drop Calculator").Cells(iRow, 3) = Round(KVA, 3)
    Sheets("Voltage Drop Calculator").Cells(iRow, 4) = Round(PF, 3)
    Sheets("Voltage Drop Calculator").Cells(iRow, 5) = Round(KW, 3)
    Sheets("Voltage Drop Calculator").Cells(iRow, 6) = "'" & WireGauge
    Sheets("Voltage Drop Calculator").Cells(iRow, 7) = NumPhase
    Sheets("Voltage Drop Calculator").Cells(iRow, 8) = CblLen
    Sheets("Voltage Drop Calculator").Cells(iRow, 9) = Round(Zeff, 5)
    Sheets("Voltage Drop Calculator").Cells(iRow, 10) = Round(VoltDrop, 3)
    Sheets("Voltage Drop Calculator").Cells(iRow, 11) = Round(VoltDropPerc, 3)
    Sheets("Voltage Drop Calculator").Cells(iRow, 12) = VSupply
    Sheets("Voltage Drop Calculator").Cells(iRow, 13) = ConduitType
    
    'Increment to Next Row And Check If We Finished The Last Row
    iRow = iRow + 1
    If iRow > LastRow Then
        GETOUT = 1
    Else: GETOUT = 0
    End If
Loop While GETOUT = 0

'*************************************************************************************'
'  Professionalize That Row Look With Borders and Centering And Adjust Column Width   '
'*************************************************************************************'

'****Border That Row
For C = 1 To 14     'Set Left Side Boarders Row By Row, Column By Column
    Sheets("Voltage Drop Calculator").Cells(iRow, C).Borders(xlEdgeLeft). _
                                                            ColorIndex = xlAutomatic
    If C < 14 Then
        Sheets("Voltage Drop Calculator").Cells(iRow, C).Borders(xlEdgeBottom). _
                                                            ColorIndex = xlAutomatic
    End If
Next C
'****Set All Cells to Correct Column Width
Sheets("Voltage Drop Calculator").Range("A:M").EntireColumn.AutoFit
'****Set Text in Cells to Center
Sheets("Voltage Drop Calculator").Range("A1:M600").VerticalAlignment = xlCenter
Sheets("Voltage Drop Calculator").Range("A1:M600").HorizontalAlignment = xlCenter

TotalTableValues

End Sub

'-------------------------------------------------------------------------------------'
'                      Clear the Worksheet for New Calculations                       '
'-------------------------------------------------------------------------------------'
Private Sub ClearSheetVD_Click()
    Sheets("Voltage Drop Calculator").Range("A1:Z900").ClearContents
    Sheets("Voltage Drop Calculator").Range("A1:Z900").ClearFormats
End Sub

'-------------------------------------------------------------------------------------'
'  This Function Will Take a Variant Input and Output A True or False if Its Number   '
'-------------------------------------------------------------------------------------'
Function CheckIfNum(ByVal NumCheck As Variant) As Boolean
    If IsNumeric(NumCheck) = False Then                   'Check if a number
        MsgBox ("**Err: Non-Numeric Value Entered. Please Re-Try.") 'Output error msg
    Else: CheckIfNum = True                               'Is number Deactivate flag
    End If
End Function

'-------------------------------------------------------------------------------------'
'---------This Funciton Set the Column Title Block For the Voltage Drop Table---------'
'-------------------------------------------------------------------------------------'
Function SetTitleBlock()

'*************************************************************************************'
'                      Variables and Declaration Constants                            '
'*************************************************************************************'

    Dim C As Integer, R As Integer

'*************************************************************************************'
'                        Write Headers for Table Columns                              '
'*************************************************************************************'
    Sheets("Voltage Drop Calculator").Range("A4:L600").NumberFormat = Text
    Sheets("Voltage Drop Calculator").Cells(6, 1) = "Load Device Description"
    Sheets("Voltage Drop Calculator").Cells(6, 2) = "Amperes"
    Sheets("Voltage Drop Calculator").Cells(6, 3) = "KVA"
    Sheets("Voltage Drop Calculator").Cells(6, 4) = "PF"
    Sheets("Voltage Drop Calculator").Cells(6, 5) = "KW"
    Sheets("Voltage Drop Calculator").Cells(6, 6) = "Gauge Size #"
    Sheets("Voltage Drop Calculator").Cells(4, 7) = "Number"
    Sheets("Voltage Drop Calculator").Cells(5, 7) = "of"
    Sheets("Voltage Drop Calculator").Cells(6, 7) = "Phases"
    Sheets("Voltage Drop Calculator").Cells(4, 8) = "Estimated"
    Sheets("Voltage Drop Calculator").Cells(5, 8) = "Cable Length"
    Sheets("Voltage Drop Calculator").Cells(6, 8) = "in Feet"
    Sheets("Voltage Drop Calculator").Cells(5, 9) = "Effective Z"
    Sheets("Voltage Drop Calculator").Cells(6, 9) = "Per 1000 ft"
    Sheets("Voltage Drop Calculator").Cells(6, 10) = "Voltage Drop (V)"
    Sheets("Voltage Drop Calculator").Cells(5, 11) = "Voltage Drop"
    Sheets("Voltage Drop Calculator").Cells(6, 11) = "Percent (%)"
    Sheets("Voltage Drop Calculator").Cells(5, 12) = "Supply"
    Sheets("Voltage Drop Calculator").Cells(6, 12) = "Voltage (V)"
    Sheets("Voltage Drop Calculator").Cells(5, 13) = "Conduit Material"
    Sheets("Voltage Drop Calculator").Cells(6, 13) = "Type"

'*************************************************************************************'
'                     Size, Border, Grey-Fill and Bold Headers                        '
'*************************************************************************************'

    '****Make Column Title Block Bold
    Sheets("Voltage Drop Calculator").Range("A4:M6").Font.Bold = True
    '****Border the Cells
    Sheets("Voltage Drop Calculator").Range("A4:M6").Borders(xlEdgeBottom). _
                                                                ColorIndex = xlAutomatic
    Sheets("Voltage Drop Calculator").Range("A4:M6").Borders(xlEdgeTop). _
                                                                ColorIndex = xlAutomatic
    C = 1               'Set Column Start at Column A
    For R = 4 To 6      'Set Left Side Boarders Row By Row, Column By Column
        Sheets("Voltage Drop Calculator").Cells(R, C).Borders(xlEdgeLeft). _
                                                                ColorIndex = xlAutomatic
        If R = 6 Then   'Have we set all row side boarders 4 to 6
            R = 3       'Reset to 3 so we increment to 4th row on next column
            C = C + 1   'Transition to next column
        End If
        If C = 15 Then  'Have we set all column side boarders
            R = 6       'Yes, Set R to 6 to exit for loop
        End If
    Next R
    
    '****Set All Cells to Correct Column Width
    Sheets("Voltage Drop Calculator").Range("A:M").EntireColumn.AutoFit
    '****Set Text in Cells to Center
    Sheets("Voltage Drop Calculator").Range("A4:M6").VerticalAlignment = xlCenter
    Sheets("Voltage Drop Calculator").Range("A4:M6").HorizontalAlignment = xlCenter
    '****Set All Title Block Cells to Grey Coloring
    Sheets("Voltage Drop Calculator").Range("A4:M6").Interior.Color = RGB(211, 211, 211)
    
End Function

Function TotalTableValues()

Dim iRow, GETOUT, Flag, LastRow As Integer
Dim ContRun As String
Dim Amps As Variant

'****Get Number of Rows With Valid Data Present
LastRow = Cells.Find(What:="*", _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row
iRow = 7
ActiveSheet.Rows(LastRow + 1).ClearContents
ActiveSheet.Rows(LastRow + 1).ClearFormats
Do
    '****Read and Store All Valid Cells in Row From HP to Conduit Material
    Amps = Sheets("Voltage Drop Calculator").Cells(iRow, 2).Value
    If CheckIfNum(Amps) = False Then
        Sheets("Voltage Drop Calculator").Cells(iRow, 2).Interior.Color = RGB(211, 100, 100)
        Exit Function
    Else
        Amps = CDbl(Amps)
        Sheets("Voltage Drop Calculator").Cells(iRow, 2).Interior.Color = xlNone
    End If
    KVA = Sheets("Voltage Drop Calculator").Cells(iRow, 3).Value
    If CheckIfNum(KVA) = False Then
        Sheets("Voltage Drop Calculator").Cells(iRow, 3).Interior.Color = RGB(211, 100, 100)
        Exit Function
    Else
        KVA = CDbl(KVA)
        Sheets("Voltage Drop Calculator").Cells(iRow, 3).Interior.Color = xlNone
    End If
    KW = Sheets("Voltage Drop Calculator").Cells(iRow, 5).Value
    If CheckIfNum(KW) = False Then
        Sheets("Voltage Drop Calculator").Cells(iRow, 5).Interior.Color = RGB(211, 100, 100)
        Exit Function
    Else
        KW = CDbl(KW)
        Sheets("Voltage Drop Calculator").Cells(iRow, 5).Interior.Color = xlNone
    End If

    KW_Total = KW_Total + KW
    KVA_Total = KVA_Total + KVA
    Ampere_Total = Ampere_Total + Amps
    
    'Increment to Next Row And Check If We Finished The Last Row
    iRow = iRow + 1
    If iRow > LastRow Then
        GETOUT = 1
    Else: GETOUT = 0
    End If
Loop While GETOUT = 0

'*************************************************************************************'
'                           Place Values in Appropriate Columns                       '
'*************************************************************************************'
    Sheets("Voltage Drop Calculator").Range("A4:L600").NumberFormat = Text
    Sheets("Voltage Drop Calculator").Cells(iRow + 5, 3) = "Total Current:" & Amperes_Total
    Sheets("Voltage Drop Calculator").Cells(iRow + 6, 3) = "Overall Total KW: " & KW_Total
    Sheets("Voltage Drop Calculator").Cells(iRow + 7, 3) = "Total kVA: " & KVA_Total

    Sheets("Voltage Drop Calculator").Range(Cells(iRow + 5, 3), Cells(iRow + 10, 3)).VerticalAlignment = xlLeft
    
End Function
