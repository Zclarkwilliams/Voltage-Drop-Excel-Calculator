'
' Author: Zachary Clark-Williams
' Last Edited: 08-31-2017
'
' Excel Voltage Drop Calculator
'
'   This "New and Improved" calculator code uses a User_Info_Panel userform like a calculator
' pulling all the data from one panel instead of a bunch of smaller input request boxes.
' ** To run properly you will need the User_Info_Panel.vba code from userform file
'

Private Sub CalculateVoltageDrop_Click()

'*************************************************************************************'
'                 Set the Column Headers Last to AutoFit All Data Entered             '
'*************************************************************************************'
Sheets("Voltage Drop Calculator").Range("A4:N600").NumberFormat = Text
SetTitleBlock

'*************************************************************************************'
'                                 Get User Information                                '
'*************************************************************************************'

User_Info_Panel.User_Info_Panel_Initialize
User_Info_Panel.Show vbModal

'*************************************************************************************'
'                                Voltage Drop Equations                               '
'*************************************************************************************'
'****Get AC Resistance and Inductive Reactiance From NEC 2017 Table 9
Select Case User_Info_Panel.ConduitType
    Case "PVC":         'Use PVC Column From Table 9
        Set ReactRange = Sheets("Table 9").Range("B7:B27")
        If User_Info_Panel.ConductType = "Copper" Then
            Set ResRange = Sheets("Table 9").Range("D7:D27")
        Else: Set ResRange = Sheets("Table 9").Range("G7:G27")
        End If
    Case "Aluminum":    'Use Aluminum Column From Table 9
        Set ReactRange = Sheets("Table 9").Range("B7:B27")
        If User_Info_Panel.ConductType = "Copper" Then
            Set ResRange = Sheets("Table 9").Range("E7:E27")
        Else: Set ResRange = Sheets("Table 9").Range("H7:H27")
        End If
    Case "Steel":       'Use Steel Column From Table 9
        Set ReactRange = Sheets("Table 9").Range("C7:C27")
        If User_Info_Panel.ConductType = "Copper" Then
            Set ResRange = Sheets("Table 9").Range("F7:F27")
        Else: Set ResRange = Sheets("Table 9").Range("I7:I27")
        End If
End Select

'Now Go To Table 9 and Get Resistance Value and Inductive Reactance Value
ACRes = Application.WorksheetFunction.Index(ResRange, _
        Application.WorksheetFunction.Match(User_Info_Panel.WireGauge, _
                                            Sheets("Table 9").Range("A7:A27").Value, 0))
InductReact = Application.WorksheetFunction.Index(ReactRange, _
              Application.WorksheetFunction.Match(User_Info_Panel.WireGauge, _
                                            Sheets("Table 9").Range("A7:A27").Value, 0))
ThetaRad = Application.Acos(User_Info_Panel.PwrFctr)
ThetaDeg = (ThetaRad * 180) / WorksheetFunction.Pi
Zeff = ACRes * Cos(ThetaRad) + InductReact * Sin(ThetaRad)
Zcond = (User_Info_Panel.CableLen / 1000) * Zeff
If PhaseNum = 1 Then    '******* Single Phase
    KVA = User_Info_Panel.Amperes * User_Info_Panel.VoltSupply / 1000
    VoltDrop = User_Info_Panel.Amperes * 2 * Zcond
Else                    '******* Three Phae
    KVA = User_Info_Panel.Amperes * User_Info_Panel.VoltSupply * Sqr(3) / 1000
    VoltDrop = User_Info_Panel.Amperes * Sqr(3) * Zcond
End If
VoltDropPerc = VoltDrop / User_Info_Panel.VoltSupply * 100
KW = User_Info_Panel.PwrFctr * KVA

'*************************************************************************************'
'                           Place Values in Appropriate Columns                       '
'*************************************************************************************'

'UpdateTable(User_Info_Panel.DevDesc, User_Info_Panel.WireGauge, CndctrType, ConduitType, Amp, KVA, PF, KW, Phase, CblLen, Zeff, VoltageDrop, VoltageDropPerc, VoltSupp)

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
    Sheets("Voltage Drop Calculator").Cells(iRow, 1) = User_Info_Panel.DevDesc
    Sheets("Voltage Drop Calculator").Cells(iRow, 2) = User_Info_Panel.Amperes
    Sheets("Voltage Drop Calculator").Cells(iRow, 3) = Round(KVA, 3)
    Sheets("Voltage Drop Calculator").Cells(iRow, 4) = Round(User_Info_Panel.PwrFctr, 3)
    Sheets("Voltage Drop Calculator").Cells(iRow, 5) = Round(KW, 3)
    Sheets("Voltage Drop Calculator").Cells(iRow, 6) = "'" & User_Info_Panel.WireGauge
    Sheets("Voltage Drop Calculator").Cells(iRow, 7) = User_Info_Panel.PhaseNum
    Sheets("Voltage Drop Calculator").Cells(iRow, 8) = Round(User_Info_Panel.CableLen, 5)
    Sheets("Voltage Drop Calculator").Cells(iRow, 9) = Round(Zeff, 5)
    Sheets("Voltage Drop Calculator").Cells(iRow, 10) = Round(VoltageDrop, 3)
    Sheets("Voltage Drop Calculator").Cells(iRow, 11) = Round(VoltageDropPerc, 3)
    Sheets("Voltage Drop Calculator").Cells(iRow, 12) = VoltSupp
    Sheets("Voltage Drop Calculator").Cells(iRow, 13) = User_Info_Panel.ConductorType
    Sheets("Voltage Drop Calculator").Cells(iRow, 14) = User_Info_Panel.ConduitType

'*************************************************************************************'
'  Professionalize That Row Look With Borders and Centering And Adjust Column Width   '
'*************************************************************************************'

'****Border That Row
For C = 1 To 15     'Set Left Side Boarders Row By Row, Column By Column
    Sheets("Voltage Drop Calculator").Cells(iRow, C).Borders(xlEdgeLeft). _
                                                            ColorIndex = xlAutomatic
    If C < 15 Then
        Sheets("Voltage Drop Calculator").Cells(iRow, C).Borders(xlEdgeBottom). _
                                                            ColorIndex = xlAutomatic
    End If
Next C
'****Set All Cells to Correct Column Width
Sheets("Voltage Drop Calculator").Range("A:N").EntireColumn.AutoFit
'****Set Text in Cells to Center
Sheets("Voltage Drop Calculator").Range("A1:N600").VerticalAlignment = xlCenter
Sheets("Voltage Drop Calculator").Range("A1:N600").HorizontalAlignment = xlCenter

TotalTableValues

'*************************************************************************************'
'           Set the Conduit and Conductors as Drop Down Lists for EOU                 '
'*************************************************************************************'

Dim WireType, PipeType, WireList, PipeList As Variant
    
' Read Conduit/Conductor Material Using To Get List To Display Correct One On Top
    WireType = Sheets("Voltage Drop Calculator").Cells(iRow, "M").Value
    PipeType = Sheets("Voltage Drop Calculator").Cells(iRow, "N").Value
    
    Select Case WireType    ' Set up Conductor List Array for List
        Case "Copper":
            WireList = Array("Copper", "Aluminum")
        Case "Aluminum":
            WireList = Array("Aluminum", "Copper")
    End Select
    Select Case PipeType    ' Set up Conduit List Array for List
        Case "PVC":
            PipeList = Array("PVC", "Aluminum", "Steel")
        Case "Aluminum":
            PipeList = Array("Aluminum", "Steel", "PVC")
        Case "Steel":
            PipeList = Array("Steel", "PVC", "Aluminum")
    End Select
    
    With Sheets("Voltage Drop Calculator").Cells(iRow, "M").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                                    xlBetween, Formula1:=Join(WireList, ",")
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    With Sheets("Voltage Drop Calculator").Cells(iRow, "N").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                                    xlBetween, Formula1:=Join(PipeList, ",")
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub

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
    Sheets("Voltage Drop Calculator").Cells(4, 13) = "Conductor"
    Sheets("Voltage Drop Calculator").Cells(5, 13) = "Material"
    Sheets("Voltage Drop Calculator").Cells(6, 13) = "Type"
    Sheets("Voltage Drop Calculator").Cells(5, 14) = "Conduit Material"
    Sheets("Voltage Drop Calculator").Cells(6, 14) = "Type"

'*************************************************************************************'
'                     Size, Border, Grey-Fill and Bold Headers                        '
'*************************************************************************************'

    '****Make Column Title Block Bold
    Sheets("Voltage Drop Calculator").Range("A4:N6").Font.Bold = True
    '****Border the Cells
    Sheets("Voltage Drop Calculator").Range("A4:N6").Borders(xlEdgeBottom). _
                                                                ColorIndex = xlAutomatic
    Sheets("Voltage Drop Calculator").Range("A4:N6").Borders(xlEdgeTop). _
                                                                ColorIndex = xlAutomatic
    C = 1               'Set Column Start at Column A
    For R = 4 To 6      'Set Left Side Boarders Row By Row, Column By Column
        Sheets("Voltage Drop Calculator").Cells(R, C).Borders(xlEdgeLeft). _
                                                                ColorIndex = xlAutomatic
        If R = 6 Then   'Have we set all row side boarders 4 to 6
            R = 3       'Reset to 3 so we increment to 4th row on next column
            C = C + 1   'Transition to next column
        End If
        If C = 16 Then  'Have we set all column side boarders
            R = 6       'Yes, Set R to 6 to exit for loop
        End If
    Next R
    
    '****Set All Cells to Correct Column Width
    Sheets("Voltage Drop Calculator").Range("A:N").EntireColumn.AutoFit
    '****Set Text in Cells to Center
    Sheets("Voltage Drop Calculator").Range("A4:N6").VerticalAlignment = xlCenter
    Sheets("Voltage Drop Calculator").Range("A4:N6").HorizontalAlignment = xlCenter
    '****Set All Title Block Cells to Grey Coloring
    Sheets("Voltage Drop Calculator").Range("A4:N6").Interior.Color = RGB(211, 211, 211)
    
End Function

'-------------------------------------------------------------------------------------'
'---This Funciton Summizes KVA, A, and KW from the table entries and displays them ---'
'-------------------------------------------------------------------------------------'
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
'             This Function Will Place Passed Values into the Display Tables          '
'-------------------------------------------------------------------------------------'

Function UpdateTable(ByRef DevDescr, Gauge, CndctrType, ConduitType As String, Amp, KVA, _
                     PF, KW, Phase, CblLen, Zeff, VoltageDrop, VoltageDropPerc, VoltSupp As Double)

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
    Sheets("Voltage Drop Calculator").Cells(iRow, 1) = DevDescr
    Sheets("Voltage Drop Calculator").Cells(iRow, 2) = User_Info_Panel.Amperes
    Sheets("Voltage Drop Calculator").Cells(iRow, 3) = Round(KVA, 3)
    Sheets("Voltage Drop Calculator").Cells(iRow, 4) = Round(PF, 3)
    Sheets("Voltage Drop Calculator").Cells(iRow, 5) = Round(KW, 3)
    Sheets("Voltage Drop Calculator").Cells(iRow, 6) = "'" & Gauge
    Sheets("Voltage Drop Calculator").Cells(iRow, 7) = Phase
    Sheets("Voltage Drop Calculator").Cells(iRow, 8) = Round(CblLen, 5)
    Sheets("Voltage Drop Calculator").Cells(iRow, 9) = Round(Zeff, 5)
    Sheets("Voltage Drop Calculator").Cells(iRow, 10) = Round(VoltageDrop, 3)
    Sheets("Voltage Drop Calculator").Cells(iRow, 11) = Round(VoltageDropPerc, 3)
    Sheets("Voltage Drop Calculator").Cells(iRow, 12) = VoltSupp
    Sheets("Voltage Drop Calculator").Cells(iRow, 13) = CndctrType
    Sheets("Voltage Drop Calculator").Cells(iRow, 14) = ConduitType

End Function

'-------------------------------------------------------------------------------------'
'  This Function Will Take a Variant Input and Output A True or False if Its Number   '
'-------------------------------------------------------------------------------------'

Private Sub Worksheet_Change(ByVal Target As Range)
    Select Case Target.Columns
        Case "A":   '   Device Description - Do Nothing
        Case "B":   '   Current Change - Re-Calculate and Update Table
            
        Case "C":   '   KVA Change - ?? This should not happen. Post Err ??
        Case "D":   '   Power Factor Change - Re-Calculate and Update Table
        Case "E":   '   KW Change - ?? This should not happen. Post Err ??
        Case "F":   '   Wire Guage Size - Validate Input and Re-Calculate and Update Table
        Case "G":   '   Phase Number Change - Re-Calculate and Update Table
        Case "H":   '   Cable Length Change - Re-Calculate and Update Table
        Case "I":   '   Zeff Change - ?? This should not happen. Post Err ??
        Case "J":   '   Calculated Voltage Drop Change ?? This should not happen. Post Err ??
        Case "K":   '   Calculated Voltage Drop % Change ?? This should not happen. Post Err ??
        Case "L":   '   Supply Voltage Change - Re-Calculate and Update Table
        Case "M":   '   Conductor Change - Re-Calculate and Update Table
        Case "N":   '   Conduit Change - Re-Calculate and Update Table
    End Select
End Sub
