'
' Author: Zachary Clark-Williams
' Last Edited: 09-12-2017
'
' Excel Voltage Drop Calculator
'
'   This "New and Improved" calculator code uses a User_Info_Panel userform like a calculator
' pulling all the data from one panel instead of a bunch of smaller input request boxes.
' ** To run properly you will need the User_Info_Panel.vba code from userform file
'

Public FLAG_FillTable As Boolean

Private Sub CalculateVoltageDrop_Click()

Dim iRow As Long
Dim HeaderSet As Variant
Dim RandXl As Collection
Dim VoltDrop, ThetaRad, Zeff, Zcond, KVA, VoltDropPerc, KW As Double


FLAG_FillTable = True

'*************************************************************************************'
'                 Set the Column Headers Last to AutoFit All Data Entered             '
'*************************************************************************************'
Sheets("Voltage Drop Calculator").Range("A4:N600").NumberFormat = Text
HeaderSet = Sheets("Voltage Drop Calculator").Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).Row
If HeaderSet <> 0 And HeaderSet <> 24 Then
    SetTitleBlock
End If

'*************************************************************************************'
'                                 Get User Information                                '
'*************************************************************************************'

User_Info_Panel.User_Info_Panel_Initialize
User_Info_Panel.Show vbModal

'   Check if the user 'X' out
If User_Info_Panel.FLAG_XedOut = True Then
    Exit Sub
End If

'*************************************************************************************'
'                                Voltage Drop Equations                               '
'*************************************************************************************'
'****Get AC Resistance and Inductive Reactiance From NEC 2017 Table 9
Set RandXl = GetXlandR(User_Info_Panel.ConductType, _
                       User_Info_Panel.ConduitType, _
                       User_Info_Panel.WireGauge)
ACRes = RandXl.Item(1)
InductReact = RandXl.Item(2)

ThetaRad = Application.Acos(User_Info_Panel.PwrFctr)
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
'                              Get Row to Print Data                                  '
'*************************************************************************************'

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
                                                
'*************************************************************************************'
'                       Place Values in Appropriate Columns                           '
'*************************************************************************************'
                                                
Call UpdateTable(User_Info_Panel.DevDesc, _
                 User_Info_Panel.Amperes, _
                 User_Info_Panel.WireGauge, _
                 KVA, _
                 User_Info_Panel.PwrFctr, _
                 KW, _
                 User_Info_Panel.PhaseNum, _
                 User_Info_Panel.CableLen, _
                 Zeff, _
                 VoltDrop, _
                 VoltageDropPerc, _
                 User_Info_Panel.VoltSupply, _
                 User_Info_Panel.ConductType, _
                 User_Info_Panel.ConduitType, _
                 iRow)

'*************************************************************************************'
'  Professionalize That Row Look With Borders and Centering And Adjust Column Width   '
'*************************************************************************************'

'****Border That Row
For c = 1 To 15     'Set Left Side Boarders Row By Row, Column By Column
    Sheets("Voltage Drop Calculator").Cells(iRow, c).Borders(xlEdgeLeft). _
                                                            ColorIndex = xlAutomatic
    If c < 15 Then
        Sheets("Voltage Drop Calculator").Cells(iRow, c).Borders(xlEdgeBottom). _
                                                            ColorIndex = xlAutomatic
    End If
Next c
'****Set All Cells to Correct Column Width
Sheets("Voltage Drop Calculator").Range("A:N").EntireColumn.AutoFit
'****Set Text in Cells to Center
Sheets("Voltage Drop Calculator").Range("A1:N600").VerticalAlignment = xlCenter
Sheets("Voltage Drop Calculator").Range("A1:N600").HorizontalAlignment = xlCenter

TotalTableValues

'*************************************************************************************'
'           Set the Conduit and Conductors as Drop Down Lists for EOU                 '
'*************************************************************************************'

Call MakeCondDropDownList(User_Info_Panel.ConductType, User_Info_Panel.ConduitType, iRow)

FLAG_FillTable = False

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

Private Sub Worksheet_Change(ByVal Target As Range)

Dim ClmChng, RowChng As Variant
Dim iRow As Long
Dim RandXl, RowVals As Collection
Dim PostChanges As Integer
Dim DevDesc, Gauge, WireType, PipeType As String
Dim Amps, PF, CblLen, Vsupp As Double

' De-Activate Flag so only needed changes are made
PostChanges = 0

' Get the Cell changed Target.Address = $Column$Row on the sheet and pull the column letter off
ClmChng = Mid$(Target.Address, 2, 1)
RowChng = Mid$(Target.Address, 4, 1)

'   We are only deleting the headers so go ahead and exit the function
If RowChng < 7 Or FLAG_FillTable = True Or User_Info_Panel.FLAG_XedOut = True Then
    Exit Sub
End If

' According to column change assessed check if in data set or outside
On Error Resume Next
LastRow = Sheets("Voltage Drop Calculator").Cells.Find(What:="Total", _
                                                       LookAt:=xlPart, _
                                                       LookIn:=xlFormulas, _
                                                       SearchOrder:=xlByRows, _
                                                       SearchDirection:=xlPrevious, _
                                                       MatchCase:=False).Row
If LastRow <> 0 Then
    Sheets("Voltage Drop Calculator").Range(Cells(LastRow - 6, 3), Cells(LastRow, 3)).ClearContents
    Sheets("Voltage Drop Calculator").Range(Cells(LastRow - 6, 3), Cells(LastRow, 3)).ClearFormats
End If
iRow = Sheets("Voltage Drop Calculator").Cells.Find(What:="", _
                                            After:=Cells(2, 7), _
                                            LookAt:=xlWhole, _
                                            LookIn:=xlValues, _
                                            SearchOrder:=xlByRows, _
                                            SearchDirection:=xlNext, _
                                            MatchCase:=False).Row + 1

' Changes made in valid spaces so update and change other cells accordingly
Select Case ClmChng
        Case "A":   Exit Sub   '   Device Description - Do Nothing
        Case "B" Or "D" Or "F" Or "G" Or "H" Or "L" Or "M" Or "N":   '   Current Change - Re-Calculate and Update Table"
                    PostChanges = 1 ' Activate Flag to edit table values
                    
                    Set RowVals = ReadRowChanged(RowChng)  ' Go Get That Row's Values
                    If RowVals.Count <> 8 Then
                        MsgBox ("**Err: Incorrect Value in Cell. Please Fix Highlighted Cell.")
                        Exit Sub
                    Else ' Unload Collection
                        DevDesc = RowVals.Item(1)   ' Var(1): Device Description
                        Amps = RowVals.Item(2)      ' Var(2): Amps
                        PF = RowVals.Item(3)        ' Var(3): Power Factor
                        Gauge = RowVals.Item(4)     ' Var(4): Wire Gauge
                        Phases = RowVals.Item(5)    ' Var(5): Number of Phases
                        CblLen = RowVals.Item(6)    ' Var(6): Est. Cable Length
                        Vsupp = RowVals.Item(7)     ' Var(7): Supply Voltage
                        WireType = RowVals.Item(8)  ' Var(8): Conduit Type
                        PipeType = RowVals.Item(9)  ' Var(9): Conductor Type
                    End If
                    
                    Set RandXl = GetXlandR(WireType, PipeType, Gauge)    'Get Resistance and Reactance
                    ACRes = RandXRange.Item(1)
                    InductReact = RandXRange.Item(2)
                    
                    ThetaRad = Application.Acos(PF)
                    ThetaDeg = (ThetaRad * 180) / WorksheetFunction.Pi
                    Zeff = ACRes * Cos(ThetaRad) + InductReact * Sin(ThetaRad)
                    Zcond = (CblLen / 1000) * Zeff
                    If Phases = 1 Then    '******* Single Phase
                        KVA = Amps * Vsupp / 1000
                        VoltDrop = Amps * 2 * Zcond
                    Else                    '******* Three Phae
                        KVA = Amps * Vsupp * Sqr(3) / 1000
                        VoltDrop = Amps * Sqr(3) * Zcond
                    End If
                    VoltDropPerc = VoltDrop / Vsupp * 100
                    KW = PF * KVA
                    
                    Call UpdateTable(DevDesc, _
                                     Gauge, _
                                     Amps, _
                                     KVA, _
                                     PF, _
                                     KW, _
                                     Phases, _
                                     CblLen, _
                                     Zeff, _
                                     VoltageDrop, _
                                     VoltageDropPerc, _
                                     Vsupp, _
                                     WireType, _
                                     PipeType, _
                                     iRow)
                    Call MakeCondDropDownList(WireType, PipeType, iRow)
                    TotalTableValues
                    
        Case "C":   Exit Sub   '   KVA Change - ?? This should not happen. Post Err ??
        'Case "D":   '   Power Factor Change - Re-Calculate and Update Table
        Case "E":   Exit Sub   '   KW Change - ?? This should not happen. Post Err ??
        'Case "F":   '   Wire Guage Size - Validate Input and Re-Calculate and Update Table
        'Case "G":   '   Phase Number Change - Re-Calculate and Update Table
        'Case "H":   '   Cable Length Change - Re-Calculate and Update Table
        Case "I":   Exit Sub   '   Zeff Change - ?? This should not happen. Post Err ??
        Case "J":   Exit Sub   '   Calculated Voltage Drop Change ?? This should not happen. Post Err ??
        Case "K":   Exit Sub   '   Calculated Voltage Drop % Change ?? This should not happen. Post Err ??
        'Case "L":   '   Supply Voltage Change - Re-Calculate and Update Table
        'Case "M":   '   Conductor Change - Re-Calculate and Update Table
        'Case "N":   '   Conduit Change - Re-Calculate and Update Table
    End Select
End Sub

'-------------------------------------------------------------------------------------'
'---------This Funciton Set the Column Title Block For the Voltage Drop Table---------'
'-------------------------------------------------------------------------------------'

Public Function SetTitleBlock()

'*************************************************************************************'
'                      Variables and Declaration Constants                            '
'*************************************************************************************'

    Dim c As Integer, R As Integer

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
    c = 1               'Set Column Start at Column A
    For R = 4 To 6      'Set Left Side Boarders Row By Row, Column By Column
        Sheets("Voltage Drop Calculator").Cells(R, c).Borders(xlEdgeLeft). _
                                                                ColorIndex = xlAutomatic
        If R = 6 Then   'Have we set all row side boarders 4 to 6
            R = 3       'Reset to 3 so we increment to 4th row on next column
            c = c + 1   'Transition to next column
        End If
        If c = 16 Then  'Have we set all column side boarders
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

Public Function UpdateTable(ByVal DevDescr As String, _
                            ByVal Gauge As String, _
                            ByVal Amp As Double, _
                            ByVal KVA As Double, _
                            ByVal PF As Double, _
                            ByVal KW As Double, _
                            ByVal Phase As Double, _
                            ByVal CblLen As Double, _
                            ByVal Zeff As Double, _
                            ByVal VoltageDrop As Double, _
                            ByVal VoltageDropPerc As Double, _
                            ByVal VoltSupp As Double, _
                            ByVal WireType As String, _
                            ByVal PipeType As String, _
                            ByVal iRow As Long)

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
    Sheets("Voltage Drop Calculator").Cells(iRow, 12) = WireType
    Sheets("Voltage Drop Calculator").Cells(iRow, 12) = PipeType

End Function

'-------------------------------------------------------------------------------------'
'-------This Funciton Generates the Conductor and Conduit Dropdown List Cells---------'
'-------------------------------------------------------------------------------------'

Public Function MakeCondDropDownList(ByVal WireType, PipeType As Variant, iRow As Long)

Dim WireList, PipeList As Variant

' Read Conduit/Conductor Material Using To Get List To Display Correct One On Top
    'WireType = Sheets("Voltage Drop Calculator").Cells(iRow, "M").Value
    'PipeType = Sheets("Voltage Drop Calculator").Cells(iRow, "N").Value
    
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
    
    ' Make Conduit drop down list
    With Sheets("Voltage Drop Calculator").Cells(iRow, "M").Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:=Join(WireList, ",")
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
    ' Make Conductor drop down list
    With Sheets("Voltage Drop Calculator").Cells(iRow, "N").Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:=Join(PipeList, ",")
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
End Function

'-------------------------------------------------------------------------------------'
'This Function Gets the Range Values to Serch in Table 9 for Resistance and Reactance '
'-------------------------------------------------------------------------------------'

Public Function GetXlandR(ByVal ConductType As String, _
                          ByVal ConduitType As String, _
                          ByVal Gauge As Variant) As Collection

Dim ResRange, ReactRange As Range
Dim ResFound, ImpFound As Variant
Dim Var As Collection

Set Var = New Collection

'****Get AC Resistance and Inductive Reactiance From NEC 2017 Table 9
    Select Case ConduitType
        Case "PVC":         'Use PVC Column From Table 9
            Set ResRange = Sheets("Table 9").Range("B7:B27")
            If ConductType = "Copper" Then
                Set ReactRange = Sheets("Table 9").Range("D7:D27")
            Else
                Set ReactRange = Sheets("Table 9").Range("G7:G27")
            End If
        Case "Aluminum":    'Use Aluminum Column From Table 9
            Set ResRange = Sheets("Table 9").Range("B7:B27")
            If ConductType = "Copper" Then
                Set ReactRange = Sheets("Table 9").Range("E7:E27")
            Else:
                Set ReactRange = Sheets("Table 9").Range("H7:H27")
            End If
        Case "Steel":       'Use Steel Column From Table 9
            Set ResRange = Sheets("Table 9").Range("C7:C27")
            If ConductType = "Copper" Then
                Set ReactRange = Sheets("Table 9").Range("F7:F27")
            Else:
                Set ReactRange = Sheets("Table 9").Range("I7:I27")
            End If
    End Select

'Now Go To Table 9 and Get Resistance Value and Inductive Reactance Value
ResFound = Application.WorksheetFunction.Index(ResRange, _
          Application.WorksheetFunction.Match(Gauge, _
                                            Sheets("Table 9").Range("A7:A27").Value, 0))
ImpFound = Application.WorksheetFunction.Index(ReactRange, _
          Application.WorksheetFunction.Match(Gauge, _
                                            Sheets("Table 9").Range("A7:A27").Value, 0))

' Set collection 1:ResVal 2:ReactVal
Var.Add ResFound
Var.Add ImpFound
Set GetXlandR = Var

End Function

Public Function ReadRowChanged(ByVal iRow As Integer) As Collection

Dim Var As Collection
Set Var = New Collection
    ' Var(1): Device Description
    ' Var(2): Amps
    ' Var(3): Power Factor
    ' Var(4): Wire Gauge
    ' Var(5): Number of Phases
    ' Var(6): Est. Cable Length
    ' Var(7): Supply Voltage
    ' Var(8): Conduit Type
    ' Var(9): Conductor Type
    
    Var.Add = Sheets("Voltage Drop Calculator").Cells(iRow, 1).Value    ' Get Device Description

    Amps = Sheets("Voltage Drop Calculator").Cells(iRow, 2).Value
    If CheckIfNum(Amps) = False Then
        Sheets("Voltage Drop Calculator").Cells(iRow, 2).Interior.Color = RGB(211, 100, 100)
        Exit Function
    Else
        Var.Add = CDbl(Amps)
        Sheets("Voltage Drop Calculator").Cells(iRow, 2).Interior.Color = xlNone
    End If
    
    WireGauge = Sheets("Voltage Drop Calculator").Cells(iRow, 6).Text
    If VBA.IsError(Application.Match(WireGauge, Sheets("Table 9").Range("A7:A27").Value, 0)) Then
        Sheets("Voltage Drop Calculator").Cells(iRow, 6).Interior.Color = RGB(211, 100, 100)
        MsgBox ("**Err: Invalid Motor Lead Gauge Size. Please Fix And Try Again.")
        Exit Function
    Else
        Sheets("Voltage Drop Calculator").Cells(iRow, 6).Interior.Color = xlNone
        Var.Add = WireGuage
    End If
    
    NumPhase = Sheets("Voltage Drop Calculator").Cells(iRow, 7).Value
    If CheckIfNum(NumPhase) = False Then
        Sheets("Voltage Drop Calculator").Cells(iRow, 7).Interior.Color = RGB(211, 100, 100)
        Exit Function
    Else
        Var.Add = CDbl(NumPhase)
        Sheets("Voltage Drop Calculator").Cells(iRow, 7).Interior.Color = xlNone
    End If
    
    CblLen = Sheets("Voltage Drop Calculator").Cells(iRow, 8).Value
    If CheckIfNum(CblLen) = False Then
        Sheets("Voltage Drop Calculator").Cells(iRow, 8).Interior.Color = RGB(211, 100, 100)
        Exit Function
    Else
        Var.Add = CDbl(CblLen)
        Sheets("Voltage Drop Calculator").Cells(iRow, 8).Interior.Color = xlNone
    End If
    
    VSupply = Sheets("Voltage Drop Calculator").Cells(iRow, 12).Value
    If CheckIfNum(VSupply) = False Then
        Sheets("Voltage Drop Calculator").Cells(iRow, 12).Interior.Color = RGB(211, 100, 100)
        Exit Function
    Else
        Var.Add = CDbl(VSupply)
        Sheets("Voltage Drop Calculator").Cells(iRow, 12).Interior.Color = xlNone
    End If

    ConductorType = Sheets("Voltage Drop Calculator").Cells(iRow, 13).Text
    If ConduitType <> "Copper" And ConduitType <> "Aluminum" Then
        MsgBox ("**Err: Invalid Conduit Type. Please Fix and Try Again.")
        Sheets("Voltage Drop Calculator").Cells(iRow, 13).Interior.Color = RGB(211, 100, 100)
        Exit Function
    Else
        Var.Add = ConductorType
        Sheets("Voltage Drop Calculator").Cells(iRow, 13).Interior.Color = xlNone
    End If

    ConduitType = Sheets("Voltage Drop Calculator").Cells(iRow, 14).Text
    If ConduitType <> "PVC" And ConduitType <> "Aluminum" And ConduitType <> "Steel" Then
        MsgBox ("**Err: Invalid Conduit Type. Please Fix and Try Again.")
        Sheets("Voltage Drop Calculator").Cells(iRow, 14).Interior.Color = RGB(211, 100, 100)
        Exit Function
    Else
        Var.Add = ConduitType
        Sheets("Voltage Drop Calculator").Cells(iRow, 13).Interior.Color = xlNone
    End If

End Function

