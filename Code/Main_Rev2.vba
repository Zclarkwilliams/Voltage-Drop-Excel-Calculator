'
' Author: Zachary Clark-Williams
' Last Edited: 09-21-2017
'
' Excel Voltage Drop Calculator
'
'   This "New and Improved" calculator code uses a User_Info_Panel userform like a calculator
' pulling all the data from one panel instead of a bunch of smaller input request boxes.
' ** To run properly you will need the User_Info_Panel.vba code from userform file
'

'*************************************************************************************'
'                                 Global Declarations                                 '
'*************************************************************************************'

Public FLAG_NoTable As Boolean
Public FLAG_ClearAll As Boolean
Public DataSetinTable As Boolean
Public FLAG_FillTable As Boolean
Public FLAG_PostChanges As Boolean
Public FLAG_PanelClosed As Boolean
Public FLAG_RowValERRORor As Boolean
Public FLAG_TitleBlockSet As Boolean

Private Sub CalculateVoltageDrop_Click()

Dim iRow As Integer
Dim RandXl As Collection
Dim VoltDrop, ThetaRad, Zeff, Zcond, KVA, VoltDropPerc, KW As Double

'*************************************************************************************'
'                                 Setup FLAGS                                         '
'*************************************************************************************'

FLAG_PanelClosed = False
FLAG_FillTable = True

'*************************************************************************************'
'                 Set the Column Headers Last to AutoFit All Data Entered             '
'*************************************************************************************'
'   Unlock cells since we no longer have data to protect
'Sheets("Voltage Drop Calculator").Unprotect Password:="HGI"
'Sheets("Voltage Drop Calculator").Cells.Locked = False
'   Make all cells text type
Sheets("Voltage Drop Calculator").Range("A4:N600").NumberFormat = Text
'   Check if we need to set up title block header
If FLAG_TitleBlockSet = False Then
    SetTitleBlock
End If

'*************************************************************************************'
'                                 Get User Information                                '
'*************************************************************************************'

User_Info_Panel.User_Info_Panel_Initialize
User_Info_Panel.Show vbModal

'   Check if the user 'X' out of Panel
If FLAG_PanelClosed = True Then
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
        Sheets("Voltage Drop Calculator").Range(Cells(LastRow - 6, 2), Cells(LastRow, 3)).ClearContents
        Sheets("Voltage Drop Calculator").Range(Cells(LastRow - 6, 2), Cells(LastRow, 3)).ClearFormats
        Sheets("Voltage Drop Calculator").Range(Cells(LastRow - 6, 4), Cells(LastRow, 4)).ClearContents
        Sheets("Voltage Drop Calculator").Range(Cells(LastRow - 6, 4), Cells(LastRow, 4)).ClearFormats
End If
Sheets("Voltage Drop Calculator").Range("A1:N600").NumberFormat = Text
iRow = Sheets("Voltage Drop Calculator").Cells.Find(What:="*", _
                                            LookAt:=xlPart, _
                                            LookIn:=xlValues, _
                                            SearchOrder:=xlByRows, _
                                            SearchDirection:=xlPrevious, _
                                            MatchCase:=False).Row + 1
                                                
'*************************************************************************************'
'                       Place Values in Appropriate Columns                           '
'*************************************************************************************'
                
Sheets("Voltage Drop Calculator").Cells(iRow, 1) = User_Info_Panel.DevDesc
Sheets("Voltage Drop Calculator").Cells(iRow, 2) = User_Info_Panel.Amperes
Sheets("Voltage Drop Calculator").Cells(iRow, 3) = Round(KVA, 2)
Sheets("Voltage Drop Calculator").Cells(iRow, 4) = Round(User_Info_Panel.PwrFctr, 2)
Sheets("Voltage Drop Calculator").Cells(iRow, 5) = Round(KW, 2)
Sheets("Voltage Drop Calculator").Cells(iRow, 6) = "'" & User_Info_Panel.WireGauge
Sheets("Voltage Drop Calculator").Cells(iRow, 7) = User_Info_Panel.PhaseNum
Sheets("Voltage Drop Calculator").Cells(iRow, 8) = Round(User_Info_Panel.CableLen, 2)
Sheets("Voltage Drop Calculator").Cells(iRow, 9) = Round(Zeff, 3)
Sheets("Voltage Drop Calculator").Cells(iRow, 10) = Round(VoltDrop, 2)
Sheets("Voltage Drop Calculator").Cells(iRow, 11) = Round(VoltageDropPerc, 2)
Sheets("Voltage Drop Calculator").Cells(iRow, 12) = User_Info_Panel.VoltSupply
Sheets("Voltage Drop Calculator").Cells(iRow, 13) = User_Info_Panel.ConductType
Sheets("Voltage Drop Calculator").Cells(iRow, 14) = User_Info_Panel.ConduitType

'   Protect Cells that should notbe changed (KVA, KW, Votlage Drop, etc.)
'Sheets("Voltage Drop Calculator").Cells(iRow, 3).Locked = True
'Sheets("Voltage Drop Calculator").Cells(iRow, 5).Locked = True
'Sheets("Voltage Drop Calculator").Cells(iRow, 9).Locked = True
'Sheets("Voltage Drop Calculator").Cells(iRow, 10).Locked = True
'Sheets("Voltage Drop Calculator").Cells(iRow, 11).Locked = True
'Sheets("Voltage Drop Calculator").Protect Password:="HGI", UserInterfaceOnly:=True

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

MakeCondDropDownList User_Info_Panel.ConductType, _
                     User_Info_Panel.ConduitType, _
                     iRow

FLAG_FillTable = False

End Sub
'-------------------------------------------------------------------------------------'
'                      Clear the Worksheet for New Calculations                       '
'-------------------------------------------------------------------------------------'
Private Sub ClearSheetVD_Click()
    '   Set flag to not run Worksheet_Chnages function
    FLAG_ClearAll = True
    FLAG_NoTable = True
    '   Unlock cells since we no longer have data to protect
    'Sheets("Voltage Drop Calculator").Unprotect Password:="HGI"
    'Sheets("Voltage Drop Calculator").Range("A7:N600").Locked = False
    '   Now clear all data from table
    Sheets("Voltage Drop Calculator").Range("A1:N600").ClearContents
    '   Clear the formats from all cells
    Sheets("Voltage Drop Calculator").Range("A1:N600").ClearFormats
    '   Set all Boarders to white for clean look.
    Sheets("Voltage Drop Calculator").Range("A1:N600").Borders.Color = RGB(255, 255, 255)
    '   Disable flag so we set title block when run again.
    FLAG_TitleBlockSet = False
    FLAG_ClearAll = False
End Sub

'-------------------------------------------------------------------------------------'
'  This Function Will Take a Variant Input and Output A True or False if Its Number   '
'-------------------------------------------------------------------------------------'

Private Sub Worksheet_Change(ByVal Target As Range)

Dim ClmChng, RowChng, NEC_Current As Variant
Dim iRow As Long
Dim RandXl, RowVals As Collection
Dim PostChanges As Integer
Dim DevDesc, Gauge, WireType, PipeType As String
Dim Amps, PF, CblLen, Vsupp As Double
Dim LastRow As Variant
Dim CurrentRange As Range

'   Clearing Sheet DO NOTHING!
If FLAG_ClearAll = True Or FLAG_PanelClosed = True Then
    Exit Sub
Else
    '   Get the Cell changed Target.Address = $Column$Row on the sheet and pull the column letter off
    ClmChng = Mid$(Target.Address, 2, 1)
    RowChng = Mid$(Target.Address, 4, 1)
    
    '   Check if we even need to make changes of if table not built and this is random
    On Error Resume Next
    LastRow = Sheets("Voltage Drop Calculator").Cells.Find(What:="*", _
                                                    LookAt:=xlPart, _
                                                    LookIn:=xlValues, _
                                                    SearchOrder:=xlByRows, _
                                                    SearchDirection:=xlPrevious, _
                                                    MatchCase:=False).Row
    If LastRow < 7 Then '   There has been no data entered! Exit NOW!
        Exit Sub
    End If
    
    '   Check if WireGuage Changed?
    If RowChng > 6 And ClmChng = "F" Then   '   Yes, Check Gauge vs. NEC Ampacity Rating
        '   Get necessary parameters
        Amps = Sheets("Voltage Drop Calculator").Cells(RowChng, "B").Value
        Gauge = Sheets("Voltage Drop Calculator").Cells(RowChng, ClmChng).Value
        Voltage = Sheets("Voltage Drop Calculator").Cells(RowChng, "L").Value
        If Voltage Is Nothing Then
            Voltage = User_Info_Panel.VoltSupply
        End If
        WireType = Sheets("Voltage Drop Calculator").Cells(RowChng, "M").Value
        If WireType Is Nothing Then
            WireType = User_Info_Panel.ConductType
        End If
        
        
        If Voltage < 2001 Then  '   If we are under 2000V -> Table 310.15(B)(16)
            If WireType = "Copper" Then   '   Focus on the Copper columns in table
                If Amps > 100 Then   '   If over 100A use 75F column
                    Set CurrentRange = Sheet3.Range("C5:C35")
                Else   '   We are under 100A use 60F Column
                    Set CurrentRange = Sheet3.Range("B5:B35")
                End If
            Else   '   Focus on the Aluminum/Copper-Clad columns in table
                If Amps > 100 Then  '   If over 100A use 75F column
                    Set CurrentRange = Sheet3.Range("F5:F35")
                Else   '   We are under 100A use 60F Column
                    Set CurrentRange = Sheet3.Range("E5:E35")
                End If
            End If
        Else                      '   If we are over 2000V -> Table 310.15(B)(17)
            If WireType = "Copper" Then   '   Focus on the Copper columns in table
                If Amps > 100 Then  '   If over 100A use 75F column
                    Set CurrentRange = Sheet3.Range("K5:K35")
                Else   '   We are under 100A use 60F Column
                    Set CurrentRange = Sheet3.Range("J5:J35")
                End If
            Else   '   Focus on the Aluminum/Copper-Clad columns in table
                If Amps > 100 Then  '   If over 100A use 75F column
                    Set CurrentRange = Sheet3.Range("N5:N35")
                Else   '   We are under 100A use 60F Column
                    Set CurrentRange = Sheet3.Range("M5:M35")
                End If
            End If
        End If
        
        '   Go get the ampacity from NEC17 Table 310.15(B)(16-17) for found conductor
        NEC_Current = Application.WorksheetFunction.Index(CurrentRange, _
                      Application.WorksheetFunction.Match(Gauge, _
                                                          Sheet3.Range("A5:A34").Value, _
                                                          0))
        '   Compare Found Current to NEC17 Ampacity
        If NEC_Current = "-" Then
            MsgBox ("Unable to verify or confirm porper NEC ampacity regulations met. Check NEC 2017 240.4(D) if necessary.")
        ElseIf Amps > CDbl(NEC_Current) Then
            MsgBox ("BEWARE, you have surpassed NEC 2017 ampacity regulations for entered conductor size!")
        End If
    End If
    If FLAG_PostChanges = False Then    '   No, check flags for needs to adjust or not
        '   We are only deleting the headers so go ahead and exit the function
        If FLAG_NoTable = True Or FLAG_FillTable = True Then
            Exit Sub
        End If
        
        ' Activate Flag to so we know we are automaing the sheet changes
        FLAG_PostChanges = True
        
        ' According to column change assessed check if in data set or outside
        If Sheets("Voltage Drop Calculator").Cells(LastRow, "C").Value = "Total kVA: " Then
            Sheets("Voltage Drop Calculator").Range(Cells(LastRow - 6, 2), Cells(LastRow, 3)).ClearContents
            Sheets("Voltage Drop Calculator").Range(Cells(LastRow - 6, 2), Cells(LastRow, 3)).ClearFormats
            Sheets("Voltage Drop Calculator").Range(Cells(LastRow - 6, 4), Cells(LastRow, 5)).ClearContents
            Sheets("Voltage Drop Calculator").Range(Cells(LastRow - 6, 4), Cells(LastRow, 5)).ClearFormats
        End If
        iRow = Sheets("Voltage Drop Calculator").Cells.Find(What:="*", _
                                                    LookAt:=xlPart, _
                                                    LookIn:=xlValues, _
                                                    SearchOrder:=xlByRows, _
                                                    SearchDirection:=xlPrevious, _
                                                    MatchCase:=False).Row
    
        ' Changes made in valid spaces so update and change other cells accordingly
        '       Case A: Device Description - Do Nothing
        '       Case B: Current Change - Re-Calculate and Update Table"
        '       Case C: KVA Change - ?? This should not happen. Post ERROR ??
        '       Case D: Power Factor Change - Re-Calculate and Update Table
        '       Case E: KW Change - ?? This should not happen. Post ERROR ??
        '       Case F: Wire Guage Size - Validate Input and Re-Calculate and Update Table
        '       Case G: Phase Number Change - Re-Calculate and Update Table
        '       Case H: Cable Length Change - Re-Calculate and Update Table
        '       Case I: Zeff Change - ?? This should not happen. Post ERROR ??
        '       Case J: Calculated Voltage Drop Change ?? This should not happen. Post ERROR ??
        '       Case K: Calculated Voltage Drop % Change ?? This should not happen. Post ERROR ??
        '       Case L: Supply Voltage Change - Re-Calculate and Update Table
        '       Case M: Conductor Change - Re-Calculate and Update Table
        '       Case N: Conduit Change - Re-Calculate and Update Table
        
        Select Case ClmChng 'ClmChng
                Case "B", "D", "F", "G", "H", "L", "M", "N":
                        '   Go Get That Row's Values
                        Set RowVals = ReadRowChanged(RowChng)
                        If FLAG_RowValERRORor = True Then
                            FLAG_RowValERRORor = False
                            FLAG_PostChanges = False
                            Exit Sub
                        Else ' Unload Row Data Collection
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
                        '   Get Resistance and Reactance
                        Set RandXl = GetXlandR(WireType, PipeType, Gauge)
                        ACRes = RandXl.Item(1)
                        InductReact = RandXl.Item(2)
                        '   All Data Collected, TIME TO MATHS
                        ThetaRad = Application.Acos(PF)
                        Zeff = ACRes * Cos(ThetaRad) + InductReact * Sin(ThetaRad)
                        Zcond = (CblLen / 1000) * Zeff
                        If Phases = 1 Then    ' Single Phase
                            KVA = Amps * Vsupp / 1000
                            VoltDrop = Amps * 2 * Zcond
                        Else                  ' Three Phase
                            KVA = Amps * Vsupp * Sqr(3) / 1000
                            VoltDrop = Amps * Sqr(3) * Zcond
                        End If
                        VoltDropPerc = VoltDrop / Vsupp * 100
                        KW = PF * KVA
                        '   Set Flag so we fill table without sheet_change re-runs
                        FLAG_FillTable = True
                        '   Go Fill table with updated values
                        UpdateTable DevDesc, _
                                    Gauge, _
                                    Amps, _
                                    KVA, _
                                    PF, _
                                    KW, _
                                    Phases, _
                                    CblLen, _
                                    Zeff, _
                                    VoltDrop, _
                                    VoltDropPerc, _
                                    Vsupp, _
                                    WireType, _
                                    PipeType, _
                                    RowChng
                        '   Make Conductor and conduit dropdown lists
                        MakeCondDropDownList WireType, _
                                             PipeType, _
                                             RowChng
                        '   ReDo Totals and Print to Table
                        TotalTableValues
                        '   Deactivate flag for table fill
                        FLAG_FillTable = False
                Case Else:   Exit Sub   '   All other edits shall be ignored
            End Select
    End If
End If
'   Disable flag for re-calc and format table
FLAG_PostChanges = False

End Sub

'-------------------------------------------------------------------------------------'
'---------This Funciton Set the Column Title Block For the Voltage Drop Table---------'
'-------------------------------------------------------------------------------------'

Public Function SetTitleBlock()

'*************************************************************************************'
'                      Variables and Declaration Constants                            '
'*************************************************************************************'

    Dim C As Integer, R As Integer

'*************************************************************************************'
'                        Write Headers for Table Columns                              '
'*************************************************************************************'
    Sheets("Voltage Drop Calculator").Range("A4:L600").NumberFormat = Text
    Sheets("Voltage Drop Calculator").Cells(5, 1) = "Load Device Description"
    Sheets("Voltage Drop Calculator").Cells(5, 2) = "Amperes"
    Sheets("Voltage Drop Calculator").Cells(5, 3) = "KVA"
    Sheets("Voltage Drop Calculator").Cells(5, 4) = "PF"
    Sheets("Voltage Drop Calculator").Cells(5, 5) = "KW"
    Sheets("Voltage Drop Calculator").Cells(5, 6) = "Gauge Size #"
    Sheets("Voltage Drop Calculator").Cells(4, 7) = "Number"
    Sheets("Voltage Drop Calculator").Cells(5, 7) = "of"
    Sheets("Voltage Drop Calculator").Cells(6, 7) = "Phases"
    Sheets("Voltage Drop Calculator").Cells(4, 8) = "Estimated"
    Sheets("Voltage Drop Calculator").Cells(5, 8) = "Cable Length"
    Sheets("Voltage Drop Calculator").Cells(6, 8) = "in Feet"
    Sheets("Voltage Drop Calculator").Cells(4, 9) = "Effective"
    Sheets("Voltage Drop Calculator").Cells(5, 9) = "Z Per"
    Sheets("Voltage Drop Calculator").Cells(6, 9) = "1000 ft"
    Sheets("Voltage Drop Calculator").Cells(4, 10) = "Voltage"
    Sheets("Voltage Drop Calculator").Cells(5, 10) = "Drop"
    Sheets("Voltage Drop Calculator").Cells(6, 10) = "(V)"
    Sheets("Voltage Drop Calculator").Cells(4, 11) = "Voltage"
    Sheets("Voltage Drop Calculator").Cells(5, 11) = "Drop"
    Sheets("Voltage Drop Calculator").Cells(6, 11) = "Percent (%)"
    Sheets("Voltage Drop Calculator").Cells(4, 12) = "Supply"
    Sheets("Voltage Drop Calculator").Cells(5, 12) = "Voltage"
    Sheets("Voltage Drop Calculator").Cells(6, 12) = "(V)"
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
    '   Border the Cells
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
    
    '   Set middle borders to grey so we don't have white boarders lines in header
    Sheets("Voltage Drop Calculator").Range("A4:N4").Borders(xlEdgeBottom). _
                                                                Color = RGB(211, 211, 211)
    Sheets("Voltage Drop Calculator").Range("A6:N6").Borders(xlEdgeTop). _
                                                                Color = RGB(211, 211, 211)
    
    '   Set All Cells to Correct Column Width
    Sheets("Voltage Drop Calculator").Range("A:N").EntireColumn.AutoFit
    '   Set Text in Cells to Center
    Sheets("Voltage Drop Calculator").Range("A4:N6").VerticalAlignment = xlCenter
    Sheets("Voltage Drop Calculator").Range("A4:N6").HorizontalAlignment = xlCenter
    '   Set All Title Block Cells to Grey Coloring
    Sheets("Voltage Drop Calculator").Range("A4:N6").Interior.Color = RGB(211, 211, 211)
    
    
    '   Protect Table Header
    'Sheets("Voltage Drop Calculator").Range("A4:N7").Locked = True
    'Sheets("Voltage Drop Calculator").Protect Password:="HGI", UserInterfaceOnly:=True
    
    '   Enable Flag saying that we are set and dont need to do this again.
    FLAG_TitleBlockSet = True
    
End Function

'-------------------------------------------------------------------------------------'
'---This Funciton Summizes KVA, A, and KW from the table entries and displays them ---'
'-------------------------------------------------------------------------------------'
Function TotalTableValues()

Dim iRow, GETOUT, Flag, LastRow As Integer
Dim ContRun As String
Dim Amps As Variant

'****Get Number of Rows With Valid Data Present
iRow = Sheets("Voltage Drop Calculator").Cells.Find(What:="*", _
                                            LookAt:=xlPart, _
                                            LookIn:=xlValues, _
                                            SearchOrder:=xlByRows, _
                                            SearchDirection:=xlPrevious, _
                                            MatchCase:=False).Row

'   Set the Start position on first row of data
iRow = 7

'ActiveSheet.Rows(LastRow + 1).ClearContents
'ActiveSheet.Rows(LastRow + 1).ClearFormats

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
    Sheets("Voltage Drop Calculator").Cells(iRow + 5, 2) = "Total Current:"
    Sheets("Voltage Drop Calculator").Cells(iRow + 5, 4) = Amperes_Total
    Sheets("Voltage Drop Calculator").Cells(iRow + 6, 2) = "Overall Total KW: "
    Sheets("Voltage Drop Calculator").Cells(iRow + 6, 4) = KW_Total
    Sheets("Voltage Drop Calculator").Cells(iRow + 7, 2) = "Total kVA: "
    Sheets("Voltage Drop Calculator").Cells(iRow + 7, 4) = KVA_Total
    
End Function

'-------------------------------------------------------------------------------------'
'  This Function Will Take a Variant Input and Output A True or False if Its Number   '
'-------------------------------------------------------------------------------------'

Function CheckIfNum(ByVal NumCheck As Variant) As Boolean
    If IsNumeric(NumCheck) = False Then                   'Check if a number
        MsgBox ("**ERROR: Non-Numeric Value Entered. Please Re-Try.") 'Output ERRORor msg
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
                            ByVal iRow As Variant)

    Sheets("Voltage Drop Calculator").Cells(iRow, 1) = DevDescr
    Sheets("Voltage Drop Calculator").Cells(iRow, 2) = Amp
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
    Sheets("Voltage Drop Calculator").Cells(iRow, 13) = WireType
    Sheets("Voltage Drop Calculator").Cells(iRow, 14) = PipeType

    '   Protect Cells that should notbe changed (KVA, KW, Votlage Drop, etc.)
    'Sheets("Voltage Drop Calculator").Cells(iRow, 3).Locked = True
    'Sheets("Voltage Drop Calculator").Cells(iRow, 5).Locked = True
    'Sheets("Voltage Drop Calculator").Cells(iRow, 9).Locked = True
    'Sheets("Voltage Drop Calculator").Cells(iRow, 10).Locked = True
    'Sheets("Voltage Drop Calculator").Cells(iRow, 11).Locked = True
    'Sheets("Voltage Drop Calculator").Protect Password:="HGI", UserInterfaceOnly:=True

End Function

'-------------------------------------------------------------------------------------'
'-------This Funciton Generates the Conductor and Conduit Dropdown List Cells---------'
'-------------------------------------------------------------------------------------'

Public Function MakeCondDropDownList(ByVal WireType As Variant, _
                                     ByVal PipeType As Variant, _
                                     ByVal iRow As Variant)

Dim WireList, PipeList As Variant

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
        .ERRORorTitle = ""
        .InputMessage = ""
        .ERRORorMessage = ""
        .ShowInput = True
        .ShowERRORor = True
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
        .ERRORorTitle = ""
        .InputMessage = ""
        .ERRORorMessage = ""
        .ShowInput = True
        .ShowERRORor = True
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

'Now Go To Table 9 and Get AC Resistance Value and Inductive Reactance Value
ResFound = Application.WorksheetFunction.Index(ResRange, _
           Application.WorksheetFunction.Match(Gauge, _
                                               Sheet5.Range("A7:A27").Value, 0))
ImpFound = Application.WorksheetFunction.Index(ReactRange, _
           Application.WorksheetFunction.Match(Gauge, _
                                               Sheet5.Range("A7:A27").Value, 0))

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
    
    Var.Add Sheets("Voltage Drop Calculator").Cells(iRow, 1).Value    ' Get Device Description

    Amps = Sheets("Voltage Drop Calculator").Cells(iRow, 2).Value
    If CheckIfNum(Amps) = False Then
        Sheets("Voltage Drop Calculator").Cells(iRow, 2).Interior.Color = RGB(211, 100, 100)
        MsgBox ("**ERROR: Invalid Current Value. Please Fix And Try Again.")
        FLAG_RowValERRORor = True
        Exit Function
    Else
        Var.Add CDbl(Amps)
        Sheets("Voltage Drop Calculator").Cells(iRow, 2).Interior.Color = xlNone
    End If
    
    PF = Sheets("Voltage Drop Calculator").Cells(iRow, 4).Value
    If CheckIfNum(PF) = False Then
        Sheets("Voltage Drop Calculator").Cells(iRow, 4).Interior.Color = RGB(211, 100, 100)
        MsgBox ("**ERROR: Invalid Power Factor Value. Please Fix And Try Again.")
        FLAG_RowValERRORor = True
        Exit Function
    Else
        Var.Add CDbl(PF)
        Sheets("Voltage Drop Calculator").Cells(iRow, 4).Interior.Color = xlNone
    End If
    
    WireGauge = Sheets("Voltage Drop Calculator").Cells(iRow, 6).Text
    If VBA.IsError(Application.Match(WireGauge, Sheets("Table 9").Range("A7:A27").Value, 0)) Then
        Sheets("Voltage Drop Calculator").Cells(iRow, 6).Interior.Color = RGB(211, 100, 100)
        MsgBox ("**ERROR: Invalid Motor Lead Gauge Size (May need to input #/0 as '#/0). Please Fix And Try Again.")
        FLAG_RowValERRORor = True
        Exit Function
    Else
        Sheets("Voltage Drop Calculator").Cells(iRow, 6).Interior.Color = xlNone
        Var.Add WireGauge
    End If
    
    NumPhase = Sheets("Voltage Drop Calculator").Cells(iRow, 7).Value
    If NumPhase <> 1 And NumPhase <> 3 Then
        Sheets("Voltage Drop Calculator").Cells(iRow, 7).Interior.Color = RGB(211, 100, 100)
        MsgBox ("**ERROR: Invalid Number of Phases. Please Fix And Try Again.")
        FLAG_RowValERRORor = True
        Exit Function
    Else
        Var.Add NumPhase
        Sheets("Voltage Drop Calculator").Cells(iRow, 7).Interior.Color = xlNone
    End If
    
    CblLen = Sheets("Voltage Drop Calculator").Cells(iRow, 8).Value
    If CheckIfNum(CblLen) = False Then
        Sheets("Voltage Drop Calculator").Cells(iRow, 8).Interior.Color = RGB(211, 100, 100)
        MsgBox ("**ERROR: Invalid Cable Length Value. Please Fix And Try Again.")
        FLAG_RowValERRORor = True
        Exit Function
    Else
        Var.Add CDbl(CblLen)
        Sheets("Voltage Drop Calculator").Cells(iRow, 8).Interior.Color = xlNone
    End If
    
    VSupply = Sheets("Voltage Drop Calculator").Cells(iRow, 12).Value
    If CheckIfNum(VSupply) = False Then
        Sheets("Voltage Drop Calculator").Cells(iRow, 12).Interior.Color = RGB(211, 100, 100)
        MsgBox ("**ERROR: Invalid Voltage Supply Value. Please Fix And Try Again.")
        FLAG_RowValERRORor = True
        Exit Function
    Else
        Var.Add CDbl(VSupply)
        Sheets("Voltage Drop Calculator").Cells(iRow, 12).Interior.Color = xlNone
    End If

    ConductorType = Sheets("Voltage Drop Calculator").Cells(iRow, 13).Text
    If ConductorType <> "Copper" And ConductorType <> "Aluminum" Then
        MsgBox ("**ERROR: Invalid Conductor Type. Please Fix and Try Again.")
        Sheets("Voltage Drop Calculator").Cells(iRow, 13).Interior.Color = RGB(211, 100, 100)
        FLAG_RowValERRORor = True
        Exit Function
    Else
        Var.Add ConductorType
        Sheets("Voltage Drop Calculator").Cells(iRow, 13).Interior.Color = xlNone
    End If

    ConduitType = Sheets("Voltage Drop Calculator").Cells(iRow, 14).Text
    If ConduitType <> "PVC" And ConduitType <> "Aluminum" And ConduitType <> "Steel" Then
        MsgBox ("**ERROR: Invalid Conduit Type. Please Fix and Try Again.")
        Sheets("Voltage Drop Calculator").Cells(iRow, 14).Interior.Color = RGB(211, 100, 100)
        FLAG_RowValERRORor = True
        Exit Function
    Else
        Var.Add ConduitType
        Sheets("Voltage Drop Calculator").Cells(iRow, 13).Interior.Color = xlNone
    End If

Set ReadRowChanged = Var

End Function

