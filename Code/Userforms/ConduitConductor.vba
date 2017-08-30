Public User_ConductorType As String
Public User_ConduitType As String

Private Sub AcceptChoices_Click()
'****Get the data from the two dropdown list's
If ConduitConductor.ConductorType.Value = "" Then
    MsgBox ("**Err: No Conductor Type Chosen. Enter Type and Try Again.")
    Exit Sub
ElseIf ConduitConductor.ConduitType.Value = "" Then
    MsgBox ("**Err: No Conduit Type Chosen. Enter Type and Try Again.")
    Exit Sub
Else
    User_ConductorType = ConduitConductor.ConductorType.Value
    User_ConduitType = ConduitConductor.ConduitType.Value
End If
Hide   'Hide the user box after click complete
End Sub

Public Sub ConduitConductor_Initialize()
    ConduitConductor.ConductorType.Value = ""
    ConduitConductor.ConduitType.Value = ""
End Sub

