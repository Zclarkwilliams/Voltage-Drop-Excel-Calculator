Public MotorRunType As Boolean

Private Sub Continuous_Click()
   MotorRunType = ContRunning.Continuous.Value = True
   Hide
End Sub

Private Sub NotContinuous_Click()
    MotorRunType = ContRunning.NotContinuous.Value = False
    Hide
End Sub

Public Sub ContRunning_Initialize()
    Me.Continuous.Value = False
    Me.NotContinuous.Value = False
End Sub
