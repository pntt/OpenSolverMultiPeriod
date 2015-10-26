Attribute VB_Name = "OSMultiPeriod"
Sub test()

Dim SolverSheet As Worksheet
Set SolverSheet = Sheets("ProcessingSchedule")

Dim myVars As Range
Set myVars = OpenSolver.GetDecisionVariables(SolverSheet)

MsgBox myVars.Areas.Count
MsgBox myVars.Address

solvePeriods = 34
solvePeriodStep = 10
counter = 1


For j = 1 To solvePeriods Step solvePeriodStep

Dim solverVars As Range
Set solverVars = Nothing

If (j + solvePeriodStep) > solvePeriods Then
    step = solvePeriods - j + 1
Else
    step = solvePeriodStep
End If

    For i = 1 To myVars.Areas.Count
    
        Set currRange = myVars.Areas(i)
        
        If solverVars Is Nothing Then
            Set solverVars = currRange.Columns(j).Resize(, step)
        Else
            Set solverVars = Union(solverVars, currRange.Columns(j).Resize(, step))
        End If
            
        Sheets("OSOut").Cells(i, 2 + counter) = solverVars.Areas(i).Address
        
    
    Next i
    
    'Set OpenSolver decision variables
    OpenSolver.SetDecisionVariables solverVars, Sheet:=SolverSheet
    
    
    'Solve OpenSolver model
    OpenSolver.RunOpenSolver Sheet:=SolverSheet
    
    
    counter = counter + 1
    
Next j

'Reset OpenSolver decision variables to the original
OpenSolver.SetDecisionVariables myVars, Sheet:=SolverSheet
    

End Sub


