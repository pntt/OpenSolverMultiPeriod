Attribute VB_Name = "OSModifyConstraints"
Sub OSModifyConstraints()

Dim SolverSheet As Worksheet
Set SolverSheet = Sheets("ProcessingSchedule")

'MsgBox OpenSolver.GetConstraintLhs(2, SolverSheet).Address

'MsgBox OpenSolver.GetConstraintLhs(2, SolverSheet).Columns(1).Resize(, 10).Address

'Number of constraints
Dim consNum As Long
consNum = OpenSolver.GetNumConstraints(SolverSheet)

'Store original contraints
Dim consLHS As Range
Dim consRHS As Range


For i = 1 To consNum
    If consLHS Is Nothing Then
        Set consLHS = OpenSolver.GetConstraintLhs((i), SolverSheet)
    Else
        Set consLHS = Union(consLHS, OpenSolver.GetConstraintLhs((i), SolverSheet))
    End If

    If consRHS Is Nothing Then
        Set consRHS = OpenSolver.GetConstraintRhs((i), (rhsString), (rhsDouble), False, SolverSheet)
    Else
        Set consRHS = Union(consRHS, OpenSolver.GetConstraintRhs((i), (rhsString), (rhsDouble), False, SolverSheet))
    End If
  
Next i

startPeriod = 1
stepSize = 5

For j = 1 To consNum
    OpenSolver.SetConstraintLhs (j), consLHS.Areas((startPeriod)).Columns(1).Resize(, stepSize), SolverSheet
    Dim someString As String
    
    OpenSolver.SetConstraintRhs (j), consLHS.Areas((j)).Columns(startPeriod).Resize(, stepSize), someString, SolverSheet
Next j

End Sub




