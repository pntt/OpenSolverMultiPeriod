Attribute VB_Name = "OSModifyConstraints"
Sub OSModifyConstraints()

Dim SolverSheet As Worksheet
Set SolverSheet = Sheets("ProcessingSchedule")

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

'For debugging - make sure both are equal
MsgBox "Unique LHS Constraints: " & consLHS.Areas.Count
MsgBox "Unique RHS Constraints: " & consRHS.Areas.Count

startPeriod = 1
stepSize = 5

Dim newLHS As Range
Dim newRHS As Range
Dim relation As RelationConsts

For k = 1 To consNum
    Set newLHS = consLHS.Areas((k)).Columns(startPeriod).Resize(, stepSize)
    Set newRHS = consRHS.Areas((k)).Columns(startPeriod).Resize(, stepSize)
    relation = OpenSolver.GetConstraintRel((k), SolverSheet)
    
    OpenSolver.UpdateConstraint (k), newLHS, relation, newRHS, Sheet:=SolverSheet
    
    'Msgbox debugging
    'MsgBox "LHS: " + newLHS.Address + " RHS: " & newRHS.Address
    'MsgBox "LHS: " & OpenSolver.GetConstraintLhs((k), SolverSheet).Address
    
    
Next k

'Output constraints to OSOut sheet for debugging
For m = 1 To consNum
    Sheets("OSOut").Cells(50 + m - 1, 1) = OpenSolver.GetConstraintLhs((i), SolverSheet).Address
    Sheets("OSOut").Cells(50 + m - 1, 2) = OpenSolver.GetConstraintRhs((i), (rhsString), (rhsDouble), False, SolverSheet).Address
Next m

'Reset OpenSolver constraints to original
For k = 1 To consNum
    OpenSolver.SetConstraintLhs (k), consLHS.Areas((k)), SolverSheet

    OpenSolver.SetConstraintRhs (k), consRHS.Areas((k)), (someString), SolverSheet

Next k

End Sub




