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

MsgBox consLHS.Areas.Count
MsgBox consRHS.Areas.Count

startPeriod = 1
stepSize = 5

Dim newLHS As Range
Dim newRHS As Range


For k = 1 To consNum
    Set newLHS = consLHS.Areas((k)).Columns(startPeriod).Resize(, stepSize)
    OpenSolver.SetConstraintLhs (k), newLHS, SolverSheet
    
    Set newRHS = consRHS.Areas((k)).Columns(startPeriod).Resize(, stepSize)
    Dim someString As String
    OpenSolver.SetConstraintRhs (k), newRHS, (someString), SolverSheet
Next k

'Output to OSOut sheet for debugging
For m = 1 To consNum
    Sheets("OSOut").Cells(50 + m - 1, 1) = consLHS.Areas((m)).Address
    Sheets("OSOut").Cells(50 + m - 1, 2) = consRHS.Areas((m)).Address
Next m




'Reset OpenSolver constraints to original
For k = 1 To consNum
    OpenSolver.SetConstraintLhs (k), consLHS.Areas((k)), SolverSheet
       
    OpenSolver.SetConstraintRhs (k), consRHS.Areas((k)), (someString), SolverSheet
    
Next k

End Sub




