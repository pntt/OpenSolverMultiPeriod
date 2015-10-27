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
    MsgBox "consLHS: " + consLHS.Address

    If consRHS Is Nothing Then
        Set consRHS = OpenSolver.GetConstraintRhs((i), (rhsString), (rhsDouble), False, SolverSheet)
    Else
        Set consRHS = Union(consRHS, OpenSolver.GetConstraintRhs((i), (rhsString), (rhsDouble), False, SolverSheet))
    End If
    
    
    MsgBox "consRHS: " + consRHS.Address
    
    
Next i


'MsgBox consLHS.Address
'MsgBox consRHS.Address
    
    


End Sub




