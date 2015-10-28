Attribute VB_Name = "OSMultiPeriod"
Sub OSMultiPeriod()

Dim SolverSheet As Worksheet
Set SolverSheet = Sheets("ProcessingSchedule")

'Clear values for decision cells
OpenSolver.GetDecisionVariables(SolverSheet).ClearContents

'Store original decision variables
Dim myVars As Range
Set myVars = OpenSolver.GetDecisionVariables(SolverSheet)

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

'Total solve periods and solve groups. Temporarily set to some quick config cells and buttons. Tidy up later!
solvePeriods = Sheets("OSMultiPeriodSolve").Range("$C$3").Value
solvePeriodStep = Sheets("OSMultiPeriodSolve").Range("$C$4").Value

For j = 1 To solvePeriods Step solvePeriodStep

    Dim solverVars As Range
    Set solverVars = Nothing
    
    'Modify step so it does not exceed total solve periods
    If (j + solvePeriodStep) > solvePeriods Then
        Step = solvePeriods - j + 1
    Else
        Step = solvePeriodStep
    End If

    'Modify each decision variable range to match current solve time period
    For i = 1 To myVars.Areas.Count
        Set currRange = myVars.Areas(i)
        
        If solverVars Is Nothing Then
            Set solverVars = currRange.Columns(j).Resize(, Step)
        Else
            Set solverVars = Union(solverVars, currRange.Columns(j).Resize(, Step))
        End If
    Next i
    
    'Set OpenSolver decision variables
    OpenSolver.SetDecisionVariables solverVars, Sheet:=SolverSheet
    
    'Update OpenSolver constraints
    Dim newLHS As Range
    Dim newRHS As Range
    Dim relation As RelationConsts
    
    For k = 1 To consNum
        Set newLHS = consLHS.Areas((k)).Columns(j).Resize(, Step)
        Set newRHS = consRHS.Areas((k)).Columns(j).Resize(, Step)
        relation = OpenSolver.GetConstraintRel((k), SolverSheet)
        
        OpenSolver.UpdateConstraint (k), newLHS, relation, newRHS, Sheet:=SolverSheet
    Next k

    'Solve OpenSolver model
    OpenSolver.RunOpenSolver Sheet:=SolverSheet
Next j

'Reset OpenSolver decision variables to the original
OpenSolver.SetDecisionVariables myVars, Sheet:=SolverSheet

'Reset OpenSolver constraints to original
For k = 1 To consNum
    OpenSolver.SetConstraintLhs (k), consLHS.Areas((k)), SolverSheet

    OpenSolver.SetConstraintRhs (k), consRHS.Areas((k)), (someString), SolverSheet
Next k

End Sub


