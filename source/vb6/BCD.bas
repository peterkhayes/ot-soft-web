Attribute VB_Name = "BCD"
'This module was programmed by Bruce Tesar.  Comments are identified with BT.
        
    '***********************************************************************************
    'BH:  I propose to add three things to this code.  Original code has been saved
    '     to make this all undoable.
    
        'I.  I have added statements that monitor the progress of the algorithm in the
        '    output file "HowIRanked" + gFileName  + ".txt"; this file can be accessed from the View menu.
        
        'II. I have added code, and an option under the Options menu, to constrain
        '    BCD to rank specific constraints above general ones.  This code is labeled
        '    "Interpolation" below.
        
        'III. I have added code that flashes a warning to the user that BCD has made
        '     an arbitrary choice in selecting best available subset, in a tied
        '     situation.  This warning is also repeated in the "HowIRanked" + gFileName  + ".txt" file.
    '***********************************************************************************
    
    Option Explicit
    
    
    Dim Subset() As Boolean
    Dim Subsetted() As Boolean              'Excluded from current stratum by the existence of
    Dim ShowMe As Long
    
    'Watch out for arbitrary rankings based on ties and warn the user.
        Dim mUpperTieFlag As Boolean
        Dim mTieFlag As Boolean
        
    'This module's copies of arrays generally used in OTSoft.
        Dim mMaximumNumberOfRivals As Long
        Dim mNumberOfConstraints As Long
        Dim mAbbrev() As String
        Dim mConstraintName() As String
        Dim mFaithfulness() As Boolean

    
Function Main(ByVal NumberOfForms As Long, NumberOfRivals() As Long, _
   WinnerViolations() As Long, RivalViolations() As Long, _
   NumberOfConstraints As Long, Abbrev() As String, ConstraintName() As String) As String
   
    'BH for why this returns a string, see the module of Form1 that calls this.

    Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long

    'BH:  localize the general OTSoft variables
        Let mNumberOfConstraints = NumberOfConstraints
        ReDim mAbbrev(mNumberOfConstraints)
        ReDim mConstraintName(mNumberOfConstraints)
        ReDim mFaithfulness(mNumberOfConstraints)
        For ConstraintIndex = 1 To mNumberOfConstraints
            Let mAbbrev(ConstraintIndex) = Abbrev(ConstraintIndex)
            Let mConstraintName(ConstraintIndex) = ConstraintName(ConstraintIndex)
        Next ConstraintIndex
        
    'BT:  Dimension the local variables.
        Dim Stratum() As Long       'BT:  Indicates stratum of each constraint (0=unranked)
        ReDim Stratum(mNumberOfConstraints)
        Dim CurrentStratum As Long  'BT:  Contains the index of the current stratum
   
    Dim InnerConstraintIndex As Long    'BH:  interpolation:  needed for subset code
    Dim StratumIndex As Long            'BH:  needed to report progress
    
    Let mMaximumNumberOfRivals = Form1.FindMaximumNumberOfRivals(NumberOfRivals())
    Dim StillInformative() As Boolean
    ReDim StillInformative(NumberOfForms, mMaximumNumberOfRivals)
    Dim Demotable() As Boolean  'BT:  Indicates that a constraint MUST be demoted
    ReDim Demotable(mNumberOfConstraints)
    ReDim Subsetted(mNumberOfConstraints)    '   a more specific version of the same Faithfulness constraint.
 
    Dim RankableMarked As Boolean  'BT:  Can a markedness constraint be ranked now
    Dim FaithIsActive As Boolean   'BT:  Is there an active faithfulness constraint
    Dim UnrankedConstraintsRemain As Boolean
    Dim AConstraintWasRanked As Boolean
    Dim Active() As Boolean
    ReDim Active(mNumberOfConstraints)
    
    'BH:  interpolation.  The Subsetted() array says which Faithfulness constraints
    '     suffer from the presence of a more specific version in the same grammar.
    
        If Form1.mnuSpecificBCD Then
            ReDim Subsetted(mNumberOfConstraints)
        End If

    'Initialize crucial variables, in case this routine gets called more than once.

    Let CurrentStratum = 0

    For ConstraintIndex = 1 To mNumberOfConstraints
        Let Stratum(ConstraintIndex) = 0
    Next ConstraintIndex

    For FormIndex = 1 To NumberOfForms
        For RivalIndex = 1 To mMaximumNumberOfRivals
            Let StillInformative(FormIndex, RivalIndex) = True
        Next RivalIndex
    Next FormIndex
    
    'Admit that there is no a priori ranking capacity in this routine.
        If Form1.mnuConstrainAlgorithmsByAPrioriRankings.Checked = True Then
            Select Case MsgBox("Sorry, you've selected the use of a priori rankings, but OTSoft doesn't implement a priori rankings for Biased Constraint Demotion.  Please contact me (bhayes@humnet.ucla.edu) if this capacity is important to you." + _
                Chr(10) + Chr(10) + _
                "Click Yes to continue without the use of a priori rankings, No to exit OTSoft.", vbYesNo + vbExclamation)
                Case vbYes
                    'do nothing
                Case vbNo
                    End
            End Select
        End If
    
    'First, make sure there is a folder for these files, a daughter of the
    '   folder in which the input file is located.
        Call Form1.CreateAFolderForOutputFiles
    
    'BH:  Print out a file to show how you did it.
        Dim FoundOne As Boolean         'Used in self-report file to indicate
                                        '    searches that found nothing.
        Let ShowMe = FreeFile
        Open gOutputFilePath + "HowIRanked" + gFileName + ".txt" For Output As #ShowMe
        Print #ShowMe, "******Application of Biased Constraint Demotion******"
        Print #ShowMe,
        Print #ShowMe, "Input file:  "; gInputFilePath + gFileName + gFileSuffix
        Print #ShowMe,
        'BH:  Interpolation:  if you used the specificity-ranking principle, say so in
        '     the output file.
            If Form1.mnuSpecificBCD.Checked = True Then
                Print #ShowMe, "Version of BCD used:  modified, gives priority to more specific Faithfulness constraints"
                Print #ShowMe,
            End If
    
    'Determine which constraints are faithfulness constraints.
        For ConstraintIndex = 1 To mNumberOfConstraints
            Let mFaithfulness(ConstraintIndex) = Form1.FaithfulnessConstraint(mConstraintName(ConstraintIndex))
        Next ConstraintIndex
    
    'BH:  interpolation.  To take the specific/general Faithfulness
    '  constraint distinction into account, you need to run the specificity
    '  detector.
    
        If Form1.mnuSpecificBCD.Checked = True Then
            Call LocateViolationSubsets(NumberOfForms, WinnerViolations(), NumberOfRivals(), RivalViolations())
        End If
          
    'BH:  Variables (module level) for detecting when BCD made an
    '  arbitrary decision to rank a particular subset of Faithfulness
    '  constraints.
    
        'mUpperTieFlag keeps track of whether any given stratum was selected
        '  arbitrarily.  Once set to True, it cannot be reset to False.
            Let mUpperTieFlag = False
        
        'TieFlag, tout court, keeps track of whether, at any given point in the
        '  construction of a stratum, the subset selection is arbitrary.  If set
        '  to True, it can be reset to False when a better subset is found.
            Let mTieFlag = False
        
        
    'Go through the Winner-Rival pairs repeatedly, looking for constraints that
    '  never prefer a loser among the pairs still being considered.

    Do
   
        'BH:  Record what stratum you're constructing, reporting this on the ShowMe file:
            Let CurrentStratum = CurrentStratum + 1
            Print #ShowMe,
            Print #ShowMe, "******Now doing Stratum #"; Trim(Str(CurrentStratum)); "******"
            Print #ShowMe,

        'BH:  If you're on any iteration of this Do loop other than the first, you may
        '  have set the local TieFlag at True on the preceding iteration, indicating
        '  an arbitrary choice.  If so, record the fact permanently by setting the
        '  upper TieFlag.
            If mTieFlag = True Then Let mUpperTieFlag = True
        
        'BT:  Initialize variables indicating which constraints must be demoted and
        'which constraints are active (prefer the winner in at least one pair).
        
        For ConstraintIndex = 1 To mNumberOfConstraints
            Let Demotable(ConstraintIndex) = False
            Let Active(ConstraintIndex) = False
            'BH:  interpolation:  reinitialize the variable that indicates constraints
            '     penalized by the presence of a more specific constraint.  This changes
            '     from stratum to stratum, as the more-specific constraints join the
            '     newly created strata.
                Let Subsetted(ConstraintIndex) = False
        Next ConstraintIndex

        'BT:  Go through all pairs of Winner vs. Rival, marking constraints preferring
        'a rival as demotable (MUST be demoted) and constraints preferring a winner
        'as active.
          
        For FormIndex = 1 To NumberOfForms
            For RivalIndex = 1 To NumberOfRivals(FormIndex)

                'BT:  Only still-informative Rivals can be learned from:
                If StillInformative(FormIndex, RivalIndex) = True Then
                    For ConstraintIndex = 1 To mNumberOfConstraints

                        'BT:  For each yet-unranked constraint
                        If Stratum(ConstraintIndex) = 0 Then
                        
                            'BT:  If it prefers the rival, mark as demotable
                            If WinnerViolations(FormIndex, ConstraintIndex) > _
                                RivalViolations(FormIndex, RivalIndex, ConstraintIndex) Then
                                Let Demotable(ConstraintIndex) = True
                                
                            'BT:  If it prefers the winner, mark as active
                            ElseIf WinnerViolations(FormIndex, ConstraintIndex) < _
                                RivalViolations(FormIndex, RivalIndex, ConstraintIndex) Then
                                Let Active(ConstraintIndex) = True
                            End If
                            
                        End If

                    Next ConstraintIndex
                End If  'BT:  StillInformative
                
            Next RivalIndex
        Next FormIndex

        'BT:  See if any markedness constraints can be ranked next; if so, rank them.
        'Meanwhile, also check if there are any active faithfulness constraints
        '(those that prefer at least one winner), in case no markedness
        'constraints can be ranked.

        'BH:  report progress
            Print #ShowMe, "Faithfulness Delay:"
        
        Let RankableMarked = False
        Let FaithIsActive = False
                
        For ConstraintIndex = 1 To mNumberOfConstraints
            If Stratum(ConstraintIndex) = 0 Then
            
                If Demotable(ConstraintIndex) = False Then
                    If mFaithfulness(ConstraintIndex) Then
                        If Active(ConstraintIndex) Then
                            Let FaithIsActive = True
                        End If
                    Else
                        Let RankableMarked = True
                        Let Stratum(ConstraintIndex) = CurrentStratum
                        'BH:  report progress
                           Print #ShowMe, "  Markedness constraint "; mAbbrev(ConstraintIndex); " prefers no losers, joins stratum #"; Trim(Str(CurrentStratum)); "."
                    End If
                End If
                
            End If
        Next ConstraintIndex
        
        'BH:  report progress
            If RankableMarked = True Then
                Print #ShowMe, "  Faithfulness constraints are excluded from stratum."
            End If
                        
        'BH:  Report progress, inspect specificity if appropriate.
            If RankableMarked = False Then
                    
                    Print #ShowMe, "  No rankable markedness constraints are available for this stratum."
                    Print #ShowMe,
                
                Select Case FaithIsActive
                    Case True
                    
                        'BH:  Interpolation:  the version of BCD that employs the
                        '  the subset criterion.  This code assigns values to the
                        '  array Subsetted(), which designates constraints that can't
                        '  be ranked because of the presence of a more-specific
                        '  Faithfulness constraint that hasn't yet been ranked.
                        
                        If Form1.mnuSpecificBCD.Checked Then
                            Print #ShowMe, "Favor Specificity:"
                            Let FoundOne = False
                            For ConstraintIndex = 1 To mNumberOfConstraints
                                'Relevant only for yet-unranked faithfulness constraints.
                                If Stratum(ConstraintIndex) = 0 Then
                                    If mFaithfulness(ConstraintIndex) = True Then
                                        For InnerConstraintIndex = 1 To mNumberOfConstraints
                                            'The blockeffect is only from non-identical Faithfulness constraints
                                            '  that haven't been ranked yet.
                                                If InnerConstraintIndex <> ConstraintIndex Then
                                                    If Stratum(InnerConstraintIndex) = 0 Then
                                                        If mFaithfulness(InnerConstraintIndex) = True Then
                                                            If Subset(InnerConstraintIndex, ConstraintIndex) = True Then
                                                                Let Subsetted(ConstraintIndex) = True
                                                                'Report progress
                                                                Print #ShowMe, "  "; mAbbrev(ConstraintIndex); " cannot be installed in this stratum, because "; mAbbrev(InnerConstraintIndex); " is more specific."
                                                                Let FoundOne = True
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                        Next InnerConstraintIndex
                                    End If
                                End If
                            Next ConstraintIndex
                            If FoundOne = False Then Print #ShowMe, "  (no cases found)"
                            Print #ShowMe,
                        End If
                        'BH:  End of interpolation
                        
                        Print #ShowMe, "Avoid The Inactive:"
                        Print #ShowMe, "  Here are the active, as-yet-unranked Faithfulness constraints available for this stratum:"
                    
                    Case False
                        Print #ShowMe, "  No active Faithfulness constraints are available for this stratum."
                End Select
            End If

        'BT:  If no markedness constraints were rankable, then some faithfulness
        'constraints must be selected and ranked next.
        
        If RankableMarked = False Then
            If FaithIsActive = True Then
            
                'BT:  Select a minimal faith set and rank it next
                Call FindMinFaithSet(NumberOfForms, NumberOfRivals, _
                    WinnerViolations, RivalViolations, mAbbrev(), Stratum, _
                    CurrentStratum, Demotable, Active, StillInformative)
                    
            Else
            
                'BT:  Rank all available constraints; either this is the last
                'stratum, or inconsistency detection is imminent
                For ConstraintIndex = 1 To mNumberOfConstraints
                    If Stratum(ConstraintIndex) = 0 Then
                        If Demotable(ConstraintIndex) = False Then
                            Let Stratum(ConstraintIndex) = CurrentStratum
                            'BH:  report progress.
                                Print #ShowMe, "Faithfulness constraint "; mAbbrev(ConstraintIndex); " joins stratum #"; Trim(Str(CurrentStratum)); " by default."
                        End If
                    End If
                Next ConstraintIndex
                
             'BH: Report progress, first including the results of the last stratum:
                Call Form1.PrintResultsOfRankingSoFar(CurrentStratum, Stratum(), ShowMe, mFaithfulness())

            End If
        End If
        
        'BT:  Determine if there are any constraints which have not yet been ranked,
        'and if any were ranked on the most recent pass.
        
        UnrankedConstraintsRemain = False
        AConstraintWasRanked = False
        For ConstraintIndex = 1 To mNumberOfConstraints
            If Stratum(ConstraintIndex) = 0 Then
                UnrankedConstraintsRemain = True
            ElseIf Stratum(ConstraintIndex) = CurrentStratum Then
                AConstraintWasRanked = True
            End If
        Next ConstraintIndex
        
        'BT:  If all constraints have been ranked, learning is successful and complete.
        'If no constraints have been ranked, learning has failed, and there is
        'no point in continuing.
        'Otherwise, continue with another pass through the data.
        
        If UnrankedConstraintsRemain = False Then

            'BT:  All constraints have been ranked, indicating success.
            'Return the constraint hierarchy and exit the function.
            
            'BH:  Report progress:
                Print #ShowMe,
                Print #ShowMe, "Ranking is complete and yields successful grammar."
            
            'BH:  Encode the result as a string, to facilitate communication across modules.
                Let Main = "True"
                Let Main = Main + Chr(9) + Trim(Str(CurrentStratum))
                For ConstraintIndex = 1 To mNumberOfConstraints
                    Let Main = Main + Chr(9) + Trim(Str(Stratum(ConstraintIndex)))
                Next ConstraintIndex
                    
            
            'Let Main.Converged = True
            'Let Main.NumberOfStrata = CurrentStratum
            'For ConstraintIndex = 1 To mNumberofconstraints
            '    Let Main.Stratum(ConstraintIndex) = Stratum(ConstraintIndex)
            'Next ConstraintIndex
            
            Call BCDExitTasks               'BH
            Exit Function

        ElseIf AConstraintWasRanked = False Then
             
            'BT:  No constraints were ranked, so the data are inconsistent.
            'Indicate failure and exit the function.
            
             'BH:  report progress
                Print #ShowMe,
                Print #ShowMe, "Ranking has failed.  This constraint set cannot derive only winners."
            
            'BH:  Encode the result as a string, to facilitate communication across modules.
                Let Main = "False"
                Let Main = Main + Chr(9) + Trim(Str(CurrentStratum))
                For ConstraintIndex = 1 To mNumberOfConstraints
                    Let Main = Main + Chr(9) + Trim(Str(Stratum(ConstraintIndex)))
                Next ConstraintIndex

            'Let Main.Converged = False
            Call BCDExitTasks     'BH
            Exit Function

        End If

        'BT:  Find out which data should be ignored henceforth, because they are explained
        'by the constraints ranked already.
        'This occurs when the Rival candidate has greater violation of a constraint
        'that has just been ranked into the new stratum.

        For ConstraintIndex = 1 To mNumberOfConstraints
            If Stratum(ConstraintIndex) = CurrentStratum Then

                For FormIndex = 1 To NumberOfForms
                   For RivalIndex = 1 To NumberOfRivals(FormIndex)
                      If RivalViolations(FormIndex, RivalIndex, ConstraintIndex) > _
                         WinnerViolations(FormIndex, ConstraintIndex) Then
                         Let StillInformative(FormIndex, RivalIndex) = False
                      End If
                   Next RivalIndex
                Next FormIndex

            End If
        Next ConstraintIndex
        
        'BH:  Report progress:
            Call Form1.PrintResultsOfRankingSoFar(CurrentStratum, Stratum(), ShowMe, mFaithfulness())
            
    Loop  'BT:  Do loop repeatedly passing through the data; exit is loop-internal

    'You should never, ever get this far.
        MsgBox "Program error.   I would appreciate your letting me know the about the problem.  Email me at bhayes@humnet.ucla.edu, specifying error #47905, and including a copy of your input file.", vbCritical

End Function  'BT:  BCD

Private Sub FindMinFaithSet(ByVal NumberOfForms As Long, NumberOfRivals() As Long, _
    WinnerViolations() As Long, RivalViolations() As Long, _
    Abbrev() As String, Stratum() As Long, _
    CurrentStratum As Long, Demotable() As Boolean, Active() As Boolean, _
    StillInformative() As Boolean)

'BT:  This subroutine returns its result by modifying the parameter Stratum().
'Stratum() is passed in containing the constraints ranked so far. This subroutine
'determines a minimal faith constraint set, and places it in the next stratum
'of the ranking, that is, in Stratum(CurrentStratum).

    Dim ConstraintIndex As Long
    Dim FormIndex As Long
    Dim RivalIndex As Long
    
    Dim RankableFaith() As Long
    ReDim RankableFaith(mNumberOfConstraints)
    Dim RFCount As Long
    
    Dim SubsetSize As Long
    Dim SubsetArray() As Long
    ReDim SubsetArray(mNumberOfConstraints)
    Dim BestSubset() As Long
    ReDim BestSubset(mNumberOfConstraints)
    Dim BestMarkCount As Long   'BT:  Number of markedness cons. freed up by best subset
    
    'BT:  Determine the rankable faithfulness constraints (and how many there are).
    'NOTE: Rankable here means that the constraint has not yet been ranked
    'and is active (i.e., might do some work if ranked next).

    Let RFCount = 0
    For ConstraintIndex = 1 To mNumberOfConstraints
        If Stratum(ConstraintIndex) = 0 And mFaithfulness(ConstraintIndex) Then
            'BH:  interpolation:  specific BCD.  Embodied in the "subsetted" condition.
            If Form1.mnuSpecificBCD = True Then
                If Demotable(ConstraintIndex) = False And Active(ConstraintIndex) = True And Subsetted(ConstraintIndex) = False Then
                    Let RFCount = RFCount + 1
                    Let RankableFaith(RFCount) = ConstraintIndex
                    'BH:  Report progress.
                        Print #ShowMe, "  "; Abbrev(ConstraintIndex)
                End If
            Else
            'BH:  Classical BCD.
                If Demotable(ConstraintIndex) = False And Active(ConstraintIndex) = True Then
                    Let RFCount = RFCount + 1
                    Let RankableFaith(RFCount) = ConstraintIndex
                    'BH:  Report progress.
                        Print #ShowMe, "  "; Abbrev(ConstraintIndex)
                End If
            End If
        End If
    Next ConstraintIndex
    
    'BT:  Try faithfulness subsets, in increasing order of subset size, keeping track
    'of how many markedness constraints are released by the best subset. Stop when
    'a subset has been found releasing at least one markedness constraint, or when
    'all subsets have been tried.
    'BT:
    'BT:  EvaluateFaithSubsets() evaluates all subsets of a given size (the second parameter).
    'It takes responsibility for choosing when more than one subset of the same size
    'releases at least one markedness constraint.
    
    'BH:  report progress.
        Print #ShowMe,
        Print #ShowMe, "Smallest Effective Faithfulness Sets"
        Print #ShowMe, "  Evaluating subsets of increasing size:"
    
    BestMarkCount = 0
    SubsetSize = 0
    Do
        Let SubsetSize = SubsetSize + 1
        'BH:  report progress
            Print #ShowMe, "    Subset size = "; Trim(Str(SubsetSize))
        Call EvaluateFaithSubsets(RFCount, SubsetSize, 0, 0, SubsetArray, BestSubset, _
            BestMarkCount, RankableFaith, NumberOfForms, NumberOfRivals, WinnerViolations, _
            RivalViolations, Stratum, CurrentStratum, Demotable, Active, StillInformative)
    Loop While (SubsetSize < RFCount) And (BestMarkCount = 0)
    
    'BT:  If no subset released any markedness constraints, then rank all active faithfulness
    'constraints next. Either there is an inconsistency, or there are some markdata pairs
    'which do not implicate any markedness constraints in any way.
    
    If BestMarkCount = 0 Then
        For ConstraintIndex = 1 To RFCount
            Let BestSubset(ConstraintIndex) = RankableFaith(ConstraintIndex)
            'BH:  report progress
                Print #ShowMe, "  No subset released any markedness constraints."
        Next ConstraintIndex
    End If
    
    'BT:  Rank the best subset.
    
    'BH:  report progress.
    Print #ShowMe,
    For ConstraintIndex = 1 To SubsetSize
        Let Stratum(BestSubset(ConstraintIndex)) = CurrentStratum
        Print #ShowMe, "  Faithfulness constraint "; Abbrev(BestSubset(ConstraintIndex)); " joins stratum #"; Trim(Str(CurrentStratum)); " as member of best subset."
        If mTieFlag = True Then
            Print #ShowMe, "  Note:  This is an arbitrary choice, arising from the tie noted above."
        End If
        
    Next ConstraintIndex
    
End Sub       'BT:  FindMinFaithSet

'BT:  Time for fun with recursive functions!
'
'This subroutine recursively constructs all subsets of size SubsetSize of the set
'RankableFaith(), which is an array of the indices of rankable faithfulness constraints.
'The first parameter, SetSize, indicates the number of elements in RankableFaith().
'Each such subset is evaluated by the function CheckMarkednessRelease(), which returns
'the number of markedness constraints released by ranking that faith subset next.
'
'Any given working subset is stored in SubsetArray(), which contains the index of
'each constraint in the subset: the index of first constraint in set is SubsetArray(1), etc.
'The parameter IndexCount indicates how many constraints are currently in the subset
'under construction; when IndexCount reaches SubsetSize, a full subset has been
'constructed.

Private Sub EvaluateFaithSubsets(SetSize As Long, SubsetSize As Long, IndexCount As Long, _
    IndexValue As Long, SubsetArray() As Long, BestSubset() As Long, _
    BestMarkCount As Long, RankableFaith() As Long, ByVal NumberOfForms As Long, _
    NumberOfRivals() As Long, WinnerViolations() As Long, RivalViolations() As Long, _
    Stratum() As Long, CurrentStratum As Long, Demotable() As Boolean, Active() As Boolean, _
    StillInformative() As Boolean)
    
    Dim NewIndexValue As Long
    Dim SubsetArrayIndex As Long
    Dim MarkCount As Long
    
    If IndexCount = SubsetSize Then
        
        'BT:  A complete subset is determined; see how many markedness constraints are released.
        Let MarkCount = CheckMarkednessRelease(SubsetSize, SubsetArray, NumberOfForms, _
            NumberOfRivals, WinnerViolations, RivalViolations, Stratum, CurrentStratum, _
            Demotable, Active, StillInformative)
        
        'BT:  If more markedness constraints are released than were released by the previous
        'best, then make the current subset of faith the new best.
        
        'BH:  One wishes to warn the user if the algorithm has arrived at a tie.
        '     Detect only authentic ties; i.e. 0-0 doesn't count:
            If MarkCount = BestMarkCount And BestMarkCount > 0 Then
                Let mTieFlag = True
                Print #ShowMe, "      This subset is in a tie with the previous best, which is retained."
            End If
        
        If MarkCount > BestMarkCount Then
            
            'BH:  Tie evaporates.  Report progress.
                Let mTieFlag = False
                Print #ShowMe, "      This subset is an improvement."
            
            
            Let BestMarkCount = MarkCount
            For SubsetArrayIndex = 1 To SubsetSize
                Let BestSubset(SubsetArrayIndex) = SubsetArray(SubsetArrayIndex)
            Next SubsetArrayIndex
        End If
        
    Else    'BT:  More constraints need to be added to the subset.
        
        'BT:  Continue construction of the subset by considering various possible
        'constraints as the next one in the subset.
        For NewIndexValue = (IndexValue + 1) To (SetSize - SubsetSize + IndexCount + 1)
            Let SubsetArray(IndexCount + 1) = RankableFaith(NewIndexValue)
            Call EvaluateFaithSubsets(SetSize, SubsetSize, IndexCount + 1, NewIndexValue, _
                SubsetArray, BestSubset, BestMarkCount, RankableFaith, NumberOfForms, _
                NumberOfRivals, WinnerViolations, RivalViolations, Stratum, _
                CurrentStratum, Demotable, Active, StillInformative)
        Next NewIndexValue
    
    End If

End Sub

Private Function CheckMarkednessRelease(SubsetSize As Long, SubsetArray() As Long, _
    ByVal NumberOfForms As Long, NumberOfRivals() As Long, WinnerViolations() As Long, _
    RivalViolations() As Long, ParStratum() As Long, ParCurrentStratum As Long, _
    ParDemotable() As Boolean, ParActive() As Boolean, ParStillInformative() As Boolean) As Long

    Dim Stratum() As Long
    ReDim Stratum(mNumberOfConstraints)
    Dim CurrentStratum As Long
    Dim Demotable() As Boolean
    ReDim Demotable(mNumberOfConstraints)
    Dim Active() As Boolean
    ReDim Active(mNumberOfConstraints)
    Dim StillInformative() As Boolean
    ReDim StillInformative(NumberOfForms, mMaximumNumberOfRivals)
    
    Dim ConstraintIndex As Long
    Dim FormIndex As Long
    Dim RivalIndex As Long
    Dim SubsetIndex As Long
    Dim RankableMarkedCount As Long
    Dim RankableMarkedCountTotal As Long
    
    'BT:  Make copies of all the relevant parameters, so that only local copies are modified
    
    Let CurrentStratum = ParCurrentStratum
    For ConstraintIndex = 1 To mNumberOfConstraints
        Let Stratum(ConstraintIndex) = ParStratum(ConstraintIndex)
        Let Demotable(ConstraintIndex) = ParDemotable(ConstraintIndex)
        Let Active(ConstraintIndex) = ParActive(ConstraintIndex)
    Next ConstraintIndex
    For FormIndex = 1 To NumberOfForms
        For RivalIndex = 1 To NumberOfRivals(FormIndex)
            Let StillInformative(FormIndex, RivalIndex) = _
                ParStillInformative(FormIndex, RivalIndex)
        Next RivalIndex
    Next FormIndex
    
    'BT:  Set the faithfulness constraints in the subset to be in the current stratum
    
    For SubsetIndex = 1 To SubsetSize
        Let Stratum(SubsetArray(SubsetIndex)) = CurrentStratum
    Next SubsetIndex
    
    'BH:  report progress
        If SubsetSize = 1 Then
            Print #ShowMe, "    Subset under consideration:  "; mAbbrev(SubsetArray(1))
        Else
            Print #ShowMe, "    Subset under consideration:"
            For SubsetIndex = 1 To SubsetSize
                Print #ShowMe, "      "; mAbbrev(SubsetArray(SubsetIndex))
            Next SubsetIndex
        End If
    
    'BT:  MAIN LOOP: keep ranking markedness constraints until no more can be ranked
    
    Let RankableMarkedCountTotal = 0
    Do
    
        'BT:  Remove any winner/rival pairs that are no longer informative
    
        For ConstraintIndex = 1 To mNumberOfConstraints
            If Stratum(ConstraintIndex) = CurrentStratum Then

                For FormIndex = 1 To NumberOfForms
                    For RivalIndex = 1 To NumberOfRivals(FormIndex)
                        If RivalViolations(FormIndex, RivalIndex, ConstraintIndex) > _
                            WinnerViolations(FormIndex, ConstraintIndex) Then
                            Let StillInformative(FormIndex, RivalIndex) = False
                        End If
                    Next RivalIndex
                Next FormIndex

            End If
        Next ConstraintIndex
    
        'BT:  Record what stratum you're constructing:
        Let CurrentStratum = CurrentStratum + 1

        'BT:  Initialize variables indicating which constraints must be demoted and
        'which constraints are active (prefer the winner in at least one pair).
        
        For ConstraintIndex = 1 To mNumberOfConstraints
            Let Demotable(ConstraintIndex) = False
            Let Active(ConstraintIndex) = False
        Next ConstraintIndex

        'BT:  Go through all pairs of Winner vs. Rival, marking constraints preferring
        'a rival as demotable (MUST be demoted) and constraints preferring a winner
        'as active.
          
        For FormIndex = 1 To NumberOfForms
            For RivalIndex = 1 To NumberOfRivals(FormIndex)

                'Only still-informative Rivals can be learned from:
                If StillInformative(FormIndex, RivalIndex) = True Then
                    For ConstraintIndex = 1 To mNumberOfConstraints

                        'BT:  For each yet-unranked constraint
                        If Stratum(ConstraintIndex) = 0 Then
                        
                            'BT:  If it prefers the rival, mark as demotable
                            If WinnerViolations(FormIndex, ConstraintIndex) > _
                                RivalViolations(FormIndex, RivalIndex, ConstraintIndex) Then
                                Let Demotable(ConstraintIndex) = True
                                
                            'BT:  If it prefers the winner, mark as active
                            ElseIf WinnerViolations(FormIndex, ConstraintIndex) < _
                                RivalViolations(FormIndex, RivalIndex, ConstraintIndex) Then
                                Let Active(ConstraintIndex) = True
                            End If
                            
                        End If

                    Next ConstraintIndex
                End If  'BT:  StillInformative
                
            Next RivalIndex
        Next FormIndex

        'BT:  See if any markedness constraints can be ranked next; if so, rank them.

        Let RankableMarkedCount = 0
        
        For ConstraintIndex = 1 To mNumberOfConstraints
            If Stratum(ConstraintIndex) = 0 Then
                If Demotable(ConstraintIndex) = False Then
                    If mFaithfulness(ConstraintIndex) Then
                    Else
                        Let RankableMarkedCount = RankableMarkedCount + 1
                        'BH:  report progress
                            Print #ShowMe, "        Markedness constraint freed up:  "; mAbbrev(ConstraintIndex)
                        Let Stratum(ConstraintIndex) = CurrentStratum
                    End If
                End If
            End If
        Next ConstraintIndex
        
        Let RankableMarkedCountTotal = RankableMarkedCountTotal + RankableMarkedCount
    
    Loop While RankableMarkedCount > 0

    'BT:  Return the number of markedness constraints
    
    Let CheckMarkednessRelease = RankableMarkedCountTotal
    
    'BH:  report progress
        Print #ShowMe, "      Number of markedness constraints freed up:  "; Trim(Str(RankableMarkedCountTotal))
            
End Function


Private Sub BCDExitTasks()

    'BH:  Two chores to do when BCD terminates.
        'BH:  Close the file that monitors progress.
            Close #ShowMe
        'BH:  Warn the user if an arbitrary decision was made with tied subsets.
            If mUpperTieFlag = True Then
                MsgBox "Caution:  The BCD algorithm has selected arbitrarily among tied Faithfulness constraint subsets." + Chr(10) + Chr(10) _
                + "You may wish to try changing the order of the Faithfulness constraints in the input file, to see whether this results in a different ranking." + Chr(10) + Chr(10) _
                + "For diagnosis, click OK, then select Show How Ranking Was Done from the View menu.", vbExclamation
            End If
        
End Sub


Sub LocateViolationSubsets(NumberOfForms As Long, WinnerViolations() As Long, NumberOfRivals() As Long, RivalViolations() As Long)

   'We need an array Subset() that says when one constraint's violations
   '  are a subset of another's.  This is common enough with, say,
   '  contextualized vs. general Faithfulness constraints

    Dim OuterConstraintIndex As Long
    Dim InnerConstraintIndex As Long
    Dim FormIndex As Long
    Dim RivalIndex As Long

    ReDim Subset(mNumberOfConstraints, mNumberOfConstraints)
   
   For OuterConstraintIndex = 1 To mNumberOfConstraints
       For InnerConstraintIndex = 1 To mNumberOfConstraints

         'We're not interested in self-comparison:
         If InnerConstraintIndex <> OuterConstraintIndex Then
            For FormIndex = 1 To NumberOfForms
               'Compare winner violations:
                    If WinnerViolations(FormIndex, OuterConstraintIndex) > WinnerViolations(FormIndex, InnerConstraintIndex) Then
                       Let Subset(OuterConstraintIndex, InnerConstraintIndex) = False
                       GoTo NotSubsetExitPoint
                    End If
               'Compare rival violations:
                    For RivalIndex = 1 To NumberOfRivals(FormIndex)
                       If RivalViolations(FormIndex, RivalIndex, OuterConstraintIndex) > RivalViolations(FormIndex, RivalIndex, InnerConstraintIndex) Then
                          Let Subset(OuterConstraintIndex, InnerConstraintIndex) = False
                          GoTo NotSubsetExitPoint
                       End If
                    Next RivalIndex
            Next FormIndex

            'If you've navigated this far, then the inner constraint is
            '  in a subset relation to the outer.  Record this.

            Let Subset(OuterConstraintIndex, InnerConstraintIndex) = True
         
            'Override this result if the inner constraint has no violations at all.
                For FormIndex = 1 To NumberOfForms
                    If WinnerViolations(FormIndex, OuterConstraintIndex) > 0 Then GoTo NullityExitPoint
                    For RivalIndex = 1 To NumberOfRivals(FormIndex)
                        If RivalViolations(FormIndex, RivalIndex, OuterConstraintIndex) > 0 Then GoTo NullityExitPoint
                    Next RivalIndex
                Next FormIndex
                'Override, you have a constraint with no violations:
                    Let Subset(OuterConstraintIndex, InnerConstraintIndex) = False

NullityExitPoint:
         
         Else
            Let Subset(OuterConstraintIndex, InnerConstraintIndex) = False  'no subset for identical
         End If

NotSubsetExitPoint:

      Next InnerConstraintIndex
   Next OuterConstraintIndex
  
   
End Sub


