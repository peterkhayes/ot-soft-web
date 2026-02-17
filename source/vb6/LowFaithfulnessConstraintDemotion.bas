Attribute VB_Name = "LowFaithfulnessConstraintDemotion"
'=======================LOW FAITHFULNESS CONSTRAINT DEMOTION=========================

Option Explicit
    
    'Module level variables
        Dim ShowMe As Long
        Dim Subset() As Boolean
        Dim mNumberOfConstraints As Long
        Dim mMaximumNumberOfRivals As Long
        Dim mFaithfulness() As Boolean

Public Function Main(ByVal NumberOfForms As Long, InputForm() As String, NumberOfRivals() As Long, _
   WinnerViolations() As Long, Rival() As String, RivalViolations() As Long, NumberOfConstraints As Long, _
   ConstraintName() As String, Abbrev() As String) As String
   
    'Hayes's Low Faithfulness Constraint Demotion.
    
    'For convenience, the output is a string; see the routine in Form1 that calls this one for rationale.
    
    'Move values of parameters into module level variables.
        Let mNumberOfConstraints = NumberOfConstraints
            
        Dim Stratum() As Long
        ReDim Stratum(mNumberOfConstraints)
      
        Dim CurrentStratum As Long
        
        Dim PutFaithfulnessLowFlag As Boolean

        Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
        Dim OuterConstraintIndex As Long, InnerConstraintIndex As Long
        Dim LocalStratumIndex As Long
      
        Dim SomeAreNonDemotible As Boolean
        Dim AllAreDemotable As Boolean
        Dim SomeAreDemotible As Boolean
        
        Dim StillInformative() As Boolean
        Let mMaximumNumberOfRivals = Form1.FindMaximumNumberOfRivals(NumberOfRivals())
        ReDim StillInformative(NumberOfForms, mMaximumNumberOfRivals)
        Dim Demotable() As Boolean              'A constraint can be demotable for various reasons.
        ReDim Demotable(mNumberOfConstraints)
        Dim FavorsLoser() As Boolean            'But favoring a loser is the worst, and is the only reason
        ReDim FavorsLoser(mNumberOfConstraints)  '   that leads to the algorithm crashing.
        Dim Subsetted() As Boolean              'Excluded from current stratum by the existence of
        ReDim Subsetted(mNumberOfConstraints)    '   a more specific version of the same Faithfulness constraint.
        Dim Active() As Boolean                 'Borrowed from Bruce Tesar:
        ReDim Active(mNumberOfConstraints)       '   mark which constraints favor a winner.
        Dim NumberOfHelpers() As Long
        ReDim NumberOfHelpers(mNumberOfConstraints)
        Dim LocalNumberOfHelpers As Long
        Dim LowestNumberOfHelpers As Long
        ReDim mFaithfulness(mNumberOfConstraints)
      
    'Read off the a priori rankings if the user so requests.
        If Form1.mnuConstrainAlgorithmsByAPrioriRankings.Checked = True Then
            'ReadAPrioriRankingsAsTable is a boolean function, which returns False if
            '   it failed to do its job.
            If ReadAPrioriRankingsAsTable(mNumberOfConstraints, Abbrev()) = False Then
                Let Form1.mnuConstrainAlgorithmsByAPrioriRankings.Checked = False
            End If
        End If
   
   'Initialize crucial variables, in case this routine gets called more
   '  than once.

      Let CurrentStratum = 0

      For ConstraintIndex = 1 To mNumberOfConstraints
         'Zero value for Stratum() means:  not yet ranked.
          Let Stratum(ConstraintIndex) = 0
          Let mFaithfulness(ConstraintIndex) = False
      Next ConstraintIndex

      For FormIndex = 1 To NumberOfForms
         For RivalIndex = 1 To mMaximumNumberOfRivals
            Let StillInformative(FormIndex, RivalIndex) = True
         Next RivalIndex
      Next FormIndex
      
    'Initializations:
        Let PutFaithfulnessLowFlag = True
        Dim FaithfulnessAllowed As Boolean
        Call LocateViolationSubsets(NumberOfForms, WinnerViolations(), NumberOfRivals(), RivalViolations())
      
    'Admit that there is no a priori ranking capacity in this routine.
    '    If mnuConstrainAlgorithmsByAPrioriRankings.Checked = True Then
    '        Select Case MsgBox("Sorry, you've selected the use of a priori rankings, but OTSoft doesn't implement a priori rankings for Low Faithfulness Constraint Demotion.  Please contact me (bhayes@humnet.ucla.edu) if this capacity is important to you." + _
    '            chr(10) + chr(10) + _
    '            "Click Yes to continue without the use of a priori rankings, No to exit OTSoft.", vbYesNo)
    '            Case vbYes
    '                'do nothing
    '            Case vbNo
    '                End
    '        End Select
    '    End If
    
    'First, make sure there is a folder for these files, a daughter of the
    '   folder in which the input file is located.
        Call Form1.CreateAFolderForOutputFiles
    
    'Produce a little file to show how you did it:
        Dim FoundOne As Boolean         'Used in self-report file to indicate searches that found nothing.
        Let ShowMe = FreeFile
        Open gOutputFilePath + "HowIRanked" + gFileName + ".txt" For Output As #ShowMe
        Print #ShowMe, "******Application of Low Faithfulness Constraint Demotion******"
        Print #ShowMe,
        Print #ShowMe, "Input file:  "; gInputFilePath + gFileName + gFileSuffix
        Print #ShowMe,
   
        For ConstraintIndex = 1 To mNumberOfConstraints
            Let mFaithfulness(ConstraintIndex) = Form1.FaithfulnessConstraint(ConstraintName(ConstraintIndex))
        Next ConstraintIndex
      
   'Go through the Winner-Rival pairs repeatedly, looking for constraints that
   '  never prefer losers among the candidates still being considered.

   Do
   
      'Record what stratum you're constructing:
            Let CurrentStratum = CurrentStratum + 1
            Print #ShowMe,
            Print #ShowMe, "******Now doing Stratum #"; Trim(Str(CurrentStratum)); "******"
            Print #ShowMe,
      
      'Initialize variables indicating demotability, activity, subset blocking,
      '  and the NumberOfHelpers() array.
          For ConstraintIndex = 1 To mNumberOfConstraints
             Let Demotable(ConstraintIndex) = False
             Let FavorsLoser(ConstraintIndex) = False
             Let Subsetted(ConstraintIndex) = False
             Let Active(ConstraintIndex) = False
             Let NumberOfHelpers(ConstraintIndex) = 10000
          Next ConstraintIndex
          
      'AVOID PREFERENCE FOR LOSERS
      
      'Go through all pairs of Winner vs. Rival, eliminating constraints that favor a loser.
        
        Print #ShowMe, "Avoid Preference For Losers:"
        Let FoundOne = False
        For FormIndex = 1 To NumberOfForms
            For RivalIndex = 1 To NumberOfRivals(FormIndex)
            'Only still-informative Rivals can be learned from:
            If StillInformative(FormIndex, RivalIndex) = True Then
                'Examine constraints to see if they prefer losers.
                For ConstraintIndex = 1 To mNumberOfConstraints
                    'We only consider constraints that haven't been ranked yet.
                    If Stratum(ConstraintIndex) = 0 Then
                    'Keep a constraint out of the current stratum if it prefers a loser.
                    'For standard Constraint Demotion, this is the only criterion.
                        If WinnerViolations(FormIndex, ConstraintIndex) > RivalViolations(FormIndex, RivalIndex, ConstraintIndex) Then
                            'Report this if it's the first time:
                                If Demotable(ConstraintIndex) = False Then
                                    Print #ShowMe, "  "; Abbrev(ConstraintIndex); " is excluded from stratum; prefers loser *["; DumbSym(Rival(FormIndex, RivalIndex)); "] for /"; DumbSym(InputForm(FormIndex)); "/."
                                    Let FoundOne = True
                                End If
                            Let Demotable(ConstraintIndex) = True
                            Let FavorsLoser(ConstraintIndex) = True
                        End If
                 End If
              Next ConstraintIndex
           End If
        Next RivalIndex
        Next FormIndex
        If FoundOne = False Then Print #ShowMe, "  Search found no unranked constraints that prefer losers."
        Print #ShowMe,
        
        
        'A PRIORI RANKINGS
        
        'If APrioriRankings are on, then demote the victims.
        '   We demote those which are a priori ranked below the yet-unranked.
            If Form1.mnuConstrainAlgorithmsByAPrioriRankings.Checked = True Then
                Print #ShowMe, "Enforce a priori rankings:"
                Let FoundOne = False
                For OuterConstraintIndex = 1 To mNumberOfConstraints
                    If Stratum(OuterConstraintIndex) = 0 Then
                        For InnerConstraintIndex = 1 To mNumberOfConstraints
                            If gAPrioriRankingsTable(OuterConstraintIndex, InnerConstraintIndex) = True Then
                                'Demote, and report this, if first time.
                                If Demotable(InnerConstraintIndex) = False Then
                                    Print #ShowMe, "  "; Abbrev(InnerConstraintIndex); " is excluded from stratum."
                                    Print #ShowMe, "    It is dominated a priori by "; Abbrev(OuterConstraintIndex); ", which has yet to be ranked."
                                    Let FoundOne = True
                                End If
                                Let Demotable(InnerConstraintIndex) = True
                            End If
                        Next InnerConstraintIndex
                    End If                  'Is the dominee yet unranked?
                Next OuterConstraintIndex
                If FoundOne = False Then
                    Print #ShowMe, "  Search found no constraints that must be demoted due to an a priori ranking."
                End If
                Print #ShowMe,
                
                'Are all constraints now demotable?  I.e. there could be a contradiction
                '   that is the combined result of loser-favoring and a priori ranking.
                    Let AllAreDemotable = True
                    For ConstraintIndex = 1 To mNumberOfConstraints
                        If Stratum(ConstraintIndex) = 0 And Demotable(ConstraintIndex) = False Then
                            Let AllAreDemotable = False
                            Exit For
                        End If
                    Next ConstraintIndex
                    'If nobody can join the new stratum, then skip all else.
                        If AllAreDemotable = True Then
                            Let SomeAreDemotible = True
                            Let SomeAreNonDemotible = False
                            GoTo SkippationPoint
                        End If
            
            End If  'Are we doing a priori rankings?
        
        'FAVOR MARKEDNESS
          
          'Is there a rankable Markedness constraint?
            Print #ShowMe, "Favor Markedness:"
            Dim ThereIsARankableMarkednessConstraint As Boolean
            Let ThereIsARankableMarkednessConstraint = False
            For ConstraintIndex = 1 To mNumberOfConstraints
                If Stratum(ConstraintIndex) = 0 And Demotable(ConstraintIndex) = False And mFaithfulness(ConstraintIndex) = False Then
                    Let ThereIsARankableMarkednessConstraint = True
                    'Report:
                        Print #ShowMe, "  "; Abbrev(ConstraintIndex); " is a Markedness constraint that favors no losers, joins new stratum."
                End If
            Next ConstraintIndex
            
            If ThereIsARankableMarkednessConstraint = True Then
                'Shut out the Faithfulness constraints
                    For InnerConstraintIndex = 1 To mNumberOfConstraints
                        If mFaithfulness(InnerConstraintIndex) = True Then
                            Let Demotable(InnerConstraintIndex) = True
                        End If
                    Next InnerConstraintIndex
                'Report progress:
                    Print #ShowMe, "  Faithfulness constraints are excluded from stratum."
                'Exit rest of the constraint-assessment code:
                    GoTo ReportStrata
            Else
                Print #ShowMe, "  There are no rankable Markedness constraints."
            
            End If
          
          'FAVOR ACTIVENESS
            
           'Find which Faithfulness constraints are active.
                
                'Report progress:
                    Print #ShowMe,
                    Print #ShowMe, "Favor Activeness:"
                
                'Set flags:
                    Dim AtLeastOneFaithfulnessConstraintIsActive As Boolean
                    Let AtLeastOneFaithfulnessConstraintIsActive = False
                    Dim AllFaithfulnessConstraintsAreActive As Boolean
                    Let AllFaithfulnessConstraintsAreActive = True
                
                For ConstraintIndex = 1 To mNumberOfConstraints
                    If mFaithfulness(ConstraintIndex) = True And Stratum(ConstraintIndex) = 0 And Demotable(ConstraintIndex) = False Then
                        For FormIndex = 1 To NumberOfForms
                           'Does the constraint favor the winner for at least one comparison?
                            For RivalIndex = 1 To NumberOfRivals(FormIndex)
                                If StillInformative(FormIndex, RivalIndex) Then
                                '  Rivals that have a violation superset of the winner will
                                '  be derived correctly no matter how what the ranking.
                                '  So don't let them designate a constraint as effective:
                                    If Superset(FormIndex, RivalIndex, RivalViolations(), WinnerViolations()) = False Then
                                        'Now you can assess effectiveness.
                                        If RivalViolations(FormIndex, RivalIndex, ConstraintIndex) > WinnerViolations(FormIndex, ConstraintIndex) Then
                                            Let Active(ConstraintIndex) = True
                                            Let AtLeastOneFaithfulnessConstraintIsActive = True
                                            Print #ShowMe, "  "; Abbrev(ConstraintIndex); " is shown to be active by ruling out *["; DumbSym(Rival(FormIndex, RivalIndex)); "] for /"; DumbSym(InputForm(FormIndex)); "/."
                                            GoTo ActivenessExitPoint  'You only need one for proof.
                                        End If
                                    End If      'not superset
                                End If          'still informative
                            Next RivalIndex
                        Next FormIndex
                    End If
ActivenessExitPoint:                            'Move on to the next constraint.
                Next ConstraintIndex
                If AtLeastOneFaithfulnessConstraintIsActive = True Then
                    'Exclude the inactive Faithfulness constraints, because at least one active
                    '  one is available:
                        For ConstraintIndex = 1 To mNumberOfConstraints
                            If mFaithfulness(ConstraintIndex) = True Then
                                If Active(ConstraintIndex) = False Then
                                    Let Demotable(ConstraintIndex) = True
                                    Let AllFaithfulnessConstraintsAreActive = False
                                    'Report progress:
                                        Print #ShowMe, "  "; Abbrev(ConstraintIndex); " is excluded from stratum because it is inactive."
                                End If
                            End If
                        Next ConstraintIndex
                        'Report progress:
                            If AllFaithfulnessConstraintsAreActive = True Then
                                Print #ShowMe, "  All unranked Faithfulness constraints are active."
                            End If
                Else
                    'All remaining constraints that favor no losers are inactive.  Put them
                    '  all in the stratum (which has better be the last stratum if the
                    '  grammar is going to work).
                    'Caution:  if you're using this algorithm with alternation data (why not), it's
                    '   possible that there will be an inactive Faithfulness constraint that also
                    '   favors a loser.  It must not be allowed to join the new stratum; instead,
                    '   it must remain Demotible, to produce a crash (rather than a wrong grammar)
                    '   if necessary.
                    
                        'Report progress:
                            Print #ShowMe, "  Only remaining rankable constraints are inactive Faithfulness constraints.  "
                            Print #ShowMe, "  All of them join the current stratum:"
                        For ConstraintIndex = 1 To mNumberOfConstraints
                            If Stratum(ConstraintIndex) = 0 And _
                                mFaithfulness(ConstraintIndex) = True And _
                                FavorsLoser(ConstraintIndex) = False And _
                                Demotable(ConstraintIndex) = False Then     'This is if it's demoted by an apriori ranking.
                                'Let Stratum(ConstraintIndex) = CurrentStratum
                                'To permit termination, we want to designate
                                '  these constraints as not demotable.
                                    Let Demotable(ConstraintIndex) = False
                                'Report progress:
                                    Print #ShowMe, "    "; Abbrev(ConstraintIndex)
                            End If
                        Next ConstraintIndex
                        'Skip the rest:
                        GoTo ReportStrata
                End If
      
          'FAVOR SPECIFICITY
          
          'Shut out the superset Faithfulness constraints.
                Print #ShowMe,
                Print #ShowMe, "Favor Specificity:"
                Let FoundOne = False
                For ConstraintIndex = 1 To mNumberOfConstraints
                    'Relevant only for yet-unranked, active faithfulness constraints.
                    If Stratum(ConstraintIndex) = 0 And Demotable(ConstraintIndex) = False And Active(ConstraintIndex) = True Then
                        If mFaithfulness(ConstraintIndex) = True Then
                            For InnerConstraintIndex = 1 To mNumberOfConstraints
                                'The blocking effect is only from non-identical Faithfulness constraints
                                '  that haven't been ranked yet.
                                    If InnerConstraintIndex <> ConstraintIndex Then
                                        If Stratum(InnerConstraintIndex) = 0 Then
                                            If mFaithfulness(InnerConstraintIndex) = True And Demotable(InnerConstraintIndex) = False Then
                                                If Subset(InnerConstraintIndex, ConstraintIndex) = True Then
                                                    Let Demotable(ConstraintIndex) = True
                                                    Print #ShowMe, "  "; Abbrev(ConstraintIndex); " is excluded from stratum because "; Abbrev(InnerConstraintIndex); " is more specific."
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
      
          'FAVOR AUTONOMY
          
          'We seek the most autonomous constraints still in contention--those
          '  which rule out a loser with the fewest helpers.
          
          Print #ShowMe, "Favor Autonomy:"
          For FormIndex = 1 To NumberOfForms
             For RivalIndex = 1 To NumberOfRivals(FormIndex)
               'Only still-informative Rivals can be learned from:
                If StillInformative(FormIndex, RivalIndex) = True Then
                   For ConstraintIndex = 1 To mNumberOfConstraints
                       'We only consider constraints that haven't been ranked yet.
                        If Stratum(ConstraintIndex) = 0 Then
                            'Don't bother if the constraint is already demotable.
                            
                            If Demotable(ConstraintIndex) = False Then
                           
                              'Start the count of helpers for this constraint/rival combo at zero.
                                  Let LocalNumberOfHelpers = 0
                              
                              'First detect if the constraint prefers the winner over the rival.
                              '  Only such constraints can receive a new (perhaps lower)
                              '  NumberOfHelpers value.
                                
                                If RivalViolations(FormIndex, RivalIndex, ConstraintIndex) > WinnerViolations(FormIndex, ConstraintIndex) Then
                                
                                   'Second, proceed further only if it is a Faithfulness constraint.
                                   '  Markedness constraints don't care about helpers.
                                    If mFaithfulness(ConstraintIndex) = True Then
                                       
                                        'Third, detect if this is a case where ranking doesn't matter: i.e.
                                        '  the rivals violations are a superset of the winners.
                                        
                                        '  If there is a superset, then it doesn't matter how the constraints
                                        '  get ranked.  Therefore, such cases provide no useful information
                                        '  about NumberOfHelpers.
                                                                            
                                        If Superset(FormIndex, RivalIndex, RivalViolations(), WinnerViolations()) = False Then
                                        
                                            'Now, look for helpers:
                                            For InnerConstraintIndex = 1 To mNumberOfConstraints
                                               If InnerConstraintIndex <> ConstraintIndex Then
                                                    If RivalViolations(FormIndex, RivalIndex, InnerConstraintIndex) > WinnerViolations(FormIndex, InnerConstraintIndex) Then
                                                        'Superset faithfulness constraints do not count here:
                                                            If Subset(ConstraintIndex, InnerConstraintIndex) = True And mFaithfulness(InnerConstraintIndex) = True Then
                                                                'do nothing
                                                            Else
                                                                'This is an authentic helper.  Assess a penalty
                                                                    Let LocalNumberOfHelpers = LocalNumberOfHelpers + 1
                                                            End If
                                                    End If      'Are we comparing with a constraint that prefers the winner?
                                                End If          'Are we comparing with a distinct constraint?
                                            Next InnerConstraintIndex  'Look everywhere for helpers.
                                            
                                            'You now know how may helpers this constraint has with respect to this
                                            '  particular form.  If it is the best so far, record it.
                                                If LocalNumberOfHelpers < NumberOfHelpers(ConstraintIndex) Then
                                                    Let NumberOfHelpers(ConstraintIndex) = LocalNumberOfHelpers
                                                    'Show progress.
                                                        Print #ShowMe, "  "; Abbrev(ConstraintIndex); " is assigned "; Trim(Str(LocalNumberOfHelpers)); " helper";
                                                        Select Case LocalNumberOfHelpers
                                                            Case 0
                                                                Print #ShowMe, "s, based on /"; DumbSym(InputForm(FormIndex)); "/ -/-> *["; DumbSym(Rival(FormIndex, RivalIndex)); "]."
                                                            Case 1
                                                                Print #ShowMe, ", based on /"; DumbSym(InputForm(FormIndex)); "/ -/-> *["; DumbSym(Rival(FormIndex, RivalIndex)); "].  The helper is:"
                                                            Case Else
                                                                Print #ShowMe, "s, based on /"; DumbSym(InputForm(FormIndex)); "/ -/-> *["; DumbSym(Rival(FormIndex, RivalIndex)); "].  The helpers are:"
                                                        End Select
                                                    'List the helpers.  Re-find them, with a more compact version of the code above.
                                                        For InnerConstraintIndex = 1 To mNumberOfConstraints
                                                            If InnerConstraintIndex <> ConstraintIndex And RivalViolations(FormIndex, RivalIndex, InnerConstraintIndex) > WinnerViolations(FormIndex, InnerConstraintIndex) Then
                                                                If Not (Subset(ConstraintIndex, InnerConstraintIndex) = True And mFaithfulness(InnerConstraintIndex) = True) Then
                                                                    Print #ShowMe, "    "; Abbrev(InnerConstraintIndex)
                                                                End If
                                                            End If
                                                        Next InnerConstraintIndex
                                                End If
                                        
                                        End If  'Is this pointless because the rival has a superset
                                                'of the winner's violations?
                                    End If      'Count helpers only for Faithfulness constraints
                                End If          'Count helpers only for constraints that disprefer a loser
                            End If              'Count helpers only for constraints not already demotable.
                        End If                    'Is this constraint in the yet-unranked pool?
                   Next ConstraintIndex         'Go through all the constraints.
                End If                          'Is this datum still informative?
             Next RivalIndex                    'Go through all the rival candidates.
          Next FormIndex                        'Go through all the input forms.

        'Now we know how many helpers the Faithfulness constraints have.
        '   We can determine the criterion that distinguishes the best;
        '   i.e. the overall lowest number of helpers.
        'If no constraint got inspected for number of helpers, then all the constraints
        '  will enter with a value of 10000, and LowestNumberOfHelpers will stay at 9999.
                  
            If PutFaithfulnessLowFlag = True Then
                Let LowestNumberOfHelpers = 9999
                For ConstraintIndex = 1 To mNumberOfConstraints
                    If mFaithfulness(ConstraintIndex) = True Then
                        If Demotable(ConstraintIndex) = False Then
                            If NumberOfHelpers(ConstraintIndex) < LowestNumberOfHelpers Then
                                Let LowestNumberOfHelpers = NumberOfHelpers(ConstraintIndex)
                            End If
                        End If
                    End If
                Next ConstraintIndex
            'Report what you've done:
                Print #ShowMe,
                If LowestNumberOfHelpers < 9999 Then
                    Print #ShowMe, "  Lowest number of helpers:  "; LowestNumberOfHelpers
                Else
                    Print #ShowMe, "  (none found; no non-superset Faithfulness constraint favors a winner)"
                End If
            End If
            
        'Install the constraints that have the fewest helpers.
            For ConstraintIndex = 1 To mNumberOfConstraints
                If Stratum(ConstraintIndex) = 0 And Demotable(ConstraintIndex) = False Then
                    If NumberOfHelpers(ConstraintIndex) = LowestNumberOfHelpers Then
                        'Let Stratum(ConstraintIndex) = CurrentStratum
                        'Report progress:
                            Print #ShowMe, "  Constraint "; Abbrev(ConstraintIndex); " joins the current stratum, having "; Trim(Str(NumberOfHelpers(ConstraintIndex))); " helpers."
                    End If
                End If
            Next ConstraintIndex

        'Demote the constraints that have too many helpers.
            For ConstraintIndex = 1 To mNumberOfConstraints
                If Stratum(ConstraintIndex) = 0 And Demotable(ConstraintIndex) = False Then
                    If NumberOfHelpers(ConstraintIndex) > LowestNumberOfHelpers Then
                        Let Demotable(ConstraintIndex) = True
                        'Report progress:
                            Print #ShowMe, "  Constraint "; Abbrev(ConstraintIndex); " is excluded from stratum because it has "; Trim(Str(NumberOfHelpers(ConstraintIndex)));
                            Select Case NumberOfHelpers(ConstraintIndex)
                                Case 1
                                    Print #ShowMe, " helper."
                                Case Else
                                    Print #ShowMe, " helpers."
                            End Select
                    End If
                End If
            Next ConstraintIndex
        
ReportStrata:
        
      'Now, the delicate matter of knowing when you're done.  Here are the three cases:

        '  I. Some but not all of the yet-unranked constraints are demotible.
        '     --> Demote appropriately and continue.
        ' II. All of the yet-unranked constraints are demotible.
        '     -->  Record a failed constraint set and exit.
        'III. None of the yet-unranked constraints are demotible.
        '        (Stratum = 0, Demotible = NO)
        '     -->  They are the lowest stratum of a working grammar.

      'Work toward proving the following assertions:
          
          Let SomeAreNonDemotible = False   'by finding a nondemotible constraint
          Let SomeAreDemotible = False      'by finding a demotible constraint
          
            Dim Buffer As String
            For ConstraintIndex = 1 To mNumberOfConstraints
                If Stratum(ConstraintIndex) = 0 Then
                    Select Case Demotable(ConstraintIndex)
                        Case False
                            Let SomeAreNonDemotible = True
                            Let Stratum(ConstraintIndex) = CurrentStratum
                        Case True
                            Let SomeAreDemotible = True
                    End Select
                End If
            Next ConstraintIndex

          'Now act on the basis of these outcomes.

          
SkippationPoint:                                'See earlier:  the check of combined NoLoserPreferrers and APrioriRanking
          If SomeAreDemotible = False Then

             'This means that no constraints are demotible.
             'The remaining constraints forms the last stratum, and you're done.
             'They have already been assigned to the right stratum, so record your results.
                
            'BH:  Encode the result as a string, to facilitate communication across modules.
                Let Main = "True"
                Let Main = Main + Chr(9) + Trim(Str(CurrentStratum))
                For ConstraintIndex = 1 To mNumberOfConstraints
                    Let Main = Main + Chr(9) + Trim(Str(Stratum(ConstraintIndex)))
                Next ConstraintIndex

             'Report progress, first including the results of the last stratum:
                    Call Form1.PrintResultsOfRankingSoFar(CurrentStratum, Stratum(), ShowMe, mFaithfulness())
                    Print #ShowMe,
                    Print #ShowMe, "Ranking is complete and yields successful grammar."

             Close #ShowMe
             Exit Function
             
          ElseIf SomeAreNonDemotible = False Then
             
             'This means that all constraints are demotible.
             'There is no hope of a working grammar.
             'But keep the work done so far, since it will be useful in diagnosis.
                
                'Give a pseudo-stratum to the unrankable constraints, for purposes
                '   of diagnostic tableaux.
                    For ConstraintIndex = 1 To mNumberOfConstraints
                        If Stratum(ConstraintIndex) = 0 Then
                            Let Stratum(ConstraintIndex) = CurrentStratum
                        End If
                    Next ConstraintIndex
                
            'Encode the result as a string, to facilitate communication across modules.
                Let Main = "False"
                Let Main = Main + Chr(9) + Trim(Str(CurrentStratum))
                For ConstraintIndex = 1 To mNumberOfConstraints
                    Let Main = Main + Chr(9) + Trim(Str(Stratum(ConstraintIndex)))
                Next ConstraintIndex

                'Despair and depart:
                    Print #ShowMe,
                    Print #ShowMe, "Ranking has failed.  This constraint set is unable to derive only winners."
                    Close #ShowMe
                    Exit Function

            'If neither I or II are true, it means that some of the
            '   constraints are demotible, and some are not.
            '   Do nothing here, just keep going.

          End If

      'Find out which data should be ignored henceforth, because already learned from.
      '   This occurs when the Rival candidate violates a constraint
      '   that has just been ranked into the new stratum.

            For ConstraintIndex = 1 To mNumberOfConstraints
                If Stratum(ConstraintIndex) = CurrentStratum Then
                    For FormIndex = 1 To NumberOfForms
                        For RivalIndex = 1 To NumberOfRivals(FormIndex)
                            If RivalViolations(FormIndex, RivalIndex, ConstraintIndex) > WinnerViolations(FormIndex, ConstraintIndex) Then
                                Let StillInformative(FormIndex, RivalIndex) = False
                            End If
                        Next RivalIndex
                    Next FormIndex
                End If
            Next ConstraintIndex

      'Report progress:
            Call Form1.PrintResultsOfRankingSoFar(CurrentStratum, Stratum(), ShowMe, mFaithfulness())

   Loop         'Loop back and form another stratum.
   

   'You should never get this far, since the loop has an exit point at
   '  convergence/nonconvergence.
      MsgBox "Program error.   I would appreciate your letting me know the about the problem.  Email me at bhayes@humnet.ucla.edu, specifying error #87906, and including a copy of your input file.", vbCritical

End Function


Function Superset(FormIndex As Long, RivalIndex As Long, RivalViolations() As Long, _
   WinnerViolations() As Long) As Boolean

    'Determine whether a rival candidate has a superset of the violations of a
    '  winning candidate.  If so, the rival can never have any bearing on
    '  ranking.
    
    Dim ConstraintIndex As Long
    
    Let Superset = True
    
    For ConstraintIndex = 1 To mNumberOfConstraints
        If WinnerViolations(FormIndex, ConstraintIndex) > RivalViolations(FormIndex, RivalIndex, ConstraintIndex) Then
            Let Superset = False
            Exit Function
        End If
    Next ConstraintIndex

End Function


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

