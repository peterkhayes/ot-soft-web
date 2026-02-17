Attribute VB_Name = "RecursiveConstraintDemotion"
'=====================CLASSICAL BATCH CONSTRAINT DEMOTION==========================

    'Module level variables
        Dim mFaithfulness() As Boolean  'Classifies constraints as faithfulness or not.


Function Main(NumberOfForms As Long, InputForm() As String, NumberOfRivals() As Long, _
   WinnerViolations() As Long, Rival() As String, RivalViolations() As Long, NumberOfConstraints As Long, _
   Abbrev() As String, ConstraintName() As String) As String
   
    On Error GoTo CheckError
    
        'I'm having trouble communicating arrays across modules.
        '   So I will cheat.  This function return a string consisting of:
        
        '   the word True or False, indicating convergence
        '   tab, plus the number of strata
        '   the stratum of each constraint, each preceded by a tab
        
        Dim Stratum() As Long
        ReDim Stratum(NumberOfConstraints)      'You never need more strata than constraints.
        ReDim mFaithfulness(NumberOfConstraints)
      
        Dim CurrentStratum As Long
   
        Dim FormIndex As Long
        Dim RivalIndex As Long
        Dim ConstraintIndex As Long
        Dim OuterConstraintIndex As Long
        Dim InnerConstraintIndex As Long
      
        Dim SomeAreNonDemotible As Boolean
        Dim SomeAreDemotible As Boolean
        Dim StillInformative() As Boolean
        Let mMaximumNumberOfRivals = Form1.FindMaximumNumberOfRivals(NumberOfRivals())
        ReDim StillInformative(NumberOfForms, mMaximumNumberOfRivals)
        Dim Demotable() As Boolean
        ReDim Demotable(NumberOfConstraints)
          
        'File number of self-progress monitoring file.
            Dim ShowMe As Long
   
   'Initialize crucial variables, in case this routine gets called more than once.

      Let CurrentStratum = 0

      For ConstraintIndex = 1 To NumberOfConstraints
          Let Stratum(ConstraintIndex) = 0
          Let mFaithfulness(ConstraintIndex) = False
      Next ConstraintIndex

      For FormIndex = 1 To NumberOfForms
         For RivalIndex = 1 To mMaximumNumberOfRivals
            Let StillInformative(FormIndex, RivalIndex) = True
         Next RivalIndex
      Next FormIndex
      
        'First, make sure there is a folder for these files, a daughter of the
        '   folder in which the input file is located.
            Call Form1.CreateAFolderForOutputFiles
    
    'Produce a little file to show how you did it:
        Dim FoundOne As Boolean         'Used in self-report file to indicate searches that found nothing.
        Let ShowMe = FreeFile
        Open gOutputFilePath + "HowIRanked" + gFileName + ".txt" For Output As #ShowMe
        Print #ShowMe, "******Application of Constraint Demotion******"
        Print #ShowMe,
        Print #ShowMe, "Input file:  "; gOutputFilePath + gFileName + gFileSuffix
        Print #ShowMe,
   
    'Look up the a priori rankings if the user has so requested.
        If Form1.mnuConstrainAlgorithmsByAPrioriRankings.Checked = True Then
            'ReadAPrioriRankingsAsTable is a boolean function, which returns False if
            '   it failed to do its job.
            If APrioriRankings.ReadAPrioriRankingsAsTable(NumberOfConstraints, Abbrev()) = False Then
                Let mnuConstrainAlgorithmsByAPrioriRankings.Checked = False
            End If
        End If
   
   'Find which constraint are faithfulness constraints.  This information is not crucial to the algorithm
   '    but it helpful for the diagnosis file:
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
      
      'Initialize demotibility.
          For ConstraintIndex = 1 To NumberOfConstraints
             Let Demotable(ConstraintIndex) = False
          Next ConstraintIndex
          
      'AVOID PREFERENCE FOR LOSERS
      
      'Go through all pairs of Winner vs. Rival, eliminating constraints that prefer a loser.
      
        Let FoundOne = False
        For FormIndex = 1 To NumberOfForms
            For RivalIndex = 1 To NumberOfRivals(FormIndex)
            'Only still-informative Rivals can be learned from:
            If StillInformative(FormIndex, RivalIndex) = True Then
                'Examine constraints to see if they prefer losers.
                For ConstraintIndex = 1 To NumberOfConstraints
                    'We only consider constraints that haven't been ranked yet.
                    If Stratum(ConstraintIndex) = 0 Then
                    'Keep a constraint out of the current stratum if it prefers a loser.
                        If WinnerViolations(FormIndex, ConstraintIndex) > _
                            RivalViolations(FormIndex, RivalIndex, ConstraintIndex) Then
                                'Report this if it's the first time:
                                    If Demotable(ConstraintIndex) = False Then
                                        Print #ShowMe, "  "; Abbrev(ConstraintIndex); " is excluded from stratum; prefers loser *["; DumbSym(Rival(FormIndex, RivalIndex)); "] for /"; DumbSym(InputForm(FormIndex)); "/."
                                        Let FoundOne = True
                                    End If
                                Let Demotable(ConstraintIndex) = True
                        End If              'Did the constraint of ConstraintIndex prefer a loser?
                    End If                  'Is the constraint of ConstraintIndex still rankable?
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
                Let FoundOne = False
                For OuterConstraintIndex = 1 To NumberOfConstraints
                    If Stratum(OuterConstraintIndex) = 0 Then
                        For InnerConstraintIndex = 1 To NumberOfConstraints
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
            End If  'Are we doing a priori rankings?
        
        
      'There is also the delicate matter of knowing when you're done.  Here
      '  are the three cases:

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
            For ConstraintIndex = 1 To NumberOfConstraints
                If Stratum(ConstraintIndex) = 0 Then
                    Select Case Demotable(ConstraintIndex)
                        Case False
                            Let SomeAreNonDemotible = True
                            'Install in stratum:
                                Let Stratum(ConstraintIndex) = CurrentStratum
                            'Show progress:
                                Print #ShowMe, "  "; Abbrev(ConstraintIndex); " favors no losers, joins new stratum."
                        Case True
                            Let SomeAreDemotible = True
                    End Select
                End If
            Next ConstraintIndex

          'Now act on the basis of these outcomes.

          If SomeAreDemotible = False Then

             'This means that no constraints are demotible.
             'The remaining constraints forms the last stratum, and you're done.
             
             'They have already been assigned to the right stratum, so
             '  record your results.
             
             'Report progress, first including the results of the last stratum:
                Call Form1.PrintResultsOfRankingSoFar(CurrentStratum, Stratum(), ShowMe, mFaithfulness())   '1
                Print #ShowMe,
                Print #ShowMe, "Ranking is complete and yields successful grammar."
            
            'Construct the output string.
                Let Main = Main + "True"
                Let Main = Main + Chr(9) + Str(CurrentStratum)
                For ConstraintIndex = 1 To NumberOfConstraints
                    Let Main = Main + Chr(9) + Trim(Str(Stratum(ConstraintIndex)))
                Next ConstraintIndex

             Close #ShowMe
             Exit Function
             
          ElseIf SomeAreNonDemotible = False Then
             
             'This means that all constraints are demotible.
             'There is no hope of a working grammar.
             'But keep the work done so far, since it will be useful in diagnosis.
                
                'Give a pseudo-stratum to the unrankable constraints, for purposes
                '   of diagnostic tableaux.
                    For ConstraintIndex = 1 To NumberOfConstraints
                        Let Buffer = ConstraintName(ConstraintIndex)
                        If Stratum(ConstraintIndex) = 0 Then
                            Let Stratum(ConstraintIndex) = CurrentStratum
                        End If
                    Next ConstraintIndex
                
                'Despair and depart:
                    Print #ShowMe,
                    Print #ShowMe, "Ranking has failed.  "
                    If Form1.mnuConstrainAlgorithmsByAPrioriRankings.Checked = True Then
                        Print #ShowMe, "At least under the a priori rankings specified, this constraint set cannot derive only winners."
                    Else
                        Print #ShowMe, "This constraint set cannot derive only winners."
                    End If
                    Close #ShowMe
                    'Construct the output string.
                        Let Main = Main + "False"
                        Let Main = Main + Chr(9) + Trim(Str(CurrentStratum))
                        For ConstraintIndex = 1 To NumberOfConstraints
                            Let Main = Main + Chr(9) + Trim(Str(Stratum(ConstraintIndex)))
                        Next ConstraintIndex
                        Exit Function

            'If neither I or II are true, it means that some of the
            '   constraints are demotible, and some are not.
            '   Do nothing here, just keep going.

          End If

      'Find out which data should be ignored henceforth, because already learned from.
      '   This occurs when the Rival candidate violates a constraint
      '   that has just been ranked into the new stratum.
            For ConstraintIndex = 1 To NumberOfConstraints
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
            Call Form1.PrintResultsOfRankingSoFar(CurrentStratum, Stratum(), ShowMe, mFaithfulness())    '2
      
   Loop         'Loop back and form another stratum.
   

   'You should never get this far, since the loop has an exit point at
   '  convergence/nonconvergence.
        MsgBox "Program error.   I would appreciate your letting me know the about the problem.  Email me at bhayes@humnet.ucla.edu, specifying error #87905, and including a copy of your input file.", vbCritical


CheckError:
    Select Case Err.Number  ' Evaluate error number.
        Case 75 ' "File access error
            MsgBox "Error.  I conjecture that the file called HowIRanked" + gFileName + ".txt already exists in " + _
                gOutputFilePath + " as a Read-Only file.  Try deleting this file (or right click, Properties, decheck Read-Only) and rerunning OTSoft.", vbExclamation
            End
        Case Else
            MsgBox "Program error.  You can ask for help at bhayes@humnet.ucla.edu.  Please send a copy of your input file with your message, specifying error number 72078.", vbCritical
    End Select


End Function


