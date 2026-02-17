Attribute VB_Name = "PrintMiniTableaux"
'This ancient bit of code has a single purpose:  print out tiny 2 x 2 tableaux to illustrate "minimal pair" ranking
'   contradictions.
   
   
   Option Explicit

    Dim mShadingFlag As Boolean

Public Sub Main(NumberOfConstraints As Long, Abbrev() As String, FormIndex As Long, RivalIndex As Long, NumberOfStrata As Long, Stratum() As Long, _
    InputForm() As String, Winner() As String, WinnerViolations() As Long, Rival() As String, RivalViolations() As Long, _
    TmpFile As Long, DocFile As Long, HTMFile As Long)

   Dim NumberOfColumns As Long
   Dim ShadingPoint As Long
   Dim ShadingFlag As Boolean
   Dim FirstColumnWidth As Long
   Dim ExclamationSite As Long

   Dim ConstraintIndex As Long
   Dim InnerConstraintIndex As Long
   Dim LocalConstraintIndex As Long
   Dim SpaceIndex As Long
   Dim StratumIndex As Long
   Dim AsteriskIndex As Long

   Dim SpacesAvailable As Long
   Dim DigitCount As Long

   Dim FatalityFlag As Boolean

   Dim Successor() As Long
   ReDim Successor(MaximumNumberOfConstraints)
   
   Dim Table() As String
   Dim NumberOfIncludedConstraints As Long

   'Print out tableaux that contain only:

   '  --the winner and the rival that make the ranking argument
   '  --the constraints that one or the other violate

     'Precede the tableau with a blank line in the screen version.

         Print #TmpFile,

     'Count how many columns will be in this table.  This is the number
     '  of constraints that prefer the winner or rival, plus one column for
     '  the input.

        Let NumberOfColumns = 1
        For ConstraintIndex = 1 To NumberOfConstraints
           If WinnerViolations(FormIndex, ConstraintIndex) <> RivalViolations(FormIndex, RivalIndex, ConstraintIndex) Then
           'If WinnerViolations(FormIndex, ConstraintIndex) > 0 Or RivalViolations(FormIndex, RivalIndex, ConstraintIndex) > 0 Then
              Let NumberOfColumns = NumberOfColumns + 1
           End If
        Next ConstraintIndex
        
    'Now you can dimension the array that will give the HTML table.
        ReDim Table(3, 3)

     'Since we aren't going to print a column for every constraint, we need
     '  to know the "successor", if any, of each printed constraint, in
     '  order to control tabs and the solid/slashed line distinction.

        GoTo CalculateSuccessors
CalculateSuccessorsReturnPoint:

     'Further, we must determine at what point the winner defeats the rival, so shading may be correctly added.

        Let ShadingPoint = -1

        For StratumIndex = 1 To NumberOfStrata
           For ConstraintIndex = 1 To NumberOfConstraints
              If Stratum(ConstraintIndex) = StratumIndex Then
                 If WinnerViolations(FormIndex, ConstraintIndex) < RivalViolations(FormIndex, RivalIndex, ConstraintIndex) Then
                    Let ShadingPoint = ConstraintIndex
                    GoTo ShadingExitPoint
                 End If
              End If
           Next ConstraintIndex
        Next StratumIndex

ShadingExitPoint:

     'Print a clue for the Word macro to turn this into a pretty table, with the right number of columns.

         Print #DocFile, "\ts" + Trim(Str(NumberOfColumns))

     'Find the width of the first column.

         Let FirstColumnWidth = 0
         If Len(Winner(FormIndex)) > FirstColumnWidth Then
            Let FirstColumnWidth = Len(Winner(FormIndex))
         End If
         If Len(Rival(FormIndex, RivalIndex)) > FirstColumnWidth Then
            Let FirstColumnWidth = Len(Rival(FormIndex, RivalIndex))
         End If

     'Begin the tableau, by printing the input form.  This can be more elegantly omitted for the screen version.

         'Print #TmpFile, "/"; DumbSym(InputForm(FormIndex)); "/:"
         Print #DocFile, "/"; SymbolTag1; InputForm(FormIndex); SymbolTag2; "/";
         Let Table(1, 1) = InputForm(FormIndex)

     'Print enough blanks to line up the constraints

         For SpaceIndex = 1 To FirstColumnWidth + 2
            Print #TmpFile, " ";
         Next SpaceIndex

     'Print the constraint labels at the top.
     '  The order of printing is:  by ranking, i.e. by the strata established earlier in the program by the Constraint Demotion
     '  algorithm.  (needs work:  what if algorithm failed?)

     '  The outer three layers in this bit of code accomplish this sorting.

         Let NumberOfIncludedConstraints = 0
         For StratumIndex = 1 To NumberOfStrata
            For ConstraintIndex = 1 To NumberOfConstraints
               If Stratum(ConstraintIndex) = StratumIndex Then

                  'The next line of code establishes if a column deserves to be printed.  This currently is set to be true if the
                  ' constraint distinguishes winner and rival in some way.

                  'If WinnerViolations(FormIndex, ConstraintIndex) > 0 Or RivalViolations(FormIndex, RivalIndex, ConstraintIndex) > 0 Then
                  If WinnerViolations(FormIndex, ConstraintIndex) <> RivalViolations(FormIndex, RivalIndex, ConstraintIndex) Then

                     Print #TmpFile, Abbrev(ConstraintIndex);
                     If Len(Abbrev(ConstraintIndex)) = 1 Then Print #TmpFile, " ";
                     Print #DocFile, Chr$(9); SmallCapTag1; Abbrev(ConstraintIndex); SmallCapTag2;
                     Let NumberOfIncludedConstraints = NumberOfIncludedConstraints + 1
                     Let Table(NumberOfIncludedConstraints + 1, 1) = Abbrev(ConstraintIndex)

                     'Obtain the proper distribution of dashed and solid lines.

                     'No line at all (or:  table end), if you're at the last
                     '  constraint:

                        If Successor(ConstraintIndex) > 0 Then

                           'Solid lines for separate strata, else dotted lines:

                           If Stratum(ConstraintIndex) = Stratum(Successor(ConstraintIndex)) Then
                              Print #DocFile, "\dl";
                              Print #TmpFile, Chr(166);    'Dotted vertical line
                           Else
                              Print #TmpFile, "|";       'Solid vertical line
                           End If
                        End If

                  End If

               End If
            Next ConstraintIndex
         Next StratumIndex

         'Now, you've printed all the constraint labels.  Terminate this line.

           Print #TmpFile,
           Print #DocFile,

     'Next, print the winner:

           Print #DocFile, "\wsF\we "; SymbolTag1; Winner(FormIndex); SymbolTag2;
           Print #TmpFile, ">"; DumbSym(Winner(FormIndex));
           For SpaceIndex = Len(Winner(FormIndex)) To FirstColumnWidth
              Print #TmpFile, " ";
           Next SpaceIndex
            'HTML gets a nice pointing finger:
                Let Table(1, 2) = PointingFinger("htmfile") + Winner(FormIndex)

     'Then print the winner's violations:
           
           Let NumberOfIncludedConstraints = 0
           For StratumIndex = 1 To NumberOfStrata
              For ConstraintIndex = 1 To NumberOfConstraints
                 If Stratum(ConstraintIndex) = StratumIndex Then
                    'If WinnerViolations(FormIndex, ConstraintIndex) > 0 Or RivalViolations(FormIndex, RivalIndex, ConstraintIndex) > 0 Then
                    If WinnerViolations(FormIndex, ConstraintIndex) <> RivalViolations(FormIndex, RivalIndex, ConstraintIndex) Then
                       Let NumberOfIncludedConstraints = NumberOfIncludedConstraints + 1
                       GoTo MiniPrintTheWinnersViolations
MiniPrintTheWinnersViolationsReturnPoint:
                    End If
                 End If
              Next ConstraintIndex
           Next StratumIndex

           Print #TmpFile,
           Print #DocFile,

     'Reset the ShadingFlag.

          Let mShadingFlag = False

     'Then print the line for the crucial rival.

             'Print the rival itself.
                'Asterisks only for discrete algorithms.
                    If gAlgorithmName <> "GLA" Then
                        Print #DocFile, "*";
                        Print #TmpFile, "*";
                    End If
                 Print #DocFile, SymbolTag1; Rival(FormIndex, RivalIndex); SymbolTag2;
                 Print #TmpFile, DumbSym(Rival(FormIndex, RivalIndex));
                    For SpaceIndex = Len(Rival(FormIndex, RivalIndex)) To FirstColumnWidth
                       Print #TmpFile, " ";
                    Next SpaceIndex
                 Let Table(1, 3) = Rival(FormIndex, RivalIndex)

                 'The fatal violations flag will keep track of whether you've found the fatal violation yet.

                    Let FatalityFlag = False

                 'Print the rival violations.

                    Let NumberOfIncludedConstraints = 0
                    For StratumIndex = 1 To NumberOfStrata
                       For ConstraintIndex = 1 To NumberOfConstraints
                          If Stratum(ConstraintIndex) = StratumIndex Then
                             'If WinnerViolations(FormIndex, ConstraintIndex) > 0 Or RivalViolations(FormIndex, RivalIndex, ConstraintIndex) > 0 Then
                             If WinnerViolations(FormIndex, ConstraintIndex) <> RivalViolations(FormIndex, RivalIndex, ConstraintIndex) Then
                                Let NumberOfIncludedConstraints = NumberOfIncludedConstraints + 1
                                GoTo MiniPrintTheRivalViolations
MiniPrintTheRivalViolationsReturnPoint:
                             End If
                          End If
                       Next ConstraintIndex
                    Next StratumIndex

                 'Terminate the rival line you've been working on.
                    Print #TmpFile,
                    Print #DocFile,

                 'Reset the mShadingFlag.
                    Let mShadingFlag = False

     'Print a clue for the Word macro to finish turning this into a pretty table.
        Print #DocFile, "\te"
        
     'Print the table.
        Call s.PrintHTMTable(Table(), HTMFile, False, False, True)
        

   Exit Sub


Stop
'---------------------------------------------------------------------------
MiniPrintTheWinnersViolations:

     'The copy intended for Word gets a tab, before you print anything:

        Print #DocFile, Chr$(9);

     'xxx fix
     Let SpacesAvailable = Len(Abbrev(ConstraintIndex))
     If SpacesAvailable = 1 Then Let SpacesAvailable = 2
     Let DigitCount = Len(Trim(Str(WinnerViolations(FormIndex, ConstraintIndex))))
           

     Select Case WinnerViolations(FormIndex, ConstraintIndex)

        'No violations:  fill the cell with blank spaces.

        Case 0
           For SpaceIndex = 1 To Len(Abbrev(ConstraintIndex))
              If mShadingFlag = True Then
                 Print #TmpFile, " ";
              Else
                 Print #TmpFile, " ";
              End If
           Next SpaceIndex
           Let Table(NumberOfIncludedConstraints + 1, 2) = "&nbsp;"

       'Violations:  print the number, centered.

        Case Else
           For SpaceIndex = 1 To Int(SpacesAvailable / 2) - DigitCount
              If mShadingFlag = True Then
                 Print #TmpFile, " ";
              Else
                 Print #TmpFile, " ";
              End If
           Next SpaceIndex

           Print #TmpFile, Trim(Str(WinnerViolations(FormIndex, ConstraintIndex)));
           
           For SpaceIndex = 1 To SpacesAvailable - (Int(SpacesAvailable / 2))
               Print #TmpFile, " ";
           Next SpaceIndex

           'Print asterisks in final version.

                If DigitCount = 1 Then
                    For AsteriskIndex = 1 To WinnerViolations(FormIndex, ConstraintIndex)
                        Print #DocFile, "*";
                        Let Table(NumberOfIncludedConstraints + 1, 2) = Table(NumberOfIncludedConstraints + 1, 2) + "*"

                    Next AsteriskIndex
                Else
                    Print #DocFile, Trim(Str(WinnerViolations(FormIndex, ConstraintIndex)))
                    Let Table(NumberOfIncludedConstraints + 1, 2) = Trim(Str(WinnerViolations(FormIndex, ConstraintIndex)))
                End If

     End Select

        Call DealWithShadingPoint(ConstraintIndex, DocFile, ShadingPoint)

        'Print dotted or solid separator, unless you're at the end of the line.
            If Successor(ConstraintIndex) > 0 Then
               If Stratum(ConstraintIndex) = Stratum(Successor(ConstraintIndex)) Then
                  'Dotted:
                      Print #TmpFile, Chr(166);
               Else
                  'Solid:
                      Print #TmpFile, "|";
               End If
            End If

    GoTo MiniPrintTheWinnersViolationsReturnPoint


Stop
'-----------------------------------------------------------------------------
MiniPrintTheRivalViolations:

    'The copy intended for Word gets a tab, before you print anything:
        Print #DocFile, Chr$(9);

    'In getting the spacing right, there are three cases to consider:  no violation, fatal violation, non-fatal violation.
            Let SpacesAvailable = Len(Abbrev(ConstraintIndex))
            If SpacesAvailable = 1 Then Let SpacesAvailable = 2
            Let DigitCount = Len(Trim(Str(RivalViolations(FormIndex, RivalIndex, ConstraintIndex))))

           Select Case RivalViolations(FormIndex, RivalIndex, ConstraintIndex)

              Case 0
                 'No violation.
                    For SpaceIndex = 1 To Len(Abbrev(ConstraintIndex))
                       If mShadingFlag = True Then
                          Print #TmpFile, " ";
                       Else
                          Print #TmpFile, " ";
                       End If
                    Next SpaceIndex
                    Let Table(NumberOfIncludedConstraints + 1, 3) = "&nbsp;"
              Case Else
                 If FatalityFlag = False And RivalViolations(FormIndex, RivalIndex, ConstraintIndex) > WinnerViolations(FormIndex, ConstraintIndex) Then

                   'Fatal violation

                       'Only one violation can be fatal, so:
                           Let FatalityFlag = True
                           Let ExclamationSite = WinnerViolations(FormIndex, ConstraintIndex)
                           For SpaceIndex = 1 To Int(SpacesAvailable / 2) - DigitCount
                              Print #TmpFile, " ";
                           Next SpaceIndex

                           'Print the violation with a !
                               Print #TmpFile, Trim(Str(RivalViolations(FormIndex, RivalIndex, ConstraintIndex)));
                               Print #TmpFile, "!";
    
                               For SpaceIndex = 1 To SpacesAvailable - (Int(SpacesAvailable / 2)) - 1
                                  Print #TmpFile, " ";
                               Next SpaceIndex

                           'Here is the fancy hooey for asterisks, in pretty copy only.  Use a numeral
                           '  if more than 9 asterisks.
                                If DigitCount = 1 Then
                                    For AsteriskIndex = 1 To ExclamationSite + 1
                                       Print #DocFile, "*";
                                       Let Table(NumberOfIncludedConstraints + 1, 3) = Table(NumberOfIncludedConstraints + 1, 3) + "*"
                                    Next AsteriskIndex
                                    Print #DocFile, "!";
                                    Let Table(NumberOfIncludedConstraints + 1, 3) = Table(NumberOfIncludedConstraints + 1, 3) + "!"
                                    For AsteriskIndex = ExclamationSite + 2 To RivalViolations(FormIndex, RivalIndex, ConstraintIndex)
                                       Print #DocFile, "*";
                                       Let Table(NumberOfIncludedConstraints + 1, 3) = Table(NumberOfIncludedConstraints + 1, 3) + "*"
                                    Next AsteriskIndex
                                    
                                Else
                                    Print #DocFile, Trim(Str(RivalViolations(FormIndex, RivalIndex, ConstraintIndex)));
                                    Let Table(NumberOfIncludedConstraints + 1, 3) = Table(NumberOfIncludedConstraints + 1, 3) + Trim(Str(RivalViolations(FormIndex, RivalIndex, ConstraintIndex)))
                                    Print #DocFile, "!";
                                    Let Table(NumberOfIncludedConstraints + 1, 3) = Table(NumberOfIncludedConstraints + 1, 3) + "!"
                                End If

                   Else

                      'Non-fatal violation
                          For SpaceIndex = 1 To Int(SpacesAvailable / 2) - DigitCount
                              Print #TmpFile, " ";
                          Next SpaceIndex
                          Print #TmpFile, Trim(Str(RivalViolations(FormIndex, RivalIndex, ConstraintIndex)));
                          For SpaceIndex = 1 To SpacesAvailable - (Int(SpacesAvailable / 2))
                             If mShadingFlag = True Then
                                Print #TmpFile, " ";
                             Else
                                Print #TmpFile, " ";
                             End If
                          Next SpaceIndex

                          'Asterisks without exclamations:
                            If DigitCount = 1 Then
                                For AsteriskIndex = 1 To RivalViolations(FormIndex, RivalIndex, ConstraintIndex)
                                    Print #DocFile, "*";
                                Next AsteriskIndex
                                Let Table(NumberOfIncludedConstraints + 1, 3) = Table(NumberOfIncludedConstraints + 1, 3) + "*"
                            Else
                                Print #DocFile, Trim(Str(RivalViolations(FormIndex, RivalIndex, ConstraintIndex)))
                                Let Table(NumberOfIncludedConstraints + 1, 3) = Trim(Str(RivalViolations(FormIndex, RivalIndex, ConstraintIndex)))
                            End If

                    End If

              End Select

              Call DealWithShadingPoint(ConstraintIndex, DocFile, ShadingPoint)

              'Print suitable separators, if not at end of line.
                 If Successor(ConstraintIndex) > 0 Then
                    If Stratum(ConstraintIndex) = Stratum(Successor(ConstraintIndex)) Then
                       Print #TmpFile, Chr(166);
                    Else
                       Print #TmpFile, "|";
                    End If
                 End If

    GoTo MiniPrintTheRivalViolationsReturnPoint

Stop
'-----------------------------------------------------------------------------
CalculateSuccessors:
         
     'Since we aren't going to print a column for every constraint, we need
     '  to know the "successor", if any, of each printed constraint, in
     '  order to control tabs and the solid/slashed line distinction.

     'Let -1 mean "no successor".

      For ConstraintIndex = 1 To NumberOfConstraints

        'Initialize as -1, and reset later if appropriate.

           Let Successor(ConstraintIndex) = -1

        'Only calculate a successor if you need it:

        If WinnerViolations(FormIndex, ConstraintIndex) <> RivalViolations(FormIndex, RivalIndex, ConstraintIndex) Then

          'Go through the constraints in order of ranking.
          'If you find a legitimate successor, replace the value -1 with
            '  the index of that successor, then get out and move
            '  on to the next constraint.

              For StratumIndex = 1 To NumberOfStrata
                 For LocalConstraintIndex = 1 To NumberOfConstraints
                    If Stratum(LocalConstraintIndex) = StratumIndex Then

                       'Only constraints violated by winner or rival need be
                       '  considered:

                        If WinnerViolations(FormIndex, LocalConstraintIndex) <> RivalViolations(FormIndex, RivalIndex, LocalConstraintIndex) Then

                          'The criterion of being in a lower stratum:

                          If Stratum(LocalConstraintIndex) > Stratum(ConstraintIndex) Then

                             Let Successor(ConstraintIndex) = LocalConstraintIndex
                             GoTo SuccessorExitPoint

                          'If in same stratum, you are a successor if you have a higher index.

                          ElseIf Stratum(LocalConstraintIndex) = Stratum(ConstraintIndex) Then

                              If LocalConstraintIndex > ConstraintIndex Then
                                 Let Successor(ConstraintIndex) = LocalConstraintIndex
                                 GoTo SuccessorExitPoint
                              End If

                          End If

                        End If   'End if for whether InnerConstraint is violated by winner or rival.
                     End If      'End if for whether InnerConstraint matches stratum under consideration.

                  Next LocalConstraintIndex
              Next StratumIndex

          End If                 'End if for whether Constraint is violated by winner or rival.

SuccessorExitPoint:
        
        Next ConstraintIndex

GoTo CalculateSuccessorsReturnPoint


End Sub


Sub DealWithShadingPoint(MyConstraintIndex As Long, DocFile As Long, ShadingPoint As Long)

   'If the shading flag was set earlier to True, then shade.

      If mShadingFlag = True Then
        Print #DocFile, "\sh";
      End If

    'If you've hit the constraint marked as the shading point (i.e. the last unshaded constraint, then alter the shading flag.

       If MyConstraintIndex = ShadingPoint Then
          Let mShadingFlag = True
       End If

End Sub

