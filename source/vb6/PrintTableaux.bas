Attribute VB_Name = "PrintTableaux"

'This module inputs the underlying forms, winners, rival candidates, and violations, and
'   prints a set of tableaux, in both draft and pretty (Word-conversion) formats.

    Option Explicit

    'Variables appear in sorted versions for tableau production.
        Dim mSortedConstraintName() As String
        Dim mSortedAbbrev() As String
        Dim mSortedStratum() As Long
        Dim mSortedWinnerViolations() As Long
        Dim mSortedRivals() As String
        Dim mSortedRivalViolations() As Long
    
    'For planning whether to reverse axes
        Dim mTotalConstraintNameLength As Long
        Dim mTotalCandidateLength() As Long
        Dim mConditionalAxisReversalFlag() As Boolean
        
    'Shading and exclamation points
        Dim mShadingChoice As Boolean
        Dim mShadingDiacriticForDoc As String
        Dim mShadingColorForHTML As String
        Dim mExclamationPointChoice As Boolean
        Dim mEx As String
        
        Dim mNumberOfStrata As Long
        

Sub Main(NumberOfForms As Long, NumberOfConstraints As Long, ConstraintName() As String, _
    Abbrev() As String, Stratum() As Long, InputForm() As String, Winner() As String, _
    WinnerFrequency() As Single, WinnerViolations() As Long, _
    MaximumNumberOfRivals As Long, NumberOfRivals() As Long, Rival() As String, RivalFrequency() As Single, RivalViolations() As Long, _
    TmpFile As Long, DocFile As Long, HTMFile As Long, _
    AlgorithmName As String, RunningFactorialTypology As Boolean, FactorialTypologyIndex As Long, _
    ShadingChoice As Boolean, ExclamationPointChoice As Boolean)
    
    
       Call CountTheStrata(Stratum())
       Call SetUpSortedArrays(NumberOfForms, NumberOfConstraints, ConstraintName(), _
           Abbrev(), Stratum(), MaximumNumberOfRivals, NumberOfRivals(), Rival(), WinnerViolations(), _
           RivalViolations())
       Call SortTheConstraints(NumberOfForms, NumberOfRivals(), NumberOfConstraints)
       Call SortTheCandidates(NumberOfForms, NumberOfRivals(), NumberOfConstraints)
       Call PickAxes(NumberOfForms, Winner(), NumberOfRivals(), NumberOfConstraints, TmpFile, DocFile)
       'Localize the tableaux-formatting choices and implement them.
            Let mShadingChoice = ShadingChoice
            Let mExclamationPointChoice = ExclamationPointChoice
            Call EstablishShadingAndExclamationPoints
       Call PrintTableaux(NumberOfConstraints, mSortedConstraintName(), mSortedAbbrev(), _
           mSortedStratum(), _
           NumberOfForms, InputForm(), Winner(), mSortedWinnerViolations(), _
           MaximumNumberOfRivals, NumberOfRivals(), mSortedRivals(), mSortedRivalViolations(), RunningFactorialTypology, FactorialTypologyIndex, _
           TmpFile, DocFile, HTMFile)
    'Save a sorted version of the input file if it was so requested.
       If Form1.mnuSaveAsTxtSortedByRank.Checked = True Then
           Call SaveSortedInputFile(NumberOfConstraints, mSortedConstraintName(), mSortedAbbrev(), _
               mSortedStratum(), NumberOfForms, InputForm(), Winner(), mSortedWinnerViolations(), _
               NumberOfRivals(), mSortedRivals(), mSortedRivalViolations(), _
               WinnerFrequency(), RivalFrequency())
       End If
       
End Sub

Public Sub InitiateHTML(FileNum As Long)

    'Print the stuff that comes at the top of the HTML output file.
    'This involves a cascading style sheet I cribbed from the web and do not fully understand.
    
    Print #FileNum, "<HEAD>"
    Print #FileNum, "<title>OTSoft " + gMyVersionNumber + " " + gFileName + gFileSuffix + "</title>"
    Print #FileNum, "<meta http-equiv=" + Chr(34) + "Content-Style-Type" + Chr(34) + " content=" + Chr(34) + "text/css" + Chr(34) + ">"
    Print #FileNum, "<link rel=" + Chr(34) + "stylesheet" + Chr(34) + " type=" + Chr(34) + "text/css" + Chr(34) + " media=" + Chr(34) + "screen" + Chr(34) + " href=" + Chr(34) + "base.css" + Chr(34) + ">"
    Print #FileNum, "<style type=" + Chr(34) + "text/css" + Chr(34) + ">"
    Print #FileNum, ".test {background-color: white; border-style: solid; border-width: 0;}.cl1 {border-right-width: thin;}"
    
    'CSS turns the specifications for a cell into a very short code starting with the letters "cl".  Here are the four specs used here
    '   for producing tableau shading and stratum separators.
        'Shading, right border:
            Print #FileNum, ".cl4 {border-right-width: 2px;border-right-color:gray;background-color: "; gShadingColor; ";}"
        'No shading, right border:
            Print #FileNum, ".cl8 {border-right-width: 2px;border-right-color:gray;background-color: white;}"
        'Shading, no right border"
            Print #FileNum, ".cl9 {background-color: "; gShadingColor; ";}"
        'No shading, no right border:
            Print #FileNum, ".cl10 {background-color: white;}"
    
    Print #FileNum, "</style>"
    Print #FileNum, "</HEAD>"
    Print #FileNum, "<body>"
    
    'Other styles not used here:
        'Print #FileNum, ".cl2 {border-right-width: medium;}"
        'Print #FileNum, ".cl3 {border-right-width: thick;}"
        'Print #FileNum, ".cl5 {border-right-width: 0;}"
        'Print #FileNum, ".cl6 {border-right-width: -5px;}"
        'Print #FileNum, ".cl7 {border-right-width: thin;}"
    
    

End Sub


Sub SetUpSortedArrays(NumberOfForms As Long, NumberOfConstraints As Long, ConstraintName() As String, _
    Abbrev() As String, Stratum() As Long, MaximumNumberOfRivals As Long, NumberOfRivals() As Long, Rivals() As String, WinnerViolations() As Long, _
    RivalViolations() As Long)

    'In order to make ordinary tableaux, the constraints have to be sorted by rank, using the
    '   strata discovered by algorithm.
    '   Here, the sorted constraints will be placed on module-level arrays, to make it easier
    '   for other procedures to access them.

        Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
        
    'Redimension them to appropriate size.
     
        ReDim mSortedConstraintName(NumberOfConstraints)
        ReDim mSortedAbbrev(NumberOfConstraints)
        ReDim mSortedStratum(NumberOfConstraints)
        ReDim mSortedWinnerViolations(NumberOfForms, NumberOfConstraints)
        ReDim mSortedRivals(NumberOfForms, MaximumNumberOfRivals) As String
        ReDim mSortedRivalViolations(NumberOfForms, MaximumNumberOfRivals, NumberOfConstraints) As Long
     
     'Transfer the data to the module-level arrays for future processing.
         For ConstraintIndex = 1 To NumberOfConstraints
             Let mSortedConstraintName(ConstraintIndex) = ConstraintName(ConstraintIndex)
             Let mSortedAbbrev(ConstraintIndex) = Abbrev(ConstraintIndex)
             Let mSortedStratum(ConstraintIndex) = Stratum(ConstraintIndex)
         Next ConstraintIndex
         For FormIndex = 1 To NumberOfForms
             For ConstraintIndex = 1 To NumberOfConstraints
                 Let mSortedWinnerViolations(FormIndex, ConstraintIndex) = WinnerViolations(FormIndex, ConstraintIndex)
             Next ConstraintIndex
             For RivalIndex = 1 To NumberOfRivals(FormIndex)
                 Let mSortedRivals(FormIndex, RivalIndex) = Rivals(FormIndex, RivalIndex)
                     For ConstraintIndex = 1 To NumberOfConstraints
                         Let mSortedRivalViolations(FormIndex, RivalIndex, ConstraintIndex) = RivalViolations(FormIndex, RivalIndex, ConstraintIndex)
                     Next ConstraintIndex
             Next RivalIndex
         Next FormIndex
    
    End Sub


Sub SortTheConstraints(NumberOfForms, NumberOfRivals() As Long, NumberOfConstraints As Long)

    'In order to make ordinary tableaux, the constraints have to be sorted by rank, using the
    '   strata discovered by algorithm.
    '   Here, the sorted constraints will be placed on module-level arrays, to make it easier
    '   for other procedures to access them.

        Dim SwapString As String, SwapInt As Long
        Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
        Dim OuterSortingIndex As Long, InnerSortingIndex As Long
        
     'Examine all constraint pairs.
         For OuterSortingIndex = 1 To NumberOfConstraints
             For InnerSortingIndex = OuterSortingIndex + 1 To NumberOfConstraints
                 'See if they're in the wrong order.
                    If mSortedStratum(InnerSortingIndex) < mSortedStratum(OuterSortingIndex) Then
                        'If so, exchange all their attributes.
                            'Constraint names.
                                Let SwapString = mSortedConstraintName(InnerSortingIndex)
                                Let mSortedConstraintName(InnerSortingIndex) = mSortedConstraintName(OuterSortingIndex)
                                Let mSortedConstraintName(OuterSortingIndex) = SwapString
                            'Constraint abbreviations.
                                Let SwapString = mSortedAbbrev(InnerSortingIndex)
                                Let mSortedAbbrev(InnerSortingIndex) = mSortedAbbrev(OuterSortingIndex)
                                Let mSortedAbbrev(OuterSortingIndex) = SwapString
                            'Stratal affiliations.
                                Let SwapInt = mSortedStratum(InnerSortingIndex)
                                Let mSortedStratum(InnerSortingIndex) = mSortedStratum(OuterSortingIndex)
                                Let mSortedStratum(OuterSortingIndex) = SwapInt
                            'Winner and rival violations
                                For FormIndex = 1 To NumberOfForms
                                    'Winner:
                                        Let SwapInt = mSortedWinnerViolations(FormIndex, InnerSortingIndex)
                                        Let mSortedWinnerViolations(FormIndex, InnerSortingIndex) = _
                                            mSortedWinnerViolations(FormIndex, OuterSortingIndex)
                                        Let mSortedWinnerViolations(FormIndex, OuterSortingIndex) = SwapInt
                                    'Rival:
                                        For RivalIndex = 0 To NumberOfRivals(FormIndex)
                                            Let SwapInt = mSortedRivalViolations(FormIndex, RivalIndex, InnerSortingIndex)
                                            Let mSortedRivalViolations(FormIndex, RivalIndex, InnerSortingIndex) = _
                                                mSortedRivalViolations(FormIndex, RivalIndex, OuterSortingIndex)
                                            Let mSortedRivalViolations(FormIndex, RivalIndex, OuterSortingIndex) = _
                                                SwapInt
                                        Next RivalIndex
                                Next FormIndex
                    End If                      'Does current constraint order match stratum order?
             Next InnerSortingIndex
         Next OuterSortingIndex
           
End Sub

Sub SortTheCandidates(NumberOfForms As Long, NumberOfRivals() As Long, NumberOfConstraints As Long)

    'A tableau is more coherent if the candidates are sorted by harmony.
    '   This procedure with with the module-level arrays that
    '   store sorted information; see top of this module for these arrays.
              
    'This can be time-consuming so should be turnable off.
        If Form1.mnuSortCandidatesByHarmony.Checked = False Then Exit Sub
              
        'Variables to store swapped material temporarily.
            Dim SwapString As String, SwapInt As Long, SwapDouble As Single
        'Indices
            Dim FormIndex As Long, ConstraintIndex As Long, SortingConstraintIndex As Long
            Dim OuterSortingIndex As Long, InnerSortingIndex As Long
            
    'Report progress.
        Let Form1.lblProgressWindow.Caption = "Sorting candidate for tableaux ..."
        DoEvents
    
    'Sort every form.
    
    For FormIndex = 1 To NumberOfForms
        'Do all pairwise comparisons.
            For OuterSortingIndex = 1 To NumberOfRivals(FormIndex)
            For InnerSortingIndex = OuterSortingIndex + 1 To NumberOfRivals(FormIndex)
            'Go through the constraints in ranking order.  N.B. They are already appropriately
            '   sorted.
                For ConstraintIndex = 1 To NumberOfConstraints
                    'Check relative harmony.
                    If mSortedRivalViolations(FormIndex, InnerSortingIndex, ConstraintIndex) _
                      < mSortedRivalViolations(FormIndex, OuterSortingIndex, ConstraintIndex) Then
                        'This pair is in the wrong order.  Fix by swapping.
                            'Swap the candidates themselves.
                                Let SwapString = mSortedRivals(FormIndex, InnerSortingIndex)
                                Let mSortedRivals(FormIndex, InnerSortingIndex) = _
                                    mSortedRivals(FormIndex, OuterSortingIndex)
                                Let mSortedRivals(FormIndex, OuterSortingIndex) = SwapString
                            'Swap their frequencies. xxx fix me
                            '    Let SwapDouble = msortedRivalFrequency(FormIndex, InnerSortingIndex)
                            '    Let msortedRivalFrequency(FormIndex, InnerSortingIndex) = _
                            '        msortedRivalFrequency(FormIndex, OuterSortingIndex)
                            '    Let msortedRivalFrequency(FormIndex, OuterSortingIndex) = SwapDouble
                            'Swap their constraint violations.
                                For SortingConstraintIndex = 1 To NumberOfConstraints
                                    Let SwapInt = mSortedRivalViolations(FormIndex, InnerSortingIndex, SortingConstraintIndex)
                                    Let mSortedRivalViolations(FormIndex, InnerSortingIndex, SortingConstraintIndex) _
                                        = mSortedRivalViolations(FormIndex, OuterSortingIndex, SortingConstraintIndex)
                                    Let mSortedRivalViolations(FormIndex, OuterSortingIndex, SortingConstraintIndex) = SwapInt
                                Next SortingConstraintIndex
                        'Since this constraint established a harmonic order, compare no further constraints.
                            Exit For
                    ElseIf mSortedRivalViolations(FormIndex, InnerSortingIndex, ConstraintIndex) _
                        > mSortedRivalViolations(FormIndex, OuterSortingIndex, ConstraintIndex) Then
                        'The candidates are already in the right order; move on to the next
                        '   pairwise comparison.
                            Exit For
                    Else
                        'Do nothing:  the candidates are tied so far, and you need to loop down
                        '   to further constraints to find out the final answer.
                    End If
                Next ConstraintIndex        'Go down the constraints in ranking order.
        Next InnerSortingIndex              'Do all pairwise comparisons
        Next OuterSortingIndex
    Next FormIndex                          'Sort every form.

End Sub

Sub PickAxes(NumberOfForms As Long, Winner() As String, NumberOfRivals() As Long, NumberOfConstraints As Long, _
    TmpFile As Long, DocFile As Long)

     
     'Check if switched-axis tableaux are appropriate.

     Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
     
     'First, determine the total string length of the constraint labels,
     '  so we can guess whether vertical or horizonal tableaux fit better on the page.

        For ConstraintIndex = 1 To NumberOfConstraints
           Let mTotalConstraintNameLength = mTotalConstraintNameLength + Len(mSortedAbbrev(ConstraintIndex)) + 1
        Next ConstraintIndex

     'Next, determine the total string length of the candidate set for each
     '   form, to see if it is sometimes shorter.
        ReDim mTotalCandidateLength(NumberOfForms)
        For FormIndex = 1 To NumberOfForms
           Let mTotalCandidateLength(FormIndex) = mTotalCandidateLength(FormIndex) + Len(Winner(FormIndex)) + 2
           For RivalIndex = 1 To NumberOfRivals(FormIndex)
              Let mTotalCandidateLength(FormIndex) = mTotalCandidateLength(FormIndex) + Len(mSortedRivals(FormIndex, RivalIndex)) + 2
           Next RivalIndex
        Next FormIndex

    'Go through the lengths.  If the candidates would fit seriously better (five fewer total characters)
    '   than the constraints, check with user about reversing the axes.
        If mTotalConstraintNameLength > 75 Then
            'Put in a diacritic inducing Landscape formatting
            'I'm not really wild about this and have turned it off for now.
            '    Print #DocFile, "\la"
            'Encode the instructions of the option buttons.
                ReDim mConditionalAxisReversalFlag(NumberOfForms)
                For FormIndex = 1 To NumberOfForms
                    If mTotalCandidateLength(FormIndex) < mTotalConstraintNameLength + 5 Then
                        If Form1.optSwitchSomeAxes.Value = True Then
                            Let mConditionalAxisReversalFlag(FormIndex) = True
                        End If
                        Exit Sub
                    End If
                Next FormIndex
        End If


End Sub


Sub EstablishShadingAndExclamationPoints()
    
    'Cancel shading if not desired for this sort of tableau.
        If mShadingChoice = True Then
            Let mShadingDiacriticForDoc = "\sh"
            Let mShadingColorForHTML = gShadingColor
        Else
            Let mShadingDiacriticForDoc = ""
            'White:
                Let mShadingColorForHTML = "FFFFFF"
        End If
    
    'Leave out the exclamation points if the tableau type is one that doesn't want them.
        If mExclamationPointChoice = True Then
            Let mEx = "!"
        Else
            Let mEx = " "
        End If
        
End Sub


Sub PrintTableaux(NumberOfConstraints As Long, ConstraintName() As String, Abbrev() As String, _
   Stratum() As Long, _
   NumberOfForms As Long, InputForm() As String, Winner() As String, WinnerViolations() As Long, _
   MaximumNumberOfRivals As Long, NumberOfRivals() As Long, Rival() As String, RivalViolations() As Long, _
   RunningFactorialTypology As Boolean, FactorialTypologyIndex As Long, _
   TmpFile As Long, DocFile As Long, HTMFile As Long)
     
     Dim FormIndex As Long
     Dim ConstraintIndex As Long
     Dim OuterConstraintIndex As Long
     Dim SpaceIndex As Long
     Dim AsteriskIndex As Long
     Dim SpacesAvailable As Long
     Dim DigitCount As Long
     Dim FirstColumnWidth As Long
     Dim FatalityFlag As Boolean
     
     Dim WinnerShadingPoint As Long
     Dim LocalRivalIsDead() As Boolean
     ReDim LocalRivalIsDead(MaximumNumberOfRivals)

     'Go through all input forms and print the tableaux that result.

    'Print an explanation of GLA-style tableaux.
        If gAlgorithmName = "GLA" And RunningFactorialTypology = False Then
            'Header:
            Print #DocFile, "\ks"
            Call PrintLevel1Header(DocFile, TmpFile, HTMFile, "Tableaux")
            Call PrintPara(DocFile, TmpFile, HTMFile, "The following are approximate tableaux for this ranking.  PARAOutputs are derived simply by sorting the constraints by PARAtheir ranking value, with no stochastic variation.")
            Call PrintPara(DocFile, TmpFile, HTMFile, "To diagnose variation, consult two things:")
            Call PrintPara(DocFile, TmpFile, HTMFile, "--The candidate frequencies (which are the generated frequencies, not the input frequencies).")
            Call PrintPara(DocFile, TmpFile, HTMFile, "--The probability that each constraint outranks the next one down, given directly after the constraint labels.")
        Else
            
            
            'xxx needs work to implement nested headings.
            
            'Print
            'Let gLevel1HeadingNumber = gLevel1HeadingNumber + 1
            'Print #HTMFile, "<p>"
            'If RunningFactorialTypology = True Then
            '    Print #TmpFile, Trim(Str(FactorialTypologyIndex)); ".";
            '    Print #HTMFile, Trim(Str(FactorialTypologyIndex)); ".";
            'End If
            'Print #TmpFile, Trim(Str(gLevel1HeadingNumber)); ". Tableaux"
            'Print #HTMFile, "<p><p><b>" + Trim(Str(gLevel1HeadingNumber)); ". Tableaux" + "</b><p><p>"
            
            'Lower heading style for factorial typology, which must report for each grammar.
            '    If RunningFactorialTypology = True Then
            '        Print #DocFile, "\h3";
            '    Else
            '        Print #DocFile, "\h1";
            '    End If
            'Print #DocFile, "Tableaux"
        End If
     
     
     'Here begins the major loop for tableau-printing:  loop through all forms.
     
     For FormIndex = 1 To NumberOfForms
        'Tableaux for wug forms only if user asked.
            'If TestWugOnly = True Then
            '    If GLA.IsAWugForm(FormIndex) = False Then
            '        GoTo ExitPoint
            '    End If
            'End If
        'Decide whether to print normal or reversed-axis tableaux:
           If Form1.optSwitchAll.Value = True Then
              Call PrintReversedAxisTableaux(FormIndex, InputForm(), Winner(), WinnerViolations(), _
                 MaximumNumberOfRivals, NumberOfRivals(), Rival(), RivalViolations(), NumberOfConstraints, Abbrev(), Stratum(), _
                 SymbolTag1, SymbolTag2, SmallCapTag1, SmallCapTag2, DocFile, TmpFile, HTMFile)
           ElseIf Form1.optSwitchSomeAxes.Value = True Then
              If mConditionalAxisReversalFlag(FormIndex) = True Then
                Call PrintReversedAxisTableaux(FormIndex, InputForm(), Winner(), WinnerViolations(), _
                   MaximumNumberOfRivals, NumberOfRivals(), Rival(), RivalViolations(), NumberOfConstraints, Abbrev(), Stratum(), _
                   SymbolTag1, SymbolTag2, SmallCapTag1, SmallCapTag2, DocFile, TmpFile, HTMFile)
              Else
                 Call PrintNormalTableaux(FormIndex, InputForm(), Winner(), WinnerViolations(), _
                    NumberOfRivals(), Rival(), RivalViolations(), _
                    NumberOfConstraints, Abbrev(), Stratum(), SymbolTag1, SymbolTag2, SmallCapTag1, _
                    SmallCapTag2, DocFile, TmpFile, HTMFile, MaximumNumberOfRivals)
              End If
           Else
              Call PrintNormalTableaux(FormIndex, InputForm(), Winner(), WinnerViolations(), _
                    NumberOfRivals(), Rival(), RivalViolations(), _
                    NumberOfConstraints, Abbrev(), Stratum(), SymbolTag1, SymbolTag2, SmallCapTag1, _
                    SmallCapTag2, DocFile, TmpFile, HTMFile, MaximumNumberOfRivals)
           End If
ExitPoint:
     Next FormIndex

End Sub



Sub PrintNormalTableaux(FormIndex As Long, InputForm() As String, _
    Winner() As String, WinnerViolations() As Long, NumberOfRivals() As Long, _
    Rival() As String, RivalViolations() As Long, _
    NumberOfConstraints As Long, Abbrev() As String, Stratum() As Long, _
    SymbolTag1 As String, SymbolTag2 As String, SmallCapTag1 As String, SmallCapTag2 As String, _
    DocFile As Long, TmpFile As Long, HTMFile As Long, MaximumNumberOfRivals As Long)
        
     Dim RivalIndex As Long, ConstraintIndex As Long, SpaceIndex As Long
     Dim FirstColumnWidth As Long, FatalityFlag As Boolean
     Dim WinnerShadingPoint As Long
         
      'Print a clue for the Word macro to turn this into a pretty table.
         Print #DocFile, "\ts"; Trim(Str(NumberOfConstraints + 1)); "\ks"
         Print #HTMFile, gHTMLTableSpecs

      'Find the width of the first column.
         Let FirstColumnWidth = 0
         If Len(Winner(FormIndex)) > FirstColumnWidth Then
            Let FirstColumnWidth = Len(Winner(FormIndex))
         End If
         For RivalIndex = 1 To NumberOfRivals(FormIndex)
            If Len(Rival(FormIndex, RivalIndex)) > FirstColumnWidth Then
               Let FirstColumnWidth = Len(Rival(FormIndex, RivalIndex))
            End If
         Next RivalIndex
       
      'Print the underlying representation in the upper left corner.
        Print #TmpFile,
        Print #TmpFile, "/"; DumbSym(InputForm(FormIndex)); "/: "
        Print #DocFile, "/"; SymbolTag1; InputForm(FormIndex); SymbolTag2; "/";
        Print #HTMFile, "<TR>"
        Print #HTMFile, "<TD>"
        Print #HTMFile, "/"; DumbSym(InputForm(FormIndex)); "/: "
        Print #HTMFile, "</TD>"
        

      'Print enough blanks to line up the constraints
         For SpaceIndex = 1 To FirstColumnWidth + 1
            Print #TmpFile, " ";
         Next SpaceIndex
         'More blanks for discrete algorithms, to cover *'s.
             If gAlgorithmName <> "GLA" Then
                Print #TmpFile, " ";
             End If

      'Print the constraint labels at the top
           For ConstraintIndex = 1 To NumberOfConstraints
              Print #TmpFile, Abbrev(ConstraintIndex);
              Print #DocFile, Chr$(9); SmallCapTag1; Abbrev(ConstraintIndex); SmallCapTag2;
              Print #HTMFile, "<TD ALIGN=Center>"
              
              'Separate them with dotted or solid lines, depending on
              '  stratum membership.  \dl is a diacritic for the Word
              '  macro that triggers conversion to a dotted line.
                If ConstraintIndex < NumberOfConstraints Then
                   If Stratum(ConstraintIndex) = Stratum(ConstraintIndex + 1) Then
                      'dotted line:
                        Print #DocFile, "\dl";
                        Print #TmpFile, Chr(166);
                   Else
                      'solid line:
                        Print #TmpFile, "|";
                        Print #HTMFile, "<p class=" + Chr(34) + "test cl8" + Chr(34) + ">";
                   End If
                Else
                   Print #TmpFile,
                   Print #DocFile,
                End If
                Print #HTMFile, Abbrev(ConstraintIndex)
                Print #HTMFile, "</TD>"
           Next ConstraintIndex
           Print #HTMFile, "</TR>"


        'Then print the winner at the beginning of the top line:
            Print #HTMFile, "<TR>"
            Print #HTMFile, "<TD>"
            'Pointy finger only for discrete algorithms.
                If gAlgorithmName <> "GLA" Then
                    Print #DocFile, PointingFinger("docfile");
                    Print #TmpFile, PointingFinger("tmpfile");
                    Print #HTMFile, PointingFinger("htmfile");
                End If
           Print #DocFile, " "; SymbolTag1; Winner(FormIndex); SymbolTag2; Chr$(9);
           Print #TmpFile, DumbSym(Winner(FormIndex));
           Print #HTMFile, " "; DumbSym(Winner(FormIndex));
              For SpaceIndex = Len(Winner(FormIndex)) To FirstColumnWidth
                 Print #TmpFile, " ";
              Next SpaceIndex

         'Find the winner's point where shading must commence.
            Let WinnerShadingPoint = FindWinnerShadingPoint(NumberOfConstraints, FormIndex, WinnerViolations(), NumberOfRivals(), RivalViolations, MaximumNumberOfRivals)
         
         'Then print the winner's violations:
            Call PrintTheWinnersViolations(FormIndex, _
                 Abbrev(), _
                 WinnerViolations(), _
                 WinnerShadingPoint, _
                 NumberOfConstraints, Stratum(), _
                 TmpFile, DocFile, HTMFile)
           
        'Then print the LocalRivals and their violations.
           For RivalIndex = 1 To NumberOfRivals(FormIndex)
                'Print the rival itself.
                    'No *'s or pointy fingers for GLA:
                        If gAlgorithmName <> "GLA" Then
                            Print #TmpFile, " ";
                        End If
                    Print #DocFile, SymbolTag1; Rival(FormIndex, RivalIndex); SymbolTag2; Chr$(9);
                    Print #TmpFile, DumbSym(Rival(FormIndex, RivalIndex));
                    Print #HTMFile, "<TR>"
                    Print #HTMFile, "<TD>"
                    Print #HTMFile, "&nbsp; &nbsp; &nbsp;" + DumbSym(Rival(FormIndex, RivalIndex));
                    For SpaceIndex = Len(Rival(FormIndex, RivalIndex)) To FirstColumnWidth
                        Print #TmpFile, " ";
                    Next SpaceIndex
              'Print the rival violations.
                    Call PrintRivalViolations(FormIndex, _
                        RivalIndex, _
                        Abbrev(), _
                        WinnerViolations(), _
                        RivalViolations(), _
                        NumberOfConstraints, Stratum(), _
                        TmpFile, DocFile, HTMFile)
           Next RivalIndex

     'Separate tableaux by a blank line in the crude file.
        Print #TmpFile,
     'Print a clue for the Word macro to finish turning this into a pretty table.
        Print #DocFile, "\te\ke"
        Print #HTMFile, "</TABLE>"
        Print #HTMFile, "<P>"
        
End Sub

Public Sub PrintTheWinnersViolations(MyForm As Long, _
    Abbrev() As String, WinnerViolations() As Long, _
    ByVal WinnerShadingPoint As Long, _
    NumberOfConstraints As Long, Stratum() As Long, _
    TmpFile As Long, DocFile As Long, HTMFile As Long)
    
    'Print the sequence of winner violations, for the top row of a tableau.

        Dim SpaceIndex As Long, AsteriskIndex As Long, ConstraintIndex As Long
        Dim SpacesAvailable As Long
        Dim FatalityFlag As Boolean
        Dim DigitCount As Long
        Dim MyCellStyle As String  'For HTM formatting.

    'Loop through the constraints
        For ConstraintIndex = 1 To NumberOfConstraints
         
             Let SpacesAvailable = Len(Abbrev(ConstraintIndex))
        
            'Html tag for column.  It varies according to shading.
                'If ConstraintIndex > WinnerShadingPoint Then
                    'Print #HTMFile, "<TD ALIGN=Center bgcolor=#" + mShadingColorForHTML + ">"
                'Else
                    Print #HTMFile, "<TD ALIGN=Center>"
                'End If
             
             'The HTML needs to know at this stage what style of cell.  There are shaded, non-shaded, right border, non-right border.
             '  These are the CSS styles defined as "cX" above.
                'Right border?
                Let MyCellStyle = ChooseCellStyle(ConstraintIndex, NumberOfConstraints, WinnerShadingPoint, Stratum(), mNumberOfStrata)
                Print #HTMFile, "<p class=" + Chr(34) + "test " + MyCellStyle + Chr(34) + ">";
             
             
             Select Case WinnerViolations(MyForm, ConstraintIndex)
                'No violations:  fill the cell with blank spaces or shading.
                    Case 0
                       For SpaceIndex = 1 To Len(Abbrev(ConstraintIndex))
                          If ConstraintIndex > WinnerShadingPoint Then
                             Print #TmpFile, " ";
                          Else
                             Print #TmpFile, " ";
                          End If
                       Next SpaceIndex
                       If ConstraintIndex > WinnerShadingPoint Then
                          Print #DocFile, mShadingDiacriticForDoc;
                       End If
                       Print #HTMFile, "&nbsp;"
               'Violations:  print the number, centered.
                    Case Is < 10
                       For SpaceIndex = 1 To Int(SpacesAvailable / 2) - 1
                          If ConstraintIndex > WinnerShadingPoint Then
                             Print #TmpFile, " ";
                          Else
                             Print #TmpFile, " ";
                          End If
                       Next SpaceIndex
                       Print #TmpFile, Trim(Str(WinnerViolations(MyForm, ConstraintIndex)));
                       For SpaceIndex = 1 To SpacesAvailable - (Int(SpacesAvailable / 2))
                          If ConstraintIndex > WinnerShadingPoint Then
                             Print #TmpFile, " ";
                          Else
                             Print #TmpFile, " ";
                          End If
                       Next ' ˜
                        'Print asterisks in the pretty version.
                              For AsteriskIndex = 1 To WinnerViolations(MyForm, ConstraintIndex)
                                 Print #DocFile, "*";
                                 Print #HTMFile, "*";
                              Next AsteriskIndex
                           If ConstraintIndex > WinnerShadingPoint Then
                              Print #DocFile, mShadingDiacriticForDoc;
                           End If
                   'Violations for large number:  print the number, centered.  Count its digits to get spacing
                   '  right, and don't do asterisks in the pretty version.
                    Case Is >= 10
                        Let DigitCount = Len(Trim(Str(WinnerViolations(MyForm, ConstraintIndex))))
                        For SpaceIndex = 1 To Int(SpacesAvailable / 2) - 1 - (DigitCount - 1)
                            Print #TmpFile, " ";
                        Next SpaceIndex
                        Print #DocFile, Trim(Str(WinnerViolations(MyForm, ConstraintIndex)));
                        Print #HTMFile, Trim(Str(WinnerViolations(MyForm, ConstraintIndex)));
                        Print #TmpFile, Trim(Str(WinnerViolations(MyForm, ConstraintIndex)));
                        For SpaceIndex = 1 To SpacesAvailable - (Int(SpacesAvailable / 2))
                            Print #TmpFile, " ";
                        Next
                        If ConstraintIndex > WinnerShadingPoint Then
                           Print #DocFile, mShadingDiacriticForDoc;
                        End If
             End Select             'How many violations?
        
               'Html column end.
                Print #HTMFile, "</TD>"
               
               
               'Print dotted or solid separator, tab separator, and line endings.
                    If ConstraintIndex < NumberOfConstraints Then
                       If Stratum(ConstraintIndex) = Stratum(ConstraintIndex + 1) Then
                          'Dotted:
                              Print #TmpFile, Chr(166);
                       Else
                          'Solid:
                              Print #TmpFile, "|";
                              Print #HTMFile, "<p class=" + Chr(34) + "test cl4" + Chr(34) + ">";
    
                       End If
                       'Tab separator:
                            Print #DocFile, Chr$(9);
                    Else
                        'Line ending.
                            Print #DocFile,
                            Print #TmpFile,
                            Print #HTMFile, "</TR>"
                    End If

    Next ConstraintIndex
    
End Sub

Sub PrintRivalViolations(ByVal MyForm As Long, ByVal MyRival As Long, _
    Abbrev() As String, WinnerViolations() As Long, _
    RivalViolations() As Long, NumberOfConstraints As Long, Stratum() As Long, _
    TmpFile As Long, DocFile As Long, HTMFile As Long)

    'Print the violations of one rival candidate.
    
        Dim SpacesAvailable As Long
        Dim ExclamationSite As Long
        Dim DigitCount As Long
        Dim ConstraintIndex As Long, SpaceIndex As Long, AsteriskIndex As Long
        Dim FatalViolationAlreadyFound As Boolean
        Dim MyCellStyle As String

    'Initialize the FatalViolationAlreadyFound flag.
        Let FatalViolationAlreadyFound = False
    
    'Go through all the constraints
        For ConstraintIndex = 1 To NumberOfConstraints
    
            'Html tag for column.  It varies according to shading.
                
                'If FatalViolationAlreadyFound Then
                '    Print #HTMFile, "<TD ALIGN=Center bgcolor=#" + mShadingColorForHTML + ">"
                'Else
                    Print #HTMFile, "<TD ALIGN=Center>"
                'End If
            
             'The HTML needs to know at this stage what style of cell.  There are shaded, non-shaded, right border, non-right border.
             '  These are the CSS styles defined as "cX" above.
                'Right border?
                Let MyCellStyle = ChooseCellStyleForRival(ConstraintIndex, NumberOfConstraints, FatalViolationAlreadyFound, Stratum(), mNumberOfStrata)
                Print #HTMFile, "<p class=" + Chr(34) + "test " + MyCellStyle + Chr(34) + ">";
            
            'Establish column width
                Let SpacesAvailable = Len(Abbrev(ConstraintIndex))

            'In getting the spacing right, there are three cases to consider:
            '   no violation, fatal violation, non-fatal violation.
                   Select Case RivalViolations(MyForm, MyRival, ConstraintIndex)
                      Case 0
                         'No violation.
                            'Diacritic for shading (possible only in the fancy file):
                                If FatalViolationAlreadyFound Then
                                     Print #DocFile, mShadingDiacriticForDoc;
                                End If
                            'Temp file:  enough spaces to fill this column.
                                For SpaceIndex = 1 To Len(Abbrev(ConstraintIndex))
                                    Print #TmpFile, " ";
                                Next SpaceIndex
                            'HTML file:  a place holder to cause there to be a border.
                               Print #HTMFile, "&nbsp;"
                      Case Else
                        'Print a violation.
                            Let DigitCount = Len(Trim(Str(RivalViolations(MyForm, MyRival, ConstraintIndex))))
                            If FatalViolationAlreadyFound = False And RivalViolations(MyForm, MyRival, ConstraintIndex) > WinnerViolations(MyForm, ConstraintIndex) Then
                                'This is a fatal violation
                                'Only one violation can be fatal, so:
                                   Let FatalViolationAlreadyFound = True
                                   Let ExclamationSite = WinnerViolations(MyForm, ConstraintIndex)
                                   For SpaceIndex = 1 To Int(SpacesAvailable / 2) - DigitCount
                                      Print #TmpFile, " ";
                                   Next SpaceIndex
                               'Print the violation with a !
                                   Print #TmpFile, Trim(Str(RivalViolations(MyForm, MyRival, ConstraintIndex)));
                                   Print #TmpFile, mEx;
                                   For SpaceIndex = 1 To SpacesAvailable - (Int(SpacesAvailable / 2)) - 1
                                      Print #TmpFile, " ";
                                   Next SpaceIndex
                               'Asterisks in the pretty copy.
                                   If DigitCount = 1 Then
                                        For AsteriskIndex = 1 To ExclamationSite + 1
                                            Print #DocFile, "*";
                                            Print #HTMFile, "*";
                                        Next AsteriskIndex
                                        Print #DocFile, mEx;
                                        Print #HTMFile, mEx;
                                        For AsteriskIndex = ExclamationSite + 2 To RivalViolations(MyForm, MyRival, ConstraintIndex)
                                            Print #DocFile, "*";
                                            Print #HTMFile, "*";
                                        Next AsteriskIndex
                                    Else
                                        Print #DocFile, Trim(Str(RivalViolations(MyForm, MyRival, ConstraintIndex)));
                                        Print #DocFile, mEx;
                                        Print #HTMFile, mEx;
                                    End If
                           Else
                              'Non-fatal violation
                                  For SpaceIndex = 1 To Int(SpacesAvailable / 2) - DigitCount
                                     If FatalViolationAlreadyFound = True Then
                                        Print #TmpFile, " ";
                                     Else
                                        Print #TmpFile, " ";
                                     End If
                                  Next SpaceIndex
                                  Print #TmpFile, Trim(Str(RivalViolations(MyForm, MyRival, ConstraintIndex)));
                                  For SpaceIndex = 1 To SpacesAvailable - (Int(SpacesAvailable / 2))
                                     If FatalViolationAlreadyFound = True Then
                                        Print #TmpFile, " ";
                                     Else
                                        Print #TmpFile, " ";
                                     End If
                                  Next SpaceIndex
                                'Asterisks without exclamations:
                                    If DigitCount = 1 Then
                                        For AsteriskIndex = 1 To RivalViolations(MyForm, MyRival, ConstraintIndex)
                                            Print #DocFile, "*";
                                            Print #HTMFile, "*";
                                        Next AsteriskIndex
                                    Else
                                        Print #DocFile, Trim(Str(RivalViolations(MyForm, MyRival, ConstraintIndex)));
                                        Print #HTMFile, Trim(Str(RivalViolations(MyForm, MyRival, ConstraintIndex)));
                                    End If
                                    If FatalViolationAlreadyFound = True Then
                                        Print #DocFile, mShadingDiacriticForDoc;
                                    End If
                            End If                  'Was the violation fatal?
                   End Select                       'Does this cell have any violations?
        
                    
                    'Html column end.
                     Print #HTMFile, "</TD>"
                    
                    'Print suitable separators and line endings.
                         If ConstraintIndex < NumberOfConstraints Then
                            'Dotted lines for constraints in the same stratum, else solid.
                                If Stratum(ConstraintIndex) = Stratum(ConstraintIndex + 1) Then
                                   Print #TmpFile, Chr(166);
                                Else
                                    Print #TmpFile, "|";
                                End If
                                Print #DocFile, Chr(9);
                         Else
                            'End the line.
                                Print #TmpFile,
                                Print #DocFile,
                                Print #HTMFile, "</TR>"
                         End If
                 
    Next ConstraintIndex

End Sub

Function FindWinnerShadingPoint(NumberOfConstraints As Long, FormIndex As Long, WinnerViolations() As Long, LocalNumberOfRivals() As Long, RivalViolations() As Long, MaximumNumberOfRivals As Long)

   Dim RivalIndex As Long, ConstraintIndex As Long
   Dim LocalRivalIsDead() As Boolean
   ReDim LocalRivalIsDead(MaximumNumberOfRivals)
   
   
   'Initialize an array used to find when all the LocalRivals are dead.
      For RivalIndex = 1 To LocalNumberOfRivals(FormIndex)
         Let LocalRivalIsDead(RivalIndex) = False
      Next RivalIndex

   'Find the spot after which all LocalRivals are dead, so winner can be shaded.

      For ConstraintIndex = 1 To NumberOfConstraints
         For RivalIndex = 1 To LocalNumberOfRivals(FormIndex)
            If RivalViolations(FormIndex, RivalIndex, ConstraintIndex) > WinnerViolations(FormIndex, ConstraintIndex) Then
               Let LocalRivalIsDead(RivalIndex) = True
            End If
         Next RivalIndex

         For RivalIndex = 1 To LocalNumberOfRivals(FormIndex)
            If LocalRivalIsDead(RivalIndex) = False Then GoTo Line1195ContinuationPoint
         Next RivalIndex
      
         'Since all the LocalRivals are dead now, you've found the point where
         '  you can start shading.  Leave the subroutine.
            Let FindWinnerShadingPoint = ConstraintIndex
            Exit Function
        
Line1195ContinuationPoint:

      Next ConstraintIndex
      
      'If you never found a shading point, set it high to avoid earlier
      '  points interfering.

         Let FindWinnerShadingPoint = 1000

End Function

Sub SaveSortedInputFile(NumberOfConstraints As Long, ConstraintName() As String, Abbrev() As String, _
   Stratum() As Long, _
   NumberOfForms As Long, InputForm() As String, Winner() As String, WinnerViolations() As Long, _
   LocalNumberOfRivals() As Long, LocalRival() As String, LocalRivalViolations() As Long, _
   LocalWinnerFrequency() As Single, LocalRivalFrequency() As Single)
     
     'Print a version of the input file, sorted by constraint ranking and
     '   also sorting the candidates by harmony.
     
        Dim FormIndex As Long
        Dim RivalIndex As Long
        Dim ConstraintIndex As Long
        Dim ReportErrorFileName As String
        Dim TextFile As Long

     'Open the file.
        Let ReportErrorFileName = gFileName + "Sorted.txt"
        Let TextFile = FreeFile
        Open gOutputFilePath + gFileName + "Sorted.txt" For Output As #TextFile
            
    'Constraint names.  These were sorted prior to calling this routine; ditto for
    '   other information.
        Print #TextFile, Chr(9); Chr(9);
        For ConstraintIndex = 1 To NumberOfConstraints
            Print #TextFile, Chr(9); ConstraintName(ConstraintIndex);
        Next ConstraintIndex
        Print #TextFile,
    
    'Constraint abbreviations:
        Print #TextFile, Chr(9); Chr(9);
        For ConstraintIndex = 1 To NumberOfConstraints
                Print #TextFile, Chr(9); Abbrev(ConstraintIndex);
        Next ConstraintIndex
        Print #TextFile,
    
    'Inputs, winners, rivals, violations.  Loop through all forms.
        For FormIndex = 1 To NumberOfForms
            'Print the input and winner.
                Print #TextFile, InputForm(FormIndex); Chr(9);
                Print #TextFile, Winner(FormIndex);
            'Print winner's frequency.
                'Zero is blank.
                If LocalWinnerFrequency(FormIndex) = 0 Then
                    Print #TextFile, Chr(9);
                Else
                    Print #TextFile, Chr(9); LocalWinnerFrequency(FormIndex);
                End If
            'Then print the winner's violations:
                For ConstraintIndex = 1 To NumberOfConstraints
                    If WinnerViolations(FormIndex, ConstraintIndex) = 0 Then
                        Print #TextFile, Chr(9);
                    Else
                        Print #TextFile, Chr(9); WinnerViolations(FormIndex, ConstraintIndex);
                    End If
                Next ConstraintIndex
                Print #TextFile,
            'Then print the LocalRivals and their frequencies and violations.
                For RivalIndex = 1 To LocalNumberOfRivals(FormIndex)
                    Print #TextFile, Chr(9); LocalRival(FormIndex, RivalIndex);
                    'Now, print frequency.  Zero is blank.
                        If LocalRivalFrequency(FormIndex, RivalIndex) = 0 Then
                            Print #TextFile, Chr(9);
                        Else
                            Print #TextFile, Chr(9); LocalRivalFrequency(FormIndex, RivalIndex);
                        End If
                    For ConstraintIndex = 1 To NumberOfConstraints
                        If LocalRivalViolations(FormIndex, RivalIndex, ConstraintIndex) = 0 Then
                            Print #TextFile, Chr(9);
                        Else
                            Print #TextFile, Chr(9); LocalRivalViolations(FormIndex, RivalIndex, ConstraintIndex);
                        End If
                    Next ConstraintIndex
                    Print #TextFile,
                Next RivalIndex
        Next FormIndex
   
    'Close the file.
        Close #TextFile
   
End Sub

Sub PrintReversedAxisTableaux(MyForm As Long, InputForm() As String, Winner() As String, WinnerViolations() As Long, _
    MaximumNumberOfRivals As Long, NumberOfRivals() As Long, Rival() As String, RivalViolations() As Long, NumberOfConstraints As Long, _
    Abbrev() As String, Stratum() As Long, _
    SymbolTag1 As String, SymbolTag2 As String, SmallCapTag1 As String, SmallCapTag2 As String, _
    DocFile As Long, TmpFile As Long, HTMFile As Long)

    Dim FirstColumnWidth As Long
    Dim SpacesAvailable As Long
    Dim WinnerShadingPoint As Long
 
    Dim ExclamationFlag() As Boolean
    Dim ExclamationSite() As Long
    
    Dim RivalIndex As Long, ConstraintIndex As Long, SpaceIndex As Long, AsteriskIndex As Long
    Dim DigitCount As Long
    
    ReDim ExclamationFlag(MaximumNumberOfRivals)
    ReDim ExclamationSite(MaximumNumberOfRivals)
    
    'xxx This needs work in that the HTML output does not have stratal separators.  It could be done with table borders and cascading
    '   style sheets, but I doubt these tableaux get used much anyway.
        
    'Find the winner's point where shading must commence.
       Let WinnerShadingPoint = FindWinnerShadingPoint(NumberOfConstraints, MyForm, WinnerViolations(), NumberOfRivals(), RivalViolations(), MaximumNumberOfRivals)

    'Print a clue for the Word macro to turn this into a pretty table.
        Print #DocFile, "\ts"; Trim(Str(NumberOfRivals(MyForm) + 2)); "\ks"
        Dim MyTable() As String
        ReDim MyTable(NumberOfRivals(MyForm) + 2, NumberOfConstraints + 1)

    'Find the width of the first column--the maximal thing that has to fit.
        Let FirstColumnWidth = 0
        For ConstraintIndex = 1 To NumberOfConstraints
            If Len(Abbrev(ConstraintIndex)) > FirstColumnWidth Then
                Let FirstColumnWidth = Len(Abbrev(ConstraintIndex))
            End If
        Next ConstraintIndex

    'Print the input form.
        Print #TmpFile,
        Print #TmpFile, "/"; DumbSym(InputForm(MyForm)); "/: "
        Print #DocFile, "/"; SymbolTag1; InputForm(MyForm); SymbolTag2; "/";
        Let MyTable(1, 1) = InputForm(MyForm)

    'Print enough blanks to line up the candidates
        For SpaceIndex = 1 To FirstColumnWidth + 1
            Print #TmpFile, " ";
        Next SpaceIndex
        Print #TmpFile, "|";
        Print #DocFile, Chr$(9);

    'Print the winner at the top:
        'Pointy finger only for discrete algorithms.
            If gAlgorithmName <> "GLA" Then
                Print #DocFile, PointingFinger("docfile");
                Print #TmpFile, PointingFinger("tmpfile");
                Let MyTable(2, 1) = PointingFinger("htmfile") + Winner(MyForm)
            Else
                Let MyTable(2, 1) = Winner(MyForm)
            End If
        Print #TmpFile, DumbSym(Winner(MyForm));
        Print #DocFile, SymbolTag1; Winner(MyForm); SymbolTag2;
        'Print a solid separator:
            Print #DocFile, Chr$(9);
            Print #TmpFile, "|";

    'Print the LocalRivals at the top:

        For RivalIndex = 1 To NumberOfRivals(MyForm)
            'Asterisks only for discrete algorithms.
                If gAlgorithmName <> "GLA" Then
                    Print #TmpFile, " ";
                End If
            Print #DocFile, SymbolTag1; Rival(MyForm, RivalIndex); SymbolTag2;
            Print #TmpFile, DumbSym(Rival(MyForm, RivalIndex));
              If RivalIndex < NumberOfRivals(MyForm) Then
                'Print a solid separator:
                    Print #TmpFile, "|";
                    Print #DocFile, Chr$(9);
              End If
            Let MyTable(RivalIndex + 2, 1) = Rival(MyForm, RivalIndex)
           Next RivalIndex
           Print #TmpFile,
           Print #DocFile,

    'Initialize the array that remembers whether to put a !
            For RivalIndex = 1 To NumberOfRivals(MyForm)
                Let ExclamationFlag(RivalIndex) = False
                Let ExclamationSite(RivalIndex) = 0
            Next RivalIndex

    'Now, one row for each constraints
            For ConstraintIndex = 1 To NumberOfConstraints
                'Print a stratum divider where appropriate.
                '   No point in doing it for GLA:  all constraints assumed in separate strata.
                If ConstraintIndex > 1 And gAlgorithmName <> "GLA" Then
                    If Stratum(ConstraintIndex) > Stratum(ConstraintIndex - 1) Then
                        'Print --- in appropriate numbers to cover the first column.
                        For SpaceIndex = 1 To FirstColumnWidth + 1
                            Print #TmpFile, "-";
                        Next SpaceIndex
                        'Print a solid separator:
                            Print #TmpFile, "|";
                        'Print --- in appropriate numbers to cover the winner's column.
                            For SpaceIndex = 1 To Len(Winner(MyForm))
                                Print #TmpFile, "-";
                            Next SpaceIndex
                            'Another one for discrete algorithms.
                                If gAlgorithmName <> "GLA" Then Print #TmpFile, "-";
                        'Print a solid separator:
                            Print #TmpFile, "|";
                        'Print --- in appropriate numbers to cover the rivals' columns.
                            For RivalIndex = 1 To NumberOfRivals(MyForm)
                               For SpaceIndex = 1 To Len(Rival(MyForm, RivalIndex))
                                  Print #TmpFile, "-";
                               Next SpaceIndex
                               If gAlgorithmName <> "GLA" Then Print #TmpFile, "-";
                                 'Print a solid separator:
                                    If RivalIndex < NumberOfRivals(MyForm) Then
                                       Print #TmpFile, "|";
                                    Else
                                       Print #TmpFile,
                                    End If
                            Next RivalIndex
                 Else
                    'Print a signal for Word to dot this line
                    Print #DocFile, "\hd";
                 End If
            End If

           'Print the constraint name:
              Print #DocFile, SmallCapTag1; Abbrev(ConstraintIndex); SmallCapTag2; Chr$(9);
              Let MyTable(1, ConstraintIndex + 1) = Abbrev(ConstraintIndex)
              Print #TmpFile, Abbrev(ConstraintIndex);
                 For SpaceIndex = Len(Abbrev(ConstraintIndex)) To FirstColumnWidth
                    Print #TmpFile, " ";
                 Next SpaceIndex
              'Print a solid separator:
                 Print #TmpFile, "|";

           'Then print the violations of the winner:
                    If gAlgorithmName = "GLA" Then
                        Let SpacesAvailable = Len(Winner(MyForm))
                    Else
                        Let SpacesAvailable = Len(Winner(MyForm)) + 1
                    End If

                   'First, locate whether there are any violations at all.
                   If WinnerViolations(MyForm, ConstraintIndex) > 0 Then
                        Let DigitCount = Len(Trim(Str(WinnerViolations(MyForm, ConstraintIndex))))
                        'Print the right number of spaces in the temp file, before and after winner violation count.
                            For SpaceIndex = 1 To Int(SpacesAvailable / 2) - DigitCount
                                Print #TmpFile, " ";
                            Next SpaceIndex
                            Print #TmpFile, Trim(Str(WinnerViolations(MyForm, ConstraintIndex)));
                            For SpaceIndex = 1 To SpacesAvailable - (Int(SpacesAvailable / 2))
                                Print #TmpFile, " ";
                            Next SpaceIndex
                        'Print asterisks for final version, using digits if it's too many.
                            Print #DocFile, AsteriskString(WinnerViolations(MyForm, ConstraintIndex))
                            Let MyTable(2, ConstraintIndex + 1) = AsteriskString(WinnerViolations(MyForm, ConstraintIndex))
                        'This was commented out earlier, and I'm not sure why.  I'm putting it back
                        '  in to see if it's ok.
                            If ConstraintIndex > WinnerShadingPoint Then
                                Print #DocFile, mShadingDiacriticForDoc;
                                Let MyTable(2, ConstraintIndex + 1) = MyTable(2, ConstraintIndex + 1) + mShadingDiacriticForDoc
                            End If
                    'Finally, cover the cases of zero violations.
                    Else
                        For SpaceIndex = 1 To SpacesAvailable
                            Print #TmpFile, " ";
                        Next SpaceIndex
                        If ConstraintIndex > WinnerShadingPoint Then
                           Print #DocFile, mShadingDiacriticForDoc;
                           Let MyTable(2, ConstraintIndex + 1) = MyTable(2, ConstraintIndex + 1) + mShadingDiacriticForDoc
                        End If
                    End If                  'Where there violations?

                'Print a solid separator or a tab:
                   Print #TmpFile, "|";
                   Print #DocFile, Chr$(9);


           'Then print the violations of the Rivals:
              For RivalIndex = 1 To NumberOfRivals(MyForm)
                 Let SpacesAvailable = Len(Rival(MyForm, RivalIndex))
                 If gAlgorithmName = "GLA" Then
                    Let SpacesAvailable = Len(Rival(MyForm, RivalIndex))
                 Else
                    Let SpacesAvailable = Len(Rival(MyForm, RivalIndex)) + 1
                 End If

                 
                 Let DigitCount = Len(Trim(Str(RivalViolations(MyForm, RivalIndex, ConstraintIndex))))
                 
                    'First, locate fatal violations, which require !.
                    If RivalViolations(MyForm, RivalIndex, ConstraintIndex) > WinnerViolations(MyForm, ConstraintIndex) And ExclamationFlag(RivalIndex) = False Then
                        Let ExclamationFlag(RivalIndex) = True
                        Let ExclamationSite(RivalIndex) = WinnerViolations(MyForm, ConstraintIndex)

                        For SpaceIndex = 1 To Int(SpacesAvailable / 2) - DigitCount
                           Print #TmpFile, " ";
                        Next SpaceIndex
                        Print #TmpFile, Trim(Str(RivalViolations(MyForm, RivalIndex, ConstraintIndex)));
                        Print #TmpFile, mEx;
                        For SpaceIndex = 1 To SpacesAvailable - (Int(SpacesAvailable / 2)) - 1
                           Print #TmpFile, " ";
                        Next SpaceIndex
                        Print #DocFile, AsterisksWithExclamationMark(RivalViolations(MyForm, RivalIndex, ConstraintIndex), ExclamationSite(RivalIndex))
                        Let MyTable(RivalIndex + 2, ConstraintIndex + 1) = AsterisksWithExclamationMark(RivalViolations(MyForm, RivalIndex, ConstraintIndex), ExclamationSite(RivalIndex))

                       'Print a solid separator or a tab:
                           If RivalIndex < NumberOfRivals(MyForm) Then
                              Print #DocFile, Chr$(9);
                              Print #TmpFile, "|";
                           Else
                              Print #DocFile,
                           End If

                    'Next, locate whether there are any violations at all.
                    ElseIf RivalViolations(MyForm, RivalIndex, ConstraintIndex) > 0 Then

                        For SpaceIndex = 1 To Int(SpacesAvailable / 2) - 1
                            Print #TmpFile, " ";
                        Next SpaceIndex
                        Print #TmpFile, Trim(Str(RivalViolations(MyForm, RivalIndex, ConstraintIndex)));
                        For SpaceIndex = 1 To SpacesAvailable - (Int(SpacesAvailable / 2))
                            Print #TmpFile, " ";
                        Next SpaceIndex
                        
                        Print #DocFile, AsteriskString(RivalViolations(MyForm, RivalIndex, ConstraintIndex))
                        Let MyTable(RivalIndex + 2, ConstraintIndex + 1) = AsteriskString(RivalViolations(MyForm, RivalIndex, ConstraintIndex))
                        
                        If ExclamationFlag(RivalIndex) = True Then
                           Print #DocFile, mShadingDiacriticForDoc;
                           Let MyTable(RivalIndex + 2, ConstraintIndex + 1) = MyTable(RivalIndex + 2, ConstraintIndex + 1) + mShadingDiacriticForDoc
                        End If

                        If RivalIndex < NumberOfRivals(MyForm) Then
                           Print #DocFile, Chr$(9);
                           Print #TmpFile, "|";
                        Else
                           Print #DocFile,
                        End If
                        
                    'Finally, cover the cases of zero violations.
                    Else
                        For SpaceIndex = 1 To SpacesAvailable
                            Print #TmpFile, " ";
                        Next SpaceIndex
                        If ExclamationFlag(RivalIndex) = True Then
                           Print #DocFile, mShadingDiacriticForDoc;
                           Let MyTable(RivalIndex + 2, ConstraintIndex + 1) = MyTable(RivalIndex + 2, ConstraintIndex + 1) + mShadingDiacriticForDoc
                        End If
                        If RivalIndex < NumberOfRivals(MyForm) Then
                           Print #DocFile, Chr$(9);
                           Print #TmpFile, "|";
                        Else
                            'End of row for pretty file.
                                Print #DocFile,
                        End If
                    End If
           Next RivalIndex

            'End of row for draft file.
                 Print #TmpFile,

        Next ConstraintIndex            'This completes a constraint, and you can go on to the next row.

     Call s.PrintHTMTable(MyTable(), HTMFile, False, False, True)
     'Print a clue for the Word macro to turn this into a pretty table.
        Print #DocFile, "\ke\te"

End Sub

Public Function AsteriskString(MyCount As Long) As String

    'Returns asterisk strings unless at least 10 violations.
    
    Dim DigitCount As Long
    Dim AsteriskIndex As Long
    Dim Buffer As String
    
    Let DigitCount = Len(Trim(Str(MyCount)))
    If DigitCount = 1 Then
        For AsteriskIndex = 1 To MyCount
            Let Buffer = Buffer + "*"
        Next AsteriskIndex
    Else
        Let Buffer = Trim(Str(MyCount))
    End If
    Let AsteriskString = Buffer

End Function

Function AsterisksWithExclamationMark(NumberOfAsterisks As Long, ExclamationSite As Long) As String

    Dim AsteriskIndex As Long, Buffer As String
    
    If NumberOfAsterisks < 10 Then
        For AsteriskIndex = 1 To ExclamationSite + 1
           Let Buffer = Buffer + "*"
        Next AsteriskIndex
        Let Buffer = Buffer + mEx
        For AsteriskIndex = ExclamationSite + 2 To NumberOfAsterisks
           Let Buffer = Buffer + "*"
        Next AsteriskIndex
    Else
        Let Buffer = Trim(Str(NumberOfAsterisks))
        Let Buffer = Buffer + mEx
    End If
    Let AsterisksWithExclamationMark = Buffer
End Function


Sub CountTheStrata(MyStratum() As Long)

    Dim i As Long, Buffer As Long
    For i = 1 To UBound(MyStratum())
        If MyStratum(i) > Buffer Then
            Let Buffer = MyStratum(i)
        End If
    Next i
    Let mNumberOfStrata = Buffer
End Sub

Function ChooseCellStyle(MyConstraint, NumberOfConstraints, ShadingPoint, Stratum() As Long, mNumberOfStrata) As String
                
    'Final column cells are un"lined" and bolded according to the Shading Point.
        If MyConstraint = NumberOfConstraints Then
            If MyConstraint > ShadingPoint Then
                'Shaded no line
                    Let ChooseCellStyle = "cl9"         'Shaded, no line
            Else
                'Unshaded no line
                    Let ChooseCellStyle = "cl10"        'Unshaded, no line
            End If
        Else
            'Nonfinal column cells are lined if the next stratum.
                'Different stratum from next:
                    If Stratum(MyConstraint) < Stratum(MyConstraint + 1) Then
                        'Shading?
                            If MyConstraint > ShadingPoint Then
                                Let ChooseCellStyle = "cl4"         'Lined and shaded
                            Else
                                Let ChooseCellStyle = "cl8"        'Lined and not shaded
                            End If
                    Else
                        'Same stratum as next:
                            'Shading?
                                If MyConstraint > ShadingPoint Then
                                    Let ChooseCellStyle = "cl9"         'Shaded, no line
                                Else
                                    Let ChooseCellStyle = "cl10"         'Unshaded, no line
                                End If
                    End If
        End If


End Function


Function ChooseCellStyleForRival(MyConstraint, NumberOfConstraints, FatalViolationAlreadyFound As Boolean, Stratum() As Long, mNumberOfStrata) As String
                
    'Final column cells are un"lined" and bolded according to the Shading Point.
        If MyConstraint = NumberOfConstraints Then
            If FatalViolationAlreadyFound Then
                'Shaded no line
                    Let ChooseCellStyleForRival = "cl9"        'Shaded, no line
            Else
                'Unshaded no line
                    Let ChooseCellStyleForRival = "cl10"        'Unshaded, no line
            End If
        Else
            'Nonfinal column cells are lined if the next stratum.
                'Different stratum from next:
                    If Stratum(MyConstraint) < Stratum(MyConstraint + 1) Then
                        'Shading?
                            If FatalViolationAlreadyFound Then
                                Let ChooseCellStyleForRival = "cl4"         'Lined and shaded
                            Else
                                Let ChooseCellStyleForRival = "cl8"        'Lined and not shaded
                            End If
                    Else
                        'Same stratum as next:
                            'Shading?
                                If FatalViolationAlreadyFound Then
                                    Let ChooseCellStyleForRival = "cl9"         'Shaded, no line
                                Else
                                    Let ChooseCellStyleForRival = "cl10"         'Unshaded, no line
                                End If
                    End If
        End If


End Function

Public Function PointingFinger(FileType As String) As String

    Select Case LCase(FileType)
        Case "doc file", "docfile", "mdocfile"
            Let PointingFinger = "\wsF\we"
        Case "tmp file", "tmpfile", "mtmpfile", "temp file", "tempfile", "mtempfile", "text file", "text"
            Let PointingFinger = ">"
        Case "htm file", "htmfile", "mhtmfile"
            Let PointingFinger = "&#9758; &nbsp;"
        Case Else
            MsgBox ("Programming error:  bad file description for the pointing finger function.  Please contact Bruce Hayes at bhayes@humnet.ucla.edu."), vbCritical
    End Select
End Function
