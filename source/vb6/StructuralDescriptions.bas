Attribute VB_Name = "StructuralDescriptions"
'STRUCTURAL DESCRIPTIONS

    'Translate any structural descriptions that appear in the input file into actual
    '  constraint violations.
        
    'A structural description is a set of "cases", arranged vertically, without gap, under the
    '  constraint name.  Thus for Turkish, a ban on round vowels would require four cases.
    
    'For a Faithfulness constraint, a "case" consists of two lists, for each of the classes
    '  of entities the difference between which defines the constraint.  Thus, for Ident(syllabic),
    '  the two lists would be the vowels and the consonants.
    
    Option Explicit

    'Module-level variables:
    
        'Read in the contents:
            Dim mColumnEntries() As String           'As the first step, store the strings that collectively
                                                    '  count as SD's.
            Dim mNumberOfEntries() As Long           'Number of strings in a column.
            Dim mMaximumNumberOfEntries As Long      'The maximum number of entries in a column.
            
        'Encode constraints:
            Dim mNumberOfConstraints As Long        'Localize these variables.
            Dim mConstraintName() As String
            Dim mAbbrev() As String
            Dim mNumberOfForms As Long
            Dim mInputForm() As String
            Dim mWinner() As String
            Dim mNumberOfRivals() As Long
            Dim mMaximumNumberOfRivals As Long
            Dim mRival() As String
            Public mWinnerViolations() As Long       'This one gets referenced by routine that called this, hence public.
            Public mRivalViolations() As Long        'This one gets referenced by routine that called this, hence public.
            
            Dim mMarkednessConstraint() As String     'A simple list of structural descriptions.
            
            Dim mFaithfulnessConstraint() As String   'Two lists, for the two classes involved in a
                                                     '   Faithfulness constraint, like Max or Dep.
            Dim mFaithfulness() As Boolean           'A simple two-way classification
            Dim mNumberOfFaithfulnessCases() As Long  'The lengths of the two lists.
            Dim mFaithfulnessType() As Long           'Is this a one-way (Max, Dep) or two-way (Ident)
                                                     '    Faithfulness constraint?
                                                     
        'Natural classes
            Dim mNaturalClasses() As String
            Dim mNumberOfMembersInNaturalClass() As Long
            Dim mFileNames() As String
            Dim mNumberOfFileNames As Long
          
Sub Main(NumberOfForms As Long, InputForm() As String, Winner() As String, WinnerViolations() As Long, _
    NumberOfRivals() As Long, Rival() As String, RivalViolations() As Long, _
    NumberOfConstraints As Long, ConstraintName() As String, Abbrev() As String)
       
    'Main routine
    
        Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long, CaseIndex As Long
    
        'Localize the passed parameters as module level variables, suitably dimensioned.
            Let mNumberOfConstraints = NumberOfConstraints
            ReDim mConstraintName(mNumberOfConstraints)
            ReDim mAbbrev(mNumberOfConstraints)
            ReDim mFaithfulness(mNumberOfConstraints)
            For ConstraintIndex = 1 To mNumberOfConstraints
                Let mConstraintName(ConstraintIndex) = ConstraintName(ConstraintIndex)
                Let mAbbrev(ConstraintIndex) = Abbrev(ConstraintIndex)
            Next ConstraintIndex
            Let mNumberOfForms = NumberOfForms
            ReDim mInputForm(mNumberOfForms)
            ReDim mWinner(mNumberOfForms)
            ReDim mWinnerViolations(mNumberOfForms, mNumberOfConstraints)
            ReDim mNumberOfRivals(mNumberOfForms)
            'Populate this array and find its maximum, to help further redimensioning.
                For FormIndex = 1 To mNumberOfForms
                    Let mNumberOfRivals(FormIndex) = NumberOfRivals(FormIndex)
                Next FormIndex
                Let mMaximumNumberOfRivals = Form1.FindMaximumNumberOfRivals(mNumberOfRivals())
            ReDim mRival(mNumberOfForms, mMaximumNumberOfRivals)
            ReDim mRivalViolations(mNumberOfForms, mMaximumNumberOfRivals, mNumberOfConstraints)
            For FormIndex = 1 To mNumberOfForms
                Let mInputForm(FormIndex) = InputForm(FormIndex)
                Let mWinner(FormIndex) = Winner(FormIndex)
                For ConstraintIndex = 1 To mNumberOfConstraints
                    Let mWinnerViolations(FormIndex, ConstraintIndex) = WinnerViolations(FormIndex, ConstraintIndex)
                Next ConstraintIndex
                For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                    Let mRival(FormIndex, RivalIndex) = Rival(FormIndex, RivalIndex)
                    For ConstraintIndex = 1 To mNumberOfConstraints
                        Let mRivalViolations(FormIndex, RivalIndex, ConstraintIndex) = RivalViolations(FormIndex, RivalIndex, ConstraintIndex)
                    Next ConstraintIndex
                Next RivalIndex
            Next FormIndex
        
        'Now that you know how many constraints there are, redimension the arrays that depend on this.
            ReDim mNumberOfEntries(mNumberOfConstraints)
            ReDim mNumberOfFaithfulnessCases(mNumberOfConstraints, 2)
            ReDim mFaithfulnessType(mNumberOfConstraints)
        
        Call TrimAndSortTheRawColumns
        Call OpenNaturalClassFiles
        'Determine which constraints are Faithfulness constraints.
            For ConstraintIndex = 1 To mNumberOfConstraints
                Let mFaithfulness(ConstraintIndex) = Form1.FaithfulnessConstraint(mConstraintName(ConstraintIndex))
            Next ConstraintIndex
        Call ParseTrimmedAndSortedColumnsAsConstraints
        Call CheckForEqualInputAndOutputLength(InputForm(), Winner(), Rival(), NumberOfRivals())
        Call CheckForConstraintErrors
        Call AssignViolations(InputForm(), Winner(), NumberOfRivals())

    Exit Sub
        
'Debug:  Print what you've learned
        
    'Print out a file for debugging.
        Dim D As Long
        Let D = FreeFile
        Open App.Path + "/debug.txt" For Output As #D
   
        For ConstraintIndex = 1 To mNumberOfConstraints
            Select Case mFaithfulness(ConstraintIndex)
                Case False
                    Print #D, "Markedness:  "; ConstraintName(ConstraintIndex)
                    If mNumberOfEntries(ConstraintIndex) = 0 Then
                        Print #D, "      (hand-entered violations)"
                    Else
                        For CaseIndex = 1 To mNumberOfEntries(ConstraintIndex)
                            Print #D, "  "; mMarkednessConstraint(ConstraintIndex, CaseIndex); ", length "; TrueLength(mMarkednessConstraint(ConstraintIndex, CaseIndex))
                        Next CaseIndex
                    End If
                Case True
                    Select Case mFaithfulnessType(ConstraintIndex)
                        Case 0
                            'A hand-entered case.
                            Print #D, "Faithfulness:  "; ConstraintName(ConstraintIndex)
                            Print #D, "      (hand-entered violations)"
                        Case 1
                            Print #D, "One way Faithfulness:  "; ConstraintName(ConstraintIndex)
                            Print #D, "    Input:"
                            For CaseIndex = 1 To mNumberOfFaithfulnessCases(ConstraintIndex, 1)
                                Print #D, "      "; mFaithfulnessConstraint(ConstraintIndex, CaseIndex, 1); ", length "; TrueLength(mFaithfulnessConstraint(ConstraintIndex, CaseIndex, 1))
                            Next CaseIndex
                            Print #D, "    Output:"
                            For CaseIndex = 1 To mNumberOfFaithfulnessCases(ConstraintIndex, 2)
                                Print #D, "      "; mFaithfulnessConstraint(ConstraintIndex, CaseIndex, 2); ", length "; TrueLength(mFaithfulnessConstraint(ConstraintIndex, CaseIndex, 2))
                            Next CaseIndex
                        Case 2
                            Print #D, "Two way Faithfulness:  "; ConstraintName(ConstraintIndex)
                            Print #D, "    Group 1:"
                            For CaseIndex = 1 To mNumberOfFaithfulnessCases(ConstraintIndex, 1)
                                Print #D, "      "; mFaithfulnessConstraint(ConstraintIndex, CaseIndex, 1); ", length "; TrueLength(mFaithfulnessConstraint(ConstraintIndex, CaseIndex, 1))
                            Next CaseIndex
                            Print #D, "    Group 2:"
                            For CaseIndex = 1 To mNumberOfFaithfulnessCases(ConstraintIndex, 2)
                                Print #D, "      "; mFaithfulnessConstraint(ConstraintIndex, CaseIndex, 2); ", length "; TrueLength(mFaithfulnessConstraint(ConstraintIndex, CaseIndex, 2))
                            Next CaseIndex
                    End Select
            End Select
        Next ConstraintIndex
   
        For FormIndex = 1 To mNumberOfForms
            Print #D, InputForm(FormIndex)
            Print #D, "  "; Winner(FormIndex); ":"
            For ConstraintIndex = 1 To NumberOfConstraints
                Print #D, "    "; mAbbrev(ConstraintIndex); Chr(9);
                Print #D, mWinnerViolations(FormIndex, ConstraintIndex)
            Next ConstraintIndex
            For RivalIndex = 1 To NumberOfRivals(FormIndex)
                Print #D, "  "; mRival(FormIndex, RivalIndex); ":"
                For ConstraintIndex = 1 To mNumberOfConstraints
                    Print #D, "    "; mAbbrev(ConstraintIndex); Chr(9);
                    Print #D, mRivalViolations(FormIndex, RivalIndex, ConstraintIndex)
                Next ConstraintIndex
            Next RivalIndex
        Next FormIndex
        Close #D
   
   ' End
   

End Sub

Sub TrimAndSortTheRawColumns()
        
    'Interpret the gRawColumns() array, read as a raw thing from the input file, and interpret it as constraints.
        
        Dim OneColumn() As String
        ReDim OneColumn(gTotalNumberOfRows)
        Dim ConstraintIndex As Long, ColumnIndex As Long, CaseIndex As Long, RowIndex As Long, ExtrasIndex As Long
    
    'Do all the constraints, finding the raw material for structural descriptions.
        For ConstraintIndex = 1 To mNumberOfConstraints
            Let mNumberOfEntries(ConstraintIndex) = 0
            'Put this all in an array, showing just one column.
                For RowIndex = 1 To gTotalNumberOfRows
                    Let OneColumn(RowIndex) = gRawColumns(ConstraintIndex, RowIndex)
                Next RowIndex
            'Interpret the column as violations if it's only numbers, else as structural descriptions.
                If AllNumbers(OneColumn(), gTotalNumberOfRows) = False Then
                    'Look at all the rows in the column.
                    For RowIndex = 1 To gTotalNumberOfRows
                        'If there is a blank, stop reading structural description cases.
                            If gRawColumns(ConstraintIndex, RowIndex) = "" Then
                                'Check to see if there is an intermediate gap; if so, warn the user.
                                    For ExtrasIndex = RowIndex + 1 To gTotalNumberOfRows
                                        If gRawColumns(ColumnIndex, ExtrasIndex) <> "" Then
                                            SendBadNews ("For the constraint " _
                                                + mConstraintName(ConstraintIndex) _
                                                + ", there is a gap in the column of structural descriptions, which OTSoft can't handle.")
                                        End If
                                    Next ExtrasIndex
                                'If no gap, you can stop here.
                                    Exit For
                            Else
                                'Augment the number of cases needed for this constraint.
                                    Let mNumberOfEntries(ConstraintIndex) = mNumberOfEntries(ConstraintIndex) + 1
                                'If this is the largest you've seen, redimension the StructuralDescription() array.
                                    If mNumberOfEntries(ConstraintIndex) > mMaximumNumberOfEntries Then
                                        Let mMaximumNumberOfEntries = mNumberOfEntries(ConstraintIndex)
                                        ReDim Preserve mColumnEntries(mNumberOfConstraints, mMaximumNumberOfEntries)
                                    End If
                                'Install the case for the structural description
                                    Let mColumnEntries(ConstraintIndex, RowIndex) = gRawColumns(ConstraintIndex, RowIndex)
                            End If
                    Next RowIndex       'Go through all the rows for this Raw Column.
                End If                  'Does this Raw Column contain structural descriptions?
        Next ConstraintIndex
    
    
End Sub
    

Function AllNumbers(MyColumn() As String, MyLength As Long) As Boolean

    'Check a column from the gRawColumns() array, and see if it consists solely of
    '   digits.  If so, it's not a constraint, it's user-entered violations.
    
    Dim RowIndex As Long, PositionIndex As Long
    Dim StringLength As Long
    Dim Buffer As String, BufferLength As Long
    
    Let AllNumbers = True
    
    For RowIndex = 1 To MyLength
        Let Buffer = MyColumn(RowIndex)
        If OnlyDigits(Buffer) = False Then
            'A non-digit.  Evidently a structural description, not hand-entered violations.
                Let AllNumbers = False
                Exit Function
        End If
    Next RowIndex
    

End Function


Function OnlyDigits(MyString As String) As Boolean

    'Does a string consist only of digits?
    '   April 2021:  arg, this did not take into account the possibility that someone would want to use negative violations.
    '       So we need to allow a leading negative sign.
    
        Dim MyLength As Long, PositionIndex As Long, LoopStart As Long, Buffer As String
        
        Let Buffer = MyString
        
        'Establish the loop start, thus accounting for initial negative signs.
            If Left(Buffer, 1) = "-" Then
                Let LoopStart = 2
            Else
                Let LoopStart = 1
            End If
        
        'Default value to return, unless falsified.
            Let OnlyDigits = True
        
        'Search for non-digits.
            Let MyLength = Len(MyString)
            For PositionIndex = LoopStart To MyLength
                Select Case Mid(Buffer, PositionIndex, 1)
                    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                        'do nothing
                    Case Else
                        'A non-digit.
                            Let OnlyDigits = False
                            Exit Function
                End Select
            Next PositionIndex

End Function


Function OnlyDigitsAndDecimalPoint(MyString As String) As Boolean

    'Does a string consist only of digits?
        Dim MyLength As Long
        Dim PositionIndex As Long
        Dim NumberOfDecimalPoints As Long
        
        Let NumberOfDecimalPoints = 0
        Let OnlyDigitsAndDecimalPoint = True
        Let MyLength = Len(MyString)
        For PositionIndex = 1 To MyLength
            Select Case Mid(MyString, PositionIndex, 1)
                Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                    'do nothing
                Case "."
                    Let NumberOfDecimalPoints = NumberOfDecimalPoints + 1
                    If NumberOfDecimalPoints > 1 Then
                        Let OnlyDigitsAndDecimalPoint = False
                        Exit Function
                    End If
                Case Else
                    'A non-digit/decimal point
                        Let OnlyDigitsAndDecimalPoint = False
                        Exit Function
            End Select
        Next PositionIndex

End Function

Sub ParseTrimmedAndSortedColumnsAsConstraints()

    'Take the raw contents of the file and turn the result into Markedness and Faithfulness
    '  constraints.
    
    Dim ConstraintIndex As Long, ColumnIndex As Long, EntryIndex As Long
    
    Dim WhatImCollectingNow As Long     'Receptivity state, for parsing Faithfulness constraints.
    Let WhatImCollectingNow = 0
    Dim CurrentCell As String           'The cell being processed.
    
    Const NothingYet As Long = 0        'The values of the receptivity state.
    Const Group1 As Long = 1
    Const Group2 As Long = 2
    Const Inputs As Long = 3
    Const Outputs As Long = 4
    
    'First redimension the arrays that will contain the constraints.
        ReDim mMarkednessConstraint(mNumberOfConstraints, mMaximumNumberOfEntries)
        ReDim mFaithfulnessConstraint(mNumberOfConstraints, mMaximumNumberOfEntries, 2)
        ReDim mFaithfulnessConstraint(mNumberOfConstraints, mMaximumNumberOfEntries, 2)
    
    'Interpret all constraints:
    
    For ConstraintIndex = 1 To mNumberOfConstraints
    
        'Markedness constraints are easy:  the column contents are simply the subcases of the
        '  structural description.
        
            If mFaithfulness(ConstraintIndex) = False Then
                For EntryIndex = 1 To mNumberOfEntries(ConstraintIndex)
                    Let mMarkednessConstraint(ConstraintIndex, EntryIndex) = mColumnEntries(ConstraintIndex, EntryIndex)
                Next EntryIndex
            
            Else
        
        'Faithfulness constraints require further interpretation of list.
        
            'Two kinds:
            '   Group 1, group 2:   symmetrical Faithfulness, like Ident
            '   Input, Output:      asymmetrical Faithfulness, like Max and Dep
            
        'Use flags to remember what you're collecting.
        
            For EntryIndex = 1 To mNumberOfEntries(ConstraintIndex)
                Let CurrentCell = mColumnEntries(ConstraintIndex, EntryIndex)
                Select Case LCase(CurrentCell)
                    'If you hit one of the official flags, change your receptivity state.
                    Case "group1", "group 1"
                        Let WhatImCollectingNow = Group1
                    Case "group2", "group 2"
                        Let WhatImCollectingNow = Group2
                    Case "input"
                        Let WhatImCollectingNow = Inputs
                    Case "output"
                        Let WhatImCollectingNow = Outputs
                    Case Else
                        'According to your receptivity state, add to the appropriate
                        '  structural description file.
                        Select Case WhatImCollectingNow
                            Case Group1
                                Let mNumberOfFaithfulnessCases(ConstraintIndex, 1) = mNumberOfFaithfulnessCases(ConstraintIndex, 1) + 1
                                Let mFaithfulnessConstraint(ConstraintIndex, mNumberOfFaithfulnessCases(ConstraintIndex, 1), 1) = CurrentCell
                                Let mFaithfulnessType(ConstraintIndex) = 2
                            Case Group2
                                Let mNumberOfFaithfulnessCases(ConstraintIndex, 2) = mNumberOfFaithfulnessCases(ConstraintIndex, 2) + 1
                                Let mFaithfulnessConstraint(ConstraintIndex, mNumberOfFaithfulnessCases(ConstraintIndex, 2), 2) = CurrentCell
                                'Needed for reporting errors:
                                Let mFaithfulnessType(ConstraintIndex) = 2
                            Case Inputs
                                Let mNumberOfFaithfulnessCases(ConstraintIndex, 1) = mNumberOfFaithfulnessCases(ConstraintIndex, 1) + 1
                                Let mFaithfulnessConstraint(ConstraintIndex, mNumberOfFaithfulnessCases(ConstraintIndex, 1), 1) = CurrentCell
                                Let mFaithfulnessType(ConstraintIndex) = 1
                            Case Outputs
                                Let mNumberOfFaithfulnessCases(ConstraintIndex, 2) = mNumberOfFaithfulnessCases(ConstraintIndex, 2) + 1
                                Let mFaithfulnessConstraint(ConstraintIndex, mNumberOfFaithfulnessCases(ConstraintIndex, 2), 2) = CurrentCell
                                'Needed for reporting errors:
                                Let mFaithfulnessType(ConstraintIndex) = 1
                            Case Else
                                'You've read a structural description with no tag to say what it is.
                                    MsgBox "Sorry, but I've found a problem in your input file. " + Chr(10) + Chr(10) _
                                    + "You've included in the structural description file an entry for a Faithfulness constraint.  In all such entries, each list of violation strings must be preceded by either Group1, Group2, Input, or Output.  But for the constraint " _
                                    + mAbbrev(ConstraintIndex) _
                                    + ", the string " + Chr(10) + Chr(10) _
                                    + mColumnEntries(ConstraintIndex, EntryIndex) + Chr(10) + Chr(10) _
                                    + " is not preceded by any of these four. " + Chr(10) + Chr(10) _
                                    + "Please exit OTSoft, and change your input file.  Use the null symbol < to mark locations of insertion and deletion.  Then start OTSoft and try again.", vbExclamation
                                    Call Form1.cmdExit_Click
                        End Select
                End Select
            Next EntryIndex
            
            End If                  'Markedness or Faithfulness?
            
    Next ConstraintIndex
    

End Sub


Sub CheckForEqualInputAndOutputLength(InputForm() As String, Winner() As String, Rival() As String, NumberOfRivals() As Long)

    'The system for assessing Faithfulness violations assumes that inputs and outputs all have
    '  the same length.  For deletion and insertion, a null symbol must be inserted.
    
    'If there is going to be auto-calculation of Faithfulness, then all input and output strings
    '  must have the same length.
    
        Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
   
    'First, check if any Faithfulness constraints are auto-calculated.  If not, you don't have
    '  to impose this requirement.
        For ConstraintIndex = 1 To mNumberOfConstraints
            If mFaithfulness(ConstraintIndex) = True Then
                If mNumberOfEntries(ConstraintIndex) > 0 Then
                    GoTo ExitPoint
                End If
            End If
        Next ConstraintIndex
        'Didn't find any, so exit this routine.
            Exit Sub

ExitPoint:
                                    
    'There is at least one Faithfulness constraint that is going to be autocalculated, so check
    '  all input and output strings for identical length.
    
        For FormIndex = 1 To mNumberOfForms
            If Len(Winner(FormIndex)) <> Len(InputForm(FormIndex)) Then
                MsgBox "Sorry, but I've found a problem in your input file. " + Chr(10) + Chr(10) _
                    + "You've included a structural description file that assesses Faithfulness violations.  For this to work, all candidates must have the same length as the input form.  " + _
                    "But the following input form:  " + Chr(10) + Chr(10) _
                    + InputForm(FormIndex) + Chr(10) + Chr(10) _
                    + "is not the same length as its surface form:  " + Chr(10) + Chr(10) _
                    + Winner(FormIndex) + Chr(10) + Chr(10) + _
                    "Please exit OTSoft, and change your input file.  Use the null symbol < to mark locations of insertion and deletion.  Then start OTSoft and try again.", vbExclamation
                    Call Form1.cmdExit_Click
            End If
            For RivalIndex = 1 To NumberOfRivals(FormIndex)
                If Len(Rival(FormIndex, RivalIndex)) <> Len(InputForm(FormIndex)) Then
                    MsgBox "Sorry, but I've found a problem in your input file. " + Chr(10) + Chr(10) _
                        + "You've included a structural description file that assesses Faithfulness violations.  For this to work, all candidates must have the same length as the input form.  " _
                        + "But the following rival candidate:  " + Chr(10) + Chr(10) _
                        + "    " + Rival(FormIndex, RivalIndex) + Chr(10) + Chr(10) _
                        + "is not the same length as its input form:  " + Chr(10) + Chr(10) _
                        + "    " + InputForm(FormIndex) + Chr(10) + Chr(10) + Chr(10) _
                        + "Please exit OTSoft, and change your input file.  Use the null symbol < to mark locations of insertion and deletion.  Then start OTSoft and try again.", vbExclamation
                        Call Form1.cmdExit_Click
                End If
            Next RivalIndex
        Next FormIndex
    

End Sub


Sub CheckForConstraintErrors()

    Dim TargetLength As Long        'For comparing lengths of structural description cases.
    Dim ConstraintIndex As Long
    Dim CaseIndex As Long, InnerCaseIndex As Long, SideIndex As Long
    
    'I. A Faithfulness constraint is bad if it lacks one side or the other.
        For ConstraintIndex = 1 To mNumberOfConstraints
            If mFaithfulness(ConstraintIndex) = True Then
                If mNumberOfFaithfulnessCases(ConstraintIndex, 1) = 0 Then
                    'If both are zero, then it's being done by hand, so no worry.
                        If mNumberOfFaithfulnessCases(ConstraintIndex, 2) > 0 Then
                            Select Case mFaithfulnessType(ConstraintIndex)
                                Case 1
                                    SendBadNews ("You've got a Faithfulness constraint, " + _
                                        mConstraintName(ConstraintIndex) + _
                                        ", that lists structural descriptions only for the output, not the input.")
                                Case 2
                                    SendBadNews ("You've got a Faithfulness constraint, " + _
                                        mConstraintName(ConstraintIndex) + _
                                        ", that lists structural descriptions only for the second, not the first group of sounds involved.")
                            End Select
                        End If      'Any cases for the first position?
                Else
                    'No cases for the *second* position:
                        If mNumberOfFaithfulnessCases(ConstraintIndex, 2) = 0 Then
                            Select Case mFaithfulnessType(ConstraintIndex)
                                Case 1
                                    SendBadNews ("You've got a Faithfulness constraint, " + _
                                        mConstraintName(ConstraintIndex) + _
                                        ", that lists structural descriptions only for the input, not the output.")
                                Case 2
                                    SendBadNews ("You've got a Faithfulness constraint, " + _
                                        mConstraintName(ConstraintIndex) + _
                                        ", that lists structural descriptions only for the first, not the second group of sounds involved.")
                            End Select
                        End If          'Any cases for the second position?
                End If              'Any cases for the first position?
            End If                  'Faithfulness constraint?
        Next ConstraintIndex        'Examine all constraints.
                    
    
    'A Faithfulness constraint is bad if any of its strings are of unequal length.
        For ConstraintIndex = 1 To mNumberOfConstraints
            If mFaithfulness(ConstraintIndex) = True Then
                'Only do this for Faithfulness constraints that have structural descriptions
                    If mNumberOfEntries(ConstraintIndex) > 0 Then
                    'Use the first one as a measure for all the others
                        Let TargetLength = TrueLength(mFaithfulnessConstraint(ConstraintIndex, 1, 1))
                        For SideIndex = 1 To 2
                            For CaseIndex = 1 To mNumberOfFaithfulnessCases(ConstraintIndex, SideIndex)
                                If TrueLength(mFaithfulnessConstraint(ConstraintIndex, CaseIndex, SideIndex)) <> TargetLength Then
                                    SendBadNews ("The strings used in formalizing Faithfulness constraints must all be the same length. " + _
                                        "But for the constraint" + Chr(10) + Chr(10) + _
                                        "   " + mConstraintName(ConstraintIndex) + Chr(10) + Chr(10) + _
                                        "I've found a difference in length between the following two strings:" + Chr(10) + Chr(10) + _
                                        "   " + mFaithfulnessConstraint(ConstraintIndex, 1, 1) + _
                                        "   " + mFaithfulnessConstraint(ConstraintIndex, CaseIndex, SideIndex))
                                End If
                            Next CaseIndex
                        Next SideIndex
                    End If                  'Does this have a structural description?
            End If                          'Is this a Faithfulness constraint?
        Next ConstraintIndex                'Go through all the constraints.
        
        
    'A constraint is bad if any of its strings are identical.
        For ConstraintIndex = 1 To mNumberOfConstraints
            If mFaithfulness(ConstraintIndex) = True Then
                For SideIndex = 1 To 2
                    For CaseIndex = 1 To mNumberOfFaithfulnessCases(ConstraintIndex, SideIndex) - 1
                        For InnerCaseIndex = CaseIndex + 1 To mNumberOfFaithfulnessCases(ConstraintIndex, SideIndex)
                            If mFaithfulnessConstraint(ConstraintIndex, CaseIndex, SideIndex) = mFaithfulnessConstraint(ConstraintIndex, InnerCaseIndex, SideIndex) Then
                                SendBadNews ("The constraint" + Chr(10) + Chr(10) + _
                                    "   " + mConstraintName(ConstraintIndex) + Chr(10) + Chr(10) + _
                                    "includes a repeated item in its list of cases.  This will cause single violations to be counted multiple times." + Chr(10) + Chr(10) + _
                                    "The repeated item is:" + Chr(10) + Chr(10) + _
                                    "   " + mFaithfulnessConstraint(ConstraintIndex, CaseIndex, SideIndex))
                            End If
                        Next InnerCaseIndex     'All possible matches
                    Next CaseIndex              'All case of one side
                Next SideIndex                  'Do both sides of the constraint.
            Else                                'Markedness constraints likewise better not have duplicates.
                For CaseIndex = 1 To mNumberOfEntries(ConstraintIndex) - 1
                    For InnerCaseIndex = CaseIndex + 1 To mNumberOfEntries(ConstraintIndex)
                        If mMarkednessConstraint(ConstraintIndex, CaseIndex) = mMarkednessConstraint(ConstraintIndex, InnerCaseIndex) Then
                            SendBadNews ("The constraint" + Chr(10) + Chr(10) + _
                                "   " + mConstraintName(ConstraintIndex) + Chr(10) + Chr(10) + _
                                "includes a repeated item in its list of cases.  This will cause single violations to be counted multiple times." + Chr(10) + Chr(10) + _
                                "The repeated item is:" + Chr(10) + Chr(10) + _
                                "   " + mMarkednessConstraint(ConstraintIndex, CaseIndex))
                        End If
                    Next InnerCaseIndex         'All possible matches
                Next CaseIndex                  'All cases
            End If                              'Is this a Faithfulness constraint?
        Next ConstraintIndex                    'Go through all the constraints.
            
    'A Faithfulness constraint is bad if it penalizes identical strings.
        For ConstraintIndex = 1 To mNumberOfConstraints
            If mFaithfulness(ConstraintIndex) = True Then
                For CaseIndex = 1 To mNumberOfFaithfulnessCases(ConstraintIndex, 1)
                    For InnerCaseIndex = 1 To mNumberOfFaithfulnessCases(ConstraintIndex, 2)
                        If mFaithfulnessConstraint(ConstraintIndex, CaseIndex, 1) = mFaithfulnessConstraint(ConstraintIndex, InnerCaseIndex, 2) Then
                            SendBadNews ("The Faithfulness constraint" + Chr(10) + Chr(10) + _
                                "   " + mConstraintName(ConstraintIndex) + Chr(10) + Chr(10) + _
                                "includes an identical item in both of its lists.  For this item, it therefore will penalize Faithfulness." + Chr(10) + Chr(10) + _
                                "(If you truly want to use a constraint that does this, please enter the violations by hand in the input file.  OTSoft is designed to treat such cases as errors, in order to protect the user.)" + Chr(10) + Chr(10) + _
                                "The repeated item is:" + Chr(10) + Chr(10) + _
                                "   " + mFaithfulnessConstraint(ConstraintIndex, CaseIndex, 1))
                        End If
                    Next InnerCaseIndex     'All cases on one side
                Next CaseIndex              'All cases on the other side
            End If                          'Faithfulness constraints only.
        Next ConstraintIndex                'Go through all the constraints.

End Sub         'CheckForConstraintErrors()

Sub AssignViolations(InputForm() As String, Winner() As String, NumberOfRivals() As Long)

    'Assign violations to all the constraints.
    
        Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long, CaseIndex As Long
        
        'Arrays to hold backups for winners and rivals, so Markedness violations can be assessed
        '  without the zero symbol
            Dim BackupInput() As String
            ReDim BackupInput(mNumberOfForms)
            Dim BackupWinner() As String
            ReDim BackupWinner(mNumberOfForms)
            Dim BackupRival() As String
            ReDim BackupRival(mNumberOfForms, mMaximumNumberOfRivals)
        
    'At various points, it is useful to alter the data themselves.  At the end, we will wish
    '  to restore the original forms.  So make backups.
    
        For FormIndex = 1 To mNumberOfForms
            Let BackupInput(FormIndex) = InputForm(FormIndex)
            Let BackupWinner(FormIndex) = Winner(FormIndex)
            For RivalIndex = 1 To NumberOfRivals(FormIndex)
                Let BackupRival(FormIndex, RivalIndex) = mRival(FormIndex, RivalIndex)
            Next RivalIndex
        Next FormIndex
    
    'To allow reference to initial and final position, prepose and postpose brackets to the forms.
        For FormIndex = 1 To mNumberOfForms
            Let InputForm(FormIndex) = "[" + InputForm(FormIndex) + "]"
            Let Winner(FormIndex) = "[" + Winner(FormIndex) + "]"
            For RivalIndex = 1 To NumberOfRivals(FormIndex)
                Let mRival(FormIndex, RivalIndex) = "[" + mRival(FormIndex, RivalIndex) + "]"
            Next RivalIndex
        Next FormIndex
    
    
    'Assign the Faithfulness violations:
    
        For ConstraintIndex = 1 To mNumberOfConstraints
            If mFaithfulness(ConstraintIndex) = True Then
                'Only do this if there are cases for this constraint; else let the hand-entered values survive.
                If mNumberOfEntries(ConstraintIndex) > 0 Then
                    'Initialize, since you're going to override.
                        For FormIndex = 1 To mNumberOfForms
                            Let mWinnerViolations(FormIndex, ConstraintIndex) = 0
                            For RivalIndex = 1 To NumberOfRivals(FormIndex)
                                Let mRivalViolations(FormIndex, RivalIndex, ConstraintIndex) = 0
                            Next RivalIndex
                        Next FormIndex
                    'Now, find the new values, for Faithfulness constraints:
                        For FormIndex = 1 To mNumberOfForms
                            'There are two kinds:  one-way and two-way:
                            If mFaithfulnessType(ConstraintIndex) = 1 Then
                                Let mWinnerViolations(FormIndex, ConstraintIndex) = AssessOneWayFaithfulness(InputForm(FormIndex), Winner(FormIndex), ConstraintIndex)
                                For RivalIndex = 1 To NumberOfRivals(FormIndex)
                                    Let mRivalViolations(FormIndex, RivalIndex, ConstraintIndex) = AssessOneWayFaithfulness(InputForm(FormIndex), mRival(FormIndex, RivalIndex), ConstraintIndex)
                                Next RivalIndex
                            Else
                                Let mWinnerViolations(FormIndex, ConstraintIndex) = AssessTwoWayFaithfulness(InputForm(FormIndex), Winner(FormIndex), ConstraintIndex)
                                For RivalIndex = 1 To NumberOfRivals(FormIndex)
                                    Let mRivalViolations(FormIndex, RivalIndex, ConstraintIndex) = AssessTwoWayFaithfulness(InputForm(FormIndex), mRival(FormIndex, RivalIndex), ConstraintIndex)
                                Next RivalIndex
                            End If      'One way vs. two-way Faithfulness.
                        Next FormIndex  'Go through all the forms
                End If                  'Does the constraint have a structural description entry?
            End If                      'Is this a Faithfulness constraint?
        Next ConstraintIndex            'Examine every constraint.
        
    'The Markedness constraints should not see the zero symbol, so delete it.
    '   This was a terrible idea because it screws up the backup file.
    '    For FormIndex = 1 To mNumberOfForms
    '        Let InputForm(FormIndex) = ZeroTrim(InputForm(FormIndex))
    '        Let Winner(FormIndex) = ZeroTrim(Winner(FormIndex))
    '        For RivalIndex = 1 To NumberOfRivals(FormIndex)
    '            Let mRival(FormIndex, RivalIndex) = ZeroTrim(mRival(FormIndex, RivalIndex))
   '         Next RivalIndex
    '    Next FormIndex
    
    'Assign Markedness violations:
    
        For ConstraintIndex = 1 To mNumberOfConstraints
            If mFaithfulness(ConstraintIndex) = False Then
                'Only do this if there are cases for this constraint; else let the hand-entered values survive.
                If mNumberOfEntries(ConstraintIndex) > 0 Then
                    'Initialize, since you're going to override.
                        For FormIndex = 1 To mNumberOfForms
                            Let mWinnerViolations(FormIndex, ConstraintIndex) = 0
                            For RivalIndex = 1 To NumberOfRivals(FormIndex)
                                Let mRivalViolations(FormIndex, RivalIndex, ConstraintIndex) = 0
                            Next RivalIndex
                        Next FormIndex
                    'Now, find the new values, for Markedness constraints.
                    '   Note that since this is Markedness, we trim out the fictional zero symbols.
                        For FormIndex = 1 To mNumberOfForms
                            Let mWinnerViolations(FormIndex, ConstraintIndex) = AssessMarkedness(ZeroTrim(Winner(FormIndex)), ConstraintIndex, mNumberOfEntries(ConstraintIndex))   'mWinnerViolations(FormIndex, ConstraintIndex) + AssessMarkedness(Winner(FormIndex), ConstraintIndex, mNumberOfEntries(ConstraintIndex))
                            For RivalIndex = 1 To NumberOfRivals(FormIndex)
                                Let mRivalViolations(FormIndex, RivalIndex, ConstraintIndex) = AssessMarkedness(ZeroTrim(mRival(FormIndex, RivalIndex)), ConstraintIndex, mNumberOfEntries(ConstraintIndex))   'mrivalviolations(FormIndex, RivalIndex, ConstraintIndex) + AssessMarkedness(Rival(FormIndex, RivalIndex), ConstraintIndex, mNumberOfEntries(ConstraintIndex))
                            Next RivalIndex
                        Next FormIndex
                End If              'Does the constraint have a structural description entry?
            End If                  'Is this a Markedness constraint?
        Next ConstraintIndex
        
        'Take away the brackets you earlier assigned.
            For FormIndex = 1 To mNumberOfForms
                Let InputForm(FormIndex) = Mid(InputForm(FormIndex), 2, Len(InputForm(FormIndex)) - 2)
                Let Winner(FormIndex) = Mid(Winner(FormIndex), 2, Len(Winner(FormIndex)) - 2)
                For RivalIndex = 1 To NumberOfRivals(FormIndex)
                    Let mRival(FormIndex, RivalIndex) = Mid(mRival(FormIndex, RivalIndex), 2, Len(mRival(FormIndex, RivalIndex)) - 2)
                Next RivalIndex
            Next FormIndex
            

        'Now that the Markedness violations are assigned, you can restore the original values
        '  of the winner and rivals from their backups.
        
        '    For FormIndex = 1 To mNumberOfForms
        '        Let InputForm(FormIndex) = BackupInput(FormIndex)
        '        Let Winner(FormIndex) = BackupWinner(FormIndex)
        '        For RivalIndex = 1 To NumberOfRivals(FormIndex)
        '            Let Rival(FormIndex, RivalIndex) = BackupRival(FormIndex, RivalIndex)
        '        Next RivalIndex
        '    Next FormIndex
                       

End Sub

Function ZeroTrim(MyString As String) As String

    'Extract the zero symbols out of any input string.
    
        Dim PositionIndex As Long
        Dim MyStringLength As Long
        Dim CurrentCharacter As String, Buffer As String
        
        Let Buffer = ""
        Let MyStringLength = Len(MyString)
        For PositionIndex = 1 To MyStringLength
            Let CurrentCharacter = Mid(MyString, PositionIndex, 1)
            Select Case CurrentCharacter
                Case "<"
                    'do nothing.
                Case Else
                    Let Buffer = Buffer + CurrentCharacter
            End Select
        Next PositionIndex
        
        Let ZeroTrim = Buffer
        
        
End Function

Function AssessMarkedness(MyForm As String, MyConstraint As Long, NumberOfCases As Long) As Long
    
    'Given a string and a set of constraint subcases, assess how many violations
    
    Dim MyCase As String        'The subcase of the constraint being applied.
    Dim FormLength As Long      'The length of the string being assessed.
    Dim CaseLength As Long      'The length of each case of violation.
    Dim UpperLimit As Long      'How far to search.
    Dim CaseIndex As Long, PositionIndex As Long
    Dim Counter As Long         'Buffer, for storing how many violations you've found.
    Dim Buffer As String        'The safely-alterable version of your candidate.
    
    'Take away the < symbol, which is not there for markedness.
        Let Buffer = ZeroTrim(MyForm)
    
    Let FormLength = Len(Buffer)
    For CaseIndex = 1 To NumberOfCases
        Let MyCase = mMarkednessConstraint(MyConstraint, CaseIndex)
        Let CaseLength = TrueLength(MyCase)
        'If FormLength >= CaseLength Then Stop
        Let UpperLimit = FormLength - CaseLength + 1
        For PositionIndex = 1 To UpperLimit
            'Is there a match in this position?
                If Match(MyCase, Mid(Buffer, PositionIndex, CaseLength)) = True Then
                    Let Counter = Counter + 1
                End If
        Next PositionIndex
    Next CaseIndex
    
    Let AssessMarkedness = Counter
    
End Function
    

Function AssessOneWayFaithfulness(MyInput As String, MyOutput As String, MyConstraint As Long) As Long
    
    'Given a string and a set of constraint subcases, assess how many violations
    
    Dim MyCaseInput As String   'The subcase of the constraint being applied:  input side.
    Dim MyCaseOutput As String  'The subcase of the constraint being applied:  output side.
    Dim Counter As Long         'Buffer, for storing how many violations you've found.
    
    Dim FormLength As Long      'The length of the string being assessed.
    Dim CaseLength As Long      'The length of each case of violation.
    Dim UpperLimit As Long      'How far to search.
    
    Dim OuterCaseIndex As Long, InnerCaseIndex As Long
    Dim PositionIndex As Long
    
    Let Counter = 0
    Let FormLength = Len(MyInput)
    
    For OuterCaseIndex = 1 To mNumberOfFaithfulnessCases(MyConstraint, 1)
        'Put into a convenient local variable the SD-case for the input side you're considering.
            Let MyCaseInput = mFaithfulnessConstraint(MyConstraint, OuterCaseIndex, 1)
        'Decide how far you have to search.
            Let CaseLength = TrueLength(MyCaseInput)
            Let UpperLimit = FormLength - CaseLength + 1
        'Go through all possible positions, looking for a match in the input.
        For PositionIndex = 1 To UpperLimit
            'Is there a match in this position?
            If Match(MyCaseInput, Mid(MyInput, PositionIndex, CaseLength)) = True Then
                'If so, check if there is a match in the corresponding output location.
                For InnerCaseIndex = 1 To mNumberOfFaithfulnessCases(MyConstraint, 2)
                    'Put into a convenient local variable the SD-case for the output side
                    '   you're considering.
                        Let MyCaseOutput = mFaithfulnessConstraint(MyConstraint, InnerCaseIndex, 2)
                    'Is there a match in the same position where there was a match in the input?
                        If Match(MyCaseOutput, Mid(MyOutput, PositionIndex, CaseLength)) = True Then
                        'Assess a violation
                            Let Counter = Counter + 1
                        'Else
                        End If              'Is there a match for the output side?
                Next InnerCaseIndex     'Examine all cases for the output side.
            End If                      'Is there a match for the input side?
        Next PositionIndex              'Examine the whole string.
    Next OuterCaseIndex                      'Examine all cases for the input side of the SD.
    
    Let AssessOneWayFaithfulness = Counter
    
End Function
    
    

Function AssessTwoWayFaithfulness(MyInput As String, MyOutput As String, MyConstraint As Long) As Long
    
    'Given a string and a set of constraint subcases, assess how many violations
    
    Dim MyCaseInput As String   'The subcase of the constraint being applied:  input side.
    Dim MyCaseOutput As String  'The subcase of the constraint being applied:  output side.
    Dim Counter As Long         'Buffer, for storing how many violations you've found.
    
    Dim FormLength As Long      'The length of the string being assessed.
    Dim CaseLength As Long      'The length of each case of violation.
    Dim UpperLimit As Long      'How far to search.
    
    Dim OuterCaseIndex As Long, InnerCaseIndex As Long, PositionIndex As Long
    
    Let Counter = 0
    Let FormLength = Len(MyInput)
    
    'One direction:
        For OuterCaseIndex = 1 To mNumberOfFaithfulnessCases(MyConstraint, 1)
            'Put into a convenient local variable the SD-case for the input side you're considering.
                Let MyCaseInput = mFaithfulnessConstraint(MyConstraint, OuterCaseIndex, 1)
            'Decide how far you have to search.
                Let CaseLength = TrueLength(MyCaseInput)
                Let UpperLimit = FormLength - CaseLength + 1
            'Go through all possible positions, looking for a match in the input.
            For PositionIndex = 1 To UpperLimit
                'Is there a match in this position?
                If (Match(MyCaseInput, Mid(MyInput, PositionIndex, CaseLength))) = True Then
                    'If so, check if there is a match in the corresponding output location.
                    For InnerCaseIndex = 1 To mNumberOfFaithfulnessCases(MyConstraint, 2)
                        'Is there a match in the same position where there was a match in the input?
                        If Match(mFaithfulnessConstraint(MyConstraint, InnerCaseIndex, 2), Mid(MyOutput, PositionIndex, CaseLength)) = True Then
                            'Assess a violation
                                Let Counter = Counter + 1
                        End If              'Is there a match for the output side?
                    Next InnerCaseIndex     'Examine all cases for the output side.
                End If                      'Is there a match for the input side?
            Next PositionIndex              'Examine the whole string.
        Next OuterCaseIndex                 'Examine all cases for the input side of the SD.
        
    'The other direction (same code, direction reversed):

        For OuterCaseIndex = 1 To mNumberOfFaithfulnessCases(MyConstraint, 1)
            'Put into a convenient local variable the SD-case for the input side you're considering.
                Let MyCaseInput = mFaithfulnessConstraint(MyConstraint, OuterCaseIndex, 1)
            'Decide how far you have to search.
                Let CaseLength = TrueLength(MyCaseInput)
                Let UpperLimit = FormLength - CaseLength + 1
            'Go through all possible positions, looking for a match in the input.
            For PositionIndex = 1 To UpperLimit
                'Is there a match in this position?
                If Match(MyCaseInput, Mid(MyOutput, PositionIndex, CaseLength)) = True Then
                    'If so, check if there is a match in the corresponding output location.
                    For InnerCaseIndex = 1 To mNumberOfFaithfulnessCases(MyConstraint, 2)
                        'Is there a match in the same position where there was a match in the input?
                        If Match(mFaithfulnessConstraint(MyConstraint, InnerCaseIndex, 2), Mid(MyInput, PositionIndex, CaseLength)) = True Then
                            'Assess a violation
                                Let Counter = Counter + 1
                        End If              'Is there a match for the output side?
                    Next InnerCaseIndex     'Examine all cases for the output side.
                End If                      'Is there a match for the input side?
            Next PositionIndex              'Examine the whole string.
        Next OuterCaseIndex                 'Examine all cases for the input side of the SD.
    
    Let AssessTwoWayFaithfulness = Counter
    
    
End Function
    
    
Sub SendBadNews(BadNews As String)

    'Post a message box explaining what's wrong with the input file, then quit.
    
    MsgBox "Sorry, but I've found a problem in your input file. " + Chr(10) + Chr(10) _
    + BadNews + Chr(10) + Chr(10) _
    + "Please exit OTSoft and make the needed changes in your input file.  Then restart OTSoft and try again.", vbCritical
    Call Form1.cmdExit_Click

End Sub

Sub OpenNaturalClassFiles()

    'There are a bunch of text files titled vo.txt (vowels), co.txt (consonants) etc.
    '   Look for their titles in the structural descriptions,
    '       open the files,
    '       and put them in a big array, mNaturalClasses(),
    '       recording how many members each natural class has in the array
    '       mNumberOfMembersInNaturalClass()
    
    Dim ConstraintIndex As Long
    Dim RowIndex As Long, PositionIndex As Long
    Dim FileNameIndex As Long
    Dim MyString As String
    Dim MyLength As Long
    Dim CandidateFileName As String
    Dim NatClassFile As Integer
    Dim FileLocation As String
    Dim Buffer As String
    Dim CurrentColumn As Long
    Dim BiggestNaturalClassSize As Long
    
    'Go through the raw columns, looking for material.
        
        For ConstraintIndex = 1 To mNumberOfConstraints
            For RowIndex = 1 To mNumberOfEntries(ConstraintIndex)
                Let MyString = gRawColumns(ConstraintIndex, RowIndex)
                Let MyLength = Len(MyString)
                For PositionIndex = 1 To MyLength
                    If Mid(MyString, PositionIndex, 1) = "\" Then
                    'Guard against \ without two following letters.
                        If PositionIndex >= MyLength - 1 Then
                            Call SendBadNews("Back slashes in input files must precede three-letter sequences designating file names for natural class files." + Chr(10) + Chr(10) + _
                                "But in your input file, this is not so for the entry:" + Chr(10) + Chr(10) + _
                                "    " + MyString)
                        Else
                            Let CandidateFileName = Mid(MyString, PositionIndex + 1, 3)
                            'Check for uniqueness.
                                For FileNameIndex = 1 To mNumberOfFileNames
                                    If mFileNames(FileNameIndex) = CandidateFileName Then GoTo ExitPoint
                                Next FileNameIndex
                                'It's unique; so record it.
                                    Let mNumberOfFileNames = mNumberOfFileNames + 1
                                    ReDim Preserve mFileNames(mNumberOfFileNames)
                                    Let mFileNames(mNumberOfFileNames) = CandidateFileName
ExitPoint:
                        End If
                    End If
                Next PositionIndex
            Next RowIndex       'Go through all the rows for this Raw Column.
        Next ConstraintIndex
        
    'Redim the array that will hold natural classes, to be the right size.
        '255 is the number of distinct symbols in ASCII, and is thus big enough.
            ReDim mNaturalClasses(255, mNumberOfFileNames)
            ReDim mNumberOfMembersInNaturalClass(mNumberOfFileNames)

    'Open all of the files found, and record their contents into the big coded array.
    
        For FileNameIndex = 1 To mNumberOfFileNames
            Let FileLocation = App.Path + "\NatClasses\" + mFileNames(FileNameIndex) + ".txt"
            'Check that this exists.
                If Dir(FileLocation) = "" Then
                    MsgBox "Sorry, but I've found a problem.  Your input file requires that the directory " + Chr(10) + Chr(10) + _
                        "    " + App.Path + "\NatClasses\" + Chr(10) + Chr(10) + _
                        "contain a natural-class file (list of sounds) named" + Chr(10) + Chr(10) + _
                        "    " + mFileNames(FileNameIndex) + ".txt" + Chr(10) + Chr(10) + _
                        "But I can't find this file.  Please install or create this file, or change your input file, before proceeding further.", vbCritical
                        Call Form1.cmdExit_Click
                End If
            'You're ok, so open it and load it into the big array.
                Let NatClassFile = FreeFile
                Open FileLocation For Input As #NatClassFile
                Do
                    If EOF(NatClassFile) Then Exit Do
                    Line Input #NatClassFile, Buffer
                    Let mNumberOfMembersInNaturalClass(FileNameIndex) = mNumberOfMembersInNaturalClass(FileNameIndex) + 1
                    Let mNaturalClasses(mNumberOfMembersInNaturalClass(FileNameIndex), FileNameIndex) = Buffer
                Loop
                Close #NatClassFile
        Next FileNameIndex
        
        'Debug:
        'Dim D As Long
        'Let D = FreeFile
        'Open "Debug.txt" For Output As #D
        'Dim i As Long
        'Dim j As Long
        'For i = 1 To mNumberOfFileNames
        '    For j = 1 To mNumberOfMembersInNaturalClass(i)
        '        Print #D, mNaturalClasses(j, i); Chr(9);
        '    Next j
        '    Print #D,
        'Next i
        'End

End Sub

Function NumericalCoding(MyString As String) As Long

    'Take a three-letter code for natural classes and locate its place in the array
    '   of that encodes natural classes.
    '   for storing the natural classes into a reasonable size array.
    
        Dim NaturalClassIndex As Long
        For NaturalClassIndex = 1 To mNumberOfFileNames
            If MyString = mFileNames(NaturalClassIndex) Then
                Let NumericalCoding = NaturalClassIndex
                Exit Function
            End If
        Next NaturalClassIndex
        
        'If you've gotten this far, there's a bug.
            MsgBox "Bug in program, can't find a natural class filename.  Contact bhayes@humnet.ucla.edu", vbCritical
            End
        
End Function

Function Match(SDString As String, TargetString As String) As Boolean

    'Take a structural description string that may include some references to natural
    '  classes, and determine if it can be matched to a linguistic string.

    Dim SDLength As Long
    Dim TargetLength As Long
    Dim SDPositionIndex As Long
    Dim TargetPositionIndex As Long
    Dim ListIndex As Long
    Dim NaturalClassIndex As Long
    
    'The TargetPositionIndex will sometimes lag behind the SDPosition index, since the latter
    '  is advanced two when there's a natural class involved.
        Let TargetPositionIndex = 0
        Let SDLength = Len(SDString)        'We need the literal length, including "\aaa" cases.
        
        
    'Scan along the length of the structural description.
    For SDPositionIndex = 1 To SDLength
        'Advance the target position analogously.
            Let TargetPositionIndex = TargetPositionIndex + 1
        'Act according to whether you have a natural class in the structural description, or
        '  a plain segment.
            If Mid(SDString, SDPositionIndex, 1) = "\" Then
                'This is a natural class.  Check each member for a possible match.
                    'Look up the natural class based on the numerical coding of the two-letter name.
                        Let NaturalClassIndex = NumericalCoding(Mid(SDString, SDPositionIndex + 1, 3))
                    'Examine the whole list for a match.
                        For ListIndex = 1 To mNumberOfMembersInNaturalClass(NaturalClassIndex)
                            If Mid(TargetString, TargetPositionIndex, 1) = _
                                mNaturalClasses(ListIndex, NaturalClassIndex) Then
                                'You've got a match.  You may continue on through the string.
                                    GoTo SafetyPoint
                            End If
                        Next ListIndex
                    'If you're here, you failed to find a match.  Return False and exit.
                        Let Match = False
                        Exit Function
SafetyPoint:
                    'Advance SDPosition index by three, since the \aaa character is four symbols long.
                        Let SDPositionIndex = SDPositionIndex + 3
            Else
                'More straightforwardly:  record failure if literal won't match to literal.
                    If Mid(SDString, SDPositionIndex, 1) <> Mid(TargetString, TargetPositionIndex, 1) Then
                        Let Match = False
                        Exit Function
                    End If
            End If
    Next SDPositionIndex
    
    'If you got this far, then you achieved a match.
        Let Match = True

End Function


Function TrueLength(MyString)

    'The length of a string in a constraint must be adjusted to deduct two symbols, for every
    '  natural-class term (\aaa).
    
    Dim Temp As Long
    Dim PositionIndex As Long
    
    Let Temp = Len(MyString)
    For PositionIndex = 1 To Len(MyString)
        If Mid(MyString, PositionIndex, 1) = "\" Then
            Let Temp = Temp - 3
        End If
    Next PositionIndex
    
    Let TrueLength = Temp
    
End Function
