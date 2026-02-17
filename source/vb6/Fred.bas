Attribute VB_Name = "Fred"
Option Explicit

'-----------------------------------------------------FRed--------------------------------------------------------------

'This is an implementation of the "FRed" ranking-argumentation algorithm, invented by Prince and Brasoveanu (2005).

        
    'Variables that it would be quite inconvenient to pass around by procedure.
        Dim mAbbrev() As String    'Inherited from outside, shared within.
        Dim mNumberOfConstraints As Long            'Ditto
        Dim mTmpFile As Long                        'The casual output file
        Dim mDocFile As Long                        'The output file for Word prettification.
        Dim mHTMFile As Long
        
    'Module-level variables about ERCs
    '   ERCs are Elementary Ranking Conditions, from which the final ranking information is deduced.
    '   They are calculated from winner-loser pairs.
        
        Dim mOriginalERCs() As String        'String of W, L and e, read off tableaux.
        Dim mNumberOfOriginalERCs As Long
        Dim mERCEvidenceReport() As String   'Justification:  who beats whom.  A string reporting all the relevant cases.
        Dim mERCOriginatingForm() As Long    'The index of the form from which the ERC arises.
        Dim mERCOriginatingRival() As Long  'The index of the rival candidate from which the ERC arises.
        Dim mERCValhalla() As String         'The final set of ERCs we seek.  This will be either the Most Informative,
                                             '  or the Skeletal, Basis of ERCs.
        Dim mValhallaSize As Long
        
    'Flags and reporting interval
        Dim mFailureFlag As Boolean          'Needed to communicate unrankability from deep within an recursive search.
        Dim mVerbose As Boolean              'Describe every step in the output file.
        Dim mUseSkeletalBasis As Boolean     'This emphasizes "pairwise" ranking arguments.
        Dim mProgressInterval As Long        'Report progress every so many steps.
    
    'Constants embodying a basic classification system for ERCs
        Const Uninformative As Long = 1
        Const Unsatisfiable As Long = 2
        Const Valid As Long = 3
        Const DuplicateCandidates As Long = 4
       
Public Sub Main(NumberOfForms As Long, InputForm() As String, Winner() As String, WinnerFrequency() As Single, _
    Rival() As String, RivalFrequency() As Single, MaximumNumberOfRivals As Long, NumberOfRivals() As Long, _
    NumberOfConstraints As Long, ConstraintName() As String, Abbrev() As String, Stratum() As Long, _
    WinnerViolations() As Long, RivalViolations() As Long, _
    RunningFactorialTypology As Boolean, FactorialTypologyIndex As Long, TmpFile As Long, DocFile As Long, HTMFile As Long)

    'Above variables should be self-explanatory except for:
    '   TmpFile and DocFile are file numbers for printout (files were already opened and written in with
    '   information from more basic algorithms).
    '   TmpFile is text format, DocFile is for prettification with a Word macro.

    'Initialize variables, in case this is not the first time.
        ReDim mOriginalERCs(0)
        Let mNumberOfOriginalERCs = 0
        ReDim mERCEvidenceReport(0)
        ReDim mERCOriginatingForm(0)
        ReDim mERCOriginatingRival(0)
        ReDim mERCValhalla(0)
        Let mValhallaSize = 0
    
    'Control whether detailed arguments will be displayed.
        If Form1.chkDetailedArguments.Value = vbChecked Then
            'This can be very time-consuming, so warn user.  4/6/11:  not any more; thanks to Adrian Brasoveanu for help.
                 Let mVerbose = True
            '    If NumberOfConstraints > 5 Then
            '        If MsgBox("Caution:  checking the Detailed Arguments box for larger input files can produce " + _
            '            "slow runs and huge output files.  Click Yes if you want to do this anyway, otherwise No.", _
            '            vbYesNo + vbExclamation) = vbNo Then
            '                Let mVerbose = False
            '                Let Form1.chkDetailedArguments.Value = vbUnchecked
            '        End If
            '    End If
        Else
            Let mVerbose = False
        End If
        
    'Control which basis of ERCs will be found:  Most Informative Basis or Skeletal Basis
        If Form1.chkMostInformativeBasis.Value = vbChecked Then
            Let mUseSkeletalBasis = False
        Else
            Let mUseSkeletalBasis = True
        End If
        
    'Announce what you're doing.
        Let Form1.lblProgressWindow.Caption = "Applying the FRed algorithm for ranking arguments..."
        'We need left-alignment and small fonts to enable good progress reporting (with tree).
            Let Form1.lblProgressWindow.Alignment = 0   '0 is left
            Let Form1.lblProgressWindow.FontSize = 8
        DoEvents
        
    'Install the constraint count and abbreviation set as module-level variables, to keep subs and functions
    '   easier to read.
        Let mNumberOfConstraints = NumberOfConstraints
        Dim ConstraintIndex As Long
        ReDim mAbbrev(mNumberOfConstraints)
        For ConstraintIndex = 1 To mNumberOfConstraints
            Let mAbbrev(ConstraintIndex) = Abbrev(ConstraintIndex)
        Next ConstraintIndex
        
    'Ditto for file numbers.
        Let mTmpFile = TmpFile
        Let mDocFile = DocFile
        Let mHTMFile = HTMFile
        
    Call PrintAHeaderForRankingReport(TmpFile, DocFile, HTMFile)
    
    'Start by forming the start-up ERC set from the data in the input file.
    '   Note:  This routine returns False if the constraints were not rankable for some reason.
        If FormTheOriginalERCs(NumberOfForms, InputForm(), Winner(), WinnerViolations(), _
            NumberOfRivals(), Rival(), RivalViolations()) = False Then Exit Sub

        
    'Launch recursion using the start-up ERC set.
        '"1" means that we are at the top level (1) of the recursion tree.
        '0 means that we are using the full ERC set and not a subset affiliated with a particular constraint.
            'Say what you are doing
            If mVerbose Then Call PrintPara(DocFile, TmpFile, HTMFile, "Recursive ranking search")
                'should be level 2 header
                
            'Begin recursion.
                Call RecursiveRoutine(mOriginalERCs(), mNumberOfOriginalERCs, "1", 0)
        
    'The recursive routine either found the right ranking or set mFailureFlag at True.
    '   Either way, report and interpret the results.
        Call ReportResults
    
    'Augment the report with other material
        If Form1.chkMiniTableaux.Value = vbChecked Then
            Call PrepareMiniTableaux(NumberOfForms, InputForm(), Winner(), WinnerFrequency(), WinnerViolations(), _
                MaximumNumberOfRivals, NumberOfRivals(), Rival(), RivalFrequency(), RivalViolations(), _
                NumberOfConstraints, Abbrev(), ConstraintName(), Stratum(), RunningFactorialTypology, _
                FactorialTypologyIndex, DocFile, TmpFile, HTMFile)
        End If
        Call PrepareHasseDiagram
       
    'Restore the progress window to normal typography and announce completion.
        Let Form1.lblProgressWindow.Alignment = 2   '2 is center
        Let Form1.lblProgressWindow.FontSize = 10
        Let Form1.lblProgressWindow.Caption = "Done."
        DoEvents
        

End Sub

Function FormTheOriginalERCs(NumberOfForms As Long, InputForm() As String, Winner() As String, WinnerViolations() As Long, _
    NumberOfRivals() As Long, Rival() As String, RivalViolations() As Long) As Boolean

    'Form the original set ERCs from which all arguments will proceed.
        
        Dim TemporaryERC As String
        Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long, ERCIndex As Long
        Dim InnerConstraintIndex As Long, InnermostConstraintIndex As Long
        Dim EvidenceString As String
        
    'Default report will be "success" unless you encounter failure below.
        Let FormTheOriginalERCs = True
        
    'Initialize the number found -- in case you're doing factorial typology.
        Let mNumberOfOriginalERCs = 0
        
    'First install some fake ERC's for a priori rankings, if appropriate.
        If Form1.mnuConstrainAlgorithmsByAPrioriRankings.Checked = True Then
            'Calculate the a priori ERCs:
                For ConstraintIndex = 1 To mNumberOfConstraints
                    For InnerConstraintIndex = 1 To mNumberOfConstraints
                        If gAPrioriRankingsTable(ConstraintIndex, InnerConstraintIndex) = True Then
                            'Make an ERC to express this a priori ranking
                                Let TemporaryERC = ""
                                For InnermostConstraintIndex = 1 To mNumberOfConstraints
                                    Select Case InnermostConstraintIndex
                                        Case ConstraintIndex
                                            'Winner preferrer.
                                                Let TemporaryERC = TemporaryERC + "W"
                                        Case InnerConstraintIndex
                                            'Loser preferrer.
                                                Let TemporaryERC = TemporaryERC + "L"
                                        Case Else
                                            'Neutral
                                                Let TemporaryERC = TemporaryERC + "e"
                                    End Select
                                Next InnermostConstraintIndex
                                'Record this ERC.
                                    Let mNumberOfOriginalERCs = mNumberOfOriginalERCs + 1
                                    ReDim Preserve mOriginalERCs(mNumberOfOriginalERCs)
                                    ReDim Preserve mERCEvidenceReport(mNumberOfOriginalERCs)
                                    Let mOriginalERCs(mNumberOfOriginalERCs) = TemporaryERC
                                    Let mERCEvidenceReport(mNumberOfOriginalERCs) = "(a priori ranking)"
                        End If
                    Next InnerConstraintIndex
                Next ConstraintIndex
            
            'Report them to the user:
                'This seems unnecessary because it pops up in the basic ERC list anyway.
                'If mVerbose Then
                '    Call PrintPara(mDocFile, mTmpFile, mHTMfile, "A number of ERC's arise because this simulation is using a priori rankings:")
                '    Dim MyTable() As String
                '    ReDim MyTable(2, mNumberOfOriginalERCs)
                '    'Print #mDocFile, "\ks"
                '    For ERCIndex = 1 To mNumberOfOriginalERCs
                '        'Print #mTmpFile, "   "; Trim(Str(ERCIndex)); "  ";
                '        'Print #mTmpFile, mOriginalERCs(ERCIndex)
                '        Let MyTable(1, ERCIndex) = Trim(Str(ERCIndex))
                '        'Print #mDocFile, Chr(9); Trim(Str(ERCIndex)); ".  ";
                '        'Print #mDocFile, CapE(mOriginalERCs(ERCIndex))
                '        Let MyTable(2, ERCIndex) = CapE(mOriginalERCs(ERCIndex))
                '    Next ERCIndex
                '    'Call s.PrintHTMTable(MyTable(), mHTMfile, False, False, False)
                '    Call s.PrintTable(mDocFile, mTmpFile, mHTMfile, MyTable(), False, False, False)
                '    'Print #mDocFile, "\ke"
                'End If
        
        End If          'Are we making ERC's to reflect a priori rankings?
    
    'Now obtain the ERC's that are calculated from the tableaux.
        For FormIndex = 1 To NumberOfForms
            For RivalIndex = 1 To NumberOfRivals(FormIndex)
                
                'Construct the current ERC in a buffer; it will then be checked in various ways before installation.
                '   Left to right by constraint index.
                '   "W" means "winner-preferrer for this competition"
                '   "L" means "loster-preferrer for this competition"
                '   "e" means "neutral for this competition"
                    Let TemporaryERC = ""
                    For ConstraintIndex = 1 To mNumberOfConstraints
                        Select Case WinnerViolations(FormIndex, ConstraintIndex) - _
                            RivalViolations(FormIndex, RivalIndex, ConstraintIndex)
                                Case Is > 0
                                    Let TemporaryERC = TemporaryERC + "L"
                                Case 0
                                    Let TemporaryERC = TemporaryERC + "e"
                                Case Else
                                    Let TemporaryERC = TemporaryERC + "W"
                        End Select
                    Next ConstraintIndex
                
                'Filter out ERCs that are uninformative or unsatisfiable.
                    Select Case ERCStatus(TemporaryERC)
                        Case Uninformative
                            'do nothing; i.e. don't add an uninformative ERC to the set.
                        Case Unsatisfiable
                            Call PrintPara(mDocFile, mTmpFile, mHTMFile, "Warning:  the constraints cannot be ranked because for input PARA/" + _
                                InputForm(FormIndex) + "/, " + "the rival candidate [" + _
                                Rival(FormIndex, RivalIndex) + "] harmonically bounds the winning candidate PARA[" + _
                                Winner(FormIndex) + "].")
                                Let FormTheOriginalERCs = False
                                Exit Function
                        Case DuplicateCandidates
                            Call PrintPara(mDocFile, mTmpFile, mHTMFile, "Warning:  the constraint cannot be ranked because for input PARA/" + _
                                ", the rival candidate [" + Rival(FormIndex, RivalIndex) + "] has the exact same violations as the winning PARAcandidate " + _
                                Winner(FormIndex) + ".")
                            Let FormTheOriginalERCs = False
                            Exit Function
                        Case Valid
                            'This ERC has some potential, but it is unique?
                                For ERCIndex = 1 To mNumberOfOriginalERCs
                                    If TemporaryERC = mOriginalERCs(ERCIndex) Then
                                        'Not unique, but you can add to the evidence supporting the ERC.
                                            Let EvidenceString = Winner(FormIndex) + " >> " + Rival(FormIndex, RivalIndex)
                                            Let mERCEvidenceReport(ERCIndex) = mERCEvidenceReport(ERCIndex) + ", " + EvidenceString
                                            GoTo ExitPoint
                                    End If
                                Next ERCIndex
                            'Since you've checked for identity for all existing original ERCs, you now know that
                            '   you have a valid, original ERC; add it to the list.
                                Let mNumberOfOriginalERCs = mNumberOfOriginalERCs + 1
                                ReDim Preserve mOriginalERCs(mNumberOfOriginalERCs)
                                ReDim Preserve mERCEvidenceReport(mNumberOfOriginalERCs)
                                Let mOriginalERCs(mNumberOfOriginalERCs) = TemporaryERC
                                Let mERCEvidenceReport(mNumberOfOriginalERCs) = "for /" + InputForm(FormIndex) + "/, " + Winner(FormIndex) + " >> " + Rival(FormIndex, RivalIndex)
                            'Remember (the first instance) of where you got this ERC from.
                                ReDim Preserve mERCOriginatingForm(mNumberOfOriginalERCs)
                                ReDim Preserve mERCOriginatingRival(mNumberOfOriginalERCs)
                                Let mERCOriginatingForm(mNumberOfOriginalERCs) = FormIndex
                                Let mERCOriginatingRival(mNumberOfOriginalERCs) = RivalIndex
                    End Select
    
ExitPoint:              'Dumping point for duplicate ERCs; keep looking.
    
            Next RivalIndex
        Next FormIndex                  'Go through the tableaux.
    
    'In some cases, you'll get no informative ERC's at all.
        If mNumberOfOriginalERCs = 0 Then
            Call PrintPara(mDocFile, mTmpFile, mHTMFile, "None of the constraints prefers a loser, so no ranking is possible or necessary.")
            Let FormTheOriginalERCs = False
            Exit Function
        End If
    
    'Report what you learned.
        'xxx needs level 2 header.
        If mVerbose Then
            Call PrintPara(mDocFile, mTmpFile, mHTMFile, "Original set of ERCs")
            'Print #mDocFile, "\ks"
            'Print #mDocFile, "\h2";
            Dim MyTable() As String
            ReDim MyTable(3, mNumberOfOriginalERCs + 1)
            Let MyTable(1, 1) = "Index"
            Let MyTable(2, 1) = "ERC"
            Let MyTable(3, 1) = "Reason"
            For ERCIndex = 1 To mNumberOfOriginalERCs
                'Call PrintPara(mDocFile, mTmpFile, -1, Trim(Str(ERCIndex)) + " " + mOriginalERCs(ERCIndex) + "    because " + mERCEvidenceReport(ERCIndex))
                Let MyTable(1, ERCIndex + 1) = Trim(Str(ERCIndex))
                Let MyTable(2, ERCIndex + 1) = mOriginalERCs(ERCIndex)
                Let MyTable(3, ERCIndex + 1) = mERCEvidenceReport(ERCIndex)
            Next ERCIndex
            'Print #mDocFile, "\ke"
            'Call s.PrintHTMTable(MyTable(), mHTMfile, True, False, False)
            Call s.PrintTable(mDocFile, mTmpFile, mHTMFile, MyTable(), True, False, False)
        End If

End Function

Sub RecursiveRoutine(MyERCSet() As String, MyERCCount As Long, _
    ProgressString As String, MyConstraint As Long)

    'The core routine of FReD.  Calls itself until done.
        
        Dim MyFusion As String, MyTotalResidueFusion As String
        Dim NumberOfe As Long, NumberOfW As Long
        Dim UseMe As Boolean
        Dim ConstraintIndex As Long, ERCIndex As Long
        Dim CheckString As String
        
        '"Send on" variables:
            Dim SendOnERCs() As String
            Dim SendOnProgressString As String, SendOnBranch As Long, SendOnERCCount As Long
             
    'Report where you are.
        If mVerbose Then
            Print #mDocFile,
            Call PrintPara(mDocFile, mTmpFile, mHTMFile, "Recursive search has now reached this location in the search tree:  " + ProgressString)
            'This includes the particular constraint, if any, whose column of W's and e's gave rise
            '   to this iteration.
                If MyConstraint > 0 Then
                    Call PrintPara(-1, mTmpFile, mHTMFile, "Current set of ERCs is based on constraint #" + Trim(Str(MyConstraint)) + ", " + mAbbrev(MyConstraint))
                    Print #mDocFile, "Current set of ERCs is based on constraint #"; SmallCapTag1; Trim(Str(MyConstraint)); ", "; mAbbrev(MyConstraint); SmallCapTag2
                End If
        End If
        
    'Print the source ERC set
        If mVerbose Then
            If ProgressString <> "1" Then
                Call PrintPara(mDocFile, mTmpFile, mHTMFile, "Working with the following ERC set:")
                Dim Table() As String
                ReDim Table(1, MyERCCount)
                For ERCIndex = 1 To MyERCCount
                    Print #mTmpFile, "      "; MyERCSet(ERCIndex)
                    Print #mDocFile, Chr(9); Chr(9); CapE(MyERCSet(ERCIndex))
                    Let Table(1, ERCIndex) = MyERCSet(ERCIndex)
                Next ERCIndex
                Call s.PrintHTMTable(Table(), mHTMFile, False, False, False)
            End If
        End If
    
    'Fuse the ERC set and report what you found.
        Let MyFusion = Fusion(MyERCSet(), MyERCCount)
        If mVerbose Then
            Print #mTmpFile,
            Call PrintPara(-1, mTmpFile, mHTMFile, "   Fusion of this ERC set is:  " + MyFusion)
            Print #mDocFile,
            Print #mDocFile, Chr(9); "Fusion of this ERC set is:  "
            Print #mDocFile, Chr(9); Chr(9); CapE(MyFusion)
        End If
        
    'Watch for failed fusions.
        If LCount(MyFusion) > 0 And WCount(MyFusion) = 0 Then
            If mVerbose Then
                Call PrintPara(mDocFile, mTmpFile, mHTMFile, "   The fused ERC set has at least one L and no W.  PARA   This means that the constraints cannot be ranked to derive the outputs specified.")
            End If
            'Document the failure and then give up.
                Let mValhallaSize = mValhallaSize + 1
                ReDim Preserve mERCValhalla(mValhallaSize)
                Let mERCValhalla(mValhallaSize) = MyFusion
                Let mFailureFlag = True
                Exit Sub
        End If
    'Watch for constraint sets that prefer no loser.
        If LCount(MyFusion) = 0 Then
            Call PrintPara(mDocFile, mTmpFile, mHTMFile, "   The fused ERC set has no L.  PARA   This means that any ranking will work; there are no ranking arguments here.")
            'Document the failure and then give up.
                Let mValhallaSize = mValhallaSize + 1
                ReDim Preserve mERCValhalla(mValhallaSize)
                Let mERCValhalla(mValhallaSize) = MyFusion
                Let mFailureFlag = True
                Exit Sub
        End If
        
    'Fuse the total residue and report:
        If mVerbose Then
            Call PrintPara(mDocFile, mTmpFile, mHTMFile, "   The following ERCs form the total information-loss residue:")
        End If
        
       '(At this point, the routine FusionOfTotalResidue takes over, reporting the individual cases it finds.)
            Let MyTotalResidueFusion = FusionOfTotalResidue(MyERCSet(), MyERCCount)
        
        If mVerbose Then
            Print #mTmpFile,
            Print #mDocFile,
            If MyTotalResidueFusion <> "" Then
                Call PrintPara(-1, mTmpFile, mHTMFile, "   Fusion of total residue:  " + MyTotalResidueFusion)
                Print #mDocFile, Chr(9); "Fusion of total residue:"
                Print #mDocFile, Chr(9); Chr(9); CapE(MyTotalResidueFusion);
                Print #mDocFile,
            End If
        End If
        
    'Determine if the total residue fusion entails the fusion of the original ERC set and report.
        
        'There two criteria to be used, depending on whether one wants the Minimal Informative Basic or the
        '   Skeletal basis.  EntailmentCheck() will use whichever one the user requested.
        'Note that "success" here is intuitively equivalent to "failure of entailment."
            
            If EntailmentCheck(MyTotalResidueFusion, MyFusion) = False Then
                'Report success, with the appropriate reason:
                    If mVerbose Then
                        If MyTotalResidueFusion = "" Then
                            Call PrintPara(-1, mTmpFile, mHTMFile, "   " + MyFusion + " has a null residue PARAand thus may be retained in the " + BasisName() + " of ERCs.")
                            Print #mDocFile, Chr(9); CapE(MyFusion); " has a null residue and thus may be retained in the "; BasisName(); " of ERCs."
                        Else
                            Call PrintPara(-1, mTmpFile, mHTMFile, "   Thus it may be retained in the " + BasisName() + " of ERCs.")
                            Print #mDocFile, Chr(9); "Thus it may be retained in the "; BasisName(); " of ERCs."
                        End If
                    End If
                'Record your success in the array for the Minimal Informative Basis:
                     'If using the minimal basis, it's now safe to replace it in the MyFusion slot.
                        If mUseSkeletalBasis = True Then
                            Let MyFusion = SkeletalBasis(MyFusion, MyTotalResidueFusion)
                        End If
                    'You must not add ERCs that already occupy this area (repetitions are common in larger simulations):
                        For ERCIndex = 1 To mValhallaSize
                            If MyFusion = mERCValhalla(ERCIndex) Then GoTo ExitPoint
                        Next ERCIndex
                    'This one is new, so add it:
                        Let mValhallaSize = mValhallaSize + 1
                        ReDim Preserve mERCValhalla(mValhallaSize)
                        'Add the ERC to Valhalla.
                            Let mERCValhalla(mValhallaSize) = MyFusion
ExitPoint:
            End If
       
       
       
     'Form the "send on" sets for recursion.
     
         'Amplify values for labeling the output.
            Let SendOnBranch = 0
         
         'Check each character column
            For ConstraintIndex = 1 To mNumberOfConstraints
                'Since each constraint can, in principle, relaunch the recursive routine, initialize the
                '   "send on" variables at the start of this loop.
                    Let SendOnERCCount = 0
                    ReDim SendOnERCs(0)
                'Count the W's, L's, and e's in this column, allowing early exit to speed execution.
                    Let UseMe = True
                    Let NumberOfe = 0
                    Let NumberOfW = 0
                    For ERCIndex = 1 To MyERCCount
                        Select Case Mid(MyERCSet(ERCIndex), ConstraintIndex, 1)
                            Case "e"
                                Let NumberOfe = NumberOfe + 1
                            Case "W"
                                Let NumberOfW = NumberOfW + 1
                            Case "L"
                                'No need to count further--this column doesn't qualify.
                                    Let UseMe = False
                                    Exit For
                        End Select
                    Next ERCIndex
                'If UseMe remains true for this constraint, then we have a character column that contains
                '   no L.  If, furthermore, there is at least one W and e, indicate that this ERC should
                '   be included in the current "send on" set.
                    If UseMe = True And NumberOfW > 0 And NumberOfe > 0 Then
                        'A new branch, so number it.
                            Let SendOnBranch = SendOnBranch + 1
                            Let SendOnProgressString = ProgressString + ", " + Trim(Str(ConstraintIndex))
                        'Since you have the label, you can report progress.
                        '   Truncate the first digit of the search tree, in order to provide a clear interpretation
                        '   for newbies.
                            Let mProgressInterval = mProgressInterval + 1
                            If mProgressInterval = 200 Then
                                Let Form1.lblProgressWindow.Caption = "Working on ranking arguments." + _
                                Chr(10) + Chr(13) + Chr(10) + Chr(13) + Mid(SendOnProgressString, 3) + _
                                Chr(10) + Chr(13) + Chr(10) + Chr(13) + "I'll be done before the first digit reaches " + Trim(Str(mNumberOfConstraints + 1))
                                DoEvents
                                Let mProgressInterval = 0
                            End If
                        'Populate this new call with an ERC set.
                            For ERCIndex = 1 To MyERCCount
                                If Mid(MyERCSet(ERCIndex), ConstraintIndex, 1) = "e" Then
                                    Let SendOnERCCount = SendOnERCCount + 1
                                    ReDim Preserve SendOnERCs(SendOnERCCount)
                                    Let SendOnERCs(SendOnERCCount) = MyERCSet(ERCIndex)
                                End If
                            Next ERCIndex
                        'Make a string to check for repeated labor.
                            Let CheckString = ""
                            For ERCIndex = 1 To SendOnERCCount
                                Let CheckString = CheckString + SendOnERCs(ERCIndex)
                            Next ERCIndex
                        'If this is a new case to check, then launch recursion:
                            If NovelSendOnERCs(CheckString) Then
                                Call RecursiveRoutine(SendOnERCs(), SendOnERCCount, SendOnProgressString, ConstraintIndex)
                            Else
                                If mVerbose Then
                                    Call PrintPara(mDocFile, mTmpFile, mHTMFile, "   ([" + SendOnProgressString + "] not novel so not checked)")
                                End If
                            End If
                            If mFailureFlag = True Then Exit Sub
                    End If
            Next ConstraintIndex
    
    
End Sub

Function NovelSendOnERCs(MyCheckString As String) As Boolean

    'We don't want to check the same set over and over again, so check for redundancy.
        
        Static CheckSet() As String
        Static NumberToCheck As Long
        Dim i As Long
        
        For i = 1 To NumberToCheck
            If MyCheckString = CheckSet(i) Then
                'It's a repeat.
                    Let NovelSendOnERCs = False
                    Exit Function
            End If
        Next i
        'It's novel.
            Let NumberToCheck = NumberToCheck + 1
            ReDim Preserve CheckSet(NumberToCheck)
            Let CheckSet(NumberToCheck) = MyCheckString
            Let NovelSendOnERCs = True

End Function


Function Fusion(ERCSet() As String, NumberOfERCs As Long)

    'Fuse a set of ERCs following the method in Prince and Brasoveanu.
    
    Dim ERCIndex As Long, ConstraintIndex As Long
    Dim Buffer As String, TempBuffer As String, MyChar As String
    Dim CheckString As String
    
    'Initialize the buffer as the null ERC -- all e.
        Let Buffer = ""
        For ConstraintIndex = 1 To mNumberOfConstraints
            Let Buffer = Buffer + "e"
        Next ConstraintIndex
        
    'Add in all the ERCs of this set, following the algebra prescribed by the inventors.
        For ERCIndex = 1 To NumberOfERCs
            Let TempBuffer = ""
            For ConstraintIndex = 1 To mNumberOfConstraints
                Select Case Mid(Buffer, ConstraintIndex, 1)
                    Case "e"
                        Let TempBuffer = TempBuffer + Mid(ERCSet(ERCIndex), ConstraintIndex, 1)
                    Case "W"
                        Let MyChar = Mid(ERCSet(ERCIndex), ConstraintIndex, 1)
                            Select Case MyChar
                                Case "e", "W"
                                    Let TempBuffer = TempBuffer + "W"
                                Case "L"
                                    Let TempBuffer = TempBuffer + "L"
                            End Select
                    Case "L"
                        Let TempBuffer = TempBuffer + "L"
                End Select
            Next ConstraintIndex
            Let Buffer = TempBuffer
        Next ERCIndex
        
        Let Fusion = Buffer
    
End Function

Function FusionOfTotalResidue(MyERCSet() As String, NumberOfERCs As Long) As String

    'See Prince and Brasoveanu p. 30.
    'Arranging an ERC set vertically, inspect each character column, which represents the W's, e's, and L's
    '   for a particular constraint.
    '   If that column consists solely of W's and e's, and there is at least one W and at least one e,
    '   then retain the rows containing e in a list (see PB, item (62), "Info Loss Configuration").
    '   The Fusion of the Total Residue is the fusion (defined elsewhere) of this list.
    
    Dim NumberOfe As Long, NumberOfW As Long
    Dim UseMe As Boolean
    Dim ConstraintIndex As Long, ERCIndex As Long
    
    Dim IncludeMe() As Boolean          'Marks ERCs for inclusion in the total residue.
    Dim SentOnERCSet() As String        'The set of ERCs that will be fused to form the total residue.
    Dim NumberOfSentOnERCs As Long
    
    ReDim IncludeMe(NumberOfERCs)
    
    'Check each character column
        For ConstraintIndex = 1 To mNumberOfConstraints
            Let UseMe = True
            'Count the W's, L's, and e's in this column, allowing early exit to speed execution.
                Let NumberOfe = 0
                Let NumberOfW = 0
                For ERCIndex = 1 To NumberOfERCs
                    Select Case Mid(MyERCSet(ERCIndex), ConstraintIndex, 1)
                        Case "e"
                            Let NumberOfe = NumberOfe + 1
                        Case "W"
                            Let NumberOfW = NumberOfW + 1
                        Case "L"
                            'No need to count further--this column doesn't qualify.
                                Let UseMe = False
                                Exit For
                    End Select
                Next ERCIndex
            'If UseMe remains true for this constraint, then we have a character column that contains
            '   no L.  If, furthermore, there is at least one W and e, indicate that this ERC should
            '   be included in the total "info loss residue".
                If UseMe = True And NumberOfW > 0 And NumberOfe > 0 Then
                    For ERCIndex = 1 To NumberOfERCs
                        If Mid(MyERCSet(ERCIndex), ConstraintIndex, 1) = "e" Then
                            Let IncludeMe(ERCIndex) = True
                        End If
                    Next ERCIndex
                End If
        Next ConstraintIndex
        
    'Collate the set of ERCs to be included in the total residue.
        Dim T() As String
        ReDim T(1, NumberOfERCs)
        For ERCIndex = 1 To NumberOfERCs
            If IncludeMe(ERCIndex) Then
                Let NumberOfSentOnERCs = NumberOfSentOnERCs + 1
                ReDim Preserve SentOnERCSet(NumberOfSentOnERCs)
                Let SentOnERCSet(NumberOfSentOnERCs) = MyERCSet(ERCIndex)
                If mVerbose Then
                    Print #mTmpFile, "      "; MyERCSet(ERCIndex)
                    Print #mDocFile, Chr(9); Chr(9); CapE(MyERCSet(ERCIndex))
                    ReDim Preserve T(1, NumberOfSentOnERCs)
                    Let T(1, NumberOfSentOnERCs) = MyERCSet(ERCIndex)
                End If
            End If
        Next ERCIndex
        If mVerbose Then Call s.PrintHTMTable(T(), mHTMFile, False, False, False)
        
    'Lastly, fuse the set of ERCs thus obtained, yielding the final result.
        If NumberOfSentOnERCs = 0 Then
            Let FusionOfTotalResidue = ""
        Else
            Let FusionOfTotalResidue = Fusion(SentOnERCSet(), NumberOfSentOnERCs)
        End If

End Function


Function EntailmentCheck(MyTotalResidueFusion As String, MyFusion As String)

    'There two criteria for entailment to be used, depending on whether one wants the Minimal Informative
    '   Basis or the Skeletal basis.
    
        Dim MyErc As String
        
        'Print a blank line in the verbose report.
            If mVerbose Then
                Print #mTmpFile,
                Print #mDocFile,
            End If
        
        'Under any version of FReD, if the Total Residue Fusion is null, there is no entailment relation.
            If MyTotalResidueFusion = "" Then
                Let EntailmentCheck = False
                'and there's really nothing worth reporting.
                Exit Function
            End If
            
        'Otherwise, you have to use the particular version appropriate to your goal.
            If mUseSkeletalBasis = True Then
                'The method for the Skeletal Basis.  See Brasoveanu and Prince (2005 34-5).
                    'First, find and report the Skeletal Basis.
                        Let MyErc = SkeletalBasis(MyFusion, MyTotalResidueFusion)
                        If mVerbose Then
                            Call PrintPara(-1, mTmpFile, mHTMFile, "   Skeletal basis of the fusion:  " + MyErc)
                            Print #mDocFile, Chr(9); "Skeletal basis of the fusion:"
                            Print #mDocFile, Chr(9); Chr(9); CapE(MyErc)
                            Print #mDocFile,
                        End If
                    'Then assess if Skeletal Basis consists of all W's and e's.
                        If LCount(MyErc) = 0 Then
                            Let EntailmentCheck = True
                            If mVerbose Then
                                Call PrintPara(-1, mTmpFile, mHTMFile, "   " + MyErc + " has no L's, so it cannot be retained in the Skeletal Basis.")
                                Print #mDocFile, Chr(9); MyErc; " has no L's, so it cannot be retained in the Skeletal Basis."
                            End If
                        Else
                            Let EntailmentCheck = False
                            If mVerbose Then
                                Call PrintPara(-1, mTmpFile, mHTMFile, "   " + MyErc + " includes at least one L and thus is not entailed by " + MyTotalResidueFusion + ".")
                                Print #mDocFile, Chr(9); CapE(MyErc); " includes at least one L and thus is not entailed by "; CapE(MyTotalResidueFusion); "."
                            End If
                        End If
            Else
                'The method for the Minimal Informative Basis, described as the default version of the
                '   algorithm by Brasoveanu and Prince.
                    If Entails(MyTotalResidueFusion, MyFusion) Then
                        Let EntailmentCheck = True
                        If mVerbose Then
                            Call PrintPara(-1, mTmpFile, mHTMFile, "   " + MyFusion + "  is entailed by " + MyTotalResidueFusion + " and thus is not a valid conclusion.")
                            Print #mDocFile, Chr(9); MyFusion; "  is entailed by "; CapE(MyTotalResidueFusion); " and thus is not a valid conclusion."
                        End If
                    Else
                        Let EntailmentCheck = False
                        'Report failure:
                        If mVerbose Then
                            Call PrintPara(-1, mTmpFile, mHTMFile, "   " + MyFusion + " is not entailed by " + MyTotalResidueFusion + ".")
                            Print #mDocFile, Chr(9); CapE(MyFusion); " is not entailed by "; MyTotalResidueFusion; "."
                        End If
                    End If
            End If

End Function
        

Function SkeletalBasis(Fusion As String, TotalResidue As String) As String

    'See Brasoveanu and Prince (2005, 35) for what's going on here.
    
    Dim i As Long
    
    If TotalResidue = "" Then
        Let SkeletalBasis = Fusion
        Exit Function
    End If
    
    For i = 1 To Len(TotalResidue)
        If Mid(TotalResidue, i, 1) = "L" Then
            Let SkeletalBasis = SkeletalBasis + "e"
        Else
            Let SkeletalBasis = SkeletalBasis + Mid(Fusion, i, 1)
        End If
    Next i
    
End Function

Function Entails(ERC1 As String, ERC2 As String) As Boolean

    Dim i As Long
    
    For i = 1 To Len(ERC1)
        Select Case Mid(ERC1, i, 1)
            Case "W"
                'W only entails itself; Prince and Brasoveanu p. 13
                    Select Case Mid(ERC2, i, 1)
                        Case "W"
                            'do nothing
                        Case Else
                            Let Entails = False
                            Exit Function
                    End Select
            Case "e"
                'e entails W and itself.
                    Select Case Mid(ERC2, i, 1)
                        Case "W", "e"
                            'do nothing
                        Case Else
                            Let Entails = False
                            Exit Function
                    End Select
            Case "L"
                'do nothing; L entails anything.
        End Select
    Next i
    
    'If you've gotten here, there is an entailment relation.
        Let Entails = True

End Function


Function ERCStatus(MyErc As String) As Long

    'Prefilter the ERCs according to the following criteria:
        'ERCs that have all e's mean a candidate that duplicates the violations of the winner; hence
        '   no ranking will work.
        'ERCs that have W's and e's but no L in them are uninformative; winner will defeat rival under any ranking.
        'ERCs that have at least one L and no W in them are crash-inducing; no ranking will work.
        'All other ERCs are "valid".
        
    'See Constants above for the four output values of this function.
    
        Dim i As Long
        Dim LCount As Long, WCount As Long, eCount As Long
        
    'Count the L's, W's, and e's:
        For i = 1 To Len(MyErc)
            Select Case Mid(MyErc, i, 1)
                Case "L"
                    Let LCount = LCount + 1
                Case "W"
                    Let WCount = WCount + 1
                Case "e"
                    Let eCount = eCount + 1
            End Select
        Next i
        
    'Classify the ERC according to these totals:
        If LCount = 0 Then
            If WCount = 0 Then
                'All e's, so a loser has exactly the same violations as a winner.
                    Let ERCStatus = DuplicateCandidates
            Else
                'All W's and e's, meaning the loser is harmonically bounded; no ranking info is available from this ERC.
                    Let ERCStatus = Uninformative
            End If
        ElseIf WCount = 0 Then
            'All L's and e's:  winner is harmonically bounded, no ranking can derive it.
                Let ERCStatus = Unsatisfiable
        Else
            'This ERC has some L's and some W's, and is hence potentially informative.
                Let ERCStatus = Valid
        End If

End Function

Function BasisName() As String

    If mUseSkeletalBasis Then
        Let BasisName = "Skeletal Basis"
    Else
        Let BasisName = "Most Informative Basis"
    End If

End Function

Sub ReportResults()

    Dim ERCIndex As Long, ConstraintIndex As Long
    Dim MyErc As String, MyConstraintSet As String
    
    Print #mTmpFile,
    If mVerbose Then
        Call PrintLevel1Header(mDocFile, mTmpFile, mHTMFile, "Ranking argumentation:  Final result")
    End If
    
    If mFailureFlag = False Then
        If mVerbose Then
            Call PrintPara(mDocFile, mTmpFile, mHTMFile, "The following set of ERCs forms the " + BasisName() + " for the ERC set as a whole, PARAand thus encapsulates the available ranking information.")
        End If
    Else
        Call PrintPara(mDocFile, mTmpFile, mHTMFile, "The constraints cannot be ranked to yield the desired outcomes.")
        If mVerbose Then
            Call PrintPara(mDocFile, mTmpFile, mHTMFile, "Here are the ERCs found up to the point of failure.")
        End If
    End If
    
    If mVerbose Then
        
        'Print the ERCs in an order that will match the intuitive order given below.
            Dim T() As String
            'Dim DocTable() As String
            Dim RowCount As Long
            
        'Header:
            ReDim T(2, 1)
            Let T(1, 1) = "This constraint"
            Let T(2, 1) = "Must dominate these constraints"
            
            For ERCIndex = 1 To mValhallaSize
                If WCount(mERCValhalla(ERCIndex)) = 1 Then
                    Print #mTmpFile, "      "; mERCValhalla(ERCIndex)
                    Print #mDocFile, Chr(9); CapE(mERCValhalla(ERCIndex))
                    Let RowCount = RowCount + 1
                    ReDim Preserve T(2, RowCount)
                    Let T(1, RowCount) = mERCValhalla(ERCIndex)
                End If
            Next ERCIndex
            For ERCIndex = 1 To mValhallaSize
                If WCount(mERCValhalla(ERCIndex)) > 1 Then
                    Print #mTmpFile, "      "; mERCValhalla(ERCIndex)
                    Print #mDocFile, Chr(9); CapE(mERCValhalla(ERCIndex))
                    Let RowCount = RowCount + 1
                    ReDim Preserve T(2, RowCount)
                    Let T(1, RowCount) = mERCValhalla(ERCIndex)
                End If
            Next ERCIndex
            'For failed runs:
                For ERCIndex = 1 To mValhallaSize
                    If WCount(mERCValhalla(ERCIndex)) = 0 Then
                        Print #mTmpFile, "      "; mERCValhalla(ERCIndex)
                        Print #mDocFile, Chr(9); Chr(9); CapE(mERCValhalla(ERCIndex))
                    Let RowCount = RowCount + 1
                    If RowCount >= 2 Then
                        ReDim Preserve T(RowCount, 1)
                    End If
                    Let T(1, RowCount) = mERCValhalla(ERCIndex)
                    End If
                Next ERCIndex
                Call s.PrintHTMTable(T(), mHTMFile, False, False, False)
            
    End If
            
        Print #mTmpFile,
        Call PrintPara(mDocFile, mTmpFile, mHTMFile, "The final rankings obtained are as follows:")
        
    
    'We report the nice simple results first, followed by the messy ones.
        'Simple:
            For ERCIndex = 1 To mValhallaSize
                Let MyErc = mERCValhalla(ERCIndex)
                Dim ERCContentCount As Long
                'Print a suitable string for the winner-preferring set, which may be just one element.
                    If WCount(MyErc) = 1 Then
                        Print #mTmpFile, "      "; ConstraintSetString(MyErc, "W", WCount(MyErc));
                        Print #mTmpFile, " >> "; ConstraintSetString(MyErc, "L", LCount(MyErc))
                        Print #mDocFile, Chr(9); SmallCapTag1; ConstraintSetString(MyErc, "W", WCount(MyErc)); SmallCapTag2;
                        Print #mDocFile, " >> "; SmallCapTag1; ConstraintSetString(MyErc, "L", LCount(MyErc)); SmallCapTag2
                        Let ERCContentCount = ERCContentCount + 1
                        ReDim Preserve T(2, ERCContentCount)
                        Let T(1, ERCContentCount) = ConstraintSetString(MyErc, "W", WCount(MyErc)) + " >> " + ConstraintSetString(MyErc, "L", LCount(MyErc))
                    End If
            Next ERCIndex
            Print #mTmpFile,
        'Messy:
            For ERCIndex = 1 To mValhallaSize
                Let MyErc = mERCValhalla(ERCIndex)
                'Print a suitable string for the winner-preferring set, which may be just one element.
                    If WCount(MyErc) > 1 Then
                        Print #mTmpFile, "      "; ConstraintSetString(MyErc, "W", WCount(MyErc));
                        Print #mTmpFile, " >> "; ConstraintSetString(MyErc, "L", LCount(MyErc))
                        Print #mDocFile, Chr(9); SmallCapTag1; ConstraintSetString(MyErc, "W", WCount(MyErc)); SmallCapTag2;
                        Print #mDocFile, " >> "; SmallCapTag1; ConstraintSetString(MyErc, "L", LCount(MyErc)); SmallCapTag2
                        Let ERCContentCount = ERCContentCount + 1
                        ReDim Preserve T(2, ERCContentCount)
                        Let T(1, ERCContentCount) = ConstraintSetString(MyErc, "W", WCount(MyErc)) + " >> " + ConstraintSetString(MyErc, "L", LCount(MyErc))
                    End If
            Next ERCIndex
            Call s.PrintHTMTable(T(), mHTMFile, False, False, False)
            
            
    
End Sub

Function WCount(MyErc As String) As Long
    Dim i As Long
    For i = 1 To Len(MyErc)
        If Mid(MyErc, i, 1) = "W" Then
            Let WCount = WCount + 1
        End If
    Next i
End Function
Function LCount(MyErc As String) As Long
    Dim i As Long
    For i = 1 To Len(MyErc)
        If Mid(MyErc, i, 1) = "L" Then
            Let LCount = LCount + 1
        End If
    Next i
End Function

Function ConstraintSetString(MyErc As String, Criterion As String, Count As Long)

    'Create a suitable string for the winner- or loser-preferring set, which may be just one element.
    
    Dim ConstraintIndex As Long
    
        Let ConstraintSetString = ""
        If Count = 1 Then
            For ConstraintIndex = 1 To mNumberOfConstraints
                If Mid(MyErc, ConstraintIndex, 1) = Criterion Then
                    Let ConstraintSetString = mAbbrev(ConstraintIndex)
                    Exit Function
                End If
            Next ConstraintIndex
        Else
            For ConstraintIndex = 1 To mNumberOfConstraints
                If Mid(MyErc, ConstraintIndex, 1) = Criterion Then
                    Let ConstraintSetString = ConstraintSetString + ", " + mAbbrev(ConstraintIndex)
                End If
            Next ConstraintIndex
            If Criterion = "L" Then
                Let ConstraintSetString = "{" + Mid(ConstraintSetString, 2) + " }"
            Else
                Let ConstraintSetString = "At least one of {" + Mid(ConstraintSetString, 2) + " }"
            End If
        End If

End Function


Sub PrintAHeaderForRankingReport(TmpFile As Long, DocFile As Long, HTMFile As Long)

    'Print a header for the a priori ranking part of the output file:
    
    Dim PrintString As String
    
    On Error GoTo CheckError
    
        Print #mDocFile, "\ks"
        
        Call PrintLevel1Header(DocFile, TmpFile, HTMFile, "Ranking Arguments, based on the Fusional Reduction Algorithm")
        
        'Report the basis striven for:
            
            Let PrintString = "This run sought to obtain the " + BasisName()
            Select Case BasisName()
                Case "Skeletal Basis"
                    Let PrintString = PrintString + ", intended to keep each final ranking argument as pithy as possible."
                Case "Most Informative Basis"
                    Let PrintString = PrintString + ", intended to minimize the set of final ranking arguments."
            End Select
            Call PrintPara(DocFile, TmpFile, HTMFile, PrintString)
            Print #mDocFile, "\ke"
         
    'First, make sure there is a folder for these files, a daughter of the
    '   folder in which the input file is located.
        Call Form1.CreateAFolderForOutputFiles
    
    'Also, prepare the beginnings of a GraphViz file, which will be the basis of a Hasse diagram:
        'This is not run when one is doing factorial typology
            If Form1.RunningFactorialTypology = False Then
                Let HasseFile = FreeFile
                Open gOutputFilePath + gFileName + "Hasse.txt" For Output As #HasseFile
                Print #HasseFile, "digraph G {"
            End If
        
        Exit Sub
        
CheckError:

    Select Case Err.Number
        Case 70 ' "File already open" error.
            MsgBox "Error.  Probably what is happening is this:  I'm trying to open the file " + _
                gOutputFilePath + gFileName + "Hasse.txt" + " for purposes making a Hasse diagram, but a file of this name is already open.  I suggest you try to find this file, close it, then click OK.", vbExclamation
            Resume
    End Select
         
End Sub

Sub PrepareHasseDiagram()

    'Step one is to clean up and annotate the ERCs.  Any pair of constraints must fit one of the following categories:
    
        Const NoArgument As Long = 0
        Const Disjunctive As Long = 1           'One of a set must dominate the loser-preferrers.
        Const Certain = 2                       'Exactly one constraint dominates the loser-preferrers
        
        Dim RankingArray() As Long
        Dim ERCIndex As Long, ConstraintIndex As Long, InnerConstraintIndex As Long
        Dim MyErc As String
        Dim Dominator As Long
        Dim ChangeWasMade As Boolean
        
        ReDim Preserve RankingArray(mNumberOfConstraints, mNumberOfConstraints)
        
        'Populate the chart by reading the ERCs.
            For ERCIndex = 1 To mValhallaSize
                Let MyErc = mERCValhalla(ERCIndex)
                If WCount(MyErc) = 1 Then
                    For ConstraintIndex = 1 To mNumberOfConstraints
                        If Mid(MyErc, ConstraintIndex, 1) = "W" Then
                            Let Dominator = ConstraintIndex
                            Exit For
                        End If
                    Next ConstraintIndex
                    For ConstraintIndex = 1 To mNumberOfConstraints
                        If Mid(MyErc, ConstraintIndex, 1) = "L" Then
                            Let RankingArray(Dominator, ConstraintIndex) = Certain
                        End If
                    Next ConstraintIndex
                Else                        'More than one W
                    For ConstraintIndex = 1 To mNumberOfConstraints
                        If Mid(MyErc, ConstraintIndex, 1) = "W" Then
                            Let Dominator = ConstraintIndex
                            For InnerConstraintIndex = 1 To mNumberOfConstraints
                                If Mid(MyErc, InnerConstraintIndex, 1) = "L" Then
                                    'I doubt that certain rankings can be overridden, but let's just make sure.
                                    If RankingArray(Dominator, InnerConstraintIndex) = NoArgument Then
                                        Let RankingArray(Dominator, InnerConstraintIndex) = Disjunctive
                                    End If
                                End If
                            Next InnerConstraintIndex
                        End If
                    Next ConstraintIndex
                End If
            Next ERCIndex
        
    'Make the Hasse file.
        'This is not run when one is doing factorial typology
            If Form1.RunningFactorialTypology = False Then
                 'Indicate the constraint abbreviations as the labels of nodes.
                    For ConstraintIndex = 1 To mNumberOfConstraints
                        Print #HasseFile, "   "; Trim(Str(ConstraintIndex));
                        'Note:   chr(34) is a double quote.
                        '        DumbSym makes constraints with phonetic symbols printable.
                        Print #HasseFile, " [label="; Chr(34); mAbbrev(ConstraintIndex); Chr(34); ",fontsize = 14]"
                        'Print #HasseFile, ",fontsize = "; Trim(Str(HasseFontSize)); "]"
                    Next ConstraintIndex
                'Print the ranking to the Hasse diagram file; each ranking becomes an arrow:
                     For ConstraintIndex = 1 To mNumberOfConstraints
                        For InnerConstraintIndex = 1 To mNumberOfConstraints
                            'Format this arc appropriately
                            If RankingArray(ConstraintIndex, InnerConstraintIndex) = Certain Then
                                'Specify the endpoints of the arc.
                                    Print #HasseFile, "   "; Trim(Str(ConstraintIndex)); " -> "; Trim(Str(InnerConstraintIndex));
                                'Format as an unlabeled solid line:
                                    Print #HasseFile, " [fontsize = 11]"
                            ElseIf RankingArray(ConstraintIndex, InnerConstraintIndex) = Disjunctive Then
                                'Specify the endpoints of the arc.
                                    Print #HasseFile, "   "; Trim(Str(ConstraintIndex)); " -> "; Trim(Str(InnerConstraintIndex));
                                'Format as a dotted line labeled "or":
                                    Print #HasseFile, " [fontsize = 11,style=dotted,label=" + Chr(34) + "or" + Chr(34) + "]"
                            End If
                        Next InnerConstraintIndex
                    Next ConstraintIndex
                'Finish off the Hasse file and close it.
                    Print #HasseFile, "}"
                    Close #HasseFile
                
                'Call the routine needed to make a Hasse diagram, using ATT's "dot.exe".
                    Call Form1.RunATTDot
                    Call Form1.InsertHasseDiagramIntoOutputFile(mDocFile, mHTMFile)
            End If      'Don't do this if it's factorial typology.

            
End Sub

Sub PrepareMiniTableaux(NumberOfForms As Long, InputForm() As String, Winner() As String, WinnerFrequency() As Single, WinnerViolations() As Long, _
    MaximumNumberOfRivals As Long, NumberOfRivals() As Long, Rival() As String, RivalFrequency() As Single, RivalViolations() As Long, _
    NumberOfConstraints As Long, Abbrev() As String, ConstraintName() As String, Stratum() As Long, _
    RunningFactorialTypology As Boolean, FactorialTypologyIndex As Long, _
    DocFile As Long, TmpFile As Long, HTMFile As Long)

    'This has little to do with Fred but it closely related to the practical requirements ranking argumentation and presentation.
    
    '   Go through the data, trying the find the tableaux that would be most useful in presenting an
    '   analysis, and call the code that prints out tableaux.
    '   Usefulness is defined thus:  has one winner-preferrer and at least one loser preferrer.
    
        Dim ConstraintIndex As Long, FormIndex As Long, RivalIndex As Long
    
        'The "sent" variables, which get filled up and sent to the standard tableau-printing routine.
            Dim SentNumberOfForms As Long
            Dim SentNumberOfConstraints As Long
            Dim SentInputForm(1) As String
            Dim SentWinner(1) As String
            Dim SentWinnerFrequency(1) As Single
            Dim SentNumberOfRivals(1) As Long
            Dim SentRival(1, 1) As String
            Dim SentRivalFrequency(1, 1) As Single
            Dim SentConstraintName() As String
            Dim SentAbbrev() As String
            Dim SentStratum() As Long
            Dim SentWinnerViolations() As Long
            Dim SentMaximumNumberOfRivals As Long
            Dim SentRivalViolations() As Long
            Dim SentTmpFile As Long, SentDocFile As Long, SentHTMFile As Long
            Dim SentAlgorithmName As String
            Dim SentRunningFactorialTypology As Boolean
            Dim SentFactorialTypologyIndex As Long
            Dim SentShadingChoice As Boolean
            Dim SentExclamationPointChoice As Boolean
            Dim NumberOfWinnerPreferrers As Long
            Dim NumberOfLoserPreferrers As Long
            Dim AtLeastOneTableauWasCreated As Boolean
        
        'Header and information.
            Call PrintLevel1Header(DocFile, TmpFile, HTMFile, "Mini-Tableaux")
            Call PrintPara(DocFile, TmpFile, HTMFile, "The following small tableaux may be useful in presenting ranking arguments. PARAThey include all winner-rival comparisons in which there is just one PARAwinner-preferring constraint and at least one loser-preferring constraint.  PARAConstraints not violated by either candidate are omitted.")
        
        'Initialize invariant variables sent to the tableau printing routine.
        '   We're doing the tableaux one at a time because they have different numbers of constraints.
            Let SentNumberOfForms = 1
            Let SentMaximumNumberOfRivals = 1
            Let SentNumberOfRivals(1) = 1
            Let SentDocFile = DocFile
            Let SentTmpFile = TmpFile
            Let SentHTMFile = HTMFile
            Let SentAlgorithmName = gAlgorithmName
            Let SentRunningFactorialTypology = RunningFactorialTypology
            Let SentFactorialTypologyIndex = FactorialTypologyIndex
            Let SentShadingChoice = False
            Let SentExclamationPointChoice = False
            
        'Go through all the data looking for winner-loser pairs whose tableaux might be helpful to the user.
        For FormIndex = 1 To NumberOfForms
            For RivalIndex = 1 To NumberOfRivals(FormIndex)
                'Do a count that is needed to enforce the criterion of usability.
                    Let NumberOfWinnerPreferrers = 0
                    Let NumberOfLoserPreferrers = 0
                    For ConstraintIndex = 1 To NumberOfConstraints
                        Select Case WinnerViolations(FormIndex, ConstraintIndex) - RivalViolations(FormIndex, RivalIndex, ConstraintIndex)
                            Case Is > 0
                                Let NumberOfLoserPreferrers = NumberOfLoserPreferrers + 1
                            Case Is < 0
                                Let NumberOfWinnerPreferrers = NumberOfWinnerPreferrers + 1
                        End Select
                    Next ConstraintIndex
                'This is useable if there was one winner-preferrer and at least one loser-preferrer.
                    If NumberOfWinnerPreferrers = 1 And NumberOfLoserPreferrers > 0 Then
                        'So transfer all the relevant data to "sent" arrays, in preparation for calling the tableaux-printing routine.
                            'Input, winner, rivals.
                                Let SentInputForm(1) = InputForm(FormIndex)
                                Let SentWinner(1) = Winner(FormIndex)
                                Let SentWinnerFrequency(1) = WinnerFrequency(FormIndex)
                                Let SentRival(1, 1) = Rival(FormIndex, RivalIndex)
                                Let SentRivalFrequency(1, 1) = RivalFrequency(FormIndex, RivalIndex)
                            'Find out how many constraints will be installed, to permit redimensioning.
                                Let SentNumberOfConstraints = 0
                                For ConstraintIndex = 1 To NumberOfConstraints
                                    If WinnerViolations(FormIndex, ConstraintIndex) > 0 Or RivalViolations(FormIndex, RivalIndex, ConstraintIndex) > 0 Then
                                        Let SentNumberOfConstraints = SentNumberOfConstraints + 1
                                    End If
                                Next ConstraintIndex
                            'Redimension the arrays that depend on this.
                                ReDim SentConstraintName(SentNumberOfConstraints)
                                ReDim SentAbbrev(SentNumberOfConstraints)
                                ReDim SentStratum(SentNumberOfConstraints)
                                ReDim SentWinnerViolations(1, SentNumberOfConstraints)
                                ReDim SentRivalViolations(1, 1, SentNumberOfConstraints)
                            'Install constraints and violations, leaving out the constraints that prefer neither winners nor losers.  This repeats the
                            '   detection of active constraints just carried out for redimensioning purposes.
                                Let SentNumberOfConstraints = 0
                                For ConstraintIndex = 1 To NumberOfConstraints
                                    If WinnerViolations(FormIndex, ConstraintIndex) > 0 Or RivalViolations(FormIndex, RivalIndex, ConstraintIndex) > 0 Then
                                        Let SentNumberOfConstraints = SentNumberOfConstraints + 1
                                        Let SentConstraintName(SentNumberOfConstraints) = ConstraintName(ConstraintIndex)
                                        Let SentAbbrev(SentNumberOfConstraints) = Abbrev(ConstraintIndex)
                                        Let SentStratum(SentNumberOfConstraints) = Stratum(ConstraintIndex)
                                        Let SentWinnerViolations(1, SentNumberOfConstraints) = WinnerViolations(FormIndex, ConstraintIndex)
                                        Let SentRivalViolations(1, 1, SentNumberOfConstraints) = RivalViolations(FormIndex, RivalIndex, ConstraintIndex)
                                    End If
                                Next ConstraintIndex
                            'Call the tableaux-printing routine                    .
                                Call PrintTableaux.Main(SentNumberOfForms, SentNumberOfConstraints, SentConstraintName(), SentAbbrev(), SentStratum(), _
                                SentInputForm(), SentWinner(), SentWinnerFrequency(), SentWinnerViolations(), _
                                SentMaximumNumberOfRivals, SentNumberOfRivals(), SentRival(), SentRivalFrequency(), SentRivalViolations(), _
                                SentTmpFile, SentDocFile, SentHTMFile, SentAlgorithmName, SentRunningFactorialTypology, SentFactorialTypologyIndex, _
                                SentShadingChoice, SentExclamationPointChoice)
                            'Record that you succeed in finding a useful tableau.
                                Let AtLeastOneTableauWasCreated = True
                        
                    End If                  'Is this a useful winner-rival pair?
            Next RivalIndex                 'Go through all rivals for this winner.
        Next FormIndex                      'Go through all forms.
                            
        
        'Report failure if necessary.
            If AtLeastOneTableauWasCreated = False Then
                Call PrintPara(DocFile, TmpFile, HTMFile, "No such tableaux could be created.  You may wish to remedy this PARAby adding more candidates to your input file.")
            End If
        
        
        
End Sub

Function CapE(MyErc As String) As String

    'The .doc file will probably show its ERCs more clearly if you use capital E.
    
    Dim i As Long
    For i = 1 To Len(MyErc)
        Select Case Mid(MyErc, i, 1)
            Case "e"
                Let CapE = CapE + "E"
            Case Else
                Let CapE = CapE + Mid(MyErc, i, 1)
        End Select
    Next i
End Function


