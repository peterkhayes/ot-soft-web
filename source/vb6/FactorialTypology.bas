Attribute VB_Name = "FactorialTypology"
'===============================FACTORIAL TYPOLOGY=================================
'==================================================================================

Option Explicit
    
    'To hold the output of Constraint Demotion.
        Private Type DiscreteRankingResult
            Converged As Boolean
            NumberOfStrata As Long
            Stratum(1000) As Long
        End Type
        
    'Module level variables
        Dim mNumberOfForms As Long
        Dim mInputForm() As String
        Dim mWinner() As String
        Dim mWinnerFrequency() As Single
        Dim mWinnerViolations() As Long
        Dim mMaximumNumberOfRivals As Long
        Dim mNumberOfRivals() As Long
        Dim mRival() As String
        Dim mRivalFrequency() As Single
        Dim mRivalViolations() As Long
        Dim mNumberOfConstraints As Long
        Dim mConstraintName() As String
        Dim mAbbrev() As String
        Dim mFaithfulness() As Boolean
        Dim mTmpFile As Long
        Dim mDocFile As Long
        Dim mHTMFile As Long
    
Public Sub Main(NumberOfForms As Long, NumberOfConstraints As Long, InputForm() As String, _
    Winner() As String, WinnerFrequency() As Single, WinnerViolations() As Long, _
    NumberOfRivals() As Long, Rival() As String, RivalFrequency() As Single, RivalViolations() As Long, _
    ConstraintName() As String, Abbrev() As String, _
    TmpFile As Long, DocFile As Long, HTMFile As Long)

   'The idea is to do an ever-larger factorial typology, adding one input at a time.
      
      Dim OldValhalla() As Long
      Dim NewValhalla() As Long
      Dim OldValhallaSize As Long
      Dim NewValhallaSize As Long
      Dim ValhallaMax As Long
      Let ValhallaMax = 100
      Dim NumberOfFormsConsidered As Long
      Dim LocalNumberOfForms As Long
       
   'Indices
      Dim FTSizeIndex As Long
      Dim TemporaryWinnerIndex As Long
      Dim FormIndex As Long, RivalIndex As Long, InnerRivalIndex As Long, ConstraintIndex As Long
      Dim ValhallaIndex As Long
      
    'Localize input parameters as module level variables.
        Let mNumberOfConstraints = NumberOfConstraints
        ReDim mAbbrev(mNumberOfConstraints)
        ReDim mConstraintName(mNumberOfConstraints)
        ReDim mFaithfulness(mNumberOfConstraints)
        For ConstraintIndex = 1 To NumberOfConstraints
            Let mConstraintName(ConstraintIndex) = ConstraintName(ConstraintIndex)
            Let mAbbrev(ConstraintIndex) = Abbrev(ConstraintIndex)
        Next ConstraintIndex
        Let mNumberOfForms = NumberOfForms
        ReDim mInputForm(mNumberOfForms)
        ReDim mWinner(mNumberOfForms)
        ReDim mWinnerFrequency(mNumberOfForms)
        ReDim mWinnerViolations(mNumberOfForms, mNumberOfConstraints)
        'This has to be one more than normal, since the winner is made into a rival, too.
            Let mMaximumNumberOfRivals = Form1.FindMaximumNumberOfRivals(NumberOfRivals()) + 1
        ReDim mRival(mNumberOfForms, mMaximumNumberOfRivals)
        ReDim mRivalFrequency(mNumberOfForms, mMaximumNumberOfRivals)
        ReDim mRivalViolations(mNumberOfForms, mMaximumNumberOfRivals, mNumberOfConstraints)
        ReDim mNumberOfRivals(mNumberOfForms)
        For FormIndex = 1 To mNumberOfForms
            Let mInputForm(FormIndex) = InputForm(FormIndex)
            Let mWinner(FormIndex) = Winner(FormIndex)
            Let mWinnerFrequency(FormIndex) = WinnerFrequency(FormIndex)
            'MsgBox "Winner is [" + mWinner(FormIndex) + "]. Winner frequency is " + Str(mWinnerFrequency(FormIndex))
            Let mNumberOfRivals(FormIndex) = NumberOfRivals(FormIndex)
            For ConstraintIndex = 1 To mNumberOfConstraints
                Let mWinnerViolations(FormIndex, ConstraintIndex) = WinnerViolations(FormIndex, ConstraintIndex)
            Next ConstraintIndex
            For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                Let mRival(FormIndex, RivalIndex) = Rival(FormIndex, RivalIndex)
                Let mRivalFrequency(FormIndex, RivalIndex) = RivalFrequency(FormIndex, RivalIndex)
                'MsgBox "Form index is " + Str(FormIndex) + " Form is " + mInputForm(FormIndex) + ".  Rival index is " + Str(RivalIndex) + ". Rival is [" + mRival(FormIndex, RivalIndex) + "] Rival frequency is " + Str(mRivalFrequency(FormIndex, RivalIndex))
                For ConstraintIndex = 1 To mNumberOfConstraints
                    Let mRivalViolations(FormIndex, RivalIndex, ConstraintIndex) = RivalViolations(FormIndex, RivalIndex, ConstraintIndex)
                Next ConstraintIndex
            Next RivalIndex
        Next FormIndex
        Let mTmpFile = TmpFile
        Let mDocFile = DocFile
        Let mHTMFile = HTMFile
        
    
    'Debug at starting point:
    
        'GoTo DebugExitPoint:
        Dim DebugFile As Long
        Let DebugFile = FreeFile
        Open gOutputFilePath + "\DebugFacTypeAtStartOfMain.txt" For Output As #DebugFile
            'Constraints:
                For ConstraintIndex = 1 To mNumberOfConstraints
                    Print #DebugFile, ConstraintIndex; vbTab;
                Next ConstraintIndex
                Print #DebugFile,
        'Inputs:
        For FormIndex = 1 To mNumberOfForms
            Print #DebugFile, "Input #"; Trim(Str(FormIndex)); "    "; mInputForm(FormIndex)
            Print #DebugFile, "      Winner:  "; mWinner(FormIndex);
            Print #DebugFile, "   Frequency:  "; mWinnerFrequency(FormIndex)
            Print #DebugFile, "         ";
            'Winner
                For ConstraintIndex = 1 To mNumberOfConstraints
                    Print #DebugFile, mWinnerViolations(FormIndex, ConstraintIndex); vbTab;
                Next ConstraintIndex
                Print #DebugFile,
            'Rivals:
                For RivalIndex = 0 To mNumberOfRivals(FormIndex)
                    Print #DebugFile, RivalIndex; "   Rival:  "; mRival(FormIndex, RivalIndex)
                    '; _
                    '    "   Frequency = "; mFrequency(FormIndex, RivalIndex);
                    Print #DebugFile, "         ";
                    'Violations:
                        For ConstraintIndex = 1 To mNumberOfConstraints
                            Print #DebugFile, mRivalViolations(FormIndex, RivalIndex, ConstraintIndex); vbTab;
                        Next ConstraintIndex
                        Print #DebugFile,
                Next RivalIndex
        Next FormIndex
        Close #DebugFile
DebugExitPoint:
    
    
    'The following is touchy:  if you've already run a factorial typology,
    '   then you'll install a winner as a duplicate candidate.  So only
    '   install winners as candidates once.
    'February 2026.  Oh dear, this seems to be only creating an error. Let us not.  The rivals were all in place.
        'If Form1.FactorialTypologyAlreadyRunOnThisFile = False Then
        '    Call InstallTheWinnersAsMereCandidates
        'End If
   
    'Print a header
      Call PrintOutputFileHeaderForFactorialTypology
    
    'Hunt for duplicate violations.  Returns true if there are some and the user asked for the amagamation fix.
        If HuntForDuplicateViolations(mNumberOfForms, mInputForm(), mNumberOfRivals, mRival(), mRivalViolations(), mDocFile, mTmpFile, mHTMFile) = True Then
            Call AmalgamateCandidates(mNumberOfForms, mMaximumNumberOfRivals, mNumberOfRivals(), mRival(), mRivalViolations(), mNumberOfConstraints)
        End If
    
    'Look up the a priori rankings if necessary:
        If Form1.mnuConstrainAlgorithmsByAPrioriRankings.Checked = True Then
            'ReadAPrioriRankingsAsTable is a boolean function, which returns False if
            '   it failed to do its job.
            If APrioriRankings.ReadAPrioriRankingsAsTable(mNumberOfConstraints, mAbbrev()) = False Then
                Let Form1.mnuConstrainAlgorithmsByAPrioriRankings.Checked = False
            End If
        End If
   
    'This was programmed long ago with gosub statements and I am weary of replacing them with subs.
    ' now they are paired goto statements.
      GoTo FindPossibleOutcomes
FindPossibleOutcomesReturnPoint:
      GoTo SetInitialValhalla
SetInitialValhallaReturnPoint:
      GoTo MainIncrementalRoutine
MainIncrementalRoutineReturnPoint:
      GoTo RestoreLosers
RestoreLosersReturnPoint:
   
    'Report results:
      
        'Standard summary:
            Call PrintOutSummaryInformationForFactorialTypology(NewValhalla(), NewValhallaSize)
        
        'Compact summaries, if desired:
            If Form1.mnuFTSumFile.Checked = True Then
                Call FSSummary(NewValhallaSize, NewValhalla())
            End If
            If Form1.mnuCompactFTFile.Checked = True Then
                Call CompactFTFile(NewValhallaSize, NewValhalla())
            End If
        't-order.
            Call TOrder(NewValhallaSize, NewValhalla())
                        
        'Grand report with details.  Later choices internal to this routine determine
        '  if tableaux get included:
            If Form1.mnuIncludeRankingInFTResults.Checked = True Then
                Call PrintOutFullFactorialTypology(NumberOfConstraints, NewValhallaSize, NewValhalla())
            End If
            
        'Since you've now loaded the file, the user might be interested in reloading
        '   an updated version.
            Let Form1.mnuReload.Caption = "Reload " + gFileName + gFileSuffix
            Let Form1.mnuReload.Visible = True
        
   Exit Sub

Stop
'-----------------------------------------------------------------------------------
FindPossibleOutcomes:

   'Usually, only some of the candidates can be generated at all.  Things will
   '  go faster if you weed out the non-generated first.

      Let LocalNumberOfForms = 1
      
      Dim RivalCounter As Long
      
      Dim LocalNumberOfRivals() As Long
      ReDim LocalNumberOfRivals(mNumberOfForms)
      Dim LocalWinnerViolations() As Long
      ReDim LocalWinnerViolations(mNumberOfForms, mNumberOfConstraints)
      Dim LocalRivalViolations() As Long
      ReDim LocalRivalViolations(mNumberOfForms, mMaximumNumberOfRivals, mNumberOfConstraints)
      Dim Loser() As String
      ReDim Loser(mNumberOfForms, mMaximumNumberOfRivals)
      Dim NumberOfLosers() As Long
      ReDim NumberOfLosers(mNumberOfForms)
      Dim LoserViolations() As Long
      ReDim LoserViolations(mNumberOfForms, mMaximumNumberOfRivals, mNumberOfConstraints)
      Dim Indicator As Boolean
              
      For FormIndex = 1 To mNumberOfForms
         Let TemporaryWinnerIndex = 0
         
         Do
            Let TemporaryWinnerIndex = TemporaryWinnerIndex + 1
            If TemporaryWinnerIndex > mNumberOfRivals(FormIndex) Then Exit Do
         
         'Go through the rivals, and kick out the ones that could not be
         '  derived under any ranking.
                  
             'Form a temporary array of winner violations
                For ConstraintIndex = 1 To mNumberOfConstraints
                   Let LocalWinnerViolations(1, ConstraintIndex) = _
                      mRivalViolations(FormIndex, TemporaryWinnerIndex, ConstraintIndex)
                Next ConstraintIndex
             
             'Form a temporary array of rival candidate violations.
             Let RivalCounter = 0
             For InnerRivalIndex = 1 To mNumberOfRivals(FormIndex)
                If InnerRivalIndex <> TemporaryWinnerIndex Then
                   'This is a true rival, not the winner.
                   Let RivalCounter = RivalCounter + 1
                   For ConstraintIndex = 1 To mNumberOfConstraints
                      Let LocalRivalViolations(1, RivalCounter, ConstraintIndex) _
                        = mRivalViolations(FormIndex, InnerRivalIndex, ConstraintIndex)
                   Next ConstraintIndex
                End If
             Next InnerRivalIndex
             
             Let LocalNumberOfRivals(1) = mNumberOfRivals(FormIndex) - 1
                          
             'Delete from the list of rivals all rivals that could never win, even on their own.
             
             'Further, form an array of Losers, to store the bad rivals for later
             '  printout.  This must be done in two ways:  a priori rankings, and not.
             
                Select Case Form1.mnuConstrainAlgorithmsByAPrioriRankings.Checked
                   Case True
                       Let Indicator = FastRCDWithAPrioriRankings(1, LocalNumberOfRivals(), LocalWinnerViolations(), LocalRivalViolations())
                   Case False
                       Let Indicator = FastRCD(1, LocalNumberOfRivals(), LocalWinnerViolations(), LocalRivalViolations())
                End Select
               
             If Indicator = False Then
                
                'Store the loser in the Loser array.
                    Let NumberOfLosers(FormIndex) = NumberOfLosers(FormIndex) + 1
                    Let Loser(FormIndex, NumberOfLosers(FormIndex)) = _
                       mRival(FormIndex, TemporaryWinnerIndex)
                    For ConstraintIndex = 1 To mNumberOfConstraints
                       Let LoserViolations(FormIndex, NumberOfLosers(FormIndex), ConstraintIndex) _
                           = mRivalViolations(FormIndex, TemporaryWinnerIndex, ConstraintIndex)
                    Next ConstraintIndex
                
                'Move every column past the target column, in the Rival violations matrix,
                '  up one.
                   For RivalIndex = TemporaryWinnerIndex To mNumberOfRivals(FormIndex) - 1
                      'Move up the actual rival.
                         Let mRival(FormIndex, RivalIndex) = mRival(FormIndex, RivalIndex + 1)
                      'Move up the violations.
                         For ConstraintIndex = 1 To mNumberOfConstraints
                            Let mRivalViolations(FormIndex, RivalIndex, ConstraintIndex) _
                               = mRivalViolations(FormIndex, RivalIndex + 1, ConstraintIndex)
                         Next ConstraintIndex
                   Next RivalIndex
                
                'Reduce the number of rivals.
                   Let mNumberOfRivals(FormIndex) = mNumberOfRivals(FormIndex) - 1
                   
                'Adjust the TemporaryWinnerIndex, which is being used to search
                '  this array, down by one, so it will target the next member.
                   Let TemporaryWinnerIndex = TemporaryWinnerIndex - 1
                
             End If
     
         Loop           'Keep examining rivals to see who is a permanent loser.
      
      Next FormIndex    'Examine all forms for permanent losers.
              
    GoTo FindPossibleOutcomesReturnPoint:

Stop
'-----------------------------------------------------------------------------------
SetInitialValhalla:

   'The Valhalla begins with the rivals for the first form, followed by -1 for
   '  all others.

       Let NewValhallaSize = mNumberOfRivals(1)
       Let OldValhallaSize = mNumberOfRivals(1)
       Let NumberOfFormsConsidered = 1
       ReDim OldValhalla(mNumberOfForms, ValhallaMax)
       ReDim NewValhalla(mNumberOfForms, ValhallaMax)
              
       For RivalIndex = 1 To mNumberOfRivals(1)
          Let NewValhalla(NumberOfFormsConsidered, RivalIndex) = RivalIndex
       Next RivalIndex
       
    GoTo SetInitialValhallaReturnPoint

Stop
'-----------------------------------------------------------------------------------
MainIncrementalRoutine:

   'We now know how many rivals there will be per form, since we've finished trimming
   '  back.  There is always one fewer rival sent to TesarSmolensky, so save repeated
   '  computation by subtracting one, for all cases, once and for all.
   
      For FormIndex = 1 To mNumberOfForms
         Let LocalNumberOfRivals(FormIndex) = mNumberOfRivals(FormIndex) - 1
      Next FormIndex
      
   'Outermost loop:  gradually increase the number of forms considered.
      
   For NumberOfFormsConsidered = 2 To mNumberOfForms
      
      'Report progress:
            Let Form1.lblProgressWindow.Caption = "Doing form #" + Format(NumberOfFormsConsidered) + " out of " + Format(mNumberOfForms)
            DoEvents
            
      'First, put the contents of NewValhalla into OldValhalla.  This will free up
      '  NewValhalla() for containing the next-sized factorial typology.
      
          For ValhallaIndex = 1 To NewValhallaSize
             For FormIndex = 1 To NumberOfFormsConsidered - 1
                Let OldValhalla(FormIndex, ValhallaIndex) = NewValhalla(FormIndex, ValhallaIndex)
             Next FormIndex
          Next ValhallaIndex
          Let OldValhallaSize = NewValhallaSize
          
      'Initialize the NewValhalla, fictitiously, by setting its length to be zero.
          Let NewValhallaSize = 0
            
      'The crucial double loop:  loop through the members of OldValhalla, and cross-
      '  classify them with the new set of rivals.
       
          For ValhallaIndex = 1 To OldValhallaSize
             For RivalIndex = 1 To mNumberOfRivals(NumberOfFormsConsidered)
                
                'Package up a special array, in which the winners are taken from
                '  OldValhalla for all forms considered earlier, and from the
                '  set of rivals, for the form being newly considered.
                
                   'Form a temporary array of winner violations.
                        'First, the winner violations from OldValhalla.
                            For FormIndex = 1 To NumberOfFormsConsidered - 1
                               For ConstraintIndex = 1 To mNumberOfConstraints
                               Let LocalWinnerViolations(FormIndex, ConstraintIndex) = _
                                  mRivalViolations(FormIndex, OldValhalla(FormIndex, ValhallaIndex), ConstraintIndex)
                               Next ConstraintIndex
                            Next FormIndex
                       
                        'Second, the winner violations from the newly considered potential winner
                            For ConstraintIndex = 1 To mNumberOfConstraints
                               Let LocalWinnerViolations(NumberOfFormsConsidered, ConstraintIndex) _
                                  = mRivalViolations(NumberOfFormsConsidered, RivalIndex, _
                                  ConstraintIndex)
                            Next ConstraintIndex
                
                   'Form a temporary array of rival candidate violations.
                        'First, the rival violations from OldValhalla.
                            For FormIndex = 1 To NumberOfFormsConsidered - 1
                               Let RivalCounter = 0
                               For InnerRivalIndex = 1 To mNumberOfRivals(FormIndex)
                                  If InnerRivalIndex <> OldValhalla(FormIndex, ValhallaIndex) Then
                                     'This is a true rival, not the winner.
                                      Let RivalCounter = RivalCounter + 1
                                      For ConstraintIndex = 1 To mNumberOfConstraints
                                         Let LocalRivalViolations(FormIndex, RivalCounter, ConstraintIndex) _
                                            = mRivalViolations(FormIndex, InnerRivalIndex, ConstraintIndex)
                                      Next ConstraintIndex
                                  End If
                               Next InnerRivalIndex
                            Next FormIndex
                       
                        'Second, the rival violations for the currently-considered form.
                            Let RivalCounter = 0
                               For InnerRivalIndex = 1 To mNumberOfRivals(NumberOfFormsConsidered)
                                  If InnerRivalIndex <> RivalIndex Then
                                  'This is a true rival, not the winner.
                                     Let RivalCounter = RivalCounter + 1
                                     For ConstraintIndex = 1 To mNumberOfConstraints
                                        Let LocalRivalViolations(NumberOfFormsConsidered, RivalCounter, ConstraintIndex) _
                                           = mRivalViolations(NumberOfFormsConsidered, InnerRivalIndex, ConstraintIndex)
                                     Next ConstraintIndex
                                  End If
                               Next InnerRivalIndex
                
                'Send this package off to the appropriate version of FastRCD, and if
                '  it likes it, add the result to NewValhalla.
                
                'We use (one version of) FastRCD, since documentation is not needed.
                
                'There are two versions of FastRCD, one for a priori rankings, and
                '   the for not.  This saves checking the value of
                '   mnuConstrainAlgorithmsByAPrioriRankings.Checked over and over.
                
                    Select Case Form1.mnuConstrainAlgorithmsByAPrioriRankings.Checked
                        Case False
                            If FastRCD(NumberOfFormsConsidered, LocalNumberOfRivals(), _
                                LocalWinnerViolations(), LocalRivalViolations()) = True Then
                               
                               'Hooray, a new member of the factorial typology!
                                   
                                   'First, designate some room for it in Valhalla
                                      Let NewValhallaSize = NewValhallaSize + 1
                                      
                                   'But this might overflow the array.  If it does, then redimension it.
                                      If NewValhallaSize > ValhallaMax Then
                                         Let ValhallaMax = 2 * ValhallaMax
                                         ReDim Preserve NewValhalla(mNumberOfForms, ValhallaMax)
                                         ReDim Preserve OldValhalla(mNumberOfForms, ValhallaMax)
                                      End If
                                      
                                    'Install the new member:
                                      For FormIndex = 1 To NumberOfFormsConsidered - 1
                                         Let NewValhalla(FormIndex, NewValhallaSize) = _
                                            OldValhalla(FormIndex, ValhallaIndex)
                                      Next FormIndex
                                      Let NewValhalla(NumberOfFormsConsidered, NewValhallaSize) = RivalIndex
                            End If
                        Case True
                            If FastRCDWithAPrioriRankings(NumberOfFormsConsidered, LocalNumberOfRivals(), _
                                LocalWinnerViolations(), LocalRivalViolations()) = True Then
                                    'Same actions as above:
                                    Let NewValhallaSize = NewValhallaSize + 1
                                    If NewValhallaSize > ValhallaMax Then
                                       Let ValhallaMax = 2 * ValhallaMax
                                       ReDim Preserve NewValhalla(mNumberOfForms, ValhallaMax)
                                       ReDim Preserve OldValhalla(mNumberOfForms, ValhallaMax)
                                    End If
                                    For FormIndex = 1 To NumberOfFormsConsidered - 1
                                       Let NewValhalla(FormIndex, NewValhallaSize) = OldValhalla(FormIndex, ValhallaIndex)
                                    Next FormIndex
                                    Let NewValhalla(NumberOfFormsConsidered, NewValhallaSize) = RivalIndex
                            End If
                    End Select
                                
             Next RivalIndex        'Check all rivals for possible addition.
          Next ValhallaIndex        'Scan through all of OldValhalla().
   Next NumberOfFormsConsidered     'Keep expanding the size of the factorial typology.

   GoTo MainIncrementalRoutineReturnPoint
   
Stop
'-----------------------------------------------------------------------------------
RestoreLosers:

    'The eternal-losers were excluded from the computation to speed it up.  But
    '  we do want them to appear in the tableaux.  So restore them, at the end
    '  of the mRival() and mRivalViolations() arrays.
    
    For FormIndex = 1 To mNumberOfForms
        For RivalIndex = 1 To NumberOfLosers(FormIndex)
           Let mNumberOfRivals(FormIndex) = mNumberOfRivals(FormIndex) + 1
           Let mRival(FormIndex, mNumberOfRivals(FormIndex)) = Loser(FormIndex, RivalIndex)
           For ConstraintIndex = 1 To mNumberOfConstraints
                Let mRivalViolations(FormIndex, mNumberOfRivals(FormIndex), ConstraintIndex) = _
                    LoserViolations(FormIndex, RivalIndex, ConstraintIndex)
           Next ConstraintIndex
        Next RivalIndex
    Next FormIndex
      
    GoTo RestoreLosersReturnPoint

End Sub


Sub PrintOutputFileHeaderForFactorialTypology()

    Dim ConstraintIndex As Long
         
    'Print a header for the factorial typology output file.

        Call PrintTopLevelHeader(mDocFile, mTmpFile, mHTMFile, "Results of Factorial Typology Search")
        
        Call PrintPara(-1, mTmpFile, mHTMFile, NiceDate + ", " + NiceTime)
        Call PrintPara(-1, mTmpFile, mHTMFile, "OTSoft version " + gMyVersionNumber + ", release date " + gMyReleaseDate)
        
        Print #mDocFile, "\cn"; NiceDate; ", "; NiceTime
        Print #mDocFile,
        Print #mDocFile, "\cnOTSoft " + gMyVersionNumber
        Print #mDocFile, "\cnRelease date " + gMyReleaseDate
        Print #mDocFile, "\cnSource file:  " + gFileName + gFileSuffix
        Call PrintPara(-1, mTmpFile, mHTMFile, "Source file:  " + gFileName + gFileSuffix)
        
    'Print a page header and diacritic to trigger a page number
        Print #mDocFile, "\hrFactorial typology for "; gFileName; gFileSuffix; Chr$(9); NiceDate; Chr$(9); "\pn"

        'List the constraints.
            Call PrintLevel1Header(mDocFile, mTmpFile, mHTMFile, "Constraints")
         
            Dim Table() As String
            ReDim Table(3, mNumberOfConstraints + 1)
            Let Table(2, 1) = "Full Name"
            Let Table(3, 1) = "Abbr."
            For ConstraintIndex = 1 To mNumberOfConstraints
                Let Table(1, ConstraintIndex + 1) = Trim(Str(ConstraintIndex)) + "."
                Let Table(2, ConstraintIndex + 1) = mConstraintName(ConstraintIndex)
                Let Table(3, ConstraintIndex + 1) = mAbbrev(ConstraintIndex)
            Next ConstraintIndex
            Call s.PrintTable(mDocFile, mTmpFile, mHTMFile, Table(), True, False, False)
            
        'Report any a priori rankings.
            If Form1.mnuConstrainAlgorithmsByAPrioriRankings.Checked = True Then
                Call Form1.PrintOutTheAprioriRankings(mTmpFile, mDocFile, mHTMFile)
            Else
                Call PrintPara(mDocFile, mTmpFile, mHTMFile, "All rankings were considered.")
            End If
            Print #mTmpFile,

        'Indicate the structure of the output file.
            If Form1.mnuIncludeRankingInFTResults.Checked = True Then
                Call PrintPara(mDocFile, mTmpFile, mHTMFile, "Summary results appear at end of file.  PARAImmediately below are reports on individual patterns generated.")
            End If

End Sub


Sub InstallTheWinnersAsMereCandidates()

   'Since there are no "winner" representations in a factorial typology,
   '  shift the status of the "winners" to mere candidates.

   Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
            
   For FormIndex = 1 To mNumberOfForms

      Let mNumberOfRivals(FormIndex) = mNumberOfRivals(FormIndex) + 1

      'Move every rival up one position.
         For RivalIndex = mNumberOfRivals(FormIndex) To 2 Step -1
            Let mRival(FormIndex, RivalIndex) = mRival(FormIndex, RivalIndex - 1)
            For ConstraintIndex = 1 To mNumberOfConstraints
               Let mRivalViolations(FormIndex, RivalIndex, ConstraintIndex) = mRivalViolations(FormIndex, RivalIndex - 1, ConstraintIndex)
            Next ConstraintIndex
         Next RivalIndex

      Let mRival(FormIndex, 1) = mWinner(FormIndex)
      For ConstraintIndex = 1 To mNumberOfConstraints
         Let mRivalViolations(FormIndex, 1, ConstraintIndex) = mWinnerViolations(FormIndex, ConstraintIndex)
      Next ConstraintIndex

   Next FormIndex
   
  
End Sub

Public Function HuntForDuplicateViolations(NumberOfForms As Long, InputForm() As String, NumberOfRivals() As Long, Rival() As String, MyRivalViolations() As Long, _
    DocFile As Long, TmpFile As Long, HTMFile As Long) As Boolean

   'If any two candidates fail to differ in their violations, then the results are not reliable.
   '    In such a situtation, ask the user if they would like to amalgamate the duplicates.

   Dim FormIndex As Long, RivalIndex As Long, InnerRivalIndex As Long, ConstraintIndex As Long
   
   Dim SameViolationsFlag As Boolean
   Dim Buffer As String
   Dim HeaderPrintedAlready As Boolean
      
   For FormIndex = 1 To NumberOfForms
      For RivalIndex = 1 To NumberOfRivals(FormIndex) - 1
         For InnerRivalIndex = RivalIndex + 1 To NumberOfRivals(FormIndex)
               
               Let SameViolationsFlag = True
               For ConstraintIndex = 1 To mNumberOfConstraints
                  If MyRivalViolations(FormIndex, InnerRivalIndex, ConstraintIndex) <> MyRivalViolations(FormIndex, RivalIndex, ConstraintIndex) Then
                     Let SameViolationsFlag = False
                     Exit For
                  End If
               Next ConstraintIndex

               'If the flag is still up, then this rival has the same number of
               '  constraint violations as the winner.  Report this in the output file,
               '  and give the user the option of quitting.
               
                  If SameViolationsFlag = True Then
                    'Print a header to cover all announcements of this problem.
                        If HeaderPrintedAlready = False Then
                            Call PrintLevel1Header(DocFile, TmpFile, HTMFile, "Candidates with duplicate violations")
                            Let Buffer = "There's a problem with the input file. F"
                            Let HeaderPrintedAlready = True
                        Else
                           Let Buffer = "Moreover, f"
                        End If
                    Let Buffer = Buffer + "or input #" + Trim(Str(FormIndex)) + ", /" + DumbSym(InputForm(FormIndex)) + "/, this candidate: PARA"
                    Let Buffer = Buffer + "   [" + DumbSym(Rival(FormIndex, InnerRivalIndex)) + "] PARA"
                    Let Buffer = Buffer + "violates the same constraints, the same number of times, as this candidate:PARA"
                    Let Buffer = Buffer + "   [" + DumbSym(Rival(FormIndex, RivalIndex)) + "]"
                    Call PrintPara(DocFile, TmpFile, HTMFile, Buffer)

                  End If
         Next InnerRivalIndex
      Next RivalIndex
   Next FormIndex

   If HeaderPrintedAlready = True Then
      Select Case MsgBox("There are candidates for the same input that have exactly the same violations.  This usually produces unreliable results. " + _
        "Pick Yes to deal with this by merging the duplicates into a single virtual candidate, No to proceed without merging, and Cancel to exit OTSoft.", _
        vbYesNoCancel + vbExclamation, "OTSoft:  Problem with your input file")
         Case vbYes
            'Return true to indicate that the candidates have to be amalgamated.
                Let HuntForDuplicateViolations = True
                Call PrintPara(DocFile, TmpFile, HTMFile, "User elected to address this problem by amalgamating candidates with a single violation profile PARAinto a single virtual candidate.")
         Case vbNo
            'do nothing
                Let HuntForDuplicateViolations = False
                Call PrintPara(DocFile, TmpFile, HTMFile, "User elected to ignore this.")
         Case vbCancel
            End
      End Select
   End If


End Function

Sub AmalgamateCandidates(NumberOfForms As Long, MaximumNumberOfRivals As Long, NumberOfRivals() As Long, Rivals() As String, RivalViolations() As Long, NumberOfConstraints As Long)

    'There are candidates that have the very same violations.  Deal with this by making them the same candidate; the candidate is given a name consisting
    '   of its ingredients separated by &.
        
    'Variables:
        'Store the amalgamated arrays.
            Dim TempRivalViolations() As Long
            Dim TempRivals() As String
            Dim TempNumberOfRivals() As Long
        'Flag
            Dim SameViolationsFlag As Boolean
        'Indices
            Dim FormIndex As Long, RivalIndex As Long, CheckMeRivalIndex As Long, ConstraintIndex As Long
        
    'Redimension the arrays.
        ReDim TempRivalViolations(NumberOfForms, MaximumNumberOfRivals, NumberOfConstraints)
        ReDim TempRivals(NumberOfForms, MaximumNumberOfRivals)
        ReDim TempNumberOfRivals(NumberOfForms)
    
   For FormIndex = 1 To NumberOfForms
        'The first rival qualifies no matter what.
            Let TempNumberOfRivals(FormIndex) = 1
            Let TempRivals(FormIndex, 1) = Rivals(FormIndex, 1)
            For ConstraintIndex = 1 To NumberOfConstraints
                Let TempRivalViolations(FormIndex, 1, ConstraintIndex) = RivalViolations(FormIndex, 1, ConstraintIndex)
            Next ConstraintIndex
        
        'Filter the remaining rivals for non-duplicatehood, one by one.
          For RivalIndex = 2 To NumberOfRivals(FormIndex)
                'Check all of the rivals installed so far for duplication.
                     For CheckMeRivalIndex = 1 To TempNumberOfRivals(FormIndex)
                           'Check to see if the violations are the same.
                                Let SameViolationsFlag = True
                                For ConstraintIndex = 1 To NumberOfConstraints
                                   If RivalViolations(FormIndex, RivalIndex, ConstraintIndex) <> TempRivalViolations(FormIndex, CheckMeRivalIndex, ConstraintIndex) Then
                                      Let SameViolationsFlag = False
                                      Exit For
                                   End If
                                Next ConstraintIndex
            
                           'If the flag is still up, then this rival has the same number of constraint violations as the winner.  Report this in the output file,
                           '  and give the user the option of quitting.
                              If SameViolationsFlag = True Then
                                'This is not a new rival.  Make an amalgamated candidate.  You have the violations already, so just con
                                    Let TempRivals(FormIndex, CheckMeRivalIndex) = TempRivals(FormIndex, CheckMeRivalIndex) + " & " + Rivals(FormIndex, RivalIndex)
                                    'Don't prove it to be a duplicate over and over, just once is ok.
                                        GoTo ResumePoint
                              End If
                     Next CheckMeRivalIndex
                'If you've gotten this far, then this is a new rival.  Install it.
                    Let TempNumberOfRivals(FormIndex) = TempNumberOfRivals(FormIndex) + 1
                    Let TempRivals(FormIndex, TempNumberOfRivals(FormIndex)) = Trim(Rivals(FormIndex, RivalIndex))
                    For ConstraintIndex = 1 To NumberOfConstraints
                        Let TempRivalViolations(FormIndex, TempNumberOfRivals(FormIndex), ConstraintIndex) = RivalViolations(FormIndex, RivalIndex, ConstraintIndex)
                    Next ConstraintIndex
ResumePoint:
          Next RivalIndex
   Next FormIndex
   
   'Install what you learned.
        For FormIndex = 1 To NumberOfForms
            Let mNumberOfRivals(FormIndex) = TempNumberOfRivals(FormIndex)
            For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                Let mRival(FormIndex, RivalIndex) = TempRivals(FormIndex, RivalIndex)
                For ConstraintIndex = 1 To NumberOfConstraints
                    Let mRivalViolations(FormIndex, RivalIndex, ConstraintIndex) = TempRivalViolations(FormIndex, RivalIndex, ConstraintIndex)
                Next ConstraintIndex
            Next RivalIndex
        Next FormIndex


End Sub



Sub PrintOutSummaryInformationForFactorialTypology(Valhalla() As Long, ValhallaPopulation As Long)

    'Print the possible outcomes in columns, four columns together per row, with inputs on the left.

        Dim ValhallaIndex As Long
        Dim NumberOfColumns As Long
        Dim ColumnIndex As Long
        Dim ColumnWidth(5) As Long
        Dim FormIndex As Long, RivalIndex As Long
        Dim RowIndex As Long
        Dim SpaceIndex As Long
        Dim FoundAtLeastOneFlag As Boolean
        Dim MaximumNumberOfGrammars As Single
        Dim PrintString As String

        Print #mTmpFile,
        
        'Header.
            Call PrintLevel1Header(mDocFile, mTmpFile, mHTMFile, "Summary Information")

        'Figure out and report how many grammars there could have been.
            Let MaximumNumberOfGrammars = Module1.Factorial(mNumberOfConstraints)
            If MaximumNumberOfGrammars = -1 Then
                Call PrintPara(mDocFile, mTmpFile, mHTMFile, "With " + Trim(Str(mNumberOfConstraints)) + " constraints, the number of logically possible grammars is too large to compute.")
            Else
                Call PrintPara(mDocFile, mTmpFile, mHTMFile, "With " + Trim(Str(mNumberOfConstraints)) + " constraints, the number of logically possible grammars is " + Trim(Str(MaximumNumberOfGrammars)) + ".")
            End If
        
        'Indicate how many output patterns there were.
            Call PrintPara(mDocFile, mTmpFile, mHTMFile, "There were " + Trim(Str(ValhallaPopulation)) + " different output patterns.")
        
        'In principle, all positive frequency forms should be marked as winners.
        '   This will take some work, and it's time to stop for now.
        '   At present, the idea is *not* to designate empirical winners if there are more than one.
            If Form1.mMoreThanOneWinner = False Then
                Print #mDocFile, "Forms marked as winners in the input file are marked with " + PointingFinger("mdocfile") + "."
                Print #mTmpFile, "Forms marked as winners in the input file are marked with " + PointingFinger("mtmpfile") + "."
                Print #mHTMFile, "Forms marked as winners in the input file are marked with " + PointingFinger("mhtmfile") + "."
            End If
            
            Print #mDocFile,
            Print #mTmpFile,
        
    'Print the typology out, four grammars at a time.
        
        For RowIndex = 0 To Int(ValhallaPopulation - 1 / 4) Step 4
            'Start a new array for html.
                Dim Table() As String
            'Print the Word conversion table-start diacritic.
                Print #mDocFile, "\ts";
                'Figure out how many columns this will be, so you can tell Word and HTML.
                   Let NumberOfColumns = 0
                   For ValhallaIndex = RowIndex To RowIndex + 4
                      If ValhallaIndex <= ValhallaPopulation Then
                         Let NumberOfColumns = NumberOfColumns + 1
                      End If
                   Next ValhallaIndex
                     Print #mDocFile, Trim(Str(NumberOfColumns))
                     Print #mDocFile, "\nt";
                     ReDim Table(NumberOfColumns, mNumberOfForms + 1)
                
             'Calculate column widths
    
                'Initialize:
                    For ColumnIndex = 1 To 5
                       Let ColumnWidth(ColumnIndex) = 0
                    Next ColumnIndex
    
                'Search for the longest:
                    For FormIndex = 1 To mNumberOfForms
                       'The input might be the longest, in column 1.
                            If Len(mInputForm(FormIndex)) > ColumnWidth(1) Then
                               Let ColumnWidth(1) = Len(mInputForm(FormIndex)) + 4
                            End If
                        'Search for a longer one.
                            For ValhallaIndex = RowIndex + 1 To RowIndex + 4
                                'Augment if you find a longer.
                                    If Len(mRival(FormIndex, Valhalla(FormIndex, ValhallaIndex))) > ColumnWidth(ValhallaIndex - RowIndex + 1) Then
                                       Let ColumnWidth(ValhallaIndex - RowIndex + 1) = Len(mRival(FormIndex, Valhalla(FormIndex, ValhallaIndex)))
                                    End If
                                'The label 'output #n' might be the longest
                                    If FormIndex > 1 Then
                                       If Len(Trim(Str(ValhallaIndex))) + 8 > ColumnWidth(ValhallaIndex - RowIndex + 1) Then
                                          Let ColumnWidth(ValhallaIndex - RowIndex + 1) = Len(Trim(Str(ValhallaIndex))) + 8
                                       End If
                                    End If
                            Next ValhallaIndex
                    Next FormIndex
    
            'Print column headers.
                For SpaceIndex = 1 To ColumnWidth(1)
                   Print #mTmpFile, " ";
                Next SpaceIndex
                Print #mDocFile, Chr$(9);
                For ValhallaIndex = RowIndex + 1 To RowIndex + 4
                   If ValhallaIndex <= ValhallaPopulation Then
                      Print #mTmpFile, "  Output #";
                      Print #mTmpFile, Trim(Str(ValhallaIndex));
                      Print #mDocFile, "Output #";
                      Print #mDocFile, Trim(Str(ValhallaIndex));
                      Let Table(ValhallaIndex - RowIndex + 1, 1) = "Output #" + Trim(Str(ValhallaIndex))
                      If ValhallaIndex < RowIndex + 4 And ValhallaIndex < ValhallaPopulation Then
                         For SpaceIndex = Len(Trim(Str(ValhallaIndex))) To ColumnWidth(ValhallaIndex - RowIndex + 1) - 8
                            Print #mTmpFile, " ";
                         Next SpaceIndex
                         Print #mDocFile, Chr$(9);
                      End If
                   End If
                Next ValhallaIndex
                Print #mTmpFile,
                Print #mDocFile,
    
             'Print all the input-output pairs, four per row.
    
             For FormIndex = 1 To mNumberOfForms
    
                'Print the names of the inputs.
                    Print #mTmpFile, "/"; DumbSym(mInputForm(FormIndex)); "/";
                    For SpaceIndex = Len(mInputForm(FormIndex)) + 2 To ColumnWidth(1)
                       Print #mTmpFile, " ";
                    Next SpaceIndex
                    Print #mDocFile, "/"; SymbolTag1; mInputForm(FormIndex); SymbolTag2; "/:"; Chr$(9);
                    Let Table(1, FormIndex + 1) = mInputForm(FormIndex)
    
                'Now the outputs.
                    For ValhallaIndex = RowIndex + 1 To RowIndex + 4
                       If ValhallaIndex <= ValhallaPopulation Then
                            'Put finger or > next to winners.
                            ' N.B. I'm not trying to show winners if there are more than one.
                            ' This needs fixed:  show winners if frequency > 0.
                                If Valhalla(FormIndex, ValhallaIndex) = 1 And Form1.mMoreThanOneWinner = False Then
                                    Print #mTmpFile, PointingFinger("mTmpFile"); DumbSym(mRival(FormIndex, Valhalla(FormIndex, ValhallaIndex))); " ";
                                    Print #mDocFile, PointingFinger("mDocFile") + " ["; SymbolTag1; mRival(FormIndex, Valhalla(FormIndex, ValhallaIndex)); SymbolTag2; "]";
                                    Let Table(ValhallaIndex - RowIndex + 1, FormIndex + 1) = PointingFinger("mHTMFile") + mRival(FormIndex, Valhalla(FormIndex, ValhallaIndex))
                                Else
                                    Print #mTmpFile, " "; DumbSym(mRival(FormIndex, Valhalla(FormIndex, ValhallaIndex))); " ";
                                    Print #mDocFile, "["; SymbolTag1; mRival(FormIndex, Valhalla(FormIndex, ValhallaIndex)); SymbolTag2; "]";
                                    Let Table(ValhallaIndex - RowIndex + 1, FormIndex + 1) = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" + mRival(FormIndex, Valhalla(FormIndex, ValhallaIndex))
                                End If
                            'Print enough spaces to fill the gap.
                                If ValhallaIndex < RowIndex + 4 And ValhallaIndex < ValhallaPopulation Then
                                    For SpaceIndex = Len(mRival(FormIndex, Valhalla(FormIndex, ValhallaIndex))) To ColumnWidth(ValhallaIndex - RowIndex + 1)
                                    Print #mTmpFile, " ";
                                    Next SpaceIndex
                                    Print #mDocFile, Chr$(9);
                                End If
                       End If
                    Next ValhallaIndex
    
                'End of line.
                   Print #mTmpFile,
                   Print #mDocFile,
                   
             Next FormIndex         'Print outputs for all forms on this row of four.
    
             'Blank line before next row.
                Print #mDocFile, "\te"
                Print #mTmpFile,
    
             'Print the html table.
                Call s.PrintHTMTable(Table(), mHTMFile, True, False, False)
    
        Next RowIndex                 'Four at a time.


    'Print out how each input did under all the grammars.
        Call PrintLevel1Header(mDocFile, mTmpFile, mHTMFile, "List of Winners")
        Call PrintPara(mDocFile, mTmpFile, mHTMFile, "The following specifies for each candidate whether there is at least one ranking that derives it:")

        'Start a new htm table.
            Dim NumberOfTableRows As Long
            Let NumberOfTableRows = 1
            ReDim Table(3, 1)
            Let Table(1, 1) = "Input"
            Let Table(2, 1) = "Candidate"
            Let Table(3, 1) = "Derivable?"
      
          For FormIndex = 1 To mNumberOfForms
          
             Print #mTmpFile, "/"; DumbSym(mInputForm(FormIndex)); "/:  "
             Print #mDocFile, "\ts2"
             Print #mDocFile, "\nt";
             Print #mDocFile, "/"; SymbolTag1; mInputForm(FormIndex); SymbolTag2; "/"; Chr$(9)
    
             For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                Print #mTmpFile, "   ["; DumbSym(mRival(FormIndex, RivalIndex)); "]:"; "  ";
                Print #mDocFile, "["; SymbolTag1; mRival(FormIndex, RivalIndex); SymbolTag2; "]"; Chr$(9);
                For SpaceIndex = Len(mRival(FormIndex, RivalIndex)) To 10
                   Print #mTmpFile, " ";
                Next SpaceIndex
                
                Let NumberOfTableRows = NumberOfTableRows + 1
                ReDim Preserve Table(3, NumberOfTableRows)
                Let Table(2, NumberOfTableRows) = mRival(FormIndex, RivalIndex)
                If RivalIndex = 1 Then
                    Let Table(1, NumberOfTableRows) = mInputForm(FormIndex)
                End If
                
                'Check if the factorial typology includes this rival.
                    Let FoundAtLeastOneFlag = False
                    For ValhallaIndex = 1 To ValhallaPopulation
                       If Valhalla(FormIndex, ValhallaIndex) = RivalIndex Then
                          Let FoundAtLeastOneFlag = True
                          Exit For
                       End If
                    Next ValhallaIndex
                    
                'Report result in the tables.
                    Select Case FoundAtLeastOneFlag
                        Case False
                            Print #mTmpFile, "no"
                            Print #mDocFile, "no"
                            Let Table(3, NumberOfTableRows) = "no"
                        Case True
                            Print #mTmpFile, "yes"
                            Print #mDocFile, "yes"
                            Let Table(3, NumberOfTableRows) = "yes"
                    End Select
             Next RivalIndex
             
            'Finish up this table
                Print #mDocFile, "\te"
          
          Next FormIndex
                
        Call s.PrintHTMTable(Table(), mHTMFile, True, False, False)
             
         
End Sub

Sub FSSummary(ValhallaPopulation As Long, Valhalla() As Long)

    Dim FormIndex As Long
    Dim ValhallaIndex As Long
    Dim f As Long
    
    'First, make sure there is a folder for these files, a daughter of the folder in which the input file is located.
        Call Form1.CreateAFolderForOutputFiles
    
    'Open a file.
        Let f = FreeFile
        Open gOutputFilePath + gFileName + "FTSum" + ".txt" For Output As #f
    
    'Print a header:
        For FormIndex = 1 To mNumberOfForms
            Print #f, "/"; mInputForm(FormIndex); "/";
            If FormIndex < mNumberOfForms Then
                Print #f, Chr(9);
            Else
                Print #f,
            End If
        Next FormIndex
    
    For ValhallaIndex = 1 To ValhallaPopulation
        For FormIndex = 1 To mNumberOfForms
            Print #f, mRival(FormIndex, Valhalla(FormIndex, ValhallaIndex));
            If FormIndex < mNumberOfForms Then
                Print #f, Chr(9);
            Else
                Print #f,
            End If
        Next FormIndex
    Next ValhallaIndex

    Close #f

End Sub

Sub CompactFTFile(ValhallaPopulation As Long, Valhalla() As Long)

    'Sometimes it is useful to know the factorial typology of a grammar that
    '  massively neutralizes different inputs, simply in terms of what can
    '  be generated, irrespective of input.
    
    'This routine collates factorial typology output sets, as lists of outputs
    '  irrespective of input.  It also provides a count, for each output set,
    '  of the number of distinct outputs that there are.  This helps in sorting.
        
        Dim f As Long       'output file number
        Dim CompactValhalla() As String
        ReDim CompactValhalla(ValhallaPopulation)
        Dim CompactPopulation As Long
        Dim NumberOfDistinctOutputs As Long
        Dim Buffer As String
        Dim ValhallaIndex As Long
        Dim InnerValhallaIndex As Long
        Dim FormIndex As Long
        Dim InnerFormIndex As Long
        
    'First, make sure there is a folder for these files, a daughter of the
    '   folder in which the input file is located.
        Call Form1.CreateAFolderForOutputFiles
    
    'Open a file.
        Let f = FreeFile
        Open gOutputFilePath + gFileName + "CompactSum" + ".txt" For Output As #f
    
    For ValhallaIndex = 1 To ValhallaPopulation
        Let Buffer = ""
        Let NumberOfDistinctOutputs = 0
        For FormIndex = 1 To mNumberOfForms
            'Find out if this output has already been recorded, derived from
            '  a different input.
                For InnerFormIndex = 1 To FormIndex - 1
                    If mRival(InnerFormIndex, Valhalla(InnerFormIndex, ValhallaIndex)) = _
                        mRival(FormIndex, Valhalla(FormIndex, ValhallaIndex)) Then
                            GoTo EscapePoint1
                    End If
                Next InnerFormIndex
            'Ok, it's passed the nonduplication test:
                Let Buffer = Buffer _
                    + mRival(FormIndex, Valhalla(FormIndex, ValhallaIndex)) _
                    + Chr(9)
                Let NumberOfDistinctOutputs = NumberOfDistinctOutputs + 1
EscapePoint1:
        Next FormIndex
        Let Buffer = _
            Trim(Str(NumberOfDistinctOutputs)) + _
            Chr(9) + Buffer
        
        'Often, the neutralization of outputs means that there is more than one
        '  grammar that derives the same surface output set.  Therefore,
        '  one must also remove duplicates at the level of grammars.  Store
        '  the distinct output sets in the array CompactValhalla().
        
            For InnerValhallaIndex = 1 To CompactPopulation
                If Buffer = CompactValhalla(InnerValhallaIndex) Then
                    GoTo EscapePoint2
                End If
            Next InnerValhallaIndex
            Let CompactPopulation = CompactPopulation + 1
            Let CompactValhalla(CompactPopulation) = Buffer
EscapePoint2:
        
    Next ValhallaIndex

    'Print what you have learned.
        For ValhallaIndex = 1 To CompactPopulation
            Print #f, CompactValhalla(ValhallaIndex)
        Next ValhallaIndex
    
    Close #f

End Sub


Sub PrintOutFullFactorialTypology(NumberOfConstraints As Long, ValhallaPopulation, Valhalla() As Long)

    'Everything gets sorted and messed around with.  So substitute local variables
    '  for the global ones.
        
        Dim LocalWinner() As String
        ReDim LocalWinner(mNumberOfForms)
        Dim LocalWinnerFrequency() As Single
        ReDim LocalWinnerFrequency(mNumberOfForms)
        Dim LocalWinnerViolations() As Long
        ReDim LocalWinnerViolations(mNumberOfForms, mNumberOfConstraints)
        Dim LocalRivalIndex As Long
        Dim LocalRival() As String
        ReDim LocalRival(mNumberOfForms, mMaximumNumberOfRivals)
        Dim LocalRivalFrequency() As Single
        ReDim LocalRivalFrequency(mNumberOfForms, mMaximumNumberOfRivals)
        Dim LocalRivalViolations() As Long
        ReDim LocalRivalViolations(mNumberOfForms, mMaximumNumberOfRivals, mNumberOfConstraints)
        Dim LocalNumberOfRivals() As Long
        ReDim LocalNumberOfRivals(mNumberOfForms)
        Dim LocalConstraintName() As String
        ReDim LocalConstraintName(mNumberOfConstraints)
        Dim LocalAbbrev() As String
        ReDim LocalAbbrev(mNumberOfConstraints)
        Dim LocalStratum() As Long
        ReDim LocalStratum(mNumberOfConstraints)
              
        Dim BackupWinner() As String
        ReDim BackupWinner(mNumberOfForms)
        Dim BackupWinnerViolations() As Long
        ReDim BackupWinnerViolations(mNumberOfForms, mNumberOfConstraints)
        Dim BackupRivalIndex As Long
        Dim BackupRival() As String
        ReDim BackupRival(mNumberOfForms, mMaximumNumberOfRivals)
        Dim BackupRivalViolations() As Long
        ReDim BackupRivalViolations(mNumberOfForms, mMaximumNumberOfRivals, mNumberOfConstraints)
        Dim BackupNumberOfRivals() As Long
        ReDim BackupNumberOfRivals(mNumberOfForms)
        Dim BackupConstraintName() As String
        ReDim BackupConstraintName(mNumberOfConstraints)
        Dim BackupAbbrev() As String
        ReDim BackupAbbrev(mNumberOfConstraints)
        Dim BackupStratum() As Long
        ReDim BackupStratum(mNumberOfConstraints)
              
        Dim TSResult As DiscreteRankingResult
        Dim Flag As Boolean
        
        Dim FirstColumnWidth As Long
        
        Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
        Dim StratumIndex As Long
        Dim ValhallaIndex As Long
        Dim SpaceIndex As Long
        
        'Table variables:
            Dim RowCount As Long, MyCellContent As String, MyDocCellContent As String
      
    'Report progress.
        Let Form1.lblProgressWindow.Caption = "Compiling output files..."
        DoEvents
   
    'First, a header:
        Call PrintLevel1Header(mDocFile, mTmpFile, mHTMFile, "Complete Listing of Output Patterns")
   
    'Reprint the outcome set that you're about to analyze in detail.
    
    For ValhallaIndex = 1 To ValhallaPopulation
    
        'Report progress.
            Let Form1.lblProgressWindow.Caption = "Listing output patterns..." + Chr(10) + Chr(10) + _
                "Finished " + Trim(Str(ValhallaIndex)) + "/" + Trim(Str(ValhallaPopulation))
            DoEvents
           
        'We need separate .doc and htm tables because they have different pointing fingers.
            Dim Table() As String
            Dim DocTable() As String
            Let RowCount = 1
            
            'Extravagant separator not needed first time.
                If ValhallaIndex > 1 Then
                    Print #mTmpFile,
                    Print #mTmpFile,
                    Print #mTmpFile, "------------------------------------------------------------------------------"
                End If
            
            Call PrintPara(mDocFile, mTmpFile, mHTMFile, "OUTPUT SET #" + Trim(Str(ValhallaIndex)) + ":")
            Print #mDocFile,
                'Need separate reports since each file type has a different pointing finger.
                    Call PrintPara(mDocFile, -1, -1, "These are the winning outputs.  " + PointingFinger("mdocfile") + " specifies candidates specified as winning candidates in the input file.")
                    Call PrintPara(-1, -1, mHTMFile, "These are the winning outputs.  " + PointingFinger("mhtmfile") + " specifies candidates specified as winning candidates in the input file.")
                    Call PrintPara(-1, mTmpFile, -1, "These are the winning outputs.  PARA" + PointingFinger("mtmpfile") + " specifies outputs marked as winning candidates in the input file.")
    
           'Header for HTML and doc tables
                ReDim Table(2, 1)
                Let Table(1, 1) = "Input"
                Let Table(2, 1) = "Output"
                ReDim DocTable(2, 1)
                Let DocTable(1, 1) = "Input"
                Let DocTable(2, 1) = "Output"
           
           'For cleanliness, find the longest input, so you can line up the arrows.
    
              Let FirstColumnWidth = 0
              For FormIndex = 1 To mNumberOfForms
                 If Len(mInputForm(FormIndex)) > FirstColumnWidth Then
                    Let FirstColumnWidth = Len(mInputForm(FormIndex))
                 End If
              Next FormIndex
    
           For FormIndex = 1 To mNumberOfForms
                'The input form:
                    Print #mTmpFile, "   /"; DumbSym(mInputForm(FormIndex)); "/";
                    For SpaceIndex = Len(mInputForm(FormIndex)) To FirstColumnWidth + 1
                       Print #mTmpFile, " ";
                    Next SpaceIndex
                    Let RowCount = RowCount + 1
                    ReDim Preserve Table(2, RowCount)
                    ReDim Preserve DocTable(2, RowCount)
                    Let Table(1, RowCount) = mInputForm(FormIndex)
                    Let DocTable(1, RowCount) = "/" + SymbolTag1 + mInputForm(FormIndex) + SymbolTag2 + "/"
                    
                'The arrow:
                    Print #mTmpFile, " -->  ";
                    'Print #mDocFile, Chr$(9); "\ys"; Chr(174); "\ye"; Chr$(9);
                'Pointy finger:
                    If Valhalla(FormIndex, ValhallaIndex) = 1 Then  'mNumberOfRivals(FormIndex) Then
                        Let MyCellContent = PointingFinger("mhtmfile")
                        Let MyDocCellContent = PointingFinger("mdocfile")
                    Else
                        Let MyCellContent = ""
                        Let MyDocCellContent = ""
                    End If
                    Print #mTmpFile, ">";
                'The candidate:
                    Print #mTmpFile, DumbSym(mRival(FormIndex, Valhalla(FormIndex, ValhallaIndex)));
                    Print #mTmpFile, " ";
                    If Valhalla(FormIndex, ValhallaIndex) = mNumberOfRivals(FormIndex) Then
                       Print #mTmpFile, " (actual)"
                    Else
                       Print #mTmpFile,
                    End If
                    Let MyDocCellContent = MyDocCellContent + "[" + SymbolTag1 + mRival(FormIndex, Valhalla(FormIndex, ValhallaIndex)) + SymbolTag2 + "]"
                    Let Table(2, RowCount) = MyCellContent + mRival(FormIndex, Valhalla(FormIndex, ValhallaIndex))
                    Let DocTable(2, RowCount) = MyDocCellContent
           Next FormIndex
           
           Print #mTmpFile,
    
            'Print out the html code.
                Call s.PrintHTMTable(Table(), mHTMFile, True, False, False)
                Call s.PrintDocTable(DocTable(), mDocFile, True, False, False)
           
           'Call the localized Recursive Constraint Demotion algorithm, to learn what is actually
           '  crucial to get this particular set of outcomes.
    
              GoTo LocalRCD
LocalRCDReturnPoint:
                   
    Next ValhallaIndex          'Go through the full factorial typology.

    'Say you're done.
        Let Form1.lblProgressWindow.Caption = "Done."
   
    Exit Sub


Stop
'-----------------------------------------------------------------------------
LocalRCD:

     'What has to be done here is to install the particular outcome of the factorial typology search
     '  as the Winners, and everybody else as the Rivals.
     'Then you call a localized version of Recurvsive Constraint Demotion and get the answer.

     'First, install the Winner:

        For FormIndex = 1 To mNumberOfForms
           Let LocalWinner(FormIndex) = mRival(FormIndex, Valhalla(FormIndex, ValhallaIndex))
           Let LocalWinnerFrequency(FormIndex) = mRivalFrequency(FormIndex, Valhalla(FormIndex, ValhallaIndex))
           For ConstraintIndex = 1 To mNumberOfConstraints
              Let LocalWinnerViolations(FormIndex, ConstraintIndex) = mRivalViolations(FormIndex, Valhalla(FormIndex, ValhallaIndex), ConstraintIndex)
           Next ConstraintIndex
        Next FormIndex

     'Then, install the local Rivals.  They consist of every Rival of the full set, except for the Winner.
     'This is done by incrementing a LocalRivalIndex every time you encounter
     '  a Rival that isn't the Winner.

        For FormIndex = 1 To mNumberOfForms
          'Initialize the LocalRivalIndex
           Let LocalRivalIndex = 0
           For RivalIndex = 1 To mNumberOfRivals(FormIndex)
              If RivalIndex <> Valhalla(FormIndex, ValhallaIndex) Then
                 Let LocalRivalIndex = LocalRivalIndex + 1
                 Let LocalRival(FormIndex, LocalRivalIndex) = mRival(FormIndex, RivalIndex)
                 Let LocalRivalFrequency(FormIndex, LocalRivalIndex) = mRivalFrequency(FormIndex, RivalIndex)
                 For ConstraintIndex = 1 To mNumberOfConstraints
                    Let LocalRivalViolations(FormIndex, LocalRivalIndex, ConstraintIndex) = _
                        mRivalViolations(FormIndex, RivalIndex, ConstraintIndex)
                 Next ConstraintIndex
              End If
           Next RivalIndex
       Next FormIndex

    'Finally, and obviously, there are one fewer local Rivals than overall
    '  Rivals, since we've deducted the winner.

        For FormIndex = 1 To mNumberOfForms
            Let LocalNumberOfRivals(FormIndex) = mNumberOfRivals(FormIndex) - 1
        Next FormIndex

    'Now you've got the data set up for a run of Constraint Demotion.

        'The algorithm is set up as a function.  Check that it's ok.
        
            Let TSResult = RecursiveConstraintDemotion(mNumberOfForms, LocalNumberOfRivals(), _
               LocalWinnerViolations(), LocalRivalViolations())
              
              If TSResult.Converged = False Then
                 MsgBox "Program error.   I would appreciate your letting me know the about the problem.  Email me at bhayes@humnet.ucla.edu, specifying error #87102, and including a copy of your input file.", vbCritical
                 End
              End If
            
           GoTo PrintAlgorithmResults
PrintAlgorithmResultsReturnPoint:

    'Now, back to the next member of Valhalla.
        GoTo LocalRCDReturnPoint

Stop
'------------------------------------------------------------------------------
PrintAlgorithmResults:

    'Print the arrangement of constraints into strata:
    
        'Set up html table
            Let RowCount = 1
            ReDim Table(3, 1)
            
        'Label the table as being the grammar.
            Call PrintPara(mDocFile, mTmpFile, mHTMFile, "Grammar:")
        
        'Header of table:
            Let Table(1, 1) = "Stratum"
            Let Table(2, 1) = "Constraint Name"
            Let Table(3, 1) = "Abbr."
            Print #mDocFile, "\ts3"
            Print #mDocFile, "\nt";
            Print #mDocFile, "Stratum"; Chr$(9); "Constraint Name"; Chr$(9); "Abbr."
            
        'Table content:
            For StratumIndex = 1 To TSResult.NumberOfStrata
                'Print what stratum this is:
                    Print #mTmpFile, "   Stratum #"; Trim(Str(StratumIndex))
                    Print #mDocFile, "Stratum #"; Trim(Str(StratumIndex)); Chr$(9);
                    Let RowCount = RowCount + 1
                    ReDim Preserve Table(3, RowCount)
                    Let Table(1, RowCount) = "Stratum #" + Trim(Str(StratumIndex))
              
                'This flag is true when you need to start a new row.
                    Let Flag = False
                
                'Find what stratum each constraint belongs to and print it out.
                    For ConstraintIndex = 1 To mNumberOfConstraints
                        If TSResult.Stratum(ConstraintIndex) = StratumIndex Then
                            'When the flag is false, do not add a new row to the table for this stratum.
                                If Flag = True Then
                                    Let RowCount = RowCount + 1
                                    ReDim Preserve Table(3, RowCount)
                                End If
                            Print #mTmpFile, "      "; mConstraintName(ConstraintIndex);
                            For SpaceIndex = Len(mConstraintName(ConstraintIndex)) To 35
                              Print #mTmpFile, " ";
                            Next SpaceIndex
                            Print #mTmpFile, "[= "; mAbbrev(ConstraintIndex); "]"
                
                            'Use the Flag to print a blank cell underneath the stratum label:
                               If Flag = True Then
                                  Print #mDocFile, "\hd"; Chr$(9);
                               End If
                               Let Flag = True
                            'Print the name and abbreviation of the constraint that belongs to this stratum.
                                Print #mDocFile, mConstraintName(ConstraintIndex); Chr$(9);
                                Print #mDocFile, mAbbrev(ConstraintIndex)
                                Let Table(2, RowCount) = mConstraintName(ConstraintIndex)
                                Let Table(3, RowCount) = mAbbrev(ConstraintIndex)
                         End If
                      Next ConstraintIndex
            Next StratumIndex
           
        'Finish up tables:
            Print #mDocFile, "\te"
            Call s.PrintHTMTable(Table(), mHTMFile, True, False, False)
            
   
    'Fill buffer arrays so that the real, permanent values don't get messed up by sorting.
                         
        For ConstraintIndex = 1 To mNumberOfConstraints
            Let BackupConstraintName(ConstraintIndex) = mConstraintName(ConstraintIndex)
            Let BackupAbbrev(ConstraintIndex) = mAbbrev(ConstraintIndex)
            Let BackupStratum(ConstraintIndex) = TSResult.Stratum(ConstraintIndex)
            Let LocalConstraintName(ConstraintIndex) = mConstraintName(ConstraintIndex)
            Let LocalAbbrev(ConstraintIndex) = mAbbrev(ConstraintIndex)
            Let LocalStratum(ConstraintIndex) = TSResult.Stratum(ConstraintIndex)
        Next ConstraintIndex
        For FormIndex = 1 To mNumberOfForms
            Let BackupWinner(FormIndex) = LocalWinner(FormIndex)
            ' BackupWinner(FormIndex)
            For ConstraintIndex = 1 To mNumberOfConstraints
                Let BackupWinnerViolations(FormIndex, ConstraintIndex) = LocalWinnerViolations(FormIndex, ConstraintIndex)
            Next ConstraintIndex
            Let BackupNumberOfRivals(FormIndex) = LocalNumberOfRivals(FormIndex)
            For RivalIndex = 1 To BackupNumberOfRivals(FormIndex)
               Let BackupRival(FormIndex, RivalIndex) = LocalRival(FormIndex, RivalIndex)
               For ConstraintIndex = 1 To mNumberOfConstraints
                   Let BackupRivalViolations(FormIndex, RivalIndex, ConstraintIndex) = LocalRivalViolations(FormIndex, RivalIndex, ConstraintIndex)
               Next ConstraintIndex
            Next RivalIndex
        Next FormIndex
    
    'If the user wants, include even the tableaux in the report.
     If Form1.mnuIncludeTableaux.Checked = True Then
        
        Print #mTmpFile,
        Call PrintPara(mDocFile, mTmpFile, mHTMFile, "Tableaux:")
        Call PrintTableaux.Main(mNumberOfForms, NumberOfConstraints, LocalConstraintName(), _
           LocalAbbrev(), LocalStratum(), mInputForm(), LocalWinner(), _
           mWinnerFrequency(), LocalWinnerViolations(), _
           mMaximumNumberOfRivals, LocalNumberOfRivals(), LocalRival(), mRivalFrequency(), LocalRivalViolations(), _
           mTmpFile, mDocFile, mHTMFile, _
           "Factorial Typology", True, ValhallaIndex, True, True)
          
    'Include ranking arguments if requested.
        
        If Form1.chkArguerOn.Value = 1 Then
            Call Fred.Main(mNumberOfForms, mInputForm(), LocalWinner(), LocalWinnerFrequency(), LocalRival(), LocalRivalFrequency(), mMaximumNumberOfRivals, LocalNumberOfRivals(), mNumberOfConstraints, _
                mConstraintName(), mAbbrev(), TSResult.Stratum(), LocalWinnerViolations(), LocalRivalViolations(), _
                True, ValhallaIndex, mTmpFile, mDocFile, mHTMFile)
        End If
    
     'Restore values from backups
     
         For ConstraintIndex = 1 To mNumberOfConstraints
             Let LocalConstraintName(ConstraintIndex) = BackupConstraintName(ConstraintIndex)
             Let LocalAbbrev(ConstraintIndex) = BackupAbbrev(ConstraintIndex)
             Let LocalStratum(ConstraintIndex) = BackupStratum(ConstraintIndex)
         Next ConstraintIndex
         For FormIndex = 1 To mNumberOfForms
             Let LocalWinner(FormIndex) = BackupWinner(FormIndex)
             For ConstraintIndex = 1 To mNumberOfConstraints
                 Let LocalWinnerViolations(FormIndex, ConstraintIndex) = BackupWinnerViolations(FormIndex, ConstraintIndex)
             Next ConstraintIndex
             Let LocalNumberOfRivals(FormIndex) = BackupNumberOfRivals(FormIndex)
             For RivalIndex = 1 To BackupNumberOfRivals(FormIndex)
                Let LocalRival(FormIndex, RivalIndex) = BackupRival(FormIndex, RivalIndex)
                For ConstraintIndex = 1 To mNumberOfConstraints
                    Let LocalRivalViolations(FormIndex, RivalIndex, ConstraintIndex) = BackupRivalViolations(FormIndex, RivalIndex, ConstraintIndex)
                Next ConstraintIndex
             Next RivalIndex
         Next FormIndex
    
    End If                  'Did user specify to include tableaux?
                
      
   GoTo PrintAlgorithmResultsReturnPoint
   
End Sub

Private Function RecursiveConstraintDemotion(NumberOfForms As Long, NumberOfRivals() As Long, _
   WinnerViolations() As Long, RivalViolations() As Long) As DiscreteRankingResult

    On Error GoTo CheckError
        
        Dim Stratum() As Long
        ReDim Stratum(mNumberOfConstraints)      'You never need more strata than constraints.
      
        Dim CurrentStratum As Long
   
        Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
        Dim OuterConstraintIndex As Long, InnerConstraintIndex As Long
      
        Dim SomeAreNonDemotible As Boolean
        Dim SomeAreDemotible As Boolean
        Dim StillInformative() As Boolean
        ReDim StillInformative(NumberOfForms, mMaximumNumberOfRivals)
        Dim Demotable() As Boolean
        ReDim Demotable(mNumberOfConstraints)
        
   'Initialize crucial variables, in case this routine gets called more than once.

      Let CurrentStratum = 0

      For ConstraintIndex = 1 To mNumberOfConstraints
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
   
    'Look up the a priori rankings if the user has so requested.
        If Form1.mnuConstrainAlgorithmsByAPrioriRankings.Checked = True Then
            'ReadAPrioriRankingsAsTable is a boolean function, which returns False if
            '   it failed to do its job.
            If APrioriRankings.ReadAPrioriRankingsAsTable(mNumberOfConstraints, mAbbrev()) = False Then
                Let Form1.mnuConstrainAlgorithmsByAPrioriRankings.Checked = False
            End If
        End If
   
   'The information obtained by this routine is not crucial to the algorithm
   '    but it helpful for the diagnosis file:
        
        For ConstraintIndex = 1 To mNumberOfConstraints
            Let mFaithfulness(ConstraintIndex) = Form1.FaithfulnessConstraint(mConstraintName(ConstraintIndex))
        Next ConstraintIndex
      
   'Go through the Winner-Rival pairs repeatedly, looking for constraints that
   '  never prefer losers among the candidates still being considered.

   Do
   
      'Record what stratum you're constructing:
            Let CurrentStratum = CurrentStratum + 1
      
      'Initialize demotibility.
          For ConstraintIndex = 1 To mNumberOfConstraints
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
                For ConstraintIndex = 1 To mNumberOfConstraints
                    'We only consider constraints that haven't been ranked yet.
                    If Stratum(ConstraintIndex) = 0 Then
                    'Keep a constraint out of the current stratum if it prefers a loser.
                        If WinnerViolations(FormIndex, ConstraintIndex) > _
                            RivalViolations(FormIndex, RivalIndex, ConstraintIndex) Then
                                'Report this if it's the first time:
                                    If Demotable(ConstraintIndex) = False Then
                                        Let FoundOne = True
                                    End If
                                Let Demotable(ConstraintIndex) = True
                        End If              'Did the constraint of ConstraintIndex prefer a loser?
                    End If                  'Is the constraint of ConstraintIndex still rankable?
                Next ConstraintIndex
            End If
            Next RivalIndex
        Next FormIndex
        
        
      'A PRIORI RANKINGS
        
        'If APrioriRankings are on, then demote the victims.
        '   We demote those which are a priori ranked below the yet-unranked.
            If Form1.mnuConstrainAlgorithmsByAPrioriRankings.Checked = True Then
                Let FoundOne = False
                For OuterConstraintIndex = 1 To mNumberOfConstraints
                    If Stratum(OuterConstraintIndex) = 0 Then
                        For InnerConstraintIndex = 1 To mNumberOfConstraints
                            If gAPrioriRankingsTable(OuterConstraintIndex, InnerConstraintIndex) = True Then
                                'Demote, and report this, if first time.
                                If Demotable(InnerConstraintIndex) = False Then
                                    Let FoundOne = True
                                End If
                                Let Demotable(InnerConstraintIndex) = True
                            End If
                        Next InnerConstraintIndex
                    End If                  'Is the dominee yet unranked?
                Next OuterConstraintIndex
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
            For ConstraintIndex = 1 To mNumberOfConstraints
                If Stratum(ConstraintIndex) = 0 Then
                    Select Case Demotable(ConstraintIndex)
                        Case False
                            Let SomeAreNonDemotible = True
                            'Install in stratum:
                                Let Stratum(ConstraintIndex) = CurrentStratum
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
             
                Let RecursiveConstraintDemotion.Converged = True
                Let RecursiveConstraintDemotion.NumberOfStrata = CurrentStratum
                For ConstraintIndex = 1 To mNumberOfConstraints
                   Let RecursiveConstraintDemotion.Stratum(ConstraintIndex) = Stratum(ConstraintIndex)
                Next ConstraintIndex
                             
             Exit Function
             
          ElseIf SomeAreNonDemotible = False Then
             
             'This means that all constraints are demotible.
             'There is no hope of a working grammar.
             'But keep the work done so far, since it will be useful in diagnosis.
                
                Let RecursiveConstraintDemotion.Converged = False
                Let RecursiveConstraintDemotion.NumberOfStrata = CurrentStratum
                
                'Give a pseudo-stratum to the unrankable constraints, for purposes
                '   of diagnostic tableaux.
                    For ConstraintIndex = 1 To mNumberOfConstraints
                        Let Buffer = mConstraintName(ConstraintIndex)
                        If Stratum(ConstraintIndex) = 0 Then
                            Let Stratum(ConstraintIndex) = CurrentStratum
                        End If
                    Next ConstraintIndex
                
                'Remember the strata:
                    For ConstraintIndex = 1 To mNumberOfConstraints
                       Let RecursiveConstraintDemotion.Stratum(ConstraintIndex) = Stratum(ConstraintIndex)
                    Next ConstraintIndex
                
                'Despair and depart:
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
            MsgBox "Program error.  You can ask for help at bhayes@humnet.ucla.edu, specifying error #44487.  Please send a copy of your input file with your message.", vbCritical
    End Select


End Function

Public Function FastRCD(ByVal NumberOfForms As Long, NumberOfRivals() As Long, _
   WinnerViolations() As Long, RivalViolations() As Long) As Boolean

   'Dimension the local variables.
   
      Dim Stratum() As Long
      ReDim Stratum(mNumberOfConstraints)
      
      Dim CurrentStratum As Long
   
      Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
   
      Dim SomeAreNonDemotible As Boolean, SomeAreDemotible As Boolean
      Dim StillInformative() As Boolean
      ReDim StillInformative(NumberOfForms, mMaximumNumberOfRivals)
      Dim Demotable() As Boolean
      ReDim Demotable(mNumberOfConstraints)
 
   'Initialize crucial variables, in case this routine gets called more
   '  than once.

      Let CurrentStratum = 0

      For ConstraintIndex = 1 To mNumberOfConstraints
         Let Stratum(ConstraintIndex) = 0
      Next ConstraintIndex

      For FormIndex = 1 To NumberOfForms
         For RivalIndex = 1 To mMaximumNumberOfRivals
            Let StillInformative(FormIndex, RivalIndex) = True
         Next RivalIndex
      Next FormIndex
      
   'Go through the Winner-Rival pairs repeatedly, looking for constraints that
   '  are never crucially violated among the pairs still being considered.

   Do
   
      'Record what stratum you're constructing:
         Let CurrentStratum = CurrentStratum + 1

      'Initialize a variable indicating demotability.
          For ConstraintIndex = 1 To mNumberOfConstraints
             Let Demotable(ConstraintIndex) = False
          Next ConstraintIndex

      'Go through all pairs of Winner vs. Rival, and learn from them
      '  so long as they are still in the informative class.
          
          For FormIndex = 1 To NumberOfForms
             For RivalIndex = 1 To NumberOfRivals(FormIndex)

               'Only still-informative Rivals can be learned from:
                If StillInformative(FormIndex, RivalIndex) = True Then
                   For ConstraintIndex = 1 To mNumberOfConstraints

                      'The crucial step of the algorithm:  demote constraints
                      '  that are violated in winners.  This is done only for
                      '  yet-unranked constraints.

                      If Stratum(ConstraintIndex) = 0 Then
                         If WinnerViolations(FormIndex, ConstraintIndex) > RivalViolations(FormIndex, RivalIndex, ConstraintIndex) Then
                            Let Demotable(ConstraintIndex) = True
                         End If
                      End If

                   Next ConstraintIndex
                End If
             Next RivalIndex
          Next FormIndex


      'You've now assessed the demotability of all yet-unranked constraints,
      '  using all the data.
      'Now, find out which constraints may be assigned to the next stratum.

      'There is also the delicate matter of knowing when you're done.  Here
      '  are the three cases:

      '  I. Some of the yet-unranked constraints are demotible.
      '        (Stratum = 0 , Demotable = True, False).  Demote them and continue.

      ' II. All of the yet-unranked constraints are demotible.
      '        (Stratum = 0, Demotable = True).       Record a failed constraint
      '                                               set and exit.

      'III. None of the yet-unranked constraints are demotible.
      '        (Stratum = 0, Demotible = False)       They are the lowest stratum
      '                                               of a working grammar.

          Let SomeAreNonDemotible = False
          Let SomeAreDemotible = False
          
          For ConstraintIndex = 1 To mNumberOfConstraints
             If Stratum(ConstraintIndex) = 0 Then
                Select Case Demotable(ConstraintIndex)
                   Case False
                      Let Stratum(ConstraintIndex) = CurrentStratum
                      Let SomeAreNonDemotible = True
                   Case True
                      Let SomeAreDemotible = True
                End Select
             End If
          Next ConstraintIndex

          'Now act on the basis of these outcomes.

          If SomeAreNonDemotible = True And SomeAreDemotible = False Then

             'This means that III is true.
             'The remaining constraints forms the last stratum, and you're done.
             'They have already been assigned to the right stratum, so
             '  set the flag of triumph and quit this subroutine.
             Let FastRCD = True
             Exit Function

          ElseIf SomeAreNonDemotible = False And SomeAreDemotible = True Then
             
             'This means that II is true.
             'There is no hope of a working grammar.
             'Ad hocly call the remaining unranked constraints a 'stratum',
             '  record failure, and go home.
                  Let FastRCD = False
                  Exit Function

              'If neither I or II are true, just keep going.

          End If

      'Find out which data should be ignored henceforth, because already learned from.
      '     This occurs when the Rival candidate violates a constraint
      '       that has just been ranked into the new stratum.

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

   Loop

   'You should never, ever get this far.
        MsgBox "Program error.   I would appreciate your letting me know the about the problem.  Email me at bhayes@humnet.ucla.edu, specifying error #84505, and including a copy of your input file.", vbCritical

End Function

Public Function FastRCDWithAPrioriRankings(ByVal NumberOfForms As Long, NumberOfRivals() As Long, _
   WinnerViolations() As Long, RivalViolations() As Long) As Boolean

   'Dimension the local variables.
   
      Dim Stratum() As Long
      ReDim Stratum(mNumberOfConstraints)
      
      Dim CurrentStratum As Long
   
      Dim FormIndex As Long, RivalIndex As Long, InnerConstraintIndex As Long, OuterConstraintIndex As Long
      Dim ConstraintIndex As Long
   
      Dim SomeAreNonDemotible As Boolean, SomeAreDemotible As Boolean
      Dim StillInformative() As Boolean
      ReDim StillInformative(NumberOfForms, mMaximumNumberOfRivals)
      Dim Demotable() As Boolean
      ReDim Demotable(mNumberOfConstraints)
 
   'Initialize crucial variables, in case this routine gets called more
   '  than once.

      Let CurrentStratum = 0

      For ConstraintIndex = 1 To mNumberOfConstraints
         Let Stratum(ConstraintIndex) = 0
      Next ConstraintIndex

      For FormIndex = 1 To NumberOfForms
         For RivalIndex = 1 To mMaximumNumberOfRivals
            Let StillInformative(FormIndex, RivalIndex) = True
         Next RivalIndex
      Next FormIndex
      
   'Go through the Winner-Rival pairs repeatedly, looking for constraints that
   '  are never crucially violated among the pairs still being considered.

   Do
   
      'Record what stratum you're constructing:
         Let CurrentStratum = CurrentStratum + 1

      'Initialize a variable indicating demotability.
          For ConstraintIndex = 1 To mNumberOfConstraints
             Let Demotable(ConstraintIndex) = False
          Next ConstraintIndex

      'Go through all pairs of Winner vs. Rival, and learn from them
      '  so long as they are still in the informative class.
          
          For FormIndex = 1 To NumberOfForms
             For RivalIndex = 1 To NumberOfRivals(FormIndex)

               'Only still-informative Rivals can be learned from:
                If StillInformative(FormIndex, RivalIndex) = True Then
                   For ConstraintIndex = 1 To mNumberOfConstraints

                      'The crucial step of the algorithm:  demote constraints
                      '  that are violated in winners.  This is done only for
                      '  yet-unranked constraints.

                      If Stratum(ConstraintIndex) = 0 Then
                         If WinnerViolations(FormIndex, ConstraintIndex) > RivalViolations(FormIndex, RivalIndex, ConstraintIndex) Then
                            Let Demotable(ConstraintIndex) = True
                         End If
                      End If

                   Next ConstraintIndex
                End If
             Next RivalIndex
          Next FormIndex

        'Since this code is for when APrioriRankings are on, demote the victims.
        '   We demote those which are a priori ranked below the yet-unranked.
                For OuterConstraintIndex = 1 To mNumberOfConstraints
                    If Stratum(OuterConstraintIndex) = 0 Then
                        For InnerConstraintIndex = 1 To mNumberOfConstraints
                            If gAPrioriRankingsTable(OuterConstraintIndex, InnerConstraintIndex) = True Then
                                Let Demotable(InnerConstraintIndex) = True
                            End If
                        Next InnerConstraintIndex
                    End If                  'Is the dominee yet unranked?
                Next OuterConstraintIndex

      'You've now assessed the demotability of all yet-unranked constraints,
      '  using all the data.
      'Now, find out which constraints may be assigned to the next stratum.

      'There is also the delicate matter of knowing when you're done.  Here
      '  are the three cases:

      '  I. Some of the yet-unranked constraints are demotible.
      '        (Stratum = 0 , Demotible = True, False). Demote them and continue.

      ' II. All of the yet-unranked constraints are demotible.
      '        (Stratum = 0, Demotible = True).         Record a failed constraint
      '                                                 set and exit.

      'III. None of the yet-unranked constraints are demotible.
      '        (Stratum = 0, Demotible = false)       They are the lowest stratum
      '                                               of a working grammar.

          Let SomeAreNonDemotible = False
          Let SomeAreDemotible = False
          
          For ConstraintIndex = 1 To mNumberOfConstraints
             If Stratum(ConstraintIndex) = 0 Then
                Select Case Demotable(ConstraintIndex)
                   Case False
                      Let Stratum(ConstraintIndex) = CurrentStratum
                      Let SomeAreNonDemotible = True
                   Case True
                      Let SomeAreDemotible = True
                End Select
             End If
          Next ConstraintIndex

          'Now act on the basis of these outcomes.

          If SomeAreNonDemotible = True And SomeAreDemotible = False Then

             'This means that III is true.
             'The remaining constraints forms the last stratum, and you're done.
             'They have already been assigned to the right stratum, so
             '  set the flag of triumph and quit this subroutine.
             Let FastRCDWithAPrioriRankings = True
             Exit Function

          ElseIf SomeAreNonDemotible = False And SomeAreDemotible = True Then
             
             'This means that II is true.
             'There is no hope of a working grammar.
             'Ad hocly call the remaining unranked constraints a 'stratum',
             '  record failure, and go home.
                  Let FastRCDWithAPrioriRankings = False
                  Exit Function

              'If neither I or II are true, just keep going.

          End If

      'Find out which data should be ignored henceforth, because already learned from.
      '     This occurs when the Rival candidate violates a constraint
      '       that has just been ranked into the new stratum.

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

   Loop

   'You should never, ever get this far.
        MsgBox "Program error.   I would appreciate your letting me know the about the problem.  Email me at bhayes@humnet.ucla.edu, specifying error #87943, and including a copy of your input file.", vbCritical

End Function

  
Sub TOrder(ValhallaSize As Long, Valhalla() As Long)

    'Calculate and print out the Anttilean t-order.
    'Strategy:  loop through forms, loop through rivals for form, loop through all other forms,
    '   loop through all rivals of these forms.  When the Valhalla contains only one such rival,
    '   print out the quadruplet Input1, Rival1, Input2, Rival2.
    
        Dim RivalFound As Long
        Dim FormIndex As Long, RivalIndex As Long, InnerFormIndex As Long, InnerRivalIndex As Long, ValhallaIndex As Long
        Dim ImplicatorForm As Long, ImplicatorRival As Long, ImplicatedForm As Long, ImplicatedRival As Long
        Dim RivalIsInFactorialTypology As Boolean
        Dim TOrderFile As Long
        
        Dim Table() As String, RowCount As Long
        Dim SingleWinnerTable() As String, SingleWinnerRowCount As Long
        
        'Report progress:
            Let Form1.lblProgressWindow.Caption = "Computing t-order ..."
            DoEvents

        
        Call PrintLevel1Header(mDocFile, mTmpFile, mHTMFile, "T-orders")
        Call PrintPara(mDocFile, mTmpFile, mHTMFile, "The t-order is the set of implications in a factorical typology.")
    
        'Open an output file.
            Let TOrderFile = FreeFile
            Open gOutputFilePath + gFileName + "TOrder" + ".txt" For Output As #TOrderFile
            Print #TOrderFile, "If this input"; Chr(9); "yields this output"; Chr(9); "then this input"; Chr(9); "yields this output"
            
        'Headers for htm table.
            ReDim Table(4, 1)
            Let RowCount = 1
            Let Table(1, 1) = "If this input"
            Let Table(2, 1) = "has this output"
            Let Table(3, 1) = "then this input"
            Let Table(4, 1) = "has this output"
        
        'Arto likes to keep track of cases that aren't implicators or implicatees.
        '   Produce an array that remembers this.
            Dim NumberOfWinnersPerInput() As Long
            ReDim NumberOfWinnersPerInput(mNumberOfForms)
            Dim Implier() As Boolean
            ReDim Implier(mNumberOfForms, mMaximumNumberOfRivals)
        
        'Find candidates that always win.
            For FormIndex = 1 To mNumberOfForms
                For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                    For ValhallaIndex = 1 To ValhallaSize
                       If Valhalla(FormIndex, ValhallaIndex) = RivalIndex Then
                          Let NumberOfWinnersPerInput(FormIndex) = NumberOfWinnersPerInput(FormIndex) + 1
                          Exit For
                       End If
                    Next ValhallaIndex
                Next RivalIndex
            Next FormIndex
            
        'If necessary to announce, specify the candidates that always win and thus are not of interest to T-order.
            Dim ThereAreSingleWinnersFlag As Boolean
                
                'Assess whether we need to do this.
                    For FormIndex = 1 To mNumberOfForms
                        If NumberOfWinnersPerInput(FormIndex) = 1 Then
                            Let ThereAreSingleWinnersFlag = True
                            Exit For
                        End If
                    Next FormIndex

                'If so, do it.
                    If ThereAreSingleWinnersFlag = True Then
                        Print #TOrderFile,
                        Print #TOrderFile, "For the following input-output pairs, no other candidate ever wins, so they are not reported separately in the t-order:"
                        Print #TOrderFile, "Input"; Chr(9); "Candidate"
                        
                        Call PrintPara(mDocFile, mTmpFile, mHTMFile, "For the following input-output pairs, no other candidate ever wins, so they are not reported separately in the t-order:")
                        ReDim SingleWinnerTable(2, 1)
                        Let SingleWinnerRowCount = 1
                        Let SingleWinnerTable(1, 1) = "Input"
                        Let SingleWinnerTable(2, 1) = "Output"
                    
                        For FormIndex = 1 To mNumberOfForms
                            If NumberOfWinnersPerInput(FormIndex) = 1 Then
                                For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                                    For ValhallaIndex = 1 To ValhallaSize
                                       If Valhalla(FormIndex, ValhallaIndex) = RivalIndex Then
                                            Print #TOrderFile, mInputForm(FormIndex); Chr(9); mRival(FormIndex, RivalIndex)
                                            Call PrintPara(mDocFile, mTmpFile, -1, mInputForm(FormIndex) + " --> " + mRival(FormIndex, RivalIndex))
                                            Let SingleWinnerRowCount = SingleWinnerRowCount + 1
                                            ReDim Preserve SingleWinnerTable(2, SingleWinnerRowCount)
                                            Let SingleWinnerTable(1, SingleWinnerRowCount) = mInputForm(FormIndex)
                                            Let SingleWinnerTable(2, SingleWinnerRowCount) = mRival(FormIndex, RivalIndex)
                                          Exit For
                                       End If
                                    Next ValhallaIndex
                                Next RivalIndex
                            End If
                        Next FormIndex
                        Call s.PrintTable(mDocFile, mTmpFile, mHTMFile, SingleWinnerTable(), True, False, False)
                    End If
        
        
        'Now, the t-order itself. Go through all inputs.
            For FormIndex = 1 To mNumberOfForms
                Let ImplicatorForm = FormIndex
                For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                    'First, affirm that this rival really *is* in the factorical typology.
                        Let RivalIsInFactorialTypology = False
                        For ValhallaIndex = 1 To ValhallaSize
                            If Valhalla(ImplicatorForm, ValhallaIndex) = RivalIndex Then
                                Let RivalIsInFactorialTypology = True
                                Let ImplicatorRival = RivalIndex
                                Exit For
                            End If
                        Next ValhallaIndex
                    'If you can so affirm, you can search the other inputs and rivals for implications.
                        If RivalIsInFactorialTypology Then
                            For InnerFormIndex = 1 To mNumberOfForms
                                'Search all the forms except for the implicator.  Also, avoid inputs that have just one output.
                                    If InnerFormIndex <> ImplicatorForm And NumberOfWinnersPerInput(InnerFormIndex) > 1 Then
                                        Let ImplicatedForm = InnerFormIndex
                                        'Now look at the rivals for this form.  We want to find if there is one unique rival that is
                                        '   derived when ImplicatorForm and ImplicatorRival are set as they currently are.
                                            For InnerRivalIndex = 1 To mNumberOfRivals(InnerFormIndex)
                                                Let ImplicatedRival = 0
                                                For ValhallaIndex = 1 To ValhallaSize
                                                    'We only want to consider valhalla members that have derived ImplicatorRival from ImplicatorForm.
                                                        If Valhalla(ImplicatorForm, ValhallaIndex) = ImplicatorRival Then
                                                            'Install the first one you find.
                                                                If ImplicatedRival = 0 Then
                                                                    Let ImplicatedRival = Valhalla(ImplicatedForm, ValhallaIndex)
                                                                Else
                                                                    'Exit if this is different from the first one you found.
                                                                    If Valhalla(ImplicatedForm, ValhallaIndex) <> ImplicatedRival Then
                                                                        GoTo ExitPoint
                                                                    End If
                                                                End If
                                                        End If              'Does this valhalla member derived ImplicatorRival from ImplicatorForm?
                                                Next ValhallaIndex
                                            'Hooray, all members of the factorial typology that have ImplicatorForm --> ImplicatorRival
                                            '   also have ImplicatedForm --> ImplicatedRival.  This is an authentic implication.
                                                'Record it.
                                                    Let RowCount = RowCount + 1
                                                    ReDim Preserve Table(4, RowCount)
                                                    Let Table(1, RowCount) = "/" + mInputForm(ImplicatorForm) + "/"
                                                    Let Table(2, RowCount) = "[" + mRival(FormIndex, ImplicatorRival) + "]"
                                                    Let Table(3, RowCount) = "/" + mInputForm(ImplicatedForm) + "/"
                                                    Let Table(4, RowCount) = "[" + mRival(ImplicatedForm, ImplicatedRival) + "]"
                                                    
                                                    Print #TOrderFile, mInputForm(ImplicatorForm); Chr(9); mRival(FormIndex, ImplicatorRival); Chr(9); _
                                                        mInputForm(ImplicatedForm); Chr(9); mRival(ImplicatedForm, ImplicatedRival)
                                                'Record that the participants are not isolates in the system of implications.
                                                    'Let ImplierOrImpliee(ImplicatorForm, ImplicatorRival) = True
                                                    'Let ImplierOrImpliee(ImplicatedForm, ImplicatedRival) = True
                                                    Let Implier(ImplicatorForm, ImplicatorRival) = True
                                                Exit For
ExitPoint:
                                            Next InnerRivalIndex    'Go through all candidates for this input.
                                    End If                          'Don't compare an input with itself.
                            Next InnerFormIndex                     'Go through all inputs.
                        End If          'Was this rival in the factorial typology?
                Next RivalIndex         'Go through all candidate outputs for this input.
            Next FormIndex              'Go through all inputs.
            Call s.PrintTable(mDocFile, mTmpFile, mHTMFile, Table(), True, False, False)
        
        'Print out the list of rivals that are not implicators or implicated.
            
            'First decide if this is even necessary.
                Dim IAmNecessary As Boolean
                For FormIndex = 1 To mNumberOfForms
                    For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                        If Implier(FormIndex, RivalIndex) = False Then
                            Let IAmNecessary = True
                            GoTo ExitPointIAmNecessary
                        End If
                    Next RivalIndex
                Next FormIndex
ExitPointIAmNecessary:
                
            If IAmNecessary Then
                Print #TOrderFile,
                Print #TOrderFile, "Nothing is implicated by these input-output pairs:"
                Print #TOrderFile, "Input"; Chr(9); "Candidate"
                
                Call PrintPara(mDocFile, mTmpFile, mHTMFile, "Nothing is implicated by these input-output pairs:")
                ReDim Table(2, 1)
                Let RowCount = 1
                Let Table(1, 1) = "Input"
                Let Table(2, 1) = "Candidate"
                
                For FormIndex = 1 To mNumberOfForms
                    For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                        If Implier(FormIndex, RivalIndex) = False Then
                            Print #TOrderFile, mInputForm(FormIndex); Chr(9); mRival(FormIndex, RivalIndex)
                            'Call PrintPara(mDocFile, mTmpFile, -1, mInputForm(FormIndex) + " --> " + mRival(FormIndex, RivalIndex))
                            Let RowCount = RowCount + 1
                            ReDim Preserve Table(2, RowCount)
                            Let Table(1, RowCount) = mInputForm(FormIndex)
                            Let Table(2, RowCount) = mRival(FormIndex, RivalIndex)
                        End If
                    Next RivalIndex
                Next FormIndex
                Call s.PrintTable(mDocFile, mTmpFile, mHTMFile, Table(), True, False, False)

            End If
        
        Call PrintPara(mDocFile, mTmpFile, mHTMFile, "For a tabbed listing of the t-order found here, see the file PARA" + gOutputFilePath + gFileName + ".")

        
        Close #TOrderFile
        
End Sub

