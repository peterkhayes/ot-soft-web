VERSION 5.00
Begin VB.Form MyMaxEnt 
   Caption         =   "Maximum entropy"
   ClientHeight    =   6615
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   8655
   Icon            =   "MyMaxEnt.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6615
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmWeightLimitations 
      Caption         =   "Limitations on weights"
      Height          =   2055
      Left            =   480
      TabIndex        =   7
      Top             =   3840
      Width           =   3255
      Begin VB.TextBox txtWeightMaximum 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Text            =   "50"
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox txtWeightMinimum 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Text            =   "0"
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Maximum"
         Height          =   375
         Left            =   960
         TabIndex        =   11
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Minimum"
         Height          =   375
         Left            =   960
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.TextBox txtDecimalPlaces 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Text            =   "3"
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txtPrecision 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Text            =   "5"
      Top             =   2280
      Width           =   855
   End
   Begin VB.PictureBox pctProgressWindow 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   4080
      ScaleHeight     =   6315
      ScaleWidth      =   4395
      TabIndex        =   2
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit to main screen"
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   3375
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run maxent"
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Report results to this many decimal places (blank = full report)"
      Height          =   495
      Left            =   1080
      TabIndex        =   6
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Number of iteration: more is more accurate"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuIncludeTableaux 
         Caption         =   "Include tableaux"
      End
      Begin VB.Menu mnuHistory 
         Caption         =   "Generate file with history of weights"
      End
      Begin VB.Menu mnuOutputHistoryFile 
         Caption         =   "Generate file with history of output probabilities"
      End
   End
   Begin VB.Menu mnuFactorialTypology 
      Caption         =   "Factorial typology"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuExploreFactoricalTypology 
         Caption         =   "Explore factorial typology through sampling"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuViewManualAsPDF 
         Caption         =   "View manual as pdf"
      End
   End
End
Attribute VB_Name = "MyMaxEnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=================================MyMaxEnt.FRM=================================
'======Weight a set of constraints given a set of input data with violations====
   
   Option Explicit

        Const mReportInterval As Long = 10000
    
    'Module-level variables for the basic data
        Dim mInputForm() As String
        Dim mWinner() As String
        Dim mNumberOfForms As Long
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
        
    'Output files.
        Dim mTmpFile As Long, mDocFile As Long, mHTMFile As Long, mTabbedFile As Long
  
    'Grammar:
        Dim mWeight() As Single
        Dim mInitialWeight() As Single
   
    'Learning:
        Dim mMinimumWeight As Single
        Dim mMaximumWeight As Single
        Dim mObservedViolations() As Single
        Dim mExpectedViolations() As Single
        Dim mSigmaSquared As Single
        Dim mSlowingFactor As Single
        Dim mUsingPriorTerm As Boolean
        Dim eHarmony() As Double    'Colin Wilson's silly but useful name.
   
   'Frequencies of forms.
        Dim mFrequency() As Single
        Dim mObservedInputProportions() As Single
        Dim mObservedRivalProportions() As Single
        Dim mPredictedRivalProportions() As Single
        Dim mTotalFrequencyForEachInput() As Single
        Dim mTotalForms As Single

    'Testing the grammar:
        Dim mErrorTerm As Single
        
    'Exploring factorial typology
        Dim mTotalNumberOfCandidates As Long
        Dim mCandidatesInOneArray() As Long
        Dim mLowestFrequencyAssigned() As Single
        Dim mHighestFrequencyAssigned() As Single
        Dim mTotalFrequencyAssigned() As Single
        Dim mProbabilityDifferenceLow() As Single
        Dim mProbabilityDifferenceHigh() As Single

    'Reporting
        Dim mSlotFiller() As Long

    'Variables for designating files:
        Dim mWeightHistoryFile As Long, mFullHistoryFile As Long, mOutputHistoryFile As Long

    'Final details:
        Dim mTimeMarker As Single
        
    'KZ: keeps track of whether or not user has cancelled learning
        Dim mblnProcessing As Boolean
    
    'KZ: in case user cancels algorithm but then runs it again without exiting the form, the output files need
    '   to be reopened.
        Dim blnOutputFilesOpen As Boolean
    

'==================================INTERFACE ITEMS==========================================

Private Sub Form_Load()
    
    Let blnOutputFilesOpen = True 'KZ: the files were opened in Main.
    
    'Center the form and put a caption on it.
        Let Me.Top = (Screen.Height - Me.Height) / 2
        Let Me.Left = (Screen.Width - Me.Width) / 2
        Let Me.Caption = "OTSoft " + gMyVersionNumber + " - MaxEnt - " + gFileName + gFileSuffix
        
    'Grab the number of training iterations from the .ini file.
        Let txtPrecision = Trim(Str(gNumberOfDataPresentations))
        
    'Let include tableaux be default.
        Let mnuIncludeTableaux.Checked = True
        
    'In fact, disk space is cheap these days, no?
        Let mnuHistory.Checked = True
        Let mnuOutputHistoryFile.Checked = True
    
        
End Sub


Private Sub Form_Resize()

    'Resize the progress window.
    '   Don't let it get too small, though; or program crashes.

    Dim NewWidth As Long, NewHeight As Long

    Let NewWidth = Me.Width - pctProgressWindow.Left - 390
    If NewWidth < 200 Then Let NewWidth = 200
    Let NewHeight = Me.Height - pctProgressWindow.Top - 780
    If NewHeight < 200 Then Let NewHeight = 200
    Let pctProgressWindow.Width = NewWidth
    Let pctProgressWindow.Height = NewHeight

End Sub



'--------------------------------------Help Menu-------------------------------------------

    'These simply call the analogous Form1 procedures.

Private Sub mnuViewManualAsPDF_Click()
    Call Form1.mnuOpenHelpAsPDF_Click
End Sub

Private Sub mnuViewManualAsWordFile_Click()
    Call Form1.mnuViewManual_Click
End Sub

Sub mnuAboutOTSoft_Click()
    Call Form1.mnuAboutOTSoft_Click
End Sub

Private Sub mnuWeightHistory_Click()
    'Ask for a wight-history file.
    
    If mnuHistory.Checked = False Then
        Let mnuHistory.Checked = True
    Else
        Let mnuHistory.Checked = False
    End If
End Sub

Private Sub mnuOutputHistory_Click()
    'Ask for a wight-history file.
    
    If mnuOutputHistoryFile.Checked = False Then
        Let mnuOutputHistoryFile.Checked = True
    Else
        Let mnuOutputHistoryFile.Checked = False
    End If
End Sub

'------------------------------Rest of Maxent Interface--------------------------------

Private Sub cmdExit_Click()

    'KZ:  This button getting clicked is an option only if the user
    '   changes her mind and wants to apply a different algorithm,
    '   or use a different file, or quit (i.e., before running the learner).
    '   See form_Unload: In order to allow such mind-changing,
    '   all files in use have to be closed:

    'Reactivate the Rank and Factorial Typology command buttons so they'll
    '   be ready.
        Let Form1.cmdRank.Enabled = True
        Let Form1.cmdFacType.Enabled = True
        
        Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'KZ: must close all files so user can start fresh and apply a different analysis if she wants.
        Close
End Sub


'================================MAIN CALCULATION ROUTINE====================================

Sub Main(NumberOfForms As Long, InputForm() As String, _
    Winner() As String, WinnerFrequency() As Single, WinnerViolations() As Long, _
    NumberOfRivals() As Long, Rival() As String, RivalFrequency() As Single, RivalViolations() As Long, _
    NumberOfConstraints As Long, ConstraintName() As String, Abbrev() As String, _
    TmpFile As Long, DocFile As Long, HTMFile As Long)

    'This routine imports the data variables and installs them as module level.
    
        Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
    
        'Localize input parameters as module level variables.
            Let mNumberOfConstraints = NumberOfConstraints
            ReDim mAbbrev(mNumberOfConstraints)
            ReDim mConstraintName(mNumberOfConstraints)
            For ConstraintIndex = 1 To NumberOfConstraints
                Let mConstraintName(ConstraintIndex) = ConstraintName(ConstraintIndex)
                Let mAbbrev(ConstraintIndex) = Abbrev(ConstraintIndex)
            Next ConstraintIndex
            Let mNumberOfForms = NumberOfForms
            ReDim mInputForm(mNumberOfForms)
            ReDim mWinner(mNumberOfForms)
            ReDim mWinnerFrequency(mNumberOfForms)
            ReDim mWinnerViolations(mNumberOfForms, mNumberOfConstraints)
            Let mMaximumNumberOfRivals = Form1.FindMaximumNumberOfRivals(NumberOfRivals())
            ReDim mRival(mNumberOfForms, mMaximumNumberOfRivals)
            ReDim mRivalFrequency(mNumberOfForms, mMaximumNumberOfRivals)
            ReDim mRivalViolations(mNumberOfForms, mMaximumNumberOfRivals, mNumberOfConstraints)
            ReDim mNumberOfRivals(mNumberOfForms)
            For FormIndex = 1 To mNumberOfForms
                Let mInputForm(FormIndex) = InputForm(FormIndex)
                Let mWinner(FormIndex) = Winner(FormIndex)
                Let mWinnerFrequency(FormIndex) = WinnerFrequency(FormIndex)
                Let mNumberOfRivals(FormIndex) = NumberOfRivals(FormIndex)
                For ConstraintIndex = 1 To mNumberOfConstraints
                    Let mWinnerViolations(FormIndex, ConstraintIndex) = WinnerViolations(FormIndex, ConstraintIndex)
                Next ConstraintIndex
                For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                    Let mRival(FormIndex, RivalIndex) = Rival(FormIndex, RivalIndex)
                    Let mRivalFrequency(FormIndex, RivalIndex) = RivalFrequency(FormIndex, RivalIndex)
                    For ConstraintIndex = 1 To mNumberOfConstraints
                        Let mRivalViolations(FormIndex, RivalIndex, ConstraintIndex) = RivalViolations(FormIndex, RivalIndex, ConstraintIndex)
                        'Debugging:  it appears the first constraint has no violations.
                        'If ConstraintIndex = 1 Then MsgBox Str(RivalViolations(FormIndex, RivalIndex, ConstraintIndex))
                    Next ConstraintIndex
                Next RivalIndex
            Next FormIndex
            Let mTmpFile = TmpFile
            Let mDocFile = DocFile
            Let mHTMFile = HTMFile
        
        
                    'Debug
                    'Dim DebugFile As Long
                    'Let DebugFile = FreeFile
                    'Open gOutputFilePath + "DebugAtEarlyStageOfMaxentRoutine.txt" For Output As #DebugFile
                    'Dim i As Long, j As Long
                    'For i = 1 To mNumberOfForms
                    '    Print #DebugFile, "Input:"; vbTab; mInputForm(i)
                    '    Print #DebugFile, vbTab; "Winner:  ["; vbTab; mWinner(i); "]"; vbTab; "Frequency"; vbTab; mWinnerFrequency(i)
                    '    For j = 1 To mNumberOfRivals(mNumberOfForms)
                    '        Print #DebugFile, vbTab; "RivalIndex:"; vbTab; Trim(Str(j)); vbTab; "Rival:"; vbTab; mRival(i, j); vbTab; "Frequency:"; vbTab; mRivalFrequency(i, j)
                    '    Next j
                    '    Print #DebugFile,
                    'Next i
                    'Close #DebugFile
        
        
    'Ready to go; show the form to the user.
        MyMaxEnt.Show
        

End Sub


Private Sub cmdRun_Click()

    'This is the primary routine for this form, which calls all other routines needed for running maxent.
    
    Dim FormIndex As Long, ConstraintIndex As Long
    
    'MsgBox Str(mHTMFile)
    
    'First, get crucial parameters from the interface.
    '   The routine ObtainValuesFromInterface does just that, but returns False if there is an error.
    '   If so, return the interface to its default state so the user can correct the error.
        If ObtainValuesFromInterface = False Then Exit Sub
        
                    'Debug
                    'Dim DebugFile As Long
                    'Let DebugFile = FreeFile
                    'Open gOutputFilePath + "DebugMaxentJustAfterClicking.txt" For Output As #DebugFile
                    'Dim i As Long, j As Long
                    'For i = 1 To mNumberOfForms
                    '    Print #DebugFile, "Input:"; vbTab; mInputForm(i)
                    '    Print #DebugFile, vbTab; "Winner:  ["; vbTab; mWinner(i); "]"; vbTab; "Frequency"; vbTab; mWinnerFrequency(i)
                    '    For j = 1 To mNumberOfRivals(mNumberOfForms)
                    '        Print #DebugFile, vbTab; "RivalIndex:"; vbTab; Trim(Str(j)); vbTab; "Rival:"; vbTab; mRival(i, j); vbTab; "Frequency:"; vbTab; mRivalFrequency(i, j)
                    '    Next j
                    '    Print #DebugFile,
                    'Next i
                    'Close #DebugFile
        
        
    'Since the user has clicked a Rank button, (s)he probably wants the settings saved.
        Call Form1.SaveUserChoices
    
    'Dimension arrays needed by the algorithm, according to the size of the problem:
        ReDim mWeight(mNumberOfConstraints)
        ReDim mInitialWeight(mNumberOfConstraints)
        ReDim mTotalFrequencyForEachInput(mNumberOfForms)
        ReDim mObservedViolations(mNumberOfConstraints)
        ReDim mExpectedViolations(mNumberOfConstraints)
        ReDim mPredictedRivalProportions(mNumberOfForms, mMaximumNumberOfRivals)
        ReDim mFrequency(mNumberOfForms, mMaximumNumberOfRivals)
        ReDim mFaithfulness(mNumberOfConstraints)
        ReDim mObservedRivalProportions(mNumberOfForms, mMaximumNumberOfRivals)
        ReDim mObservedInputProportions(mNumberOfForms)

    'KZ: the variable mblnProcessing keeps track of whether or not the algorithm is running:
    
    If mblnProcessing = True Then
        'KZ: when the algorithm is running, clicking this button again stops it
            Let mblnProcessing = False 'KZ: return button to "not-running" state and shut down the learner
    Else    'KZ: otherwise, run the algorithm
        
        'KZ: let the user know this is now the cancel button:
            cmdRun.Caption = "Cancel"
            mblnProcessing = True
        
        'Clear the progress window.
            pctProgressWindow.Cls
        'Start timing.
            Let mTimeMarker = Timer
        'Report progress.
            pctProgressWindow.Print "Learning..."
        
        'MaxEntPreliminaries:
        'If all goes well in the preliminary operations, MaxEntPreliminaries() will return True, and you can continue.
            If MaxEntPreliminaries() = False Then
                'Something went wrong in the preliminaries.  Go back to ur-state.
                    Let mblnProcessing = False
                'Reactivate the various buttons that start things off.
                    Let Form1.cmdRank.Enabled = True
                    Let Form1.cmdFacType.Enabled = True
                    Let cmdRun.Caption = "&Run Maxent"
                Exit Sub
            End If
            
        Call PrintAHeader("Maximum Entropy")
    
        'At the moment, two or three duelling versions.
        'This is the stage where the global variables get localized for this routine.
         
         Call MaxEntCore(mInputForm())
         
            'KZ: This algorithm checks periodically to see if the button
            'has been clicked again (to cancel). If user cancels,
            'mblnProcessing gets set to False.  BH:  not yet. xxx
        
        'KZ: don't bother with this part if user has cancelled:
            If mblnProcessing = True Then
                Call PrintMaxentResults(mInputForm(), mWeight(), "Weights")
                Call PrintTableaux(mNumberOfConstraints, mAbbrev(), mWeight(), mNumberOfForms, mInputForm(), mNumberOfRivals(), mRival(), mRivalViolations(), mObservedRivalProportions(), _
                    eHarmony(), mPredictedRivalProportions(), mDocFile, mTmpFile, mHTMFile)
                Call PrintFinalDetails
                Close
            End If
        
        'Close output files.
            Close #mTmpFile
            Close #mDocFile
            'Print #mHTMFile, "</BODY>"
            Close #mHTMFile

            Let blnOutputFilesOpen = False 'KZ
            
        'KZ: get rid of "cancel" caption:
            Let cmdRun.Caption = "&Run Maxent"
        
        If mblnProcessing = True Then 'KZ: only do this if still processing
            Let mblnProcessing = False   'KZ
            'Announce you're done.
                Let Form1.lblProgressWindow.Caption = "I'm done."
                'Reactivate the Rank and Factorial Typology buttons of the main form.
                    Let Form1.cmdRank.Enabled = True
                    Let Form1.cmdFacType.Enabled = True
            'Guide user to the View Results button.
                Let Form1.cmdViewResults.Default = True
                Let Form1.cmdViewResults.FontSize = 10
                Let Form1.cmdViewResults.FontBold = True
            'Get ready to View Results.
                Form1.cmdViewResults.SetFocus
                Let gHasTheProgramBeenRun = True
            'Get out.
                Unload Me
        Else
            'KZ: extra stuff to do if user cancelled:
                pctProgressWindow.Cls
                pctProgressWindow.Print "Learning Cancelled."
                Close #mWeightHistoryFile
                Close #mFullHistoryFile
        End If
        
    End If  'Is mblnProcessing true?  (i.e. Rank or Cancel)

End Sub


Function ObtainValuesFromInterface() As Boolean
    
    'Precision value and sigma from the interface.
        Let gNumberOfDataPresentations = Val(txtPrecision.Text)
    
    'Prior term:  Turned off, with no interface access.  I would like to retry it.
    '    If Trim(txtSigma.Text) = "" Then
            
    'RESTORE ME LATER:
            
            Let mUsingPriorTerm = False
    
    'For now:
        'Let mUsingPriorTerm = True
        Let mSigmaSquared = 1
        'Old"
        'Let mSigmaSquared = 10 ^ Val(txtSigma.Text)
    
    
    'Lower bound on weights
        If Trim(txtWeightMinimum) = "" Then
            MsgBox "Please specify a minimum weight.  A typically-used value is zero.", vbExclamation
            Let ObtainValuesFromInterface = False
            Exit Function
        Else
            Let mMinimumWeight = Val(Trim(txtWeightMinimum))
        End If
    'Upper bound on weights
        If Trim(txtWeightMaximum) = "" Then
            MsgBox "Please specify a maximum weight.  A typically-used value is 50.", vbExclamation
            Let ObtainValuesFromInterface = False
            Exit Function
        Else
            Let mMaximumWeight = Val(Trim(txtWeightMaximum.Text))
        End If
        
    'All is well, so the program may proceed.
        Let ObtainValuesFromInterface = True

End Function

Function MaxEntPreliminaries() As Boolean

   Dim ConstraintIndex As Long
   
   'Execute preliminary actions needed by the maxent algorithm.
   
        Call InstallWinnerAsRival
        Call CalculateObservedProportions
        Call CalculateObservedViolations
        Call InitialWeights
        Call CalculateSlowingFactor

        If mnuHistory.Checked = True Or mnuOutputHistoryFile.Checked = True Then Call SetUpHistory
    
    'All is well, so set value as True and exit.
        Let MaxEntPreliminaries = True
        Exit Function

End Function

Sub SetUpHistory()

    'Setting up tracing of weights and outputs.
        
        Dim ConstraintIndex As Long, FormIndex As Long, RivalIndex As Long
        
    'Weights:
    If mnuHistory.Checked = True Then
        'First, make sure there is a folder for these files, a daughter of the
        '   folder in which the input file is located.
            Call Form1.CreateAFolderForOutputFiles
        'Now open the file.
            Let mWeightHistoryFile = FreeFile
            Open gOutputFilePath + gFileName + "HistoryOfWeights" + ".txt" For Output As #mWeightHistoryFile
            For ConstraintIndex = 1 To mNumberOfConstraints
                Print #mWeightHistoryFile, Chr$(9); mAbbrev(ConstraintIndex);
            Next ConstraintIndex
            Print #mWeightHistoryFile,
            'Also, the initial weights:
                Print #mWeightHistoryFile, "0";
                For ConstraintIndex = 1 To mNumberOfConstraints
                   Print #mWeightHistoryFile, Chr$(9); mInitialWeight(ConstraintIndex);
                Next ConstraintIndex
                Print #mWeightHistoryFile,
    End If
    
    'Outputs:
    If mnuOutputHistoryFile.Checked = True Then
        'First, make sure there is a folder for these files, a daughter of the
        '   folder in which the input file is located.
            Call Form1.CreateAFolderForOutputFiles
        'Now open the file.
            Let mOutputHistoryFile = FreeFile
            Open gOutputFilePath + gFileName + "HistoryOfOutputProbabilities" + ".txt" For Output As #mOutputHistoryFile
        'Header:  everything horizontal
            For FormIndex = 1 To mNumberOfForms
                If FormIndex > 1 Then
                    Print #mOutputHistoryFile, vbTab;
                End If
                Print #mOutputHistoryFile, mInputForm(FormIndex);
                For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                    Print #mOutputHistoryFile, vbTab; mRival(FormIndex, RivalIndex);
                Next RivalIndex
            Next FormIndex
            Print #mOutputHistoryFile,
    End If

End Sub


Sub InstallWinnerAsRival()

    'The file-reading apparatus creates separate arrays for winners and rivals.
    '   Amalgamate these into just rivals, with the "winners" (which may be actually tied) in row zero.
    '   Also, amalgamate the WinnerFrequency and RivalFrequency arrays, which were earlier read separately.
    
    Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
    
    For FormIndex = 1 To mNumberOfForms
        Let mFrequency(FormIndex, 0) = mWinnerFrequency(FormIndex)
        Let mRival(FormIndex, 0) = mWinner(FormIndex)
        For ConstraintIndex = 1 To mNumberOfConstraints
            Let mRivalViolations(FormIndex, 0, ConstraintIndex) = mWinnerViolations(FormIndex, ConstraintIndex)
        Next ConstraintIndex
        For RivalIndex = 1 To mNumberOfRivals(FormIndex)
            Let mFrequency(FormIndex, RivalIndex) = mRivalFrequency(FormIndex, RivalIndex)
        Next RivalIndex
    Next FormIndex
    
                    'Debug
                    'Dim DebugFile As Long
                    'Let DebugFile = FreeFile
                    'Open gOutputFilePath + "DebugAtInstallWinnerAsRival.txt" For Output As #DebugFile
                    'Dim i As Long, j As Long
                    'For i = 1 To mNumberOfForms
                    '    Print #DebugFile, "Input:"; vbTab; mInputForm(i)
                    '    Print #DebugFile, vbTab; "Winner:  ["; vbTab; mWinner(i); "]"; vbTab; "Frequency"; vbTab; mWinnerFrequency(i)
                    '    For j = 0 To mNumberOfRivals(mNumberOfForms)
                    '        Print #DebugFile, vbTab; "RivalIndex:"; vbTab; Trim(Str(j)); vbTab; "Rival:"; vbTab; mRival(i, j); vbTab; "Frequency:"; vbTab; mRivalFrequency(i, j)
                    '    Next j
                    '    Print #DebugFile,
                    'Next i
                    'Close #DebugFile
    
    
End Sub

Sub InitialWeights()

    'Set the weights at their initial value.
    
    Dim ConstraintIndex As Long
    
    'The weight Colin uses is 1.
    '0 produces more legible grammars, since worthless constraints don't have to be demoted within the
    '   time available.
        For ConstraintIndex = 1 To mNumberOfConstraints
            'Let mWeight(ConstraintIndex) = 1
            Let mWeight(ConstraintIndex) = 0
        Next ConstraintIndex
    
    'Save the initial ranking values for future reporting.
        For ConstraintIndex = 1 To mNumberOfConstraints
            Let mInitialWeight(ConstraintIndex) = mWeight(ConstraintIndex)
        Next ConstraintIndex
    
End Sub

Sub CalculateObservedProportions()

    'We need to know the aggregated output frequency across candidates for each input, and what proportion of each input's share
    '   is taken by each candidate.
    
    Dim FormIndex As Long, RivalIndex As Long
    
        For FormIndex = 1 To mNumberOfForms
            'Total over this form.
                'August 2025:  I don't know why this should start at zero; it adds a bad candidate
                'For RivalIndex = 0 To mNumberOfRivals(FormIndex)
                For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                    Let mTotalFrequencyForEachInput(FormIndex) = mTotalFrequencyForEachInput(FormIndex) + mFrequency(FormIndex, RivalIndex)
                Next RivalIndex
            'Increment the grand total.
                Let mTotalForms = mTotalForms + mTotalFrequencyForEachInput(FormIndex)
            'Find the proportions within this form, avoiding crash if the denominator is zero.
                If mTotalFrequencyForEachInput(FormIndex) = 0 Then
                    'For RivalIndex = 0 To mNumberOfRivals(FormIndex)
                    For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                        Let mObservedRivalProportions(FormIndex, RivalIndex) = 0
                    Next RivalIndex
                Else
                    'For RivalIndex = 0 To mNumberOfRivals(FormIndex)
                    For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                        Let mObservedRivalProportions(FormIndex, RivalIndex) = mFrequency(FormIndex, RivalIndex) / mTotalFrequencyForEachInput(FormIndex)
                    Next RivalIndex
                End If
        Next FormIndex
        
        'Warn the user and exit if there are no data.
            If mTotalForms = 0 Then
                MsgBox "Sorry, but the total frequency of forms in your data is zero, making learning impossible.  Please correct your input file before running the program, which will now exit.", vbExclamation
                Close
                End
            End If
        
        'Since there are some forms, you can compute the proportion per input.
            For FormIndex = 1 To mNumberOfForms
                Let mObservedInputProportions(FormIndex) = mTotalFrequencyForEachInput(FormIndex) / mTotalForms
            Next FormIndex
    
End Sub

Sub CalculateObservedViolations()

    'Following Goodman (2002), this is done with raw frequency, not proportions.
    
    Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
    
    For FormIndex = 1 To mNumberOfForms
        'August 2025; why start at zero?  Winners are not read separately any more.
        'For RivalIndex = 0 To mNumberOfRivals(FormIndex)
        For RivalIndex = 1 To mNumberOfRivals(FormIndex)
            For ConstraintIndex = 1 To mNumberOfConstraints
                Let mObservedViolations(ConstraintIndex) = mObservedViolations(ConstraintIndex) + mFrequency(FormIndex, RivalIndex) * mRivalViolations(FormIndex, RivalIndex, ConstraintIndex)
            Next ConstraintIndex
        Next RivalIndex
    Next FormIndex
    
End Sub

Sub CalculateSlowingFactor()

    'The slowing factor, in Generalized Iterative Scaling, is the greatest number
    '   of total violations (summed across constraints) for any candidate.
    '   Source:  Joshua Goodman, ACL 2002.
    
    Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
    Dim Buffer As Long
    
    'The original guys cited by Goodman, Chen and Rosenfeld, use a constraint-specified slowing factor.
    '   Might this help?
         
    For FormIndex = 1 To mNumberOfForms
        'We use zero as start value for this loop, since the winner has been folded in as rival #0.
        'August 2025:  why start at zero?
        'For RivalIndex = 0 To mNumberOfRivals(FormIndex)
        For RivalIndex = 1 To mNumberOfRivals(FormIndex)
            Let Buffer = 0
            'Sum the violations for this candidate.
                For ConstraintIndex = 1 To mNumberOfConstraints
                    Let Buffer = Buffer + mRivalViolations(FormIndex, RivalIndex, ConstraintIndex)
                Next ConstraintIndex
            'Buffer now contains the total number of violations for this candidate.  Install it
            '   as the new mSlowingFactor if it is bigger than the existing one.
                If Buffer > mSlowingFactor Then
                    Let mSlowingFactor = Buffer
                End If
        Next RivalIndex
    Next FormIndex
    
End Sub


Sub MaxEntCore(mInputForm() As String)

    'The core algorithm.
        
    'The amount that a weight gets adjusted.
        Dim Delta As Single

    'Variables for where we are in the run.
        Dim LearningStageIndex As Long
        
    'Probability variables
        Dim MySlope As Single
        Dim MyProbability As Single
        Dim MyObjectiveFunction As Single
        Dim StoredObjectiveFunction As Single
        Dim SquaredSumOfWeights As Single
        
    'Reporting progress
        Dim NoChangeSequence As Long            'Detect long static period, indicating an exit is appropriate.
        
    'Indices etc.:
        Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
        Dim DummyCounter As Integer     'KZ: every time gets to gReportingFrequency it checks if
                                        'the user pressed cancel.  BH:  needs to be implemented here.
                                    
    
    'Prepare to debug:
        'Dim DebugFile As Long
        'Let DebugFile = FreeFile
        'Open gInputFilePath + "\debug.txt" For Output As #DebugFile
    
    'Print header for debugging file.
        'For ConstraintIndex = 1 To mNumberOfConstraints
        '    Print #DebugFile, Chr(9); mAbbrev(ConstraintIndex);
        'Next ConstraintIndex
        'Print #DebugFile,
        'Print #DebugFile, "observed";
        'For ConstraintIndex = 1 To mNumberOfConstraints
        '    Print #DebugFile, Chr(9); mObservedViolations(ConstraintIndex);
        'Next ConstraintIndex
        'Print #DebugFile,
        'End
        
    'Do as many learning stages as there are.
        
        Let StoredObjectiveFunction = ObjectiveFunction()
        
        For LearningStageIndex = 1 To gNumberOfDataPresentations
        
            If mblnProcessing = False Then Exit Sub
        
            Call CalculateExpectedViolations(LearningStageIndex)
            
            'Modify the weight of each constraint according to its Observed and Expected violation counts.
                'Avoid negative weights unless user specifies.
                   
             For ConstraintIndex = 1 To mNumberOfConstraints
                'Calculate the change that must be made to the weight of this constraint.
                '   This comes in two flavors, depending on whether a Gaussian prior is being used.
                   
                   If mUsingPriorTerm Then
                        'Call CalculateExpectedViolations
                            Let Delta = DeltaUsingPrior(ConstraintIndex)
                        'Perform the update.  xxx What is this 1000 for?
                             Let mWeight(ConstraintIndex) = mWeight(ConstraintIndex) - Delta / 1000
                   Else
                        'The basic update rule is taken from Goodman (2002, 10; ex. (2)).
                        'He, and other people, don't say what to do when there are
                        '   no observed violations.  What is the standard approach to this problem?
                        '   At the moment, we're assuming that unviolated constraints are given the maximum
                        '   weight, so we can just ignore them here.
                                                        
                            'I think this is actually a bug (12/2/25):  you have to let the constraint become very powerful,
                            '   so give a fictional very low positive value.
                            If mObservedViolations(ConstraintIndex) = 0 Then
                                Let mObservedViolations(ConstraintIndex) = 0.000000001
                            End If
                            
                            'Now, the update:
                                Let Delta = Log((mObservedViolations(ConstraintIndex)) _
                                    / (mExpectedViolations(ConstraintIndex))) / mSlowingFactor
                                Let mWeight(ConstraintIndex) = mWeight(ConstraintIndex) - Delta
                   End If
                   
                'Some notes from Claire, 9/5/25:
                'L1:
                '   at each update, take the modified weight and subtract from it DecayRate * w, where
                '       DecayRate:  how strong is the prior?  Hard to tell if it is related to good old sigma and mu.
                '       w is the constraint weight
                
                'L2:
                '   at each update, take the modified weight and subtract from it DecayRate / 2 * w^2, where
                
                'These hold of zero mu; for non-zero mu, replace with w - mu
                
                'Don't let an update create a weight below the specified minimum (from interface)
                    If mWeight(ConstraintIndex) < mMinimumWeight Then
                        Let mWeight(ConstraintIndex) = mMinimumWeight
                    End If
                    
                'Prevent crashes with weight maximum.
                    If mWeight(ConstraintIndex) > mMaximumWeight Then
                        Let mWeight(ConstraintIndex) = mMaximumWeight
                    End If
                    
             Next ConstraintIndex
             
             'Ok, you've just changed the weights, so report.
                If mnuHistory.Checked = True Then
                    Print #mWeightHistoryFile, Trim(Str(LearningStageIndex));
                    For ConstraintIndex = 1 To mNumberOfConstraints
                        Print #mWeightHistoryFile, Chr$(9); mWeight(ConstraintIndex);
                    Next ConstraintIndex
                    Print #mWeightHistoryFile,
                End If
             
            'Debug:  Is this helping?
                GoTo ResumePoint
                'Compute the objective function.
                    Let MyObjectiveFunction = ObjectiveFunction()
                'Report nonimprovements, for debugging.
                    If MyObjectiveFunction < StoredObjectiveFunction Then
                        Let NoChangeSequence = 0
                    ElseIf MyObjectiveFunction = StoredObjectiveFunction Then
                        'MsgBox "Iteration " + Str(LearningStageIndex) + ":  last step made no improvement."
                        Let NoChangeSequence = NoChangeSequence + 1
                        If NoChangeSequence = 20 Then GoTo ExitPoint
                    Else
                        Let NoChangeSequence = 0
                        'MsgBox "Iteration " + Str(LearningStageIndex) + ":  last step improved by " + Str(MyObjectiveFunction - StoredObjectiveFunction)
                    End If
                'Remember the last value.
                    Let StoredObjectiveFunction = MyObjectiveFunction
ResumePoint:
                
            'Report progress.
            '   To do this, you need to compute a number of diagnostics.
                If (LearningStageIndex) / mReportInterval = Int(LearningStageIndex / mReportInterval) Then
                    Let MyProbability = LogProbabilityOfData(mWeight())
                    'I don't this this is needed given what follows; we're using a specified number of trials.  But it is reported to user.
                        Let MySlope = Slope(mExpectedViolations(), mObservedViolations())
                        'MsgBox Str(MySlope)
                    Let MyObjectiveFunction = ObjectiveFunction()
                    Call MaxEntReportProgress(LearningStageIndex, MyObjectiveFunction, MySlope)
                    DoEvents
                    'Stop
                End If
                
            'Exit if magnitude of slope is small.
                'Let MySlope = Slope(mExpectedViolations(), mObservedViolations())
                'Debug:
                '    Let MyProbability = LogProbabilityOfData(mWeight())
                '    MsgBox "Slope:  " + Str(MySlope) + ".   Data probability:  " + Str(MyProbability)
                'If MySlope < 0.01 Then GoTo ExitPoint
                
SkipPoint:
                
        Next LearningStageIndex     'Do as many learning stages as the user requested.

ExitPoint:                          'Go here on failure, so you can check the Timer.
                                    'Also go here if 100 iterations make no improvement.

        Call QuickiePrintout(mInputForm())
        
    'Determine how long learning took.
        Let mTimeMarker = Timer - mTimeMarker
      
    'Close the debugging file.
    '    Close #DebugFile
        
   Exit Sub

DebugMe:

    'Debug:
        
        Print #3, "expected";
        For ConstraintIndex = 1 To mNumberOfConstraints
            Print #3, Chr(9); mExpectedViolations(ConstraintIndex);
        Next ConstraintIndex
        
        Print #3, Chr(9); "o - e";
        For ConstraintIndex = 1 To mNumberOfConstraints
            Print #3, Chr(9); mObservedViolations(ConstraintIndex) - mExpectedViolations(ConstraintIndex);
        Next ConstraintIndex
        
        Print #3, Chr(9); "weight";
        For ConstraintIndex = 1 To mNumberOfConstraints
            Print #3, Chr(9); mWeight(ConstraintIndex);
        Next ConstraintIndex
        Print #3,
        
        Return


End Sub

Function Slope(MyExpectedViolations() As Single, MyObservedViolations() As Single) As Single

    'Not in use, but a more principled way to determine termination.
    
    Dim ConstraintIndex As Long
    
    For ConstraintIndex = 1 To mNumberOfConstraints
        Let Slope = Slope + (MyExpectedViolations(ConstraintIndex) - MyObservedViolations(ConstraintIndex)) ^ 2
    Next ConstraintIndex

End Function


Function ObjectiveFunction() As Single

    'Compute the objective function, which depends on whether there is a prior.
    
        Dim SquaredSumOfWeights As Single
        Dim ConstraintIndex As Long
    
        Let ObjectiveFunction = LogProbabilityOfData(mWeight())
        
        'Additional calculations if using a prior:
            If mUsingPriorTerm Then
                'Compute the summed squares of the weights and subtract it.
                    For ConstraintIndex = 1 To mNumberOfConstraints
                        Let SquaredSumOfWeights = SquaredSumOfWeights + mWeight(ConstraintIndex) ^ 2
                    Next ConstraintIndex
                    Let ObjectiveFunction = ObjectiveFunction - (SquaredSumOfWeights / mSigmaSquared / 2)
            End If

End Function

Function LogProbabilityOfData(LocalWeights() As Single) As Single

    'To find the probability of the data under a grammar, you find the probability of each candidate,
    '   and multiply by its real (not fractional) frequency.
    
    'The probability of the data is computed as a log probability, since it is very low.
    
    Dim CandidateLogProbabilities() As Single
    Dim Buffer As Single
    Dim Harmony As Single
    Dim Z As Single
    
    ReDim eHarmony(mNumberOfForms, mMaximumNumberOfRivals)
    ReDim CandidateLogProbabilities(mNumberOfForms, mMaximumNumberOfRivals)
    
    Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
        
    'Calculate harmony, exponentiated harmony, Z, and finally candidate probabilities.
        For FormIndex = 1 To mNumberOfForms
            
            'Initialize Z, which is calculated separately for each input.
                Let Z = 0
            'August 2025:  why start at zero?
            'For RivalIndex = 0 To mNumberOfRivals(FormIndex)
            For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                'Initialize Harmony, which is computed for each candidate.
                    Let Harmony = 0
                'For this candidate, form Harmony, the dot product of weights and violations.
                    For ConstraintIndex = 1 To mNumberOfConstraints
                        Let Harmony = Harmony + LocalWeights(ConstraintIndex) * mRivalViolations(FormIndex, RivalIndex, ConstraintIndex)
                    Next ConstraintIndex
                'Exponentiate the negative.
                    Let eHarmony(FormIndex, RivalIndex) = Exp(-Harmony)
                'Sum exponentiated scores over rivals.
                    Let Z = Z + eHarmony(FormIndex, RivalIndex)
            
            Next RivalIndex
            
            'Calculate log probabilities
                'August 2025:  why start at zero?
                'For RivalIndex = 0 To mNumberOfRivals(FormIndex)
                For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                    'This causes crashes for zero-frequency candidates:  if the weights get high enough, then
                    '   ExponentiatedScores() becomes zero.
                    'But we don't even need this number--it's not part of the input data.
                    If mFrequency(FormIndex, RivalIndex) > 0 Then
                        Let CandidateLogProbabilities(FormIndex, RivalIndex) = Log(eHarmony(FormIndex, RivalIndex) / Z)
                    End If
                Next RivalIndex
                
                
        Next FormIndex
        
    'Calculate probability of input data.
        For FormIndex = 1 To mNumberOfForms
            'August 2025:  why start at zero?
            'For RivalIndex = 0 To mNumberOfRivals(FormIndex)
            For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                Let Buffer = Buffer + CandidateLogProbabilities(FormIndex, RivalIndex) * mFrequency(FormIndex, RivalIndex)
            Next RivalIndex
        Next FormIndex
        
    'Probability is switched in sign.
        Let LogProbabilityOfData = Buffer


End Function


Sub CalculateExpectedViolations(TrialIndex As Long)

    'The expected violations are based on the current grammar and are the dot product of candidate frequencies and violations.
    
    Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
    
    'This routine always needs the following; which is only called by it.
        Call CalculatePredictedRivalProportions(TrialIndex)
    
    'Initialize, since each reweighting produces a new set of expectations.
        For ConstraintIndex = 1 To mNumberOfConstraints
            Let mExpectedViolations(ConstraintIndex) = 0
        Next ConstraintIndex
    
    'Perform the calculation.
        For ConstraintIndex = 1 To mNumberOfConstraints
            For FormIndex = 1 To mNumberOfForms
                'August 2025:  why start at zero?
                'For RivalIndex = 0 To mNumberOfRivals(FormIndex)
                For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                        Let mExpectedViolations(ConstraintIndex) = mExpectedViolations(ConstraintIndex) + _
                            mTotalFrequencyForEachInput(FormIndex) * mPredictedRivalProportions(FormIndex, RivalIndex) * mRivalViolations(FormIndex, RivalIndex, ConstraintIndex)
                Next RivalIndex
            Next FormIndex
            'Debug:
                If mExpectedViolations(ConstraintIndex) = 0 Then Stop
        Next ConstraintIndex
    
    Exit Sub
   

End Sub


Sub CalculatePredictedRivalProportions(TrialNumber As Long)

    'Given a grammar, what proportions (probabilities) does it assign to each candidate?
    'This is needed to calculate expected violations, and is also of interest as an output of the algorithm.
    

    Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
    Dim MyScore As Single, MyScoreSum As Single
    
    ReDim eHarmony(mNumberOfForms, mMaximumNumberOfRivals)
    
    For FormIndex = 1 To mNumberOfForms
            'August 2025:  why start at zero?
            'For RivalIndex = 0 To mNumberOfRivals(FormIndex)
            For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                'Find dot product of violations and weights.
                    For ConstraintIndex = 1 To mNumberOfConstraints
                        Let MyScore = MyScore + mWeight(ConstraintIndex) * mRivalViolations(FormIndex, RivalIndex, ConstraintIndex)
                    Next ConstraintIndex
                'Exponentiate by the negative.
                    Let eHarmony(FormIndex, RivalIndex) = SafeExp(-MyScore)
                'Reinitialize MyScore
                    Let MyScore = 0
                'Sum exponentiated scores over rivals.
                    Let MyScoreSum = MyScoreSum + eHarmony(FormIndex, RivalIndex)
            Next RivalIndex
        'Calculate and print probabilities
            'August 2025:  why start at zero?
            'For RivalIndex = 0 To mNumberOfRivals(FormIndex)
            For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                Let mPredictedRivalProportions(FormIndex, RivalIndex) = eHarmony(FormIndex, RivalIndex) / MyScoreSum
            Next RivalIndex
        'Initialize for the next form.
            Let MyScoreSum = 0
    Next FormIndex
    
    'If requested, add to the history file for output probabilities.
        If mnuOutputHistoryFile.Checked = True Then
            For FormIndex = 1 To mNumberOfForms
                    'Trial number.
                        Print #mOutputHistoryFile, Trim(Str(TrialNumber)); vbTab;
                    For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                        Print #mOutputHistoryFile, vbTab; Trim(Str(mPredictedRivalProportions(FormIndex, RivalIndex)));
                    Next RivalIndex
            Next FormIndex
            Print #mOutputHistoryFile,
        End If
    
End Sub

Function SafeExp(MyDouble As Single) As Single

    'We need to return zero as the exp(value) when VB would overflow.
    
    On Error GoTo ErrorPoint
    
    Let SafeExp = Exp(MyDouble)
    Exit Function
    
ErrorPoint:
    Let SafeExp = 0
    
End Function


Function DeltaUsingPrior(MyConstraint As Long)

    'Find the adjustment of a constraint weight when a Gaussian prior is being used.
    '   This is calculated using the formula in Goodman (2002, p. 12).
    
    Dim Expected As Single
    Dim Observed As Single
    Dim MyWeight As Single
    Dim Delta As Single
    Dim Derivative As Single
    Dim FunctionValue As Single
    Dim IterationIndex As Long
    Dim OldDelta As Single
    
    'We need to solve the equation given by Goodman, for which we can use, as he suggests (Goodman n.d., p. 3),
    '   Newton's method.
    
    'The equation to be solved, alas, comes in two versions, depending on which paper by Goodman you read.
    'One version (2002, 12) is:
    
    '   Observed(i) = Expected(i) * exp(Delta * FPoundSign) + (MyWeight + Delta)/SigmaSquared
    
    'Where in the present program:
    '   Observed(i) is given by mObservedViolations(MyConstraint)
    '   Expected(i) is given by mExpectedViolations(MyConstraint)
    '   FPoundSign is given by mSlowingFactor.
    '   MyWeight is given by mWeight(MyConstraint)
    '   SigmaSquared is given by mSigmaSquared
    
    'A later version of the update formula (Goodman, no date, p. 3) is:
    
    '   Observed(i) - MyWeight/SigmaSquared = Expected(i) * exp(Delta * FPoundSign)
    
    'i.e.
    
    '   Observed(i) = Expected(i) * exp(Delta * FPoundSign) + MyWeight/SigmaSquared
    
    'thus omitting the second appearance of Delta from the first version.
    
    'Assuming the first of these, we need to put a zero on one side of the equals sign, thus:
    
    '    0 = Expected(i) * exp(Delta * FPoundSign) + (MyWeight + Delta)/SigmaSquared - Observed(i)
    
    'We'll use Newton's Method to calculate Delta.
    
   On Error GoTo ErrorPoint
   
    'Let the starting point be zero.  If this is going well, the weight updates should eventually become
    '   very small, so this is a sensible starting point.
        Let Delta = 0
        
    'Localize some variables for clarity and speed.
        Let Expected = mExpectedViolations(MyConstraint)
        Let Observed = mObservedViolations(MyConstraint)
        Let MyWeight = mWeight(MyConstraint)
        
    'Experiment:  let's smooth these, just like in the no-Gaussian version.
        'Let Expected = Expected + 0.0001
        'Let Observed = Observed + 0.0001
        
    'Open a debug file.
        'Close
        'Open App.Path + "/debug.txt" For Output As #10
        
    'Let's try 100 iterations of Newton's Method.
    '    For IterationIndex = 1 To 10000
    
    'New:  insist on convergence.
        Let IterationIndex = 1
        Do
            Let IterationIndex = IterationIndex + 1
            
            'Debug:  save the value that caused a crash.
                Let OldDelta = Delta
            
            'To employ Newton's method, we calculate the function value and the derivative.
                Let FunctionValue = (Expected * Exp(Delta * mSlowingFactor)) + ((MyWeight + Delta) / mSigmaSquared) - Observed
                Let Derivative = (Expected * mSlowingFactor * Exp(Delta * mSlowingFactor)) + (1 / mSigmaSquared)
            
            'The formula for Newton's Method, from http://en.wikipedia.org/wiki/Newton's_method#Description_of_the_method
                Let Delta = Delta - (FunctionValue / Derivative)
                If Abs(FunctionValue) < 0.00001 Then
                    'MsgBox "perfect zero achieved"
                    Exit Do
                End If
            'Debug:
                'If IterationIndex / 1000 = Int(IterationIndex / 1000) Then
            '        Print #10, "It"; Chr(9); "Delta"; Chr(9); Delta; Chr(9); "Function"; Chr(9); FunctionValue; Chr(9); "Deriv"; Chr(9); Derivative
                'End If
                
            'Debug:
                If IterationIndex > 100000 Then Stop
            
            Loop
        'Next IterationIndex
        
        'MsgBox "Function reached" + Str(FunctionValue)
        Let DeltaUsingPrior = Delta
        
        'Debug:
        '    Dim Dummy as single
        '    Let Dummy = MyWeight
        '    Let Dummy = FunctionValue
        '    Let Dummy = Derivative
        Exit Function
        
ErrorPoint:

        MsgBox "Error.  Function value is " + Str(FunctionValue) + ".  Delta is " + Str(Delta), vbCritical
        Let DeltaUsingPrior = Log((mObservedViolations(MyConstraint) + 0.0001) / (mExpectedViolations(MyConstraint) + 0.0001)) / mSlowingFactor

        'Debug:
            'Close
            'End

End Function


Sub MaxEntReportProgress(CycleIndex As Long, ObjectiveFunction As Single, MySlope As Single)

   'Print out the results so far.

      Dim ConstraintIndex As Long
      Dim SquaredSum As Single
         
   'Print the results so far.

        pctProgressWindow.Cls
        
        'Cycle number:
            pctProgressWindow.Print "Completed learning cycle #"; CycleIndex&
            pctProgressWindow.Print
            
        'Objective function:
            If mUsingPriorTerm Then
                pctProgressWindow.Print "Objective function:  " + NDecimalPlaces(ObjectiveFunction)
                pctProgressWindow.Print
            Else
                pctProgressWindow.Print "Log probability of data:  " + NDecimalPlaces(ObjectiveFunction)
                pctProgressWindow.Print
            End If
        
        'Slope (needed long term?):
            pctProgressWindow.Print "Slope of gradient:  " + Trim(Str(MySlope))
            pctProgressWindow.Print
        
        'Weights:
            pctProgressWindow.Print "Current weights:"
            For ConstraintIndex = 1 To mNumberOfConstraints
                pctProgressWindow.Print s.RightJustifiedFill(NDecimalPlaces(mWeight(ConstraintIndex)), 11);
                pctProgressWindow.Print "   "; mAbbrev(ConstraintIndex)
            Next ConstraintIndex
            pctProgressWindow.Print
            
            
            pctProgressWindow.Print
            pctProgressWindow.Print
            pctProgressWindow.Print
            pctProgressWindow.Print
            

End Sub


'=================================PRINTING==================================
'===========================================================================

Sub QuickiePrintout(mInputForm() As String)

    'A quickie printout--needs more work.
    
    Dim ConstraintIndex As Long, FormIndex As Long, RivalIndex As Long
    Dim MyScore As Single, MyScoreSum As Single
    Dim LongestInputLength As Long, LongestCandidateLength As Long
    
        'Constraints:
            Call PrintLevel1Header(mDocFile, mTmpFile, mHTMFile, "Constraints and weights")
            
            Dim MyTable() As String
            ReDim MyTable(2, mNumberOfConstraints)
            Print #mTmpFile,
            For ConstraintIndex = 1 To mNumberOfConstraints
                Print #mTmpFile, s.RightJustifiedFill(NDecimalPlaces(mWeight(ConstraintIndex)), 8); Chr(9); mConstraintName(ConstraintIndex)
                Let MyTable(1, ConstraintIndex) = mConstraintName(ConstraintIndex)
                Let MyTable(2, ConstraintIndex) = s.RightJustifiedFill(NDecimalPlaces(mWeight(ConstraintIndex)), 8)
            Next ConstraintIndex
            
            Call s.PrintHTMTable(MyTable(), mHTMFile, False, False, True)
            Print #mTmpFile,
        
        'For pretty printing, find longest input and candidate.
            Let LongestInputLength = 0
            Let LongestCandidateLength = 0
            For FormIndex = 1 To mNumberOfForms
                If Len(mInputForm(FormIndex)) > LongestInputLength Then
                    Let LongestInputLength = Len(mInputForm(FormIndex))
                End If
                For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                    If Len(mRival(FormIndex, RivalIndex)) > LongestCandidateLength Then
                        Let LongestCandidateLength = Len(mRival(FormIndex, RivalIndex))
                    End If
                Next RivalIndex
            Next FormIndex
                
        'Scores, exponentiated scores, summed exponentiated scores, and finally candidate probabilities.
            ReDim ExponentiatedScores(mNumberOfForms, mMaximumNumberOfRivals)
            ReDim CandidateLogProbabilities(mNumberOfForms, mMaximumNumberOfRivals)
                
            Call PrintLevel1Header(mDocFile, mTmpFile, mHTMFile, "Inputs, candidates, input frequencies, input proportions, predicted probabilities")
            
            Dim Table() As String
            Dim RowCount As Long
            
            
            For FormIndex = 1 To mNumberOfForms
                'Recalculate with final grammar:
                    'August 2025:  why start at zero?
                    'For RivalIndex = 0 To mNumberOfRivals(FormIndex)
                    For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                        'Find dot product of violations and weights.
                            For ConstraintIndex = 1 To mNumberOfConstraints
                                Let MyScore = MyScore + mWeight(ConstraintIndex) * mRivalViolations(FormIndex, RivalIndex, ConstraintIndex)
                            Next ConstraintIndex
                        'Exponentiate by the negative.
                            Let ExponentiatedScores(FormIndex, RivalIndex) = Exp(-MyScore)
                        'Reinitialize MyScore
                            Let MyScore = 0
                        'Sum exponentiated scores over rivals.
                            Let MyScoreSum = MyScoreSum + ExponentiatedScores(FormIndex, RivalIndex)
                    Next RivalIndex
                'Start a table for this form.
                    ReDim Table(5, 1)
                    Let RowCount = 1
                    Let Table(1, 1) = "Inputs"
                    Let Table(2, 1) = "Candidates"
                    Let Table(3, 1) = "Input frequencies"
                    Let Table(4, 1) = "Input proportions"
                    Let Table(5, 1) = "Predicted probabilities"

                'Calculate and print probabilities
                    'August 2025:  why start at zero?  Because zero is a different diacritic here.
                    For RivalIndex = 0 To mNumberOfRivals(FormIndex)
                    'For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                        Let mPredictedRivalProportions(FormIndex, RivalIndex) = ExponentiatedScores(FormIndex, RivalIndex) / MyScoreSum
                        'Table
                            Let RowCount = RowCount + 1
                            ReDim Preserve Table(5, RowCount)
                            If RivalIndex = 0 Then
                                Let Table(1, RowCount) = mInputForm(FormIndex)
                            End If
                            Let Table(2, RowCount) = mRival(FormIndex, RivalIndex)
                            Let Table(3, RowCount) = mFrequency(FormIndex, RivalIndex)
                            Let Table(4, RowCount) = s.RightJustifiedFill(NDecimalPlaces(mObservedRivalProportions(FormIndex, RivalIndex)), 10)
                            Let Table(5, RowCount) = NDecimalPlaces(mPredictedRivalProportions(FormIndex, RivalIndex))
                    Next RivalIndex
                
                'Print the table.
                    Call s.PrintTable(mDocFile, mTmpFile, mHTMFile, Table(), True, False, True)
                    Print #mTmpFile,
                'Initialize for the next form.
                    Let MyScoreSum = 0
            
            Next FormIndex
            
        Call PrintPara(mDocFile, mTmpFile, mHTMFile, "Probability of data = " + Str(LogProbabilityOfData(mWeight())))

End Sub




Function PrintMaxentResults(mInputForm() As String, RankingValue() As Single, ThingFound As String)

   'Print out the results of a numerical algorithm.
        'The RankingValue() array can be either GLA ranking values or Maximum Entropy weights;
        '   and ThingFound (must be grammatically plural) verbally identifies which one it is.
        'Within this code, for historical reasons, the variable is called RankingValue().

      ReDim mSlotFiller(mNumberOfConstraints)
      Dim LocalRankingValue() As Single
      ReDim LocalRankingValue(mNumberOfConstraints)

      Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long, InnerConstraintIndex As Long
      Dim ThresholdIndex As Long
      Dim SpaceIndex As Long
      Dim i As Long, j As Long
      
      Dim Swappant As Single, SwapInt As Long
      
      Dim Difference As Single
      
      Dim OldValsFile As Long

      
   'Sort the constraints by their ranking values.  Dec. 2025:  Let's not.

      For ConstraintIndex = 1 To mNumberOfConstraints
         Let mSlotFiller(ConstraintIndex) = ConstraintIndex
         Let LocalRankingValue(ConstraintIndex) = RankingValue(ConstraintIndex)
      Next ConstraintIndex

      'For i = 1 To mNumberOfConstraints
      '   For j = 1 To i - 1
      '      If LocalRankingValue(j) < LocalRankingValue(i) Then
      '          Let Swappant = LocalRankingValue(i)
      '          Let LocalRankingValue(i) = LocalRankingValue(j)
      '          Let LocalRankingValue(j) = Swappant
      '          Let SwapInt = mSlotFiller(i)
      '          Let mSlotFiller(i) = mSlotFiller(j)
      '          Let mSlotFiller(j) = SwapInt
      '      End If
      '   Next j
      'Next i

   'Print the results of the algorithm.

      Print #mDocFile, "\ks"
      
      'ThingFound is either weights or ranking values.
            Call PrintLevel1Header(mDocFile, mTmpFile, mHTMFile, ThingFound + " Found")
      
      Print #mDocFile, "\ts2"
       
      Dim MyTable() As String
      ReDim MyTable(2, mNumberOfConstraints + 1)
      Let MyTable(1, 1) = "Constraints"
      Let MyTable(2, 1) = ThingFound + " Found"
      
      For ConstraintIndex = 1 To mNumberOfConstraints
         Print #mDocFile, SmallCapTag1; mConstraintName(mSlotFiller(ConstraintIndex)); SmallCapTag2; Chr(9);
         Print #mDocFile, NDecimalPlaces(RankingValue(mSlotFiller(ConstraintIndex)))
         Print #mTmpFile, FillStringTo(NDecimalPlaces(RankingValue(mSlotFiller(ConstraintIndex))), 10); NDecimalPlaces(RankingValue(mSlotFiller(ConstraintIndex)));
         Print #mTmpFile, "   "; mConstraintName(mSlotFiller(ConstraintIndex))
         Let MyTable(1, ConstraintIndex + 1) = mConstraintName(mSlotFiller(ConstraintIndex))
         Let MyTable(2, ConstraintIndex + 1) = NDecimalPlaces(RankingValue(mSlotFiller(ConstraintIndex)))
      Next ConstraintIndex
      
        Call PrintHTMTable(MyTable(), mHTMFile, True, False, True)
        Print #mDocFile, "\te\ke"
      
    'We also want a plain output that can be processed by Excel.
    '   Weights should be reported in straight constraint order, to facilitate
    '   comparison over multiple runs.
    
        Let mTabbedFile = FreeFile
        Open gOutputFilePath + gFileName + "TabbedOutput.txt" For Output As #mTabbedFile
        
        'Top:  constraints and weights
            Print #mTabbedFile, "Input"; Chr(9); "Candidate"; Chr(9); "Freq. in input file"; Chr(9); "Target proportion"; Chr(9); "Predicted proportion";
            For ConstraintIndex = 1 To mNumberOfConstraints
               Print #mTabbedFile, Chr(9); mAbbrev(ConstraintIndex);
            Next ConstraintIndex
            Print #mTabbedFile,
        'Weights
            Print #mTabbedFile, Chr(9); Chr(9); Chr(9); Chr(9); "Weights:";
            For ConstraintIndex = 1 To mNumberOfConstraints
               Print #mTabbedFile, Chr(9); NDecimalPlaces(mWeight(ConstraintIndex));
            Next ConstraintIndex
            Print #mTabbedFile,

        'Inputs, candidates, scores, etc.:
            For FormIndex = 1 To mNumberOfForms
                'August 2025:  why start at zero?
                'For RivalIndex = 0 To mNumberOfRivals(FormIndex)
                For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                    Print #mTabbedFile, mInputForm(FormIndex);
                    Print #mTabbedFile, Chr(9); mRival(FormIndex, RivalIndex);
                    Print #mTabbedFile, Chr(9); mFrequency(FormIndex, RivalIndex);
                    Print #mTabbedFile, Chr(9); NDecimalPlaces(mObservedRivalProportions(FormIndex, RivalIndex));
                    Print #mTabbedFile, Chr(9); NDecimalPlaces(mPredictedRivalProportions(FormIndex, RivalIndex));
                    For ConstraintIndex = 1 To mNumberOfConstraints
                        Print #mTabbedFile, Chr(9); s.ZeroToNull(mRivalViolations(FormIndex, RivalIndex, ConstraintIndex));
                    Next ConstraintIndex
                    Print #mTabbedFile,
                Next RivalIndex
            Next FormIndex
        
      
   'Print a file to save results if you want to run it further.

        'First, make sure there is a folder for these files, a daughter of the
        '   folder in which the input file is located.
            Call Form1.CreateAFolderForOutputFiles
      
        'Print the file.
            Let OldValsFile = FreeFile
            Open gOutputFilePath + gFileName + "MostRecentWeights.txt" For Output As #OldValsFile
            For ConstraintIndex = 1 To mNumberOfConstraints
               Print #OldValsFile, mAbbrev(ConstraintIndex); Chr(9); RankingValue(ConstraintIndex)
            Next ConstraintIndex

End Function


Sub PrintAHeader(MyAlgorithmName As String)

    
    MsgBox "Trying to print; mhtmfile is " + Str(mHTMFile)
    
     'The files are supposed to be opened in Form1.  Somehow, they got closed again.  Since this is now a low-priority routine, I will just put in a hack here.
    
       'The quick, draft output:
            Let mTmpFile = FreeFile
            Open gOutputFilePath + gFileName + "DraftOutput.txt" For Output As #mTmpFile
            'Initialize the header numbers, in case this isn't the first run.
                Let gLevel1HeadingNumber = 0
        'The output for pretty Word file conversion:
            Let mDocFile = FreeFile
            Open gOutputFilePath + gFileName + "QualityOutput.txt" For Output As #mDocFile
        'The HTML output:
            Let mHTMFile = FreeFile
            Open gOutputFilePath + "ResultsFor" + gFileName + ".htm" For Output As #mHTMFile
            'Call PrintTableaux.InitiateHTML(mHTMFile)
    
    'Print a header for the output file:
        Print #mDocFile, "\ft\tiResult of Applying ";
        Print #mDocFile, MyAlgorithmName;
        Print #mDocFile, " to "; gFileName; gFileSuffix
        Print #mDocFile,
        Print #mDocFile, "\cn"; NiceDate; ", "; NiceTime
        Print #mDocFile,
        Print #mDocFile, "\cn"; "OTSoft " + gMyVersionNumber + ", release date " + gMyReleaseDate
        Print #mDocFile,
        Print #mTmpFile, "Result of Applying ";
        Print #mTmpFile, MyAlgorithmName;
        Print #mTmpFile, " to "; gFileName; gFileSuffix
        Print #mTmpFile,
        Print #mHTMFile, "<b>Result of Applying ";
        Print #mHTMFile, MyAlgorithmName;
        Print #mHTMFile, " to "; gFileName; gFileSuffix; "</b><p><p>"
        
        Print #mTmpFile,
        Call PrintPara(-1, mTmpFile, mHTMFile, "OTSoft " + gMyVersionNumber + ", release date " + gMyReleaseDate)
        Call PrintPara(-1, mTmpFile, mHTMFile, NiceDate + ", " + NiceTime)
        Print #mTmpFile,
        
        
        Call PrintPara(mDocFile, mTmpFile, mHTMFile, "For more detailed examination of results, please use a spreadsheet program to open the file " + _
            "PARATabbedOutput.txt, located in the folder FilesFor" + gFileName + ".")
          
    'Print a header and diacritic to trigger a page number
        Print #mDocFile, "\hrGLA Results for "; gFileName; gFileSuffix; Chr$(9); NiceDate; Chr$(9); "\pn"


End Sub


Function NDecimalPlaces(ALong As Single) As String

    Dim Buffer As String
    
    Select Case Val(txtDecimalPlaces)
        Case 0
            Let Buffer = Format(ALong, "##,##0")
        Case 1
            Let Buffer = Format(ALong, "##,##0.0")
        Case 2
            Let Buffer = Format(ALong, "##,##0.00")
        Case 3
            Let Buffer = Format(ALong, "##,##0.000")
        Case 4
            Let Buffer = Format(ALong, "##,##0.0000")
        Case 5
            Let Buffer = Format(ALong, "##,##0.00000")
        Case 6
            Let Buffer = Format(ALong, "##,##0.000000")
        Case 7
            Let Buffer = Format(ALong, "##,##0.0000000")
        Case 8
            Let Buffer = Format(ALong, "##,##0.00000000")
        Case 9
            Let Buffer = Format(ALong, "##,##0.000000000")
        Case 10
            Let Buffer = Format(ALong, "##,##0.0000000000")
        Case Else
            Let Buffer = Str(ALong)
    End Select
        
    Let NDecimalPlaces = Trim(Buffer)

End Function

Sub PrintFinalDetails()
    
    Dim ConstraintIndex As Long
    
    'For sorting:
        Dim mSlotFiller() As Long
        Dim LocalRankingValue() As Single
        ReDim mSlotFiller(mNumberOfConstraints)
        ReDim LocalRankingValue(mNumberOfConstraints)
        Dim i As Long, j As Long
        Dim Swappant As Single, SwapInt As Long
    
    Dim LearningStageIndex As Long      'for learning schedule table
    Dim Buffer As String                'to help with spacing
 
    'Note the accuracy achieved, and the time needed.  xxx This needs more work.
       'Print #mDocFile,
       ' 'First, a header:
       '     Print #mDocFile, "\h1Testing the Grammar:  Details"
       '
       '     Let gLevel1HeadingNumber = gLevel1HeadingNumber + 1
       '     Print #TmpFile, Trim(gLevel1HeadingNumber); ". Testing the Grammar:  Details"
       '     Print #TmpFile,
       ' 'Results:
       '     Print #mDocFile, "The grammar was tested for "; Trim(Str(gCyclesToTest)); " cycles."
       '     Print #TmpFile, "   The grammar was tested for "; Trim(Str(gCyclesToTest)); " cycles."
       '     'Print #mDocFile, "Average error per candidate:  "; ndecimalplaces(100 * mErrorTerm / mTotalNumberOfRivals); " percent"
       '     'Print #TmpFile, "   Average error per candidate:  "; ndecimalplaces(100 * mErrorTerm / mTotalNumberOfRivals); " percent"
            
            
    'Learning time:
            Call PrintPara(mDocFile, mTmpFile, mHTMFile, "Learning time:  " + NDecimalPlaces(mTimeMarker / 60) + " minutes")
    
    'Print the details of the learning simulation.
        Print #mDocFile,
        Print #mTmpFile,
        Print #mDocFile, "\ks"
        
        'Call PrintLevel1Header(mDocFile, mTmpFile, mHTMFile, "Parameter Values Used by the Maxent Algorithm")
    
        
           

End Sub


Private Sub mnuIncludeTableaux_Click()
    If mnuIncludeTableaux.Checked = False Then
        Let mnuIncludeTableaux.Checked = True
        Let IncludeTableauxInGLAOutput = True
    Else
        Let mnuIncludeTableaux.Checked = False
        Let IncludeTableauxInGLAOutput = False
    End If
End Sub


Sub PrintTableaux(NumberOfConstraints As Long, Abbrev() As String, Weight() As Single, NumberOfForms As Long, InputForms() As String, NumberOfRivals() As Long, Rivals() As String, RivalViolations() As Long, _
    PredictedRivalProportions() As Single, eHarmony() As Double, ObservedRivalProportions() As Single, DocFile As Long, TmpFile As Long, HTMFile As Long)
    
    Dim Table() As String
    Dim RowCount As Long
    Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
    Dim MyHarmony As Single
    
    'Main header
        Call PrintLevel1Header(DocFile, TmpFile, HTMFile, "Tableaux")
    
    'Go through the forms, producing one tableau for each.
        For FormIndex = 1 To NumberOfForms
            
            'Tableaux header:
                Let RowCount = 1
                ReDim Table(NumberOfConstraints + 6, RowCount)
                Let Table(1, 1) = "Input"
                Let Table(2, 1) = "Candidate"
                Let Table(3, 1) = "Harmony"
                Let Table(4, 1) = "exp(-H)"
                Let Table(5, 1) = "Predicted"
                Let Table(6, 1) = "Observed"
                For ConstraintIndex = 1 To NumberOfConstraints
                    Let Table(ConstraintIndex + 6, 1) = Abbrev(mSlotFiller(ConstraintIndex))
                Next ConstraintIndex
                
            'Weights
                Let RowCount = 2
                ReDim Preserve Table(NumberOfConstraints + 6, RowCount)
                For ConstraintIndex = 1 To NumberOfConstraints
                    Let Table(ConstraintIndex + 6, RowCount) = NDecimalPlaces(Weight(mSlotFiller(ConstraintIndex)))
                Next ConstraintIndex
                
            'Violations and score computation.
                'August 2025:  why start at zero?
                'For RivalIndex = 0 To NumberOfRivals(FormIndex)
                For RivalIndex = 1 To NumberOfRivals(FormIndex)
                    'Expand the table.
                        Let RowCount = RowCount + 1
                        ReDim Preserve Table(NumberOfConstraints + 6, RowCount)
                    'Give the input and rival.
                        If RivalIndex = 0 Then
                            Let Table(1, RowCount) = InputForms(FormIndex)
                        End If
                        Let Table(2, RowCount) = Rivals(FormIndex, RivalIndex)
                    'Calculate the harmony
                        Let MyHarmony = 0
                        For ConstraintIndex = 1 To NumberOfConstraints
                            Let MyHarmony = MyHarmony + Weight(ConstraintIndex) * RivalViolations(FormIndex, RivalIndex, ConstraintIndex)
                        Next ConstraintIndex
                    'Print the numbers that calculate the predicted score.
                        Let Table(3, RowCount) = NDecimalPlaces(MyHarmony)
                        Let Table(4, RowCount) = NDecimalPlaces(CSng(eHarmony(FormIndex, RivalIndex)))
                        Let Table(5, RowCount) = NDecimalPlaces(ObservedRivalProportions(FormIndex, RivalIndex))
                        Let Table(6, RowCount) = NDecimalPlaces(PredictedRivalProportions(FormIndex, RivalIndex))
                    'Constraint violations:
                        For ConstraintIndex = 1 To NumberOfConstraints
                            If RivalViolations(FormIndex, RivalIndex, mSlotFiller(ConstraintIndex)) > 0 Then
                                Let Table(ConstraintIndex + 6, RowCount) = AsteriskString(RivalViolations(FormIndex, RivalIndex, mSlotFiller(ConstraintIndex)))
                            End If
                        Next ConstraintIndex
                Next RivalIndex
            
        'Final printout of this tableau.
            Call s.PrintTable(DocFile, TmpFile, HTMFile, Table(), True, False, True)
            
        Next FormIndex
    
End Sub

'===================================Factorial Typology=========================================

    'A little hokey, since it just samples from weights in the range 0-5.  Not in effect or accessible right now.


Private Sub mnuExploreFactoricalTypology_Click()
    
    ReDim mFrequency(mNumberOfForms, mMaximumNumberOfRivals)
    ReDim mWeight(mNumberOfConstraints)
    ReDim mPredictedRivalProportions(mNumberOfForms, mMaximumNumberOfRivals)

    'Since the user has clicked a Rank button, (s)he probably wants the settings saved.
        Call Form1.SaveUserChoices
        Call InstallWinnerAsRival
        Call FormCandidatesInOneArray
        Call ComputeFactorialTypology

End Sub
Sub FormCandidatesInOneArray()

    Dim FormIndex As Long, RivalIndex As Long, i As Long, j As Long
    
    'Make one single array so we can generalize across all the candidates.
        For FormIndex = 1 To mNumberOfForms
            ReDim Preserve mCandidatesInOneArray(2, mTotalNumberOfCandidates)
            'Record the rivals.  Winner is in the zero slot.
                'August 2025:  why start at zero?
                'For RivalIndex = 0 To mNumberOfRivals(FormIndex)
                For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                    Let mTotalNumberOfCandidates = mTotalNumberOfCandidates + 1
                    ReDim Preserve mCandidatesInOneArray(2, mTotalNumberOfCandidates)
                    Let mCandidatesInOneArray(1, mTotalNumberOfCandidates) = FormIndex
                    Let mCandidatesInOneArray(2, mTotalNumberOfCandidates) = RivalIndex
                Next RivalIndex
        Next FormIndex
        
    'Now that you know the total number of candidates, redimension some arrays.
        ReDim mLowestFrequencyAssigned(mTotalNumberOfCandidates)
        ReDim mHighestFrequencyAssigned(mTotalNumberOfCandidates)
        ReDim mTotalFrequencyAssigned(mTotalNumberOfCandidates)
        ReDim mProbabilityDifferenceLow(mTotalNumberOfCandidates, mTotalNumberOfCandidates)
        ReDim mProbabilityDifferenceHigh(mTotalNumberOfCandidates, mTotalNumberOfCandidates)
        
    'The lowest frequency array must be initialized at one.
        For i = 1 To mTotalNumberOfCandidates
            Let mLowestFrequencyAssigned(i) = 1
        Next i
        For i = 1 To mTotalNumberOfCandidates
            For j = 1 To mTotalNumberOfCandidates
                Let mProbabilityDifferenceLow(i, j) = 1
            Next j
        Next i
        

End Sub

Sub ComputeFactorialTypology()

    Dim MyHarmony As Single, LocalCandidate As Long
    Dim NumberOfTrials As Long, WeightMinimum As Single, WeightMaximum As Single, WeightRange As Single
    Dim ProbabilityDifference As Single, OuterProbability As Single, InnerProbability As Single
    Dim IterationIndex As Long, FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long, CandidateIndex As Long, i As Long
    Dim InnerCandidateIndex As Long
    Dim OutFile As Long
    
    Randomize
    
    Let NumberOfTrials = Val(txtPrecision.Text)
    Let WeightMinimum = Val(txtWeightMinimum.Text)
    Let WeightMaximum = Val(txtWeightMaximum.Text)
    Let WeightRange = WeightMaximum - WeightMinimum
    
    For IterationIndex = 1 To NumberOfTrials
    
        'Assign a random set of weights.  For now, 0-5 at random.
            For ConstraintIndex = 1 To mNumberOfConstraints
                Let mWeight(ConstraintIndex) = WeightMinimum + WeightRange * Rnd()
            Next ConstraintIndex
            
        'Use these to compute a set of output probabilities.  Since the variables are module-level, nothing needs to be passed.
            Call CalculatePredictedRivalProportions(IterationIndex)
            
        'Install these in the various arrays, keeping track of whether improvement has been made.
            Let LocalCandidate = 0
            For FormIndex = 1 To mNumberOfForms
                'August 2025:  why start at zero?
                'For RivalIndex = 0 To mNumberOfRivals(FormIndex)
                For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                    Let LocalCandidate = LocalCandidate + 1
                    'Update total frequency (will be divided by number of trials to get a mean)
                        Let mTotalFrequencyAssigned(LocalCandidate) = mTotalFrequencyAssigned(LocalCandidate) + mPredictedRivalProportions(FormIndex, RivalIndex)
                    'Update lowest frequency.
                        If mPredictedRivalProportions(FormIndex, RivalIndex) < mLowestFrequencyAssigned(LocalCandidate) Then
                            Let mLowestFrequencyAssigned(LocalCandidate) = mPredictedRivalProportions(FormIndex, RivalIndex)
                            'Report it:
                                'Print #OutFile, IterationIndex;
                                'For i = 1 To LocalCandidate
                                '    Print #OutFile, Chr(9);
                                'Next i
                                'Print #OutFile, "Low: "; Chr(9); Trim(Str(mLowestFrequencyAssigned(LocalCandidate)))
                        End If
                    'Update highest frequency.
                        If mPredictedRivalProportions(FormIndex, RivalIndex) > mHighestFrequencyAssigned(LocalCandidate) Then
                            Let mHighestFrequencyAssigned(LocalCandidate) = mPredictedRivalProportions(FormIndex, RivalIndex)
                            'Report it:
                                'Print #OutFile, IterationIndex;
                                'For i = 1 To LocalCandidate
                                '    Print #OutFile, Chr(9);
                                'Next i
                                'Print #OutFile, "High: "; Chr(9); Trim(Str(mHighestFrequencyAssigned(LocalCandidate)))
                        End If
                Next RivalIndex
            Next FormIndex
            
        'Now look at differences.  Compare the probability of all distinct outputs that don't share input.
            For CandidateIndex = 1 To mTotalNumberOfCandidates - 1
                'Look up and store the probability the grammar assigns to this candidate.
                    Let OuterProbability = mPredictedRivalProportions(mCandidatesInOneArray(1, CandidateIndex), mCandidatesInOneArray(2, CandidateIndex))
                'Find another candidate.
                    For InnerCandidateIndex = CandidateIndex + 1 To mTotalNumberOfCandidates
                        'They must not have the same input form.
                            If mCandidatesInOneArray(1, CandidateIndex) <> mCandidatesInOneArray(1, InnerCandidateIndex) Then
                                Let InnerProbability = mPredictedRivalProportions(mCandidatesInOneArray(1, InnerCandidateIndex), mCandidatesInOneArray(2, InnerCandidateIndex))
                                Let ProbabilityDifference = OuterProbability - InnerProbability
                                'Install if it's a new record.
                                    If ProbabilityDifference < mProbabilityDifferenceLow(CandidateIndex, InnerCandidateIndex) Then
                                        Let mProbabilityDifferenceLow(CandidateIndex, InnerCandidateIndex) = ProbabilityDifference
                                    ElseIf ProbabilityDifference > mProbabilityDifferenceHigh(CandidateIndex, InnerCandidateIndex) Then
                                        Let mProbabilityDifferenceHigh(CandidateIndex, InnerCandidateIndex) = ProbabilityDifference
                                    End If
                            End If
                    Next InnerCandidateIndex
            Next CandidateIndex
            
            
        'Report progress.
            If Int(IterationIndex / 1000) = IterationIndex / 1000 Then
                pctProgressWindow.Cls
                pctProgressWindow.Print "Completed learning cycle #"; Str(IterationIndex)
                pctProgressWindow.Print "xxxxx"
                pctProgressWindow.Print "xxxxx"
                pctProgressWindow.Print "xxxxx"
                pctProgressWindow.Print "xxxxx"
                pctProgressWindow.Print
                pctProgressWindow.Print
                pctProgressWindow.Print
                pctProgressWindow.Print
            End If

        
    Next IterationIndex
        'Dim mProbabilityDifferenceLow() As Single
        'Dim mProbabilityDifferenceHigh() As Single
        
        
    'Print the results.
        
        Let OutFile = FreeFile
        Open gOutputFilePath + gFileName + "StochasticFactorialTypology.txt" For Output As #OutFile
        
        
        Print #OutFile, "Maxima and minima:"
        Print #OutFile, "Input"; Chr(9); "Candidate"; Chr(9); "Min"; Chr(9); "Max"; Chr(9); "Ave."
        'Maxima and minima:
            For CandidateIndex = 1 To mTotalNumberOfCandidates
                Print #OutFile, mInputForm(mCandidatesInOneArray(1, CandidateIndex));
                Print #OutFile, Chr(9); mRival(mCandidatesInOneArray(1, CandidateIndex), mCandidatesInOneArray(2, CandidateIndex));
                Print #OutFile, Chr(9); NDecimalPlaces(mLowestFrequencyAssigned(CandidateIndex));
                Print #OutFile, Chr(9); NDecimalPlaces(mHighestFrequencyAssigned(CandidateIndex));
                Print #OutFile, Chr(9); NDecimalPlaces(mTotalFrequencyAssigned(CandidateIndex) / NumberOfTrials)
            Next CandidateIndex
            
            Print #OutFile,
            Print #OutFile, "Differences:"
            Print #OutFile,
        
        'Differences:
            Print #OutFile, "1st IO Pair"; Chr(9); "2nd IO pair"; Chr(9); "Min diff."; Chr(9); "Max diff."; Chr(9); "Range"
            For CandidateIndex = 1 To mTotalNumberOfCandidates - 1
                For InnerCandidateIndex = CandidateIndex + 1 To mTotalNumberOfCandidates
                    'They must not have the same input form.
                        If mCandidatesInOneArray(1, CandidateIndex) <> mCandidatesInOneArray(1, InnerCandidateIndex) Then
                            Print #OutFile, mInputForm(mCandidatesInOneArray(1, CandidateIndex));
                            Print #OutFile, "-->";
                            Print #OutFile, mRival(mCandidatesInOneArray(1, CandidateIndex), mCandidatesInOneArray(2, CandidateIndex));
                            Print #OutFile, Chr(9);
                            Print #OutFile, mInputForm(mCandidatesInOneArray(1, InnerCandidateIndex));
                            Print #OutFile, "-->";
                            Print #OutFile, mRival(mCandidatesInOneArray(1, InnerCandidateIndex), mCandidatesInOneArray(2, InnerCandidateIndex));
                            Print #OutFile, Chr(9); NDecimalPlaces(mProbabilityDifferenceLow(CandidateIndex, InnerCandidateIndex));
                            Print #OutFile, Chr(9); NDecimalPlaces(mProbabilityDifferenceHigh(CandidateIndex, InnerCandidateIndex));
                            Print #OutFile, Chr(9); NDecimalPlaces(mProbabilityDifferenceHigh(CandidateIndex, InnerCandidateIndex) - mProbabilityDifferenceLow(CandidateIndex, InnerCandidateIndex))
                        End If
                Next InnerCandidateIndex
            Next CandidateIndex
        
        Close #OutFile
        MsgBox "Maxent factorial typology complete.  You can find the results in the output file folder (" + gOutputFilePath + " under the name " + gFileName + "StochasticFactorialTypology.txt."
    
        Unload Me
   
        
End Sub

Function PriorTerm(LocalWeights() As Single) As Single

    'This routine is currently not used.  Delete at some point.
    
    'The prior is computed from the weights, sigma, and mu.
    '   Sigma comes from the interface.
    '   Mu is currently assumed to be zero.
    'Ultimately, these should be readable from an input file.
    
    Dim ConstraintIndex As Long
    Dim Buffer As Single
    
    'Compute the summed squares of the weights.
        For ConstraintIndex = 1 To mNumberOfConstraints
            Let Buffer = Buffer + mWeight(ConstraintIndex) ^ 2
        Next ConstraintIndex
    
    'Divide by twice mSigmaSquared
        Let PriorTerm = Buffer / 2 / mSigmaSquared
    
End Function

