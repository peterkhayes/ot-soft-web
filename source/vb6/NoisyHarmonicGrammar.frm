VERSION 5.00
Begin VB.Form NoisyHarmonicGrammar 
   Caption         =   "Noisy Harmonic Grammar"
   ClientHeight    =   7965
   ClientLeft      =   3060
   ClientTop       =   1815
   ClientWidth     =   9045
   Icon            =   "NoisyHarmonicGrammar.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7965
   ScaleWidth      =   9045
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkExponentialNHG 
      Caption         =   "Employ Exponential HNG"
      Height          =   255
      Left            =   480
      TabIndex        =   22
      Top             =   2400
      Width           =   4095
   End
   Begin VB.CheckBox chkNoiseForZeroCells 
      Caption         =   "Include noise even in cells that have no violation"
      Height          =   255
      Left            =   480
      TabIndex        =   21
      Top             =   1320
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CheckBox chkLateNoise 
      Caption         =   "Add noise to candidates, after harmony calculation"
      Height          =   255
      Left            =   480
      TabIndex        =   20
      Top             =   2040
      Width           =   4215
   End
   Begin VB.CheckBox chkNegativeWeightsOK 
      Caption         =   "Allow constraint weights to go negative"
      Height          =   255
      Left            =   480
      TabIndex        =   19
      Top             =   1680
      Width           =   4215
   End
   Begin VB.CheckBox chkNoiseIsAddedAfterMultiplication 
      Caption         =   "Apply noise after multiplication of weights by violations"
      Height          =   255
      Left            =   480
      TabIndex        =   18
      Top             =   960
      Width           =   4335
   End
   Begin VB.CheckBox chkNoiseAppliesToTableauCells 
      Caption         =   "Apply noise  by tableau cell, not by constraint"
      Height          =   255
      Left            =   480
      TabIndex        =   17
      Top             =   600
      Width           =   3495
   End
   Begin VB.TextBox txtValueThatImplementsAPrioriRankings 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3720
      TabIndex        =   14
      Text            =   "20"
      Top             =   4080
      Width           =   975
   End
   Begin VB.PictureBox pctProgressWindow 
      Height          =   7215
      Left            =   4920
      ScaleHeight     =   7155
      ScaleWidth      =   3795
      TabIndex        =   10
      Top             =   240
      Width           =   3855
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit to main screen"
      Height          =   615
      Left            =   840
      TabIndex        =   9
      Top             =   6960
      Width           =   1695
   End
   Begin VB.TextBox txtTimesToTestGrammar 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3720
      TabIndex        =   5
      Text            =   "2000"
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox txtLowerPlasticity 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3720
      TabIndex        =   4
      Text            =   "0.002"
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox txtUpperPlasticity 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3720
      TabIndex        =   3
      Text            =   "2"
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox txtNumberOfCycles 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3720
      TabIndex        =   1
      Text            =   "5000"
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Run Noisy HG"
      Default         =   -1  'True
      Height          =   615
      Left            =   3000
      TabIndex        =   0
      Top             =   6960
      Width           =   1695
   End
   Begin VB.Label lblExact 
      Alignment       =   1  'Right Justify
      Caption         =   "Learning data presented in exact proportions"
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   5160
      Width           =   4575
   End
   Begin VB.Label lblConstraintsRankedAPrioriMustDifferBy 
      Caption         =   "Constraints ranked a priori must differ  by"
      Height          =   255
      Left            =   600
      TabIndex        =   15
      Top             =   4080
      Width           =   2895
   End
   Begin VB.Label lblAPrioriRankings 
      Alignment       =   1  'Right Justify
      Caption         =   "A priori rankings in effect"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   6240
      Width           =   4455
   End
   Begin VB.Label lblCustomRank 
      Alignment       =   1  'Right Justify
      Caption         =   "Custom initial rankings in effect"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   5880
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label lblCustomLearningSchedule 
      Alignment       =   1  'Right Justify
      Caption         =   "Custom learning schedule in effect"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   5520
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   9000
      Y1              =   55
      Y2              =   55
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   0
      X2              =   9000
      Y1              =   50
      Y2              =   50
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Number of times to test grammar"
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Label lblFinalMark 
      Alignment       =   1  'Right Justify
      Caption         =   "Final  plasticity"
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label lblInitMark 
      Alignment       =   1  'Right Justify
      Caption         =   "Initial  plasticity"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Number of times to go through forms"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   3015
   End
   Begin VB.Menu mnuInitialRankings 
      Caption         =   "&Initial rankings"
      Begin VB.Menu mnuUseDefaultInitialRankingValues 
         Caption         =   "Use default initial weights (all zero)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuUseSeparateMarkFaithInitialRankings 
         Caption         =   "Use separate initial weights for Markedness and Faithfulness"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuUseFullyCustomizedInitialRankingValues 
         Caption         =   "Use fully customized initial weights"
      End
      Begin VB.Menu mnuUsePreviousResultsAsInitialRankingValues 
         Caption         =   "Use results of previous run as initial weights"
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSpecifySeparateMarkFaithInitialRankings 
         Caption         =   "Specify &separate initial weights for Markedness and Faithfulness"
      End
      Begin VB.Menu mnuFullyCustomizedInitialRankingValues 
         Caption         =   "Edit file for fully customized initial weights"
      End
   End
   Begin VB.Menu mnuLearningSchedule 
      Caption         =   "&Learning schedule"
      Begin VB.Menu mnuUseCustomLearningSchedule 
         Caption         =   "&Use custom learning schedule from file"
      End
      Begin VB.Menu mnuSepLearningSchedule 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFileWithCustomLearningSchedule 
         Caption         =   "&Edit file with custom learning schedule"
      End
   End
   Begin VB.Menu mnuAPrioriRankings 
      Caption         =   "A Priori Rankings"
      Begin VB.Menu mnuDoGLAWithAPrioriRankings 
         Caption         =   "Run GLA constrained by a priori rankings"
      End
      Begin VB.Menu mnuSepAPriori 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditAprioriRankings 
         Caption         =   "Make or edit a file containing a priori rankings"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuIncludeTableaux 
         Caption         =   "Include tableaux in output file"
      End
      Begin VB.Menu mnuPairwiseRankingProbabilities 
         Caption         =   "Include pairwise ranking probabilities in output"
      End
      Begin VB.Menu mnuSepOptions 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGenerateHistory 
         Caption         =   "Print file with history of ranking values"
      End
      Begin VB.Menu mnuFullHistory 
         Caption         =   "Print file with history of all actions "
      End
      Begin VB.Menu mnuExactProportions 
         Caption         =   "Present data to GLA in exact proportions"
      End
      Begin VB.Menu mnuTestWugOnly 
         Caption         =   "Test wug forms only"
      End
      Begin VB.Menu mnuDemiGaussians 
         Caption         =   "Use positive demi-Gaussians"
      End
      Begin VB.Menu mnuResolveTiesBySkipping 
         Caption         =   "Resolve ties by skipping trial"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuViewManualAsPDF 
         Caption         =   "View manual as Adobe &PDF file"
      End
      Begin VB.Menu mnuViewManualAsWordFile 
         Caption         =   "View manual as Word file"
      End
      Begin VB.Menu mnuAboutOTSoft 
         Caption         =   "About OTSoft"
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "NoisyHarmonicGrammar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=================================NoisyHarnomicGrammar.FRM=================================
'======Weight a set of constraints given a set of input data with violations===============
'===========Use the stochastic algorithm described by Joe Pater and Paul Boersma===========
'=========================in a variety of different forms==================================
   
   Option Explicit
    
    'Localized versions of general OTSoft variables.  These get filled by values taken from Form1
        Dim mInputForm() As String
        Dim mWinner() As String
        Dim mNumberOfForms As Long
        Dim mWinnerFrequency() As Single
        Dim mWinnerViolations() As Long
        Dim mNumberOfRivals() As Long
        Dim mMaximumNumberOfRivals As Long
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
        Dim mTabbedFile As Long
   
   'Frequencies of forms.
      Dim mFrequency() As Single

   'Other variables needed by NHG:
      Dim mWeight() As Single
      Dim mInitialWeight() As Single
      Dim mFrequencyShare() As Single           'Frequency share in the input file.
      Dim mActualFrequencyShare() As Long       'Frequency share as it actually emerged in stochastic learning
      Dim mActualFrequencyPerInput() As Long    'Summing the latter over all rival candidates.
      
    'Different noise values are appropriate for the normal vs. exponential calculations.
        Dim mNoise As Single
        
    'KZ: allows random selection of exemplar:
      Dim mFrequencyInterval() As Single
      
      'Save the number of data presentations so that, when a custom schedule
      ' is use, it can be reported to the user without saving it in OTSoftRememberUserChoices.txt
        Dim mReportedNumberOfDataPresentations  As Long
      
    'Learning schedule parameters.
        Dim mNumberOfLearningStages As Long
        Dim mTrialsPerLearningStage() As Long
        Dim mblnUseCustomLearningSchedule As Boolean
        Dim CustomPlastFaith() As Single
        Dim CustomPlastMark() As Single
                 
    'Testing the grammar:
        Dim mPercentageGenerated() As Single   'Two dimensions:  forms, rivals.
        Dim mTotalNumberOfRivals As Long
        Dim mErrorTerm As Single
        
    'Monitoring the process:
        Dim mNumberOfTimesWeightsWentBelowZero As Long
        Dim mNumberOfChancesForWeightsToGoBelowZero As Long
        Dim mNumberOfTies As Long
        Dim mNumberOfChancesForTie As Long
        'Dim mTieCriterion as single
        Const mcNoWinner As Long = -10000000

    'A priori rankings:
        Dim mUseAPrioriRankings As Boolean
        Dim mNumericalEquivalentOfStrictRanking As Single
      
    'Likelihood, to report
        Dim mLogLikelihood As Single
        Dim mZeroPredictionWarning As Boolean
   
   'Variables for designating files:
      Dim mSimpleHistoryFile As Long
      Dim mFullHistoryFile As Long

   'Final details:
        Dim mTimeMarker As Single
        
    'KZ: keeps track of whether or not user has cancelled learning
        Dim mblnProcessing As Boolean
    
    'KZ: in case user cancels algorithm but then runs it again without exiting the form, the output files need
    '   to be reopened.
        Dim blnOutputFilesOpen As Boolean
    
    'Variables for presenting data in exact frequencies.
        Private Type LearningDatum
            FormIndex As Long
            RivalIndex As Long
        End Type
        Dim mDataPresentationArray() As LearningDatum
        Dim mTotalFrequency As Long
        
    'Monitor and report learning
        Dim mTieWarning As Boolean
        Dim mInfinityWarning As Boolean
        
'==================================INTERFACE ITEMS==========================================

Sub Main(NumberOfForms As Long, InputForm() As String, _
    Winner() As String, WinnerFrequency() As Single, WinnerViolations() As Long, _
    NumberOfRivals() As Long, Rival() As String, RivalFrequency() As Single, RivalViolations() As Long, _
    NumberOfConstraints As Long, ConstraintName() As String, Abbrev() As String, _
    TmpFile As Long, DocFile As Long, HTMFile As Long)

    Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
    
    Let blnOutputFilesOpen = True 'KZ: the files were opened in Main.
    
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
        Let mMaximumNumberOfRivals = Form1.FindMaximumNumberOfRivals(NumberOfRivals())
        ReDim mRival(mNumberOfForms, mMaximumNumberOfRivals)
        ReDim mRivalFrequency(mNumberOfForms, mMaximumNumberOfRivals)
        ReDim mRivalViolations(mNumberOfForms, mMaximumNumberOfRivals, mNumberOfConstraints)
        ReDim mNumberOfRivals(mNumberOfForms)
        'Localize the mNumberOfRivals() array and use it to calculate the maximum, needed for redimensioning.
            For FormIndex = 1 To mNumberOfForms
                Let mNumberOfRivals(FormIndex) = NumberOfRivals(FormIndex)
            Next FormIndex
        For FormIndex = 1 To mNumberOfForms
            Let mInputForm(FormIndex) = InputForm(FormIndex)
            Let mWinner(FormIndex) = Winner(FormIndex)
            Let mWinnerFrequency(FormIndex) = WinnerFrequency(FormIndex)
            For ConstraintIndex = 1 To mNumberOfConstraints
                Let mWinnerViolations(FormIndex, ConstraintIndex) = WinnerViolations(FormIndex, ConstraintIndex)
            Next ConstraintIndex
            For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                Let mRival(FormIndex, RivalIndex) = Rival(FormIndex, RivalIndex)
                Let mRivalFrequency(FormIndex, RivalIndex) = RivalFrequency(FormIndex, RivalIndex)
                For ConstraintIndex = 1 To mNumberOfConstraints
                    Let mRivalViolations(FormIndex, RivalIndex, ConstraintIndex) = RivalViolations(FormIndex, RivalIndex, ConstraintIndex)
                Next ConstraintIndex
            Next RivalIndex
        Next FormIndex
        Let mTmpFile = TmpFile
        Let mDocFile = DocFile
        Let mHTMFile = HTMFile
    
    'Put a caption on the form.
        Let NoisyHarmonicGrammar.Caption = "OTSoft " + gMyVersionNumber + " - Noisy Harmonic Grammar - " + gFileName + gFileSuffix
    
    'Put on the interface some memorized values for the most crucial parameters:
        Let txtNumberOfCycles.Text = Trim(Str(gNumberOfDataPresentations))
        Let txtUpperPlasticity.Text = Trim(Str(gCoarsestPlastMark))
        Let txtLowerPlasticity.Text = Trim(Str(gFinestPlastMark))
        Let txtTimesToTestGrammar.Text = Trim(Str(gCyclesToTest))
        
    'Copy the a priori rankings setting to this part of the program and interface.
        If Form1.mnuConstrainAlgorithmsByAPrioriRankings.Checked = True Then
            Let mnuDoGLAWithAPrioriRankings.Checked = True
            Let lblAPrioriRankings.Visible = True
            Let txtValueThatImplementsAPrioriRankings.Visible = True
            Let lblConstraintsRankedAPrioriMustDifferBy.Visible = True
        Else
            Let mnuDoGLAWithAPrioriRankings.Checked = False
            Let lblAPrioriRankings.Visible = False
            Let txtValueThatImplementsAPrioriRankings.Visible = False
            Let lblConstraintsRankedAPrioriMustDifferBy.Visible = False
        End If
        
    'If a file with initial weights exists, give the user the option of using it.
        If Dir(gOutputFilePath + gFileName + "InitialRankings.txt") <> "" Then
            'There is a custom file, so use these values.
                Let mnuUseFullyCustomizedInitialRankingValues.Visible = True
                Let mnuUseFullyCustomizedInitialRankingValues.Checked = True
                Let lblCustomRank.Visible = True
                Let lblCustomRank.Caption = "Using customized initial weights from file"
                Let InitialRankingChoice = FullyCustomized
                Let mnuUseDefaultInitialRankingValues.Checked = False
        Else
            'Just use the ordinary values--all 100.
                Let mnuUseFullyCustomizedInitialRankingValues.Visible = False
                Let mnuUseDefaultInitialRankingValues.Checked = True
                Let InitialRankingChoice = AllSame
                Let mnuUseFullyCustomizedInitialRankingValues.Visible = False
                Let mnuUseFullyCustomizedInitialRankingValues.Checked = False
                Let lblCustomRank.Visible = False
        End If
        
    'Mark in the relevant menu item whether the users wants to include tableaux.
        Select Case IncludeTableauxInGLAOutput
            Case True
                Let mnuIncludeTableaux.Checked = True
            Case False
                Let mnuIncludeTableaux.Checked = False
        End Select
        
    'Mark in the relevant menu item whether the users wants to use exact input proportions.
        Select Case gExactProportionsForGLAEtc
            Case True
                Let mnuExactProportions.Checked = True
                Let lblExact.Visible = True
            Case False
                Let mnuExactProportions.Checked = False
                Let lblExact.Visible = False
        End Select
        
'        Public gNHGNoiseIsAddedAfterMultiplication As Boolean
    
    'Fill in the check boxes about what particular variety of NHG you would like to run.
        Select Case gNHGLateNoise
            Case True
                Let chkLateNoise.Value = vbChecked
            Case False
                Let chkLateNoise.Value = vbUnchecked
        End Select
        Select Case gNHGNegativeWeightsOK
            Case True
                Let chkNegativeWeightsOK.Value = vbChecked
            Case False
                Let chkNegativeWeightsOK.Value = vbUnchecked
        End Select
        Select Case gNHGNoiseAppliesToTableauCells
            Case True
                Let chkNoiseAppliesToTableauCells.Value = vbChecked
            Case False
                Let chkNoiseAppliesToTableauCells.Value = vbUnchecked
        End Select
        Select Case gNHGNoiseForZeroCells
            Case True
                Let chkNoiseForZeroCells.Value = vbChecked
            Case False
                Let chkNoiseForZeroCells.Value = vbUnchecked
        End Select
        Select Case gNHGNoiseIsAddedAfterMultiplication
            Case True
                Let chkNoiseIsAddedAfterMultiplication.Value = vbChecked
            Case False
                Let chkNoiseIsAddedAfterMultiplication.Value = vbUnchecked
        End Select
        Select Case gExponentialNHG
            Case True
                Let chkExponentialNHG.Value = vbChecked
            Case False
                Let chkExponentialNHG.Value = vbUnchecked
        End Select
        Select Case gDemiGaussianNHG
            Case True
                Let mnuDemiGaussians.Checked = True
            Case False
                Let mnuDemiGaussians.Checked = False
        End Select
        Select Case gResolveTiesBySkipping
            Case True
                Let mnuResolveTiesBySkipping.Checked = vbChecked
            Case False
                Let mnuResolveTiesBySkipping.Checked = vbUnchecked
        End Select
        
        
    'We're ready to go, so let the user see the form.
        NoisyHarmonicGrammar.Show
    
End Sub





Private Sub chkNoiseIsAddedAfterMultiplication_Click()
    'When the user checks this, then the choice of adding noise to zero-violation cells also becomes visible.
        If chkNoiseIsAddedAfterMultiplication.Value = vbChecked Then
            Let chkNoiseForZeroCells.Visible = True
        Else
            Let chkNoiseForZeroCells.Value = vbUnchecked
            Let chkNoiseForZeroCells.Visible = False
        End If
        
End Sub

Private Sub chkExponentialNHG_Click()
        
    'Transfer the choice to the global variable.
        If chkExponentialNHG.Value = vbChecked Then
            Let gExponentialNHG = True
        Else
            Let gExponentialNHG = False
        End If

End Sub

Private Sub mnuDemigaussians_Click()
        
    'Transfer the choice to the global variable.
        If mnuDemiGaussians.Checked = True Then
            Let gDemiGaussianNHG = False
            Let mnuDemiGaussians.Checked = False
        Else
            Let gDemiGaussianNHG = True
            Let mnuDemiGaussians.Checked = True
        End If

End Sub


Private Sub Form_Unload(Cancel As Integer)

    'KZ: must close all files so user can start fresh and apply a different analysis if she wants.
        Close

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




'---------------------------------Initial Weights Menu-------------------------------------------

Private Sub mnuUseDefaultInitialRankingValues_Click()

    'Check, and set choice so as to use the custom values.
        Let mnuUseDefaultInitialRankingValues.Checked = True
        Let InitialRankingChoice = AllSame
        Let lblCustomRank.Visible = False
    
    'Cancel rival methods
        Let mnuUseSeparateMarkFaithInitialRankings.Checked = False
        Let mnuUseFullyCustomizedInitialRankingValues.Checked = False
        Let mnuUsePreviousResultsAsInitialRankingValues.Checked = False
        
End Sub


Private Sub mnuUseSeparateMarkFaithInitialRankings_Click()

    'KZ: check this box and alert user that custom values are being used
        Let mnuUseSeparateMarkFaithInitialRankings.Checked = True
        Let lblCustomRank.Visible = True
        Let lblCustomRank.Caption = "Separate initial weights for Markedness and Faithfulness"
        Let InitialRankingChoice = MarkednessFaithfulness
        
    'Decheck other boxes.
        Let mnuUseDefaultInitialRankingValues.Checked = False
        Let mnuUseFullyCustomizedInitialRankingValues.Checked = False
        Let mnuUsePreviousResultsAsInitialRankingValues.Checked = False

End Sub


Private Sub mnuUseFullyCustomizedInitialRankingValues_Click()

    'Add the check mark, and alter the flag variable, to that a fully
    '    customized initial ranking schedule will be used.
    
    Let mnuUseFullyCustomizedInitialRankingValues.Checked = True
    Let InitialRankingChoice = FullyCustomized
    Let lblCustomRank.Caption = "Using customized initial weights from file"
    Let lblCustomRank.Visible = True
    
    'Suppress the other choices.
        Let mnuUsePreviousResultsAsInitialRankingValues.Checked = False
        Let mnuUseSeparateMarkFaithInitialRankings.Checked = False
        Let mnuUseDefaultInitialRankingValues.Checked = False
    
End Sub

Private Sub mnuUsePreviousResultsAsInitialRankingValues_Click()
    
    'Check to see if there really is a file.
        If Dir(gOutputFilePath + gFileName + "MostRecentWeights.txt") <> "" Then
            Let mnuUsePreviousResultsAsInitialRankingValues.Checked = True
            Let InitialRankingChoice = ValuesFromPreviousRun
            Let lblCustomRank.Visible = True
            Let lblCustomRank.Caption = "Using result of previous run as initial weights"
            'Deactivate other choices.
                Let mnuUseDefaultInitialRankingValues.Checked = False
                Let mnuUseSeparateMarkFaithInitialRankings.Checked = False
                Let mnuUseFullyCustomizedInitialRankingValues.Checked = False
        Else
            MsgBox "Sorry, I can't find the file that stores the weights from the previous run."
        End If
    
End Sub

Private Sub mnuFullyCustomizedInitialRankingValues_Click()

    'Control production of file with fully customized initial weights.
        
    'Make the file by calling a function.  If it returns True, then that means
    '   it succeeded, so let the user know it will be used.
        If MakeFileForFullyCustomizedInitialWeights() = True Then
            Let mnuUseFullyCustomizedInitialRankingValues.Visible = True
            Let mnuUseFullyCustomizedInitialRankingValues.Checked = True
            'And let OTSoft know too:
                Let InitialRankingChoice = FullyCustomized
                Let lblCustomRank.Caption = "Using customized initial weights from file"
                Let lblCustomRank.Visible = True
            'Decheck the other choices.
                Let mnuUseDefaultInitialRankingValues.Checked = False
                Let mnuUsePreviousResultsAsInitialRankingValues.Checked = False
                Let mnuUseSeparateMarkFaithInitialRankings.Checked = False
        End If
        
End Sub

Private Sub mnuSpecifySeparateMarkFaithInitialRankings_Click()

    'KZ: allow user to pick separate values for markedness vs. faithfulness
    '  in the initial weights of constraints.
 
    frmInitialRankings.Show
    
    If blnCustomRankCreated = True Then
        'KZ: if there already are custom ranks in effect, allow user to edit them
        Let frmInitialRankings.txtMark.Text = gCustomRankMark
        Let frmInitialRankings.txtFaith.Text = gCustomRankFaith
    End If  'KZ: (otherwise, defaults are loaded)
    
    Let frmInitialRankings.Visible = True
    
End Sub

Function MakeFileForFullyCustomizedInitialWeights()

    'Open and edit a file for customized initial weights.
    
        Dim IRFile As Long
        Dim ConstraintIndex As Long
        
        'First, make sure there is a folder for this file, a daughter of the
        '   folder in which the input file is located.
            Call Form1.CreateAFolderForOutputFiles
        
        If Dir(gOutputFilePath + gFileName + "InitialRankings.txt") = "" Then
            'You have to make the file anew.
                Let IRFile = FreeFile
                Open gOutputFilePath + gFileName + "InitialRankings.txt" For Output As IRFile
                For ConstraintIndex = 1 To mNumberOfConstraints
                    Print #IRFile, mAbbrev(ConstraintIndex); vbTab; "100"
                Next ConstraintIndex
                Close IRFile
            'If you make it, surely you want to use it.
                Call mnuUseFullyCustomizedInitialRankingValues_Click
        End If

        'Now you know it exists, one way or the other.  Edit it.
            If Dir(gExcelLocation) <> "" Then
                'Shell to Excel.
                Dim Dummy As Long
                Let Dummy = _
                    Shell(gExcelLocation + " " + Chr(34) + gOutputFilePath + gFileName + "InitialRankings.txt" + Chr(34), vbNormalFocus)
            Else
                'Whatever Windows says.
                Call UseWindowsPrograms.TryShellExecute(gOutputFilePath + gFileName + "InitialRankings.txt")
            End If
            
        'Report success.
            Let MakeFileForFullyCustomizedInitialWeights = True
            
        Exit Function

CheckError:

    MsgBox "Program error:  I was unable to edit a file with initial weights.  Please report this bug to Bruce Hayes " + _
        "(bhayes@humnet.ucla.edu), specifying error #11198."
    Let MakeFileForFullyCustomizedInitialWeights = False
    
End Function

'-------------------------------Learning Schedule Menu--------------------------------------------

Private Sub mnuUseCustomLearningSchedule_Click()
    'Toggle the choice of using Custom learning schedule.
        If mnuUseCustomLearningSchedule.Checked = False Then
            Let mnuUseCustomLearningSchedule.Checked = True
            Let mblnUseCustomLearningSchedule = True
        'Alter the interface to show what's happened.
            Let txtNumberOfCycles.Text = "(cust.)"
            Let txtUpperPlasticity = "(cust.)"
            Let txtLowerPlasticity = "(cust.)"
            Let lblCustomLearningSchedule.Caption = "Customized learning schedule in effect"
            Let lblCustomLearningSchedule.Visible = True
        Else
            Let mnuUseCustomLearningSchedule.Checked = False
            Let lblCustomLearningSchedule.Visible = False
            Let mblnUseCustomLearningSchedule = False
            Let lblCustomLearningSchedule.Visible = False
            'Reinstall on the interface some memorized values for the most crucial parameters:
                Let txtNumberOfCycles.Text = Trim(Str(gNumberOfDataPresentations))
                Let txtUpperPlasticity.Text = Trim(Str(gCoarsestPlastMark))
                Let txtLowerPlasticity.Text = Trim(Str(gFinestPlastMark))
                Let txtTimesToTestGrammar.Text = Trim(Str(gCyclesToTest))
        End If
End Sub

Private Sub mnuEditFileWithCustomLearningSchedule_Click()

    'Edit the custom learning schedule.
    
        Dim CusLSched As Long
        Dim MessageBoxResult As Long
    
    'Make sure there is a folder for this file.
        Call Form1.CreateAFolderForOutputFiles
    
    'If a custom learning schedule file doesn't exist, prepare a default one.
        If Dir(gOutputFilePath + gFileName + "CustomLearningSchedule.txt") = "" Then
            Let CusLSched = FreeFile
            Open gOutputFilePath + gFileName + "CustomLearningSchedule.txt" For Output As CusLSched
            Print #CusLSched, "Trials"; vbTab; "PlastMark"; vbTab; "PlastFaith"
            Print #CusLSched, "15000"; vbTab; "2"; vbTab; "2"; vbTab; "2"; vbTab; "2"
            Print #CusLSched, "15000"; vbTab; ".2"; vbTab; ".2"; vbTab; "2"; vbTab; "2"
            Print #CusLSched, "15000"; vbTab; ".02"; vbTab; ".02"; vbTab; "2"; vbTab; "2"
            Print #CusLSched, "15000"; vbTab; ".002"; vbTab; ".002"; vbTab; "2"; vbTab; "2"
            Close #CusLSched
        End If
        
    'Now let the user edit the schedule, using Excel if it's findable, else Notepad or
    '   whatever Windows wants.
        If Dir(gExcelLocation) <> "" Then
            'Shell to Excel.
            Dim Dummy As Long
            Let Dummy = _
                Shell(gExcelLocation + " " + Chr(34) + gOutputFilePath + gFileName + "CustomLearningSchedule.txt" + Chr(34), vbNormalFocus)
        Else
            'Whatever Windows says.
            Let MessageBoxResult = MsgBox("To edit your custom learning schedule, it would be most convenient to use Excel or some other spreadsheet program to edit your custom learning schedule file.  But I can't find your program.  The place where I'm looking is" + _
                Chr(10) + Chr(10) + _
                gExcelLocation + _
                Chr(10) + Chr(10) + _
                "But I can't find it there.  You can specify the correct location by opening the file" + _
                Chr(10) + Chr(10) + _
                App.Path + "\OTSoftAuxiliarySoftwareLocations.txt" + _
                Chr(10) + Chr(10) + _
                "and typing it in." + _
                Chr(10) + Chr(10) + _
                "Click Yes to edit OTSoftAuxiliarySoftwareLocations.txt (OTSoft will exit), No to edit your custom learning schedule file with whatever program your computer currently uses to open a .txt file.", vbYesNo)
            Select Case MessageBoxResult
                Case vbYes
                    Call UseWindowsPrograms.TryShellExecute(App.Path + "\OTSoftAuxiliarySoftwareLocations.txt")
                    'Call UseWindowsPrograms.TryShellExecute(gSafePlaceToWriteTo + "\OTSoftAuxiliarySoftwareLocations.txt")
                    Call Form1.cmdExit_Click
                Case vbNo
                    Call UseWindowsPrograms.TryShellExecute(gOutputFilePath + gFileName + "CustomLearningSchedule.txt")
            End Select
        End If
        
'Former version:  gSafePlaceToWriteTo + "\OTSoftAuxiliarySoftwareLocations.txt" + _

    'Since you edited it, you probably want to use it.
        Let mnuUseCustomLearningSchedule.Checked = True
        Let mblnUseCustomLearningSchedule = True
        Let txtNumberOfCycles.Text = "(cust.)"
        Let txtUpperPlasticity = "(cust.)"
        Let txtLowerPlasticity = "(cust.)"
        Let lblCustomLearningSchedule.Caption = "Customized learning schedule in effect"
        Let lblCustomLearningSchedule.Visible = True

End Sub

'---------------------------------A Priori Rankings Menu-------------------------------------------

Private Sub mnuDoGLAWithAPrioriRankings_Click()
    
    'Toggle the use of a priori rankings.
    
    If mnuDoGLAWithAPrioriRankings.Checked = False Then
        'Don't do anything if there's no file.
            If Dir(gOutputFilePath + "\" + gFileName + "apriori.txt") = "" Then
                MsgBox "I can't use a priori rankings until you make a file that specifies them.  You can do this by selecting " + _
                    Chr(34) + _
                    "Make a file for a priori rankings" + _
                    Chr(34) + _
                    " from the " + _
                    Chr(34) + _
                    "A priori rankings" + Chr(34) + " menu."
                Exit Sub
            End If
        Let mnuDoGLAWithAPrioriRankings.Checked = True
        Let Form1.mnuConstrainAlgorithmsByAPrioriRankings.Checked = True
        Let lblAPrioriRankings.Visible = True
        Let txtValueThatImplementsAPrioriRankings.Visible = True
        Let lblConstraintsRankedAPrioriMustDifferBy.Visible = True
    Else
        Let mnuDoGLAWithAPrioriRankings.Checked = False
        Let Form1.mnuConstrainAlgorithmsByAPrioriRankings.Checked = False
        Let lblAPrioriRankings.Visible = False
        Let txtValueThatImplementsAPrioriRankings.Visible = False
        Let lblConstraintsRankedAPrioriMustDifferBy.Visible = False
    End If
    
End Sub

Private Sub mnuEditAprioriRankings_Click()
    'Make or edit file for a priori rankings.
        Call APrioriRankings.PrintOutTemplateForAPrioriRankings(mNumberOfConstraints, mAbbrev())
    'If they did this, assume they want to use it.
        Let mnuDoGLAWithAPrioriRankings.Checked = True
        Let Form1.mnuConstrainAlgorithmsByAPrioriRankings.Checked = True
        Let lblAPrioriRankings.Visible = True
End Sub



'---------------------------------------Options Menu-------------------------------------------

Private Sub mnuIncludeTableaux_Click()
    If mnuIncludeTableaux.Checked = False Then
        Let mnuIncludeTableaux.Checked = True
        Let IncludeTableauxInGLAOutput = True
    Else
        Let mnuIncludeTableaux.Checked = False
        Let IncludeTableauxInGLAOutput = False
    End If
End Sub

Private Sub mnuPairwiseRankingProbabilities_Click()
    If mnuPairwiseRankingProbabilities.Checked = False Then
        Let mnuPairwiseRankingProbabilities.Checked = True
    Else
        Let mnuPairwiseRankingProbabilities.Checked = False
    End If
End Sub
Private Sub mnuGenerateHistory_Click()
    If mnuGenerateHistory.Checked = False Then
        Let mnuGenerateHistory.Checked = True
    Else
        Let mnuGenerateHistory.Checked = False
    End If
End Sub
Private Sub mnuFullHistory_Click()
    If mnuFullHistory.Checked = False Then
        Let mnuFullHistory.Checked = True
    Else
        Let mnuFullHistory.Checked = False
    End If
End Sub

Private Sub mnuExactProportions_Click()

    'Toggle the Exact proportions (data for GLA) menu item.
        If mnuExactProportions.Checked = False Then
            Let mnuExactProportions.Checked = True
            'Save this information in the global boolean variable, for the ini file:
                Let gExactProportionsForGLAEtc = True
            'Indicate this on the interface
                Let lblExact.Visible = True
            'Close off the option of custom learning schedule; this is unimplemented for
            '   exact proportions.
                Let mnuUseCustomLearningSchedule.Checked = False
                Let mblnUseCustomLearningSchedule = False
        Else
            Let mnuExactProportions.Checked = False
            'Save this information in the global boolean variable, for the ini file:
                Let gExactProportionsForGLAEtc = False
                Let lblExact.Visible = False
        End If

End Sub

Private Sub mnuTestWugOnly_Click()
    'Toggle
        If mnuTestWugOnly.Checked = True Then
            Let mnuTestWugOnly.Checked = False
            Let TestWugOnly = False
        Else
            Let mnuTestWugOnly.Checked = True
            Let TestWugOnly = True
        End If
End Sub

Private Sub mnuResolveTiesBySkipping_Click()
    'Toggle
        If mnuResolveTiesBySkipping.Checked = True Then
            Let mnuResolveTiesBySkipping = False
            Let gResolveTiesBySkipping = False
        Else
            Let mnuResolveTiesBySkipping.Checked = True
            Let gResolveTiesBySkipping = True
        End If
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

Private Sub mnuExit_Click()
    Close
    End
End Sub

'------------------------------Rest of GLA Interface--------------------------------
Private Sub cmdRun_Click()
    
    Call RunNoisyHarmonicGrammar

End Sub

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


'================================MAIN CALCULATION ROUTINE====================================

Private Sub RunNoisyHarmonicGrammar()

    'This is the primary routine for this form, which calls all other routines
    '   needed for running NoisyHarmonicGrammar.
    
    'Dimension arrays needed by the algorithm, according to the size of the problem:
        ReDim mWeight(mNumberOfConstraints)
        ReDim mInitialWeight(mNumberOfConstraints)
        ReDim mFrequencyShare(mNumberOfForms, mMaximumNumberOfRivals)
        ReDim mActualFrequencyShare(mNumberOfForms, mMaximumNumberOfRivals)
        ReDim mActualFrequencyPerInput(mNumberOfForms)
        
    'KZ: third dimension gives the lower(inclusive) and upper(exclusive)
    '   bounds of the interval in [0,1] assigned to that form.
        ReDim mFrequencyInterval(mNumberOfForms, mMaximumNumberOfRivals, 1)
        ReDim mFrequency(mNumberOfForms, mMaximumNumberOfRivals)
        ReDim mFaithfulness(mNumberOfConstraints) 'KZ

    'KZ: the variable mblnProcessing keeps track of whether or not the algorithm is running:
    If mblnProcessing = True Then
        'KZ: when the algorithm is running, clicking this button again stops it
            Let mblnProcessing = False 'KZ: return button to "not-running" state and shut down the learner
    Else    'KZ: otherwise, run the algorithm
        'KZ: let the user know this is now the cancel button:
            Let cmdRun.Caption = "Cancel"
            Let mblnProcessing = True
        'Cover a weird contingency:  user cancelled learning, then started it again without exiting to the main form (KZ)
            If blnOutputFilesOpen = False Then
                'First, make sure there is a folder for these files, a daughter of the
                '   folder in which the input file is located.
                    Call Form1.CreateAFolderForOutputFiles
                'The draft file:
                    Let mTmpFile = FreeFile
                    Open gOutputFilePath + gFileName + "DraftOutput.txt" For Output As #mTmpFile
                'Initialize the header numbers, in case this isn't the first run.
                    Let gLevel1HeadingNumber = 0
                'The fancy file:
                    Let mDocFile = FreeFile
                    Open gOutputFilePath + gFileName + "QualityOutput.txt" For Output As #mDocFile
            End If
        
        Call ObtainInformationFromMainWindow
        
                    'Debug
                    'Dim DebugFile As Long
                    'Let DebugFile = FreeFile
                    'Open gOutputFilePath + "DebugAtEarlyStageOfNHGRoutine.txt" For Output As #DebugFile
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
        
        'Save the choices that the user put in the checkboxes.
            If chkLateNoise.Value = vbChecked Then
                Let gNHGLateNoise = True
            Else
                Let gNHGLateNoise = False
            End If
            If chkNegativeWeightsOK.Value = vbChecked Then
                Let gNHGNegativeWeightsOK = True
            Else
                Let gNHGNegativeWeightsOK = False
            End If
            If chkNoiseAppliesToTableauCells.Value = vbChecked Then
                Let gNHGNoiseAppliesToTableauCells = True
            Else
                Let gNHGNoiseAppliesToTableauCells = False
            End If
            If chkNoiseForZeroCells.Value = vbChecked Then
                Let gNHGNoiseForZeroCells = True
            Else
                Let gNHGNoiseForZeroCells = False
            End If
            If chkNoiseIsAddedAfterMultiplication.Value = vbChecked Then
                Let gNHGNoiseIsAddedAfterMultiplication = True
            Else
                Let gNHGNoiseIsAddedAfterMultiplication = False
            End If
            Call Form1.SaveUserChoices

        'Caution users about high noise in Exponential NHG.
            If Val(txtUpperPlasticity.Text) >= 0.5 Then
                If gExponentialNHG Then
                    MsgBox "Caution:  for Exponential Noisy Harmony Grammar, it may produce aberrant results to use an upper plasticity value as high as " _
                        + txtUpperPlasticity.Text + ". I suggest you terminate this run and enter a value no higher than 0.1, then click the Run button again."
                    Let cmdRun.Caption = "Run"
                    Let mblnProcessing = False
                    Exit Sub
                End If
            End If
            
        'Caution users about the nonstochastic version of the theory:
            If chkNoiseIsAddedAfterMultiplication.Value = vbChecked Then
                If chkNoiseForZeroCells.Value = vbChecked Then
                    If chkNoiseAppliesToTableauCells.Value <> vbChecked Then
                        If chkLateNoise <> vbChecked Then
                            If chkNegativeWeightsOK = vbChecked Then
                                MsgBox "Caution:  by the choices you have made, the result will be non-stochastic (noise has no effect, " + _
                                "one single winner per input)."
                            End If
                        End If
                    End If
                End If
            End If
        
        'Clear the progress window.
            pctProgressWindow.Cls
        'Start timing.
            Let mTimeMarker = Timer
        'Report progress.
            pctProgressWindow.Print "Learning..."
        
        'NoisyHarmonicGrammarPreliminaries:
        'If all goes will in the preliminary operations, NoisyHarmonicGrammarPreliminaries() will return True, and you can continue.
            If NoisyHarmonicGrammarPreliminaries() = False Then
                'Something went wrong in the preliminaries.  Go back to ur-state.
                    Let mblnProcessing = False
                'Reactivate the various buttons that start things off, and give up.
                    Let Form1.cmdRank.Enabled = True
                    Let Form1.cmdFacType.Enabled = True
                    Let cmdRun.Caption = "&Run NHG"
                Exit Sub
            End If
            
        Call NoisyHarmonicGrammarCore    'KZ: NoisyHarmonicGrammarCore checks periodically to see if the button
                                         'has been clicked again (to cancel). If user cancels,
                                         'mblnProcessing gets set to False.
        
        'KZ: don't bother with this part if user has cancelled:
            If mblnProcessing = True Then
                Call PrintAHeader("Noisy Harmonic Grammar")
                Call PrintNHGResults(mWeight(), "Weights")
                'If you're using a priori rankings, look them up and implement them  as initial values.
                    If mUseAPrioriRankings = True Then
                        Call Form1.PrintOutTheAprioriRankings(mTmpFile, mDocFile, mHTMFile)
                    End If
                'NHGTestGrammar returns a value, namely the degree of error.
                '   This permits it to be used for hill-climbing learning.
                '   The second parameter is True if one wants a printed report.
                '   KZ: NHGTestGrammar also checks for cancellation.
                    Let mErrorTerm = NHGTestGrammar(mWeight(), True, gCyclesToTest)
                Call PrepareApproximateTableaux     'This also prints a table
                                                    '  converting weights to probability.
                Call PrintFinalDetails
            End If
        
        'Close output files.
            Close #mTmpFile
            Close #mDocFile
            Print #mHTMFile, "</BODY>"
            Close #mHTMFile
            Let blnOutputFilesOpen = False 'KZ
            
        'KZ: get rid of "cancel" caption:
            Let cmdRun.Caption = "&Run NHG"
        
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
            
            'KZ: I'm changing this from Hide to Unload. Otherwise,
            'the form doesn't get reloaded again when changing to a
            'new file, and the arrays don't get re-dimensioned.
                Unload Me
            
        Else
            
            'KZ: extra stuff to do if user cancelled:
                pctProgressWindow.Cls
                pctProgressWindow.Print "Learning Cancelled."
                Close #mSimpleHistoryFile
                Close #mFullHistoryFile
        
        End If
        
    End If  'Is mblnProcessing true?  (i.e. Rank or Cancel)

End Sub


Sub ObtainInformationFromMainWindow()
  
    'Other parameters are obtained from the interface of the main OTSoft window.
    '   This routine only grabs the parameters from the NHG window, and therefore
    '   has to run after the user has clicked Run.
    
    'Ask the Noisy Harmonic Grammar interface how many trials are wanted.
        Let gCyclesToTest = Val(NoisyHarmonicGrammar.txtTimesToTestGrammar.Text)
         
    'Determine if apriori rankings are in effect
        If mnuDoGLAWithAPrioriRankings.Checked = True Then
            Let mUseAPrioriRankings = True
        Else
            Let mUseAPrioriRankings = False
        End If

End Sub



Function NoisyHarmonicGrammarPreliminaries() As Boolean

   'Execute preliminary actions needed by Noisy Harmonic Grammar learning.
   
    'Variables of convenience for reading the file with initial weights.
        Dim MyLine As String, IRFile As Long
        Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
    
    'To assign plasticity, initial weights, and noise differently to Markedness
    '   and Faithfulness constraints, you need to look up their status.
        For ConstraintIndex = 1 To mNumberOfConstraints
            Let mFaithfulness(ConstraintIndex) = Form1.FaithfulnessConstraint(mConstraintName(ConstraintIndex))
        Next ConstraintIndex
        
    'Note:  important to run Randomize just once; else a subtle bug.
        Randomize
        
        Call InstallWinnerAsRival
        
        'This is not working.
        'Call DebugMe(mNumberOfForms, mInputForm(), mMaximumNumberOfRivals, mNumberOfRivals(), mRival(), mRivalViolations(), mFrequency(), _
            mNumberOfConstraints, mConstraintName())
        
        Call FrequencyThresholds
        Call SetTheNoise
        
        'I am weary of converting all my ancient gosubs and am turning them into GoTo's so that they
        '   will run in VB2008.
            GoTo InitialWeights
InitialWeightsReturnPoint:

    'If you're presenting data in exact proportions, form a chart for this purpose.
        If mnuExactProportions.Checked = True Then
            Call FormDataPresentationArray
        End If
      
    'Set up the possibility of tracing the history of ranking.
    '   Not done if its just running more tests on an existing grammar.
        If gNumberOfDataPresentations > 0 Or mReportedNumberOfDataPresentations > 0 Then
            If mnuGenerateHistory.Checked = True Then
                GoTo OpenExcelFile
OpenExcelFileReturnPoint:
            End If
            If mnuFullHistory.Checked = True Then
                GoTo OpenFullHistoryFile
OpenFullHistoryFileReturnPoint:
            End If
        End If
    
    'If you're using a priori rankings, look them up and implement them as initial weights.
        If mUseAPrioriRankings = True Then
            'Vet the value that implements an a priori ranking numerically.
                If GoodDecimal(txtValueThatImplementsAPrioriRankings.Text) = False Then
                    MsgBox "Please enter a valid number in the box labeled " + _
                        Chr(34) + "Constraints ranked a priori must differ by" + Chr(34)
                    Let NoisyHarmonicGrammarPreliminaries = False
                    Exit Function
                End If
            Let mNumericalEquivalentOfStrictRanking = Val(txtValueThatImplementsAPrioriRankings)
            'ReadAPrioriRankingsAsTable is a boolean function, which returns False if
            '   it failed to do its job.
            If APrioriRankings.ReadAPrioriRankingsAsTable(mNumberOfConstraints, mAbbrev) = True Then
                'This must happen after the history files are opened, since this routine records
                '   itself when history is being taken.
                    Call AdjustAPrioriRankings_Up
                    'xxx needs to be done later or deleted.
            Else
                Let Form1.mnuConstrainAlgorithmsByAPrioriRankings.Checked = False
            End If
            
        End If
        
    'Determine the schedule for the learning parameters.  If an error, this will return as False.
        If DetermineLearningSchedule() = False Then
            Let NoisyHarmonicGrammarPreliminaries = False
            Exit Function
        End If

    'Record the initial weights, as part of notating history of the run.
        Call RecordInitialWeights
      
    'All is well, so set value as True and exit.
        Let NoisyHarmonicGrammarPreliminaries = True

        Exit Function


Stop
'------------------------------------------------------------------------------
InitialWeights:

    'Set the weight of every constraint at Boersma's arbitrary value of 100.
    '   KZ: or, if user wants, use custom values for faithfulness and markedness
    '   or from file--useful for running more.
    
    Select Case InitialRankingChoice
        Case AllSame
            'Let's go with zero as the default.
                For ConstraintIndex = 1 To mNumberOfConstraints
                    Let mWeight(ConstraintIndex) = 0
                Next ConstraintIndex
        Case MarkednessFaithfulness
            'Invoke user's choice for markedness and Faithfulness.
                For ConstraintIndex = 1 To mNumberOfConstraints
                    If mFaithfulness(ConstraintIndex) = True Then
                        Let mWeight(ConstraintIndex) = gCustomRankFaith
                    Else
                        Let mWeight(ConstraintIndex) = gCustomRankMark
                    End If
                Next ConstraintIndex
        Case FullyCustomized, ValuesFromPreviousRun
            'Read the file with customized choices, or choices saved from last run.
                Let IRFile = FreeFile
                'Custom values or old values?  Open appropriate file.
                    Select Case InitialRankingChoice
                        Case FullyCustomized
                            'Don't ever try to open a nonexistent file.
                                If Dir(gOutputFilePath + gFileName + "InitialRankings.txt") <> "" Then
                                    Open gOutputFilePath + gFileName + "InitialRankings.txt" For Input As IRFile
                                Else
                                    MsgBox "Sorry, I can't find the file containing your initial weights.  It is supposed to be at " + _
                                        gOutputFilePath + gFileName + "InitialRankings.txt.  Click OK to continue."
                                    Let NoisyHarmonicGrammarPreliminaries = False
                                    Exit Function
                                End If
                        Case ValuesFromPreviousRun
                            'Don't ever try to open a nonexistent file.
                                If Dir(gOutputFilePath + gFileName + "MostRecentWeights.txt") <> "" Then
                                    Open gOutputFilePath + gFileName + "MostRecentWeights.txt" For Input As IRFile
                                Else
                                    MsgBox "Sorry, I can't find the file containing your most recent rankings.  It is supposed to be at " + _
                                        gOutputFilePath + gFileName + "MostRecentWeights.txt.  Click OK to continue."
                                    Let NoisyHarmonicGrammarPreliminaries = False
                                    Exit Function
                                End If
                    End Select
                'Read the values off the file.
                    For ConstraintIndex = 1 To mNumberOfConstraints
                        Line Input #IRFile, MyLine
                        Let MyLine = Trim(MyLine)
                        Let mWeight(ConstraintIndex) = Val(Trim(s.Residue(MyLine)))
                    Next ConstraintIndex
                Close IRFile
                
    End Select
    
    'Save the initial weights for future reporting.
        For ConstraintIndex = 1 To mNumberOfConstraints
            Let mInitialWeight(ConstraintIndex) = mWeight(ConstraintIndex)
        Next ConstraintIndex
    
    'There are no more Return statements anymore, so just use a GoTo.
        GoTo InitialWeightsReturnPoint

Stop
'-----------------------------------------------------------------------------
OpenExcelFile:

    'History of weights, for an Excel file.
    '  Open file, and print a header:

    'First, make sure there is a folder for these files, a daughter of the
    '   folder in which the input file is located.
        Call Form1.CreateAFolderForOutputFiles
        
        Let mSimpleHistoryFile = FreeFile
        Open gOutputFilePath + gFileName + "History" + ".xls" For Output As #mSimpleHistoryFile
        Print #mSimpleHistoryFile, mAbbrev(1);
        For ConstraintIndex = 2 To mNumberOfConstraints
            Print #mSimpleHistoryFile, Chr$(9); mAbbrev(ConstraintIndex);
        Next ConstraintIndex
        Print #mSimpleHistoryFile,

    'There are no more Return statements anymore, so just use a GoTo.
        GoTo OpenExcelFileReturnPoint

Stop
'-----------------------------------------------------------------------------
OpenFullHistoryFile:

    'Complete history of how ranking was done.
    '  Open file, and print a header:

       On Error GoTo OpenFullHistoryFileErrorPoint
        
        'First, make sure there is a folder for these files, a daughter of the
        '   folder in which the input file is located.
            Call Form1.CreateAFolderForOutputFiles
        
        Let mFullHistoryFile = FreeFile
        Open gOutputFilePath + gFileName + "FullHistory" + ".xls" For Output As #mFullHistoryFile
        'Header:
            Print #mFullHistoryFile, "Trial"; vbTab; "Input"; vbTab; "Generated"; vbTab; "Heard"; vbTab;
            
            'Third column, adjustment number, only if apriori rankings being used
                If mUseAPrioriRankings = True Then
                    Print #mFullHistoryFile, "Adj. Num"; vbTab;
                End If
            
            'Header for each constraint:  delta, then new value.
                For ConstraintIndex = 1 To mNumberOfConstraints
                    Print #mFullHistoryFile, mAbbrev(ConstraintIndex); vbTab; "now"; vbTab;
                Next ConstraintIndex
            'End of line
                Print #mFullHistoryFile,
                
    'There are no more Return statements anymore, so just use a GoTo.
        GoTo OpenFullHistoryFileReturnPoint

OpenFullHistoryFileErrorPoint:

    Select Case Err.Number
        Case 70
            MsgBox "Error.  I conjecture that the file " + gOutputFilePath + gFileName + "FullHistory" + ".xls" + _
                " is already open.  Please close this file, then run OTSoft again."
            End
        Case Else
            MsgBox "Program error.  For help please contact me at bhayes@humnet.ucla.edu, enclosing your input file and specifying error #14897."
    End Select

End Function

Sub SetTheNoise()

    'It needs to be different if you are using exponentiation.
        If gExponentialNHG Then
            Let mNoise = 0.1
        Else
            Let mNoise = 1
        End If
    
End Sub


Sub FrequencyThresholds()

   'Forms must be input to Noisy Harmonic Grammar learning in frequencies that match their real-world counts.
   '  We can do this by keep track of proportions.
   '  This is done on a candidate-by-candidate basis.
   
    Dim FormIndex As Long, RivalIndex As Long
    Dim TotalFrequencies As Single
    Dim PreviousUpperBoundOfInterval As Single

   'Loop through all forms and rivals, and find the sum of the frequencies of all candidates.
        Let TotalFrequencies = 0
        For FormIndex = 1 To mNumberOfForms
            'Dec. 2025 I think the 0 is an error -- we have already folded in the winner as a rival.
                 'For RivalIndex = 0 To mNumberOfRivals(FormIndex)
            For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                Let TotalFrequencies = TotalFrequencies + mFrequency(FormIndex, RivalIndex)
            Next RivalIndex
        Next FormIndex

   'Go through all candidates, dividing their frequency by the total.
      'KZ: I'm assigning each rival/winner an interval in [0,1]. When
      'selecting an exemplar, generate a random number and see whose interval
      'it falls in.
      Let PreviousUpperBoundOfInterval = 0
      For FormIndex = 1 To mNumberOfForms
            'I think the 0 is an error -- we have already folded in the winner as a rival.
                 'For RivalIndex = 0 To mNumberOfRivals(FormIndex)
         For RivalIndex = 1 To mNumberOfRivals(FormIndex)
             If mFrequency(FormIndex, RivalIndex) > 0 Then
                Let mFrequencyShare(FormIndex, RivalIndex) = mFrequency(FormIndex, RivalIndex) / TotalFrequencies
                'KZ: lower bound of interval:
                Let mFrequencyInterval(FormIndex, RivalIndex, 0) = PreviousUpperBoundOfInterval
                'KZ: upper bound of interval:
                Let mFrequencyInterval(FormIndex, RivalIndex, 1) = PreviousUpperBoundOfInterval _
                 + mFrequencyShare(FormIndex, RivalIndex)
                'KZ: update PrevUpperBound:
                Let PreviousUpperBoundOfInterval = PreviousUpperBoundOfInterval _
                 + mFrequencyShare(FormIndex, RivalIndex)
             End If
         Next RivalIndex
      Next FormIndex
      
      'Now:  you can produce a random number, go through the thresholds,
      '  and find the candidate with the highest threshold that is less
      '  than the random number that is generated.

End Sub


Sub FormDataPresentationArray()

    'Form an array, DataPresentationArray(), in which the rows are presentations,
    'and the columns are:
    '   form
    '   rival
    
    'The number of rows for each form/rival pair is the mFrequency() of that pair,
    '   as it appears in the input file.
    
    'This array is randomized each time it is run through, and learning data
    '   are selected from it.  The idea is to make the presentation
    '   frequencies exact, but the presentation order random.
    
    Dim FormIndex As Long, RivalIndex As Long, FrequencyIndex As Long
    
    For FormIndex = 1 To mNumberOfForms
                'Dec. 2025 I think the 0 is an error -- we have already folded in the winner as a rival.
                 'For RivalIndex = 0 To mNumberOfRivals(FormIndex)
        For RivalIndex = 1 To mNumberOfRivals(FormIndex)
            'Check that this frequency is an integer.  Report -1 and exit if not.
                If mFrequency(FormIndex, RivalIndex) <> Int(mFrequency(FormIndex, RivalIndex)) Then
                    MsgBox "Caution:  you've requested exact frequency matching, but OTSoft can do this only if frequencies are all whole numbers.  Please fix your input file first if you want to use this method.  When you click ok, OTSoft will proceed with inexact frequency matching."
                    Let mTotalFrequency = -1
                    Exit Sub
                End If
            'Add enough rows to the array, specifying this learning datum.
                For FrequencyIndex = 1 To mFrequency(FormIndex, RivalIndex)
                    Let mTotalFrequency = mTotalFrequency + 1
                    ReDim Preserve mDataPresentationArray(mTotalFrequency)
                    Let mDataPresentationArray(mTotalFrequency).FormIndex = FormIndex
                    Let mDataPresentationArray(mTotalFrequency).RivalIndex = RivalIndex
                Next FrequencyIndex
        Next RivalIndex
    Next FormIndex
    
    Exit Sub
    
    'Debug:
    '    Dim DebugFile As Long
    '    Let DebugFile = FreeFile
    '    Open gInputFilePath + "debug.txt" For Output As #DebugFile
    '    For FrequencyIndex = 1 To mTotalFrequency
    '        Print #DebugFile, FrequencyIndex; vbTab; mDataPresentationArray(FrequencyIndex).FormIndex;
    '        Print #DebugFile, vbTab; mDataPresentationArray(FrequencyIndex).RivalIndex
    '    Next FrequencyIndex
    '    Close #DebugFile

End Sub

Function DetermineLearningSchedule() As Boolean

   'Get a learning schedule (trials, plasticity/noise for faithfulness/markedness)
   '    either off a file, or by interpreting the interface.
   
    Dim LearningStageIndex As Long
   
    If mblnUseCustomLearningSchedule Then
       'Open the file with the learning schedule and read it.
            'First make sure it exists.
            If Dir(gOutputFilePath + gFileName + "CustomLearningSchedule.txt") = "" Then
                'No file.  Explain the problem and exit.
                    MsgBox "You've requested a custom learning schedule, but I can't find the file containing it.  It's supposed to be at" + _
                        gOutputFilePath + gFileName + "CustomLearningSchedule.txt" + _
                        "You can edit such a file by selecting this choice on the Learning Schedule menu." + _
                        "Click OK to return to the GLA screen."
                    Let DetermineLearningSchedule = False
                    Exit Function
            Else    'The file does exist, so you're ok to proceed.
                'Apparatus for opening the file.
                    Dim CusLSched As Long
                    Dim MyLine As String
                    Dim Buffer As String
                    Let CusLSched = FreeFile
                    'Don't ever try to open a nonexistent file.
                        If Dir(gOutputFilePath + gFileName + "CustomLearningSchedule.txt") <> "" Then
                            Open gOutputFilePath + gFileName + "CustomLearningSchedule.txt" For Input As CusLSched
                        Else
                            MsgBox "Sorry, I can't find the file that has your custom learning schedule.  It is supposed to be in " + _
                                gOutputFilePath + gFileName + "CustomLearningSchedule.txt.  Click OK to continue."
                            Let DetermineLearningSchedule = False
                            Exit Function
                        End If
                'Read the initial line with headings.
                    Line Input #CusLSched, MyLine
                'Loop through the rest of the file line by line.
                    Do While Not EOF(CusLSched)
                        'Grab a line.
                            Line Input #CusLSched, MyLine
                        'Each line is the parameters for another learning stage, so redimension.
                            Let mNumberOfLearningStages = mNumberOfLearningStages + 1
                            ReDim Preserve mTrialsPerLearningStage(mNumberOfLearningStages)
                            ReDim Preserve CustomPlastMark(mNumberOfLearningStages)
                            ReDim Preserve CustomPlastFaith(mNumberOfLearningStages)
                        'Chomp up the line, recording the values.  Warn the user and
                        '   exit if there are any blanks.
                            Let Buffer = s.Chomp(MyLine)
                                If Trim(Buffer) = "" Then GoTo BlankCell
                                Let mTrialsPerLearningStage(mNumberOfLearningStages) = Val(Buffer)
                                Let MyLine = s.Residue(MyLine)
                            Let Buffer = s.Chomp(MyLine)
                                If Trim(Buffer) = "" Then GoTo BlankCell
                                Let CustomPlastMark(mNumberOfLearningStages) = Val(Buffer)
                                Let MyLine = s.Residue(MyLine)
                            Let Buffer = s.Chomp(MyLine)
                                If Trim(Buffer) = "" Then GoTo BlankCell
                                Let CustomPlastFaith(mNumberOfLearningStages) = Val(Buffer)
                                Let MyLine = s.Residue(MyLine)
                    Loop
                'You're done with this file now.
                    Close #CusLSched
                'Determine the total number of learning trials, for report purposes.
                    'This should not be the "real" value--since we don't want it saved
                    '   to OTSoftRememberUserChoices.txt.
                    Let mReportedNumberOfDataPresentations = 0
                    For LearningStageIndex = 1 To mNumberOfLearningStages
                        Let mReportedNumberOfDataPresentations = mReportedNumberOfDataPresentations + mTrialsPerLearningStage(LearningStageIndex)
                    Next LearningStageIndex
                'Alter the interface to show what's happened.
                    Let txtNumberOfCycles.Text = "(cust.)"
                    Let txtUpperPlasticity = "(cust.)"
                    Let txtLowerPlasticity = "(cust.)"
            End If                      'Does there exist a file with a learning schedule?
    Else                                'Normal or custom learning schedule?
        'Construct a schedule off the interface.
            Let mNumberOfLearningStages = 4
            ReDim Preserve mTrialsPerLearningStage(4)
            ReDim Preserve CustomPlastMark(4)
            ReDim Preserve CustomPlastFaith(4)
            
            'How many cycles to run.
                Let gNumberOfDataPresentations = Val(txtNumberOfCycles.Text)
            'Trials per stage:
                Dim MyTrialsPerLearningStage As Long
                Let MyTrialsPerLearningStage = Int(Val(txtNumberOfCycles.Text) / 4)
                
            'If the user want presentation of data in exact proportions to their frequencies,
            '   then round up to an exact multiple of the total number of learning data.
                If mnuExactProportions.Checked = True Then
                    Dim MyFactor As Single
                    If mTotalFrequency = -1 Then
                        'Cancel if the count is not integerial.
                            Let mnuExactProportions.Checked = False
                            Let lblExact.Visible = False
                    Else
                        'Proceed.  Round MyTrialsPerLearningStage up to an exact multiple.
                            Let MyFactor = Int(MyTrialsPerLearningStage / mTotalFrequency - 0.0001) + 1
                            Let MyTrialsPerLearningStage = mTotalFrequency * MyFactor
                            Let gNumberOfDataPresentations = 4 * MyTrialsPerLearningStage
                            Let txtNumberOfCycles.Text = Str(gNumberOfDataPresentations)
                    End If
                End If

            'Place total presentations in the dummy variable for reporting purposes.
                Let mReportedNumberOfDataPresentations = gNumberOfDataPresentations
            
            'Plasticity:
                'Initial and final plasticities:
                    Let CustomPlastMark(1) = Val(txtUpperPlasticity)
                    Let CustomPlastMark(4) = Val(txtLowerPlasticity)
                'Interpolate other two geometrically.
                    Let CustomPlastMark(2) = (CustomPlastMark(1) * CustomPlastMark(1) * CustomPlastMark(4)) ^ (1 / 3)
                    Let CustomPlastMark(3) = (CustomPlastMark(1) * CustomPlastMark(4) * CustomPlastMark(4)) ^ (1 / 3)
            'Assign remaining material, which is invariant across the five stages.
                For LearningStageIndex = 1 To 4
                    'All stages have same number of trials.
                        Let mTrialsPerLearningStage(LearningStageIndex) = MyTrialsPerLearningStage
                    'Faithfulness plasticity same as Markedness.
                        Let CustomPlastFaith(LearningStageIndex) = CustomPlastMark(LearningStageIndex)
                Next LearningStageIndex
                
            'Assign the right value to the old variables, permitting them to be remembered.
                Let gCoarsestPlastMark = CustomPlastMark(1)
                Let gFinestPlastMark = CustomPlastMark(4)

            
    End If      'Learning schedule from file or from interface?
    
    'All is well, so return True.
        Let DetermineLearningSchedule = True
        Exit Function
        
BlankCell:
    'Go here if there's a problem with the learning schedule file.
        MsgBox "Sorry, your input file for the learning schedule contains one or more blank cells.  Please fix before continuing."
        Let DetermineLearningSchedule = False
        Exit Function
    
End Function


Sub InstallWinnerAsRival()

    'The file-reading apparatus creates separate arrays for winners and rivals.  Amalgamate these into just rivals, with
    '  the "winners" (which may be actually tied) in row zero.
    
    'Also, amalgamate the WinnerFrequency and RivalFrequency arrays, which were earlier read separately.
    
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
    
End Sub



Sub AdjustAPrioriRankings_Up()

    'The idea is to make every pair of constraints in an a priori relation at least 20 units apart.
    
    'This routine has no danger of running forever, because the a priori rankings have been pre-vetted.
    
    'It is called initially, and also whenever a constraint is raised.
    
    Dim KeepGoing As Boolean    'If you change anything, keep looking to detect further necessary changes.
    Dim Margin As Single        'The difference in weights of two constraints:  enough?
    Dim OuterConstraintIndex As Long, InnerConstraintIndex As Long
    Dim HistoryConstraintIndex As Long
    
    Do
        Let KeepGoing = False
        For OuterConstraintIndex = 1 To mNumberOfConstraints
            For InnerConstraintIndex = 1 To mNumberOfConstraints
                If gAPrioriRankingsTable(OuterConstraintIndex, InnerConstraintIndex) = True Then
                    Let Margin = mWeight(OuterConstraintIndex) - mWeight(InnerConstraintIndex)
                    'Amazingly, the addition of x to y in Visual Basic does not necessarily
                    '   make x + y greater than y by a margin of x.
                    'Let's try to fix this, by saying not x, but ever so slightly less
                    '   than x.  Sheesh.
                    If Margin < mNumericalEquivalentOfStrictRanking - 0.0001 Then
                        Let mWeight(OuterConstraintIndex) = mWeight(InnerConstraintIndex) + _
                            mNumericalEquivalentOfStrictRanking
                        Let KeepGoing = True
                        'Report progress, in the correct column of the spreadsheet.
                        If mnuFullHistory.Checked = True Then
                            'Caption, in first two columns:
                                Print #mFullHistoryFile, "Apriori";
                                Print #mFullHistoryFile, vbTab; mAbbrev(OuterConstraintIndex); " >> "; mAbbrev(InnerConstraintIndex); vbTab;
                            'The change:
                                For HistoryConstraintIndex = 1 To mNumberOfConstraints
                                    If HistoryConstraintIndex = OuterConstraintIndex Then
                                        Print #mFullHistoryFile, vbTab; mNumericalEquivalentOfStrictRanking - Margin;
                                        Print #mFullHistoryFile, vbTab; mWeight(OuterConstraintIndex);
                                    Else
                                        Print #mFullHistoryFile, vbTab; vbTab;
                                    End If
                                Next HistoryConstraintIndex
                            'Finish the line with a carriage return:
                                Print #mFullHistoryFile,
                        End If
                    End If  'Was the margin small enough to justify change?
                End If      'Was there an a priori ranking?
            Next InnerConstraintIndex
        Next OuterConstraintIndex
        If KeepGoing = False Then Exit Do
    Loop
    
    'Stop
    'Debug this routine:
    '    Dim f As Long
    '    Let f = FreeFile
    '    Open gOutputFilePath + "DebugAPrioriRankingValuesFor" + gFileName + ".txt" For Output As f
    '    For OuterConstraintIndex = 1 To mNumberofconstraints
    '        Print #f, mabbrev(OuterConstraintIndex); vbtab;
    '        Print #f, mWeight(OuterConstraintIndex)
    '    Next OuterConstraintIndex
    '    Close f

End Sub

Sub AdjustAPrioriRankings_Down()

    'The idea is to make every pair of constraints in an a priori relation at least 20 units apart.
    
    'This routine is called whenever a constraint is lowered.
    
    Dim Margin As Single
    Dim KeepGoing As Boolean
    Dim OuterConstraintIndex As Long, InnerConstraintIndex As Long
    Dim HistoryConstraintIndex As Long
    
    Do
        Let KeepGoing = False
        For OuterConstraintIndex = 1 To mNumberOfConstraints
            For InnerConstraintIndex = 1 To mNumberOfConstraints
                If gAPrioriRankingsTable(OuterConstraintIndex, InnerConstraintIndex) = True Then
                    Let Margin = mWeight(OuterConstraintIndex) - mWeight(InnerConstraintIndex)
                    If Margin < mNumericalEquivalentOfStrictRanking Then
                        Let mWeight(InnerConstraintIndex) = mWeight(OuterConstraintIndex) - _
                            mNumericalEquivalentOfStrictRanking
                        Let KeepGoing = True
                        'Report progress
                            If mnuFullHistory.Checked = True Then
                                Print #mFullHistoryFile, "Apriori";
                                Print #mFullHistoryFile, vbTab; mAbbrev(OuterConstraintIndex); " >> "; mAbbrev(InnerConstraintIndex); vbTab;
                                For HistoryConstraintIndex = 1 To mNumberOfConstraints
                                    If HistoryConstraintIndex = InnerConstraintIndex Then
                                        Print #mFullHistoryFile, vbTab; ThreeDecPlaces(Margin - mNumericalEquivalentOfStrictRanking);
                                        Print #mFullHistoryFile, vbTab; ThreeDecPlaces(mWeight(InnerConstraintIndex));
                                    Else
                                        Print #mFullHistoryFile, vbTab; vbTab;
                                    End If
                                Next HistoryConstraintIndex
                                Print #mFullHistoryFile,
                            End If
                    End If
                End If
            Next InnerConstraintIndex
        Next OuterConstraintIndex
        If KeepGoing = False Then Exit Do
    Loop
    

End Sub

Sub RecordInitialWeights()

    'If you're reporting progress, record the initial weights.
    
    Dim ConstraintIndex As Long
    
    If mnuFullHistory.Checked = True Then
        Print #mFullHistoryFile, "(Initial)"; vbTab; vbTab; vbTab; vbTab;
        For ConstraintIndex = 1 To mNumberOfConstraints
            Print #mFullHistoryFile, vbTab; mWeight(ConstraintIndex); vbTab;
        Next ConstraintIndex
        Print #mFullHistoryFile,
    End If

End Sub

Sub NoisyHarmonicGrammarCore()

    'Noisy Harmonic Grammar learning, from Boersma and Pater (2008)
   
    'Learning variables:
        Dim SelectedExemplarForm As Long
        Dim SelectedExemplarRival As Long
      
        Dim LocalWinner As Long

        Dim PlastMark As Single
        Dim PlastFaith As Single
        
        Dim MyWinnerViols As Long
        Dim MyTrainingDataViols As Long

    'Indices etc.:
        Dim CycleIndex As Long, FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
        Dim HistoryConstraintIndex As Long
        Dim i As Long, j As Long
      
        Dim DummyCounter As Long     'KZ: every time gets to 200 it checks if the user pressed cancel.
                                    
    'Variables for selecting learning data at random.
        Dim RandomNumber As Single        'KZ
        Dim blnExemplarChosen As Boolean  'KZ
      
    'Variable for reporting what you did in detail.
        Dim Increment As Single             'For remembering plasticity used.
    
    'Variables for time
        Dim TimeSinceLastReport As Single
        Let TimeSinceLastReport = Timer
        
    'Variables for where we are in the run.
        Dim LearningStageIndex As Long
        Dim TotalCycles As Long             'For monitoring progress.
        
    'For exact frequency presentation:
        Dim NumberSelectedFromArray As Long

    'Initialize counters.
        Let DummyCounter = 0        'Interrupts execution for various purposes
                                    '   (progress report, cancellation)
        Let TotalCycles = 0         'Reports total progress to user.
        'Prevent crashes with default gReportingFrequency:
            If gReportingFrequency = 0 Then Let gReportingFrequency = 200
        
    'Get a progress report on the screen with the initial weights.
        'You need to fake some parameters for the first display.
            Let PlastMark = CustomPlastMark(1)
            Let PlastFaith = CustomPlastFaith(1)
        'Now it's safe to report progress.
            Call NoisyHarmonicGrammarReportProgress(0, PlastMark, PlastFaith)

    'Do as many learning stages as there are.
        
        For LearningStageIndex = 1 To mNumberOfLearningStages
        
            'Establish the plasticities that will hold for
            '   this learning stage.  Use local variables that might be
            '    bit faster than looking them up repeatedly in an array.
                Let PlastMark = CustomPlastMark(LearningStageIndex)
                Let PlastFaith = CustomPlastFaith(LearningStageIndex)
                    
            'Go through the cycles of this learning stage.
            For CycleIndex = 1 To mTrialsPerLearningStage(LearningStageIndex)
            
                'mblnProcessing (from KZ) lets you interrupt the GLA by clicking
                '   the Run button--relabeled Cancel during a run.
                    If mblnProcessing = False Then
                        GoTo ExitPoint
                    Else
                
                    'We're not canceled, so execute the current learning cycle.
                        
                    'Select an exemplar for learning with, respecting the frequencies, and generate
                    '    a form stochastically from the same underlying representation.
                        GoTo SelectExemplars
SelectExemplarsReturnPoint:
                        
                    'Find the winner for this input using a stochastically perturbed version of your current grammar.
                        Let LocalWinner = GenerateAForm(SelectedExemplarForm)
                        'MsgBox Str(SelectedExemplarForm) + " " + Str(LocalWinner)

                    'Compare the exemplar with the generated form, and adjust the
                    '   weights of the constraints accordingly.
                    'This comes in two flavors, depending on whether a priori
                    '   rankings are in effect.
                        If mUseAPrioriRankings = False Then
                            GoTo WeightAdjustment
                        Else
                            'Not implemented
                            Stop
                            GoTo WeightAdjustmentWithAprioriRankings
                        End If
WeightAdjustmentReturnPoint:
            
                    'Check if it's time to do tasks that need done every so often.
                        If DummyCounter < gReportingFrequency Then
                            'No.  Just increment the counter.
                                Let DummyCounter = DummyCounter + 1
                        Else
                            'Yes.  Do these important tasks.

                            'Check for cancellation by user.
                                'KZ: checks after every gReportingFrequency cycles.
                                DoEvents    'KZ: passes control back to the operating environment
                                            'to check if there are any pending events--one such
                                            'event would be a second click of cmdRun, which would
                                            'cancel the running of the algorithm.
                            
                            'Update the history file, using (for now) same gReportingFrequency cycle interval.
                                If mnuGenerateHistory.Checked = True Then
                                    For ConstraintIndex = 1 To mNumberOfConstraints
                                       Print #mSimpleHistoryFile, mWeight(ConstraintIndex); Chr$(9);
                                    Next ConstraintIndex
                                    Print #mSimpleHistoryFile,
                                End If
                            
                            'Keep track of total cycles for reporting progress.
                                Let TotalCycles = TotalCycles + gReportingFrequency
                            
                            'Report progress if 2 seconds have elapsed.
                                If Timer - TimeSinceLastReport > 2 Then
                                    Call NoisyHarmonicGrammarReportProgress(TotalCycles, PlastMark, PlastFaith)
                                    Let TimeSinceLastReport = Timer
                                End If

                            'Reset the DummyCounter to start the next interval of gReportingFrequency.
                                Let DummyCounter = 0
                                
                        End If      'Has it been gReportingFrequency cycles?
                End If              'Is mblnProcessing False, so we can keep going?
            Next CycleIndex         'Go on to the next cycle of this learning stage.
        Next LearningStageIndex     'Go on to the next learning stage.

ExitPoint:                          'Go here on failure, so you can check the Timer.

    'Determine how long learning took.
        Let mTimeMarker = Timer - mTimeMarker
      
    'Close the debugging file.
        'Close #DebugFile
        
    'This is the end of the main part of NoisyHarmonicGrammarCore; the rest are subroutines.
   
   Exit Sub

Stop
'-----------------------------------------------------------------------------
SelectExemplars:

    'This is done in two ways:  exact matching from input file, or stochastically.
    
    If mnuExactProportions.Checked = True Then
        
        'Exact method:  pick from a randomized mDataPresentationArray().
                
            'If you've gone through the whole array, start over
                Let NumberSelectedFromArray = NumberSelectedFromArray + 1
                If NumberSelectedFromArray > mTotalFrequency Then
                    Let NumberSelectedFromArray = 1
                End If
                
            'If you're just starting on this array, randomize it.
                If NumberSelectedFromArray = 1 Then
                    Call ShuffleDataPresentationArray
                End If
                
            'Grab the next item off the array.
                Let SelectedExemplarForm = mDataPresentationArray(NumberSelectedFromArray).FormIndex
                Let SelectedExemplarRival = mDataPresentationArray(NumberSelectedFromArray).RivalIndex
            'Keep a record of the learning frequencies
                Let mActualFrequencyShare(mDataPresentationArray(NumberSelectedFromArray).FormIndex, mDataPresentationArray(NumberSelectedFromArray).RivalIndex) = mActualFrequencyShare(mDataPresentationArray(NumberSelectedFromArray).FormIndex, mDataPresentationArray(NumberSelectedFromArray).RivalIndex) + 1
                Let mActualFrequencyPerInput(mDataPresentationArray(NumberSelectedFromArray).FormIndex) = mActualFrequencyPerInput(mDataPresentationArray(NumberSelectedFromArray).FormIndex) + 1

            'Debug:
                'Print #DebugFile, "Form:  "; mInputForm(SelectedExemplarForm); vbtab; "Rival:  "; mRival(SelectedExemplarForm, SelectedExemplarRival)
            
    Else

        'Stochastic method:  Go through the possible outputs, and select one at
        '    random, according to the frequencies.
           
             Let blnExemplarChosen = False
             Let RandomNumber = Rnd()
             For FormIndex = 1 To mNumberOfForms
                'Dec. 2025 I think the 1 is an error -- we have already folded in the winner as a rival.
                 'For RivalIndex = 0 To mNumberOfRivals(FormIndex)
                 For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                     'Don't bother with zero items.
                     If mFrequencyShare(FormIndex, RivalIndex) > 0 Then
                         If RandomNumber >= mFrequencyInterval(FormIndex, RivalIndex, 0) _
                         And RandomNumber < mFrequencyInterval(FormIndex, RivalIndex, 1) Then
                             Let SelectedExemplarForm = FormIndex
                             Let SelectedExemplarRival = RivalIndex
                             Let mActualFrequencyShare(FormIndex, RivalIndex) = mActualFrequencyShare(FormIndex, RivalIndex) + 1
                             Let mActualFrequencyPerInput(FormIndex) = mActualFrequencyPerInput(FormIndex) + 1
                             Let blnExemplarChosen = True
                             Exit For
                         End If
                     End If
                 Next RivalIndex
                 If blnExemplarChosen Then
                     Exit For
                 End If
             Next FormIndex
             
    End If                          'Use exact vs. stochastic method.
    
    GoTo SelectExemplarsReturnPoint


Stop
'------------------------------------------------------------------------------
WeightAdjustment:

   'Compare the "empirical" input with the currently generated form, and adjust the weights accordingly.

   'Do this only if the two are different, and also that there really is a local winner.
      If LocalWinner <> SelectedExemplarRival And LocalWinner <> mcNoWinner Then
        
        'Report progress if requested.  First, list the input and the two crucial candidates.
            If mnuFullHistory.Checked = True Then
                Print #mFullHistoryFile, mInputForm(SelectedExemplarForm); vbTab; mRival(SelectedExemplarForm, LocalWinner); vbTab; mRival(SelectedExemplarForm, SelectedExemplarRival);
            End If
            
        'Compute the adjustments.
            For ConstraintIndex = 1 To mNumberOfConstraints
                'Load into local variable for legibility:
                    Let MyWinnerViols = mRivalViolations(SelectedExemplarForm, LocalWinner, ConstraintIndex)
                    Let MyTrainingDataViols = mRivalViolations(SelectedExemplarForm, SelectedExemplarRival, ConstraintIndex)
                'First see if any change will be needed.
                    If MyWinnerViols = MyTrainingDataViols Then
                        'No change for this constraint.  Report thus.
                            If mnuFullHistory.Checked = True Then
                                Print #mFullHistoryFile, vbTab; vbTab;
                            End If
                    Else
                            'The update rule for noisy harmonic grammar:  subtract the LocalWinner's violations from the SelectedExamplarRival's,
                            '   multiply by the plasticity, and add to the weight.
                            '   KZ: plasticity depends on whether it's a faithfulness constraint or a markedness constraint:
                                If mFaithfulness(ConstraintIndex) = True Then
                                    Let mWeight(ConstraintIndex) = mWeight(ConstraintIndex) + PlastFaith * (MyWinnerViols - MyTrainingDataViols)
                                    Let mNumberOfChancesForWeightsToGoBelowZero = mNumberOfChancesForWeightsToGoBelowZero + 1
                                    'Don't let a weight go below zero, unless user checked the box allowing this.
                                        If mWeight(ConstraintIndex) < 0 Then
                                            'Report the number of times this happened, mostly for debugging purposes.
                                                Let mNumberOfTimesWeightsWentBelowZero = mNumberOfTimesWeightsWentBelowZero + 1
                                            If chkNegativeWeightsOK.Value <> vbChecked Then
                                                'We're ok with negative weights in Exponential NHG.
                                                    If gExponentialNHG = False Then
                                                        Let mWeight(ConstraintIndex) = 0
                                                    End If
                                            End If
                                        End If
                                        
                                    'Report progress if requested.
                                         If mnuFullHistory.Checked = True Then
                                            Print #mFullHistoryFile, vbTab; PlastFaith; vbTab; mWeight(ConstraintIndex);
                                         End If
                                Else
                                    Let mWeight(ConstraintIndex) = mWeight(ConstraintIndex) + PlastMark * (MyWinnerViols - MyTrainingDataViols)
                                    Let mNumberOfChancesForWeightsToGoBelowZero = mNumberOfChancesForWeightsToGoBelowZero + 1
                                    'Repeat code used to deal with negative weights.
                                        If mWeight(ConstraintIndex) < 0 Then
                                            Let mNumberOfTimesWeightsWentBelowZero = mNumberOfTimesWeightsWentBelowZero + 1
                                            If chkNegativeWeightsOK.Value <> vbChecked Then
                                                If gExponentialNHG = False Then
                                                    Let mWeight(ConstraintIndex) = 0
                                                End If
                                            End If
                                        End If
                                    'Report progress if requested.
                                         If mnuFullHistory.Checked = True Then
                                            Print #mFullHistoryFile, vbTab; PlastMark; vbTab; mWeight(ConstraintIndex);
                                         End If
                                End If
                    End If
            Next ConstraintIndex
         
         'If you're reporting your actions, generate a carriage return to complete the line.
            If mnuFullHistory.Checked = True Then
                Print #mFullHistoryFile,
            End If

      End If        'Are local winner and empirical form the same?

    GoTo WeightAdjustmentReturnPoint
    
Stop
'---------------------------------------------------------------------------------------------
WeightAdjustmentWithAprioriRankings:

   'Compare the "empirical" input with the currently generated form, and
   '    adjust weights accordingly.  While you do this, adjust
   '    any other weights that need to be in order to keep the
   '    a priori rankings duly respected.

    'xxx not done
    Stop

   'Do this only if the generated form and observed form are different.
      If LocalWinner <> SelectedExemplarRival Then
        
            If mnuFullHistory.Checked = True Then
                Print #mFullHistoryFile, mRival(SelectedExemplarForm, LocalWinner); vbTab; mRival(SelectedExemplarForm, SelectedExemplarRival);
            End If
                
        'Examine all constraint violations.
        
         For ConstraintIndex = 1 To mNumberOfConstraints
            
            Select Case mRivalViolations(SelectedExemplarForm, LocalWinner, ConstraintIndex) _
               - mRivalViolations(SelectedExemplarForm, SelectedExemplarRival, ConstraintIndex)
               
               Case Is > 0

                    'The (wrong) LocalWinner violates more.  To improve the grammar,
                    '  *strengthen* this constraint, to punish the wrong local winner.
                      
                    'KZ: plasticity adjustment depends on whether it's a faithfulness
                    'constraint or a markedness constraint:
                        If mFaithfulness(ConstraintIndex) = True Then
                            Let mWeight(ConstraintIndex) = mWeight(ConstraintIndex) + PlastFaith
                            'Remember the change, for reporting progress:
                                Let Increment = PlastFaith
                        Else    'It's a markedness constraint that has to be adjusted.
                            Let mWeight(ConstraintIndex) = mWeight(ConstraintIndex) + PlastMark
                            Let Increment = PlastMark
                        End If
                   
                    'Report progress if requested.
                        'Since the a priori rankings are interleaved, I've adjusted this
                        '   code so it reports just one change per line.
                         If mnuFullHistory.Checked = True Then
                            'Caption:  the generated and heard forms, plus the number of what's being added:
                                Print #mFullHistoryFile, mRival(SelectedExemplarForm, LocalWinner); vbTab; mRival(SelectedExemplarForm, SelectedExemplarRival);
                            'Find the column in which the change should be recorded:
                            For HistoryConstraintIndex = 1 To mNumberOfConstraints
                                If HistoryConstraintIndex = ConstraintIndex Then
                                    'The increment, and the new value.
                                        Print #mFullHistoryFile, vbTab; Increment; vbTab; mWeight(ConstraintIndex);
                                Else
                                    'Column gaps:
                                        Print #mFullHistoryFile, vbTab; vbTab;
                                End If
                            Next HistoryConstraintIndex
                            'Carriage return, since line is now done:
                                Print #mFullHistoryFile,
                         End If         'Does user want progress reported?
                        
                    'Enforce the apriori rankings as needed:
                        Call AdjustAPrioriRankings_Up
                        
               Case Is < 0

                    'The (wrong) LocalWinner violates less.  To improve the grammar,
                    '  *weaken* this constraint, so it will not punish the correct form
                    '  so much.
    
                        If mFaithfulness(ConstraintIndex) = True Then
                            Let mWeight(ConstraintIndex) = mWeight(ConstraintIndex) - PlastFaith
                            Let Increment = -1 * PlastFaith
                        Else
                            Let mWeight(ConstraintIndex) = mWeight(ConstraintIndex) - PlastMark
                            Let Increment = -1 * PlastMark
                        End If
                        
                    'Report progress if requested.  Same code as above.
                        If mnuFullHistory.Checked = True Then
                           Print #mFullHistoryFile, mRival(FormIndex, LocalWinner); vbTab; mRival(FormIndex, SelectedExemplarRival);
                           For HistoryConstraintIndex = 1 To mNumberOfConstraints
                               If HistoryConstraintIndex = ConstraintIndex Then
                                   Print #mFullHistoryFile, vbTab; Increment; vbTab; mWeight(ConstraintIndex);
                               Else
                                   Print #mFullHistoryFile, vbTab; vbTab;
                               End If
                           Next HistoryConstraintIndex
                           Print #mFullHistoryFile,
                        End If
                        
                    'Enforce the apriori rankings as needed:
                        Call AdjustAPrioriRankings_Down
                    
            End Select              'Which had more violations, winner or rival?
         Next ConstraintIndex       'Examine all constraints.
      End If                        'Are the generated form and observed form different?

    GoTo WeightAdjustmentReturnPoint
    

End Sub

Function GenerateAForm(MyFormIndex As Long) As Long

   'Go through the candidates, keeping a local best, until you have a winner.
   'This is complicated, since it implements eight different flavors of NHG, labeled A-G.
   'This code is used both for GLA-type learning and for testing the final grammar.

    'Learning variables:
        'An array of perturbed weights:
            Dim LocalWeight() As Single
            ReDim LocalWeight(mNumberOfConstraints)
        'For perturbation after multiplication by violation count, we need an array of perturbations:
            Dim Perturbation() As Single
            ReDim Perturbation(mNumberOfConstraints)
            'After we multiply by violations and add the noise:
                Dim ConstraintSpecificHarmonyContribution As Single
        Dim MyGaussian As Single
        
    'Track winners.
        Dim BestHarmony As Single
        Dim MyHarmony As Single
        Dim LocalWinner As Long
        
    'Variables for handling ties.
        Dim TiedCandidate() As Boolean      'yes or no; are you in a tie for first?
        ReDim TiedCandidate(mNumberOfRivals(MyFormIndex))
        Dim NumberOfTiedCandidates As Long
        Dim TiesOnThisIteration As Long     'for overall counting/monitoring
        Dim TieWinner As Long               'to pick a winner at random from ties
        'Keep the program from crashing on permanent ties, which can happen when it uses the method of trial cancellation.
            Dim ConsecutiveTies As Long
        
    'For printout, description of what done.
        Dim NarrativeReport
        
        Dim RivalIndex As Long, InnerRivalIndex As Long, ConstraintIndex As Long, CandidateIndex As Long
    
RestartPoint:
    
    
    'If needed, calculate the perturbed weights or perturbation that are used consistently across all candidates for this input.
        'This is not done if perturbation takes place on a cell-by-cell basis:
         If chkNoiseAppliesToTableauCells.Value <> vbChecked Then
            'For post-multiplication noise, we need to record the perturbation value itself, so it doesn't get multiplied.
            If chkNoiseIsAddedAfterMultiplication.Value = vbChecked Then
                For ConstraintIndex = 1 To mNumberOfConstraints
                    Let Perturbation(ConstraintIndex) = Gaussian
                Next ConstraintIndex
            Else
            'For premultiplication noise, we can just form a set of local perturbed weights.
                For ConstraintIndex = 1 To mNumberOfConstraints
                    'We should use a small noise value for Exponential NHG, as Boersma and Pater point out.  The variable mNoise has already
                    '   been set and handles this.
                         Let LocalWeight(ConstraintIndex) = mWeight(ConstraintIndex) + (mNoise * Gaussian)
                    Let mNumberOfChancesForWeightsToGoBelowZero = mNumberOfChancesForWeightsToGoBelowZero + 1
                    'Code for dealing with negative weights.
                        If LocalWeight(ConstraintIndex) < 0 Then
                            Let mNumberOfTimesWeightsWentBelowZero = mNumberOfTimesWeightsWentBelowZero + 1
                            If chkNegativeWeightsOK.Value <> vbChecked Then
                                'We do allow negative literal weights in Exponential NHG.
                                If gExponentialNHG = False Then
                                    Let LocalWeight(ConstraintIndex) = 0
                                End If
                            End If
                        End If
                Next ConstraintIndex
            End If
         End If
    
    'Start with the first candidate (0) as default winner.
        Let LocalWinner = 0
    'Set an impossibly lax initial criterion for best harmony (since lower is better).
        Let BestHarmony = 1000000000
    'For the selected exemplar form, ponder each rival in turn.
                'I think the 0 is an error -- we have already folded in the winner as a rival.
                 'For RivalIndex = 0 To mNumberOfRivals(FormIndex)
        For RivalIndex = 1 To mNumberOfRivals(MyFormIndex)
            'Compute the harmony of this candidate (dot product of violations time weights), with various options for perturbation.
                'Initialize the sum that will accumulate across constraints.
                    Let MyHarmony = 0
                'If you are perturbing the weights at the very end, skip ahead.
                    If chkLateNoise.Value <> vbChecked Then
                        'Go through all the constraints.
                            For ConstraintIndex = 1 To mNumberOfConstraints
                                'There are four options, in a two-by-two array.
                                    'First, the weights are perturbed prior to multiplication by violation count.
                                        If chkNoiseIsAddedAfterMultiplication.Value <> vbChecked Then
                                            If chkNoiseAppliesToTableauCells <> vbChecked Then
                                                If gExponentialNHG = True Then
                                                    'A-prime:  weights are perturbed on a constraint-by-constraint basis, but with an exponential function.
                                                    '   This is Boersma and Pater's "exponential NHG".
                                                        Let MyHarmony = MyHarmony + Exp(LocalWeight(ConstraintIndex)) * mRivalViolations(MyFormIndex, RivalIndex, ConstraintIndex)
                                                        'There is an issue here:  are negative harmonies anything to worry about?  Aaron says no.
                                                            'If MyHarmony < 0 Then Stop
                                                Else
                                                    'A:  This is classical NHG.
                                                    '   Weights are perturbed on a constraint-by-constraint basis.
                                                    '   Use the LocalWeights() calculated once-and-for-all above for this particular exemplar input form.
                                                    '   Note:  the possibility of negative weights (recording them, fixing them) is already dealt with above.
                                                        Let MyHarmony = MyHarmony + LocalWeight(ConstraintIndex) * mRivalViolations(MyFormIndex, RivalIndex, ConstraintIndex)
                                                End If
                                                    
                                            Else
                                                'B:  weights are re-perturbed each time you evaluate a tableau cell.
                                                    'So find a perturbed weight, a local one.
                                                        Let MyGaussian = Gaussian
                                                        Let LocalWeight(ConstraintIndex) = mWeight(ConstraintIndex) + MyGaussian
                                                    'Deal with negative weights.
                                                        'Note first that if there are zero violations, multiplication by zero will remove the problem by itself.
                                                            If Not mRivalViolations(MyFormIndex, RivalIndex, ConstraintIndex) = 0 Then
                                                                'So we do have to deal with it:
                                                                    Let mNumberOfChancesForWeightsToGoBelowZero = mNumberOfChancesForWeightsToGoBelowZero + 1
                                                                    If LocalWeight(ConstraintIndex) < 0 Then
                                                                        Let mNumberOfTimesWeightsWentBelowZero = mNumberOfTimesWeightsWentBelowZero + 1
                                                                        If chkNegativeWeightsOK.Value <> vbChecked Then
                                                                            Let LocalWeight(ConstraintIndex) = 0
                                                                        End If
                                                                    End If
                                                            End If
                                                    'Total up harmony, in whatever way matches user's choice for exponential.
                                                        If gExponentialNHG = True Then
                                                            'B-prime:  the exponential version.
                                                            Let MyHarmony = MyHarmony + Exp(LocalWeight(ConstraintIndex)) * mRivalViolations(MyFormIndex, RivalIndex, ConstraintIndex)
                                                        Else
                                                            'B per se:  the nonexponential version.
                                                            Let MyHarmony = MyHarmony + LocalWeight(ConstraintIndex) * mRivalViolations(MyFormIndex, RivalIndex, ConstraintIndex)
                                                        End If
                                            End If
                                        Else
                                    'Second, weights are perturbed after multiplication by violation count.
                                            If chkNoiseAppliesToTableauCells.Value <> vbChecked Then
                                                'C-D:  weights are perturbed on a constraint-by-constraint basis, so use the Perturbations() calculated earlier, in order
                                                '   to have the same perturbation for each candidate.
                                                    'Decide what you want to do about a cell with no violations -- either cell has violations, or user wants noise anyway.
                                                        If mRivalViolations(MyFormIndex, RivalIndex, ConstraintIndex) = 0 Then
                                                            'Suppose the user wants no noise for zero-violation cells.
                                                                If chkNoiseForZeroCells.Value <> vbChecked Then
                                                                    Let ConstraintSpecificHarmonyContribution = 0                'This is D.
                                                                Else
                                                                    'Here, the user *does* want noise for zero-violation cells.
                                                                    'Total up harmony, in whatever way matches user's choice for exponential.
                                                                        If gExponentialNHG = True Then
                                                                            Let ConstraintSpecificHarmonyContribution = Exp((mWeight(ConstraintIndex) * mRivalViolations(MyFormIndex, RivalIndex, ConstraintIndex)) _
                                                                                + Perturbation(ConstraintIndex))    'This is C, in exponential version.
                                                                        Else
                                                                            'With no violations, the harmony contribution simply *is* the noise.
                                                                                Let ConstraintSpecificHarmonyContribution = Perturbation(ConstraintIndex)     'This is C, in its nonexponential version.
                                                                        End If
                                                                End If
                                                        Else
                                                            'The number of violations in the cell is non-zero (perhaps negative).
                                                            'Apply the normal procedures for this kind of NHG.
                                                                If gExponentialNHG = True Then
                                                                    Let ConstraintSpecificHarmonyContribution = Exp((mWeight(ConstraintIndex) * mRivalViolations(MyFormIndex, RivalIndex, ConstraintIndex)) _
                                                                        + Perturbation(ConstraintIndex))
                                                                Else
                                                                    Let ConstraintSpecificHarmonyContribution = (mWeight(ConstraintIndex) * mRivalViolations(MyFormIndex, RivalIndex, ConstraintIndex)) _
                                                                        + Perturbation(ConstraintIndex)
                                                                End If
                                                        End If
                                                    
                                                    'Checking to make sure we avoid a "negative-weight" problem.
                                                        'This would be bad for harmonic bounding, for negative-weighted constraints become rewards.
                                                        'Monitor the situation for final report.
                                                            Let mNumberOfChancesForWeightsToGoBelowZero = mNumberOfChancesForWeightsToGoBelowZero + 1
                                                        'See notes for why we think this is the true way to keep weights from going negative.
                                                            'None of this matters if there are zero violations, and we need to avoid dividing by zero.
                                                            'No:  we suspect this zero-violation part is problematic, perhaps even allowing harmonically bounded
                                                            '   candidates to win.
                                                                If mRivalViolations(MyFormIndex, RivalIndex, ConstraintIndex) <> 0 Then
                                                                    'The case of non-zero violations.
                                                                    'Observe that that following expression, when multiplied by the number of violations,
                                                                    '   (which, N.B., might be negative!), yields the ConstraintSpecificHarmony contribution.
                                                                    '   It is in effect, the "weight" being used here; if it is below zero, then we have a reward, and
                                                                    '      lose harmonic bounding.  Kaplan's experience running this routine confirms this.
                                                                        If mWeight(ConstraintIndex) + (Perturbation(ConstraintIndex) / mRivalViolations(MyFormIndex, RivalIndex, ConstraintIndex)) < 0 Then
                                                                            'Monitor this:
                                                                                Let mNumberOfTimesWeightsWentBelowZero = mNumberOfTimesWeightsWentBelowZero + 1
                                                                            'Restore nonnegativity:
                                                                                If chkNegativeWeightsOK.Value <> vbChecked Then
                                                                                    Let ConstraintSpecificHarmonyContribution = 0
                                                                                End If
                                                                        End If
                                                                Else
                                                                    'The case of zero violations.  If the user permitted noise in a violation-free cell,
                                                                    '   that noise could take us below zero.
                                                                        If chkNoiseForZeroCells.Value = vbChecked Then
                                                                            If chkNegativeWeightsOK.Value <> vbChecked Then
                                                                                Let ConstraintSpecificHarmonyContribution = 0
                                                                            End If
                                                                        End If
                                                                            

                                                                End If
                                                        
                                                    'Use this ConstraintSpecificHarmonyContribution to compute the harmony of this candidate.
                                                        Let MyHarmony = MyHarmony + ConstraintSpecificHarmonyContribution
                                            
                                            Else
                                                'E-F:  weights are perturbed on a cell-by-cell basis. Here, you can just add the random factor in directly.
                                                    'Decide what you want to do about a cell with no violations -- either cell has violations, or user wants noise anyway.
                                                        'Kaplan-notice:  this was set at "> 0", not "<> 0", which is needed for your case.
                                                        If mRivalViolations(MyFormIndex, RivalIndex, ConstraintIndex) <> 0 Or chkNoiseForZeroCells.Value = vbChecked Then
                                                            If gExponentialNHG = True Then
                                                                Let ConstraintSpecificHarmonyContribution = Exp((mWeight(ConstraintIndex) * mRivalViolations(MyFormIndex, RivalIndex, ConstraintIndex)) _
                                                                + Gaussian)                                  'This is E, exponential version.
                                                            Else
                                                                Let ConstraintSpecificHarmonyContribution = (mWeight(ConstraintIndex) * mRivalViolations(MyFormIndex, RivalIndex, ConstraintIndex)) _
                                                                + Gaussian                                   'This is E.
                                                            End If
                                                        Else
                                                            'This represents non-assignment of perturbation to violation-less cells, when that is what the user wants.
                                                                Let ConstraintSpecificHarmonyContribution = 0              'This is F.
                                                        End If
                                                    
                                                    'Don't let a harmony contribution go below zero, unless user checked the box allowing this.
                                                        Let mNumberOfChancesForWeightsToGoBelowZero = mNumberOfChancesForWeightsToGoBelowZero + 1
                                                        
                                                        
                                                        'We suspect this is wrong, and are doing it as above.
                                                        'If ConstraintSpecificHarmonyContribution < 0 Then
                                                        '    'Monitor this:
                                                        '        Let mNumberOfTimesWeightsWentBelowZero = mNumberOfTimesWeightsWentBelowZero + 1
                                                        '    'Restore nonnegativity:
                                                        '        If chkNegativeWeightsOK.Value <> vbChecked Then
                                                        '            Let ConstraintSpecificHarmonyContribution = 0
                                                        '        End If
                                                        'Else
                                                        '    'MsgBox "Final harmony contribution is:  " + Str(ConstraintSpecificHarmonyContribution)
                                                        'End If
                                                        
                                                        'Repeating the same procedure we used for D-E
                                                            If mRivalViolations(MyFormIndex, RivalIndex, ConstraintIndex) <> 0 Then
                                                                'Here is our guess about what avoids a "reward-not-constraint", which is the bad thing.
                                                                    If mWeight(ConstraintIndex) + (Perturbation(ConstraintIndex) / mRivalViolations(MyFormIndex, RivalIndex, ConstraintIndex)) < 0 Then
                                                                        'Monitor this:
                                                                            Let mNumberOfTimesWeightsWentBelowZero = mNumberOfTimesWeightsWentBelowZero + 1
                                                                        'Restore nonnegativity:
                                                                            If chkNegativeWeightsOK.Value <> vbChecked Then
                                                                                Let ConstraintSpecificHarmonyContribution = 0
                                                                            End If
                                                                    End If
                                                            End If
                                                        
                                                    'Use this ConstraintSpecificHarmonyContribution to compute the harmony of this candidate.
                                                        Let MyHarmony = MyHarmony + ConstraintSpecificHarmonyContribution

                                            End If              'Constraint-granularity or cell-granularity?
                                        
                                        End If                  'Are we perturbing before or after multiplication by violation count?
                            
                            Next ConstraintIndex                'Accumulate harmony across all constraints.
            
                    Else
                        'G. Lastly, there is the possibility of perturbing the weights at the very end.  We'll do this with a separate loop.
                            For ConstraintIndex = 1 To mNumberOfConstraints
                                If gExponentialNHG = True Then
                                    'This is G, exponential version.
                                        Let MyHarmony = MyHarmony + Exp(mWeight(ConstraintIndex) * mRivalViolations(MyFormIndex, RivalIndex, ConstraintIndex) + Gaussian)
                                Else
                                    'This is G, linear version.
                                        Let MyHarmony = MyHarmony + mWeight(ConstraintIndex) * mRivalViolations(MyFormIndex, RivalIndex, ConstraintIndex) + Gaussian
                                End If
                            Next ConstraintIndex
                    End If
                        
            'If it is the best harmony, record it as such and adjust the criterion.
                'For some reason I was playing with treating small margins as ties, but now I think we can just go with plain ties.
                'Monitor the potential of ties, first with a denominator. This is just for watching.
                    Let mNumberOfChancesForTie = mNumberOfChancesForTie + 1
                'Check if you have a victory or a tie this time.
                    'Let mTieCriterion = 0.00001, 'Let mTieCriterion = 1E-18
                    'If BestHarmony - MyHarmony > mTieCriterion Then
                    If BestHarmony > MyHarmony Then
                        'This is outright victory:
                            Let LocalWinner = RivalIndex
                            Let BestHarmony = MyHarmony
                            'No longer a tie if it was before; reset variables.
                                Let TiesOnThisIteration = 0
                                Let NumberOfTiedCandidates = 1  'I.e. total number of winners you must search if tied.  It's only a tie if this ends up 2 or more.
                                'Cancel the status as tied of anyone who was tied before.
                                    'Dec. 2025 I think the 1 is an error -- we have already folded in the winner as a rival.
                                    'For InnerRivalIndex = 0 To mNumberOfRivals(MyFormIndex)
                                    For InnerRivalIndex = 1 To mNumberOfRivals(MyFormIndex)
                                        Let TiedCandidate(InnerRivalIndex) = False
                                    Next InnerRivalIndex
                                'This *will* be a tie if anyone else later matches it.
                                    Let TiedCandidate(RivalIndex) = True
                                
                    'ElseIf BestHarmony - MyHarmony > -1 * mTieCriterion Then
                    ElseIf BestHarmony = MyHarmony Then
                        'This is a tie; record info.
                            'This diagnoses a real mess if you are doing Exponential NHG.
                                If gExponentialNHG Then
                                    'Only issue this message box warning once.
                                        If mTieWarning = False Then
                                            MsgBox "Caution:  a candidate has just yield two tied winners. A possible cause of this in Exponential HNG is that a weight has gone so low that its exponentiated version is represented as zero by the programming language; though there can be other causes. In any event, the result of this run should not be considered reliable."
                                            Let mTieWarning = True
                                        End If
                                End If
                            Let TiesOnThisIteration = TiesOnThisIteration + 1
                            Let NumberOfTiedCandidates = NumberOfTiedCandidates + 1
                            Let TiedCandidate(RivalIndex) = True
                    Else
                        'Do nothing; this is a loss.
                    End If
          
          Next RivalIndex       'Consider all the rivals for this input.
          
    'If, after considering all candidates, the process generated a tie, pick a method to follow.
        'This is a real possibility, since non-negative cutoff can give rise to zero-zero ties.
        If NumberOfTiedCandidates > 1 Then
            'Keep track of ties globally.
                Let mNumberOfTies = mNumberOfTies + TiesOnThisIteration
            'If user wishes, resolve the tie by skipping this trial.
                If gResolveTiesBySkipping Then
                    'We don't want to do this forever!
                        Let ConsecutiveTies = ConsecutiveTies + 1
                        If ConsecutiveTies = 100 Then
                            'Warn the user the first time this happens.
                                If mInfinityWarning = False Then
                                    MsgBox "I'm stuck in an endless loop.  In trying to find a winner for " + mInputForm(MyFormIndex) + " I have run the grammar 100 consecutive times, obtaining a tie each time. I cannot tell you what the problem is, but perhaps it resides in the constraint set. I am terminating the endless loop now, but the result the program obtains should not be considered reliable.  Please click Ok to continue."
                                End If
                            'Record the NoWinner result with an absurd winner index (minus billion) and get out.
                                Let GenerateAForm = mcNoWinner
                                Let mInfinityWarning = True
                                Exit Function
                        End If
                    GoTo RestartPoint
                End If
            'The other option is to select a winner randomly. This procedure has emerged (Kaplan) as quite problematic; it destroys harmonic bounding.
                'This rolls the dice, setting a criterion.
                    Let TieWinner = Int(Rnd() * NumberOfTiedCandidates)
                'Initialize the counter that will find the TieWinner'th member of the list of those tied for first.
                    Let NumberOfTiedCandidates = -1
                For CandidateIndex = 0 To mNumberOfRivals(MyFormIndex)
                    If TiedCandidate(CandidateIndex) = True Then
                        Let NumberOfTiedCandidates = NumberOfTiedCandidates + 1
                        If NumberOfTiedCandidates = TieWinner Then
                            Let LocalWinner = CandidateIndex
                            Exit For
                        End If
                    End If
                Next CandidateIndex
                
        Else
            'There is an unambiguous winner.  The LocalWinner you already have in place will suffice.  But let the
            '   program off the hook on consecutive ties.
                Let ConsecutiveTies = 0
        End If
                
    'Return the value of the winner you selected
        Let GenerateAForm = LocalWinner
        
        'MsgBox "local winner is " + Str(LocalWinner)


End Function

Function Gaussian() As Single

    'This algorithm for producing random values from the Gaussian
    '   distribution give you two values at once.  The static variables
    '   below remember the second value, and that you have one available.
    
    'N.B. In the GLA code, I defer to Paul Boersma's choice of a standard deviation of 2 for the noise.
    'Here, in the interest of explaining my own work, I use noise with standard deviation of one.
    'So this routine is microscopically different from the one appearing in the Stochastic OT (GLA) code.

        Dim fac As Single, r As Single, v1 As Single, v2 As Single
        Static blnValuedAlreadyStored As Boolean
        Static StoredValue As Single
        
        If blnValuedAlreadyStored = False Then
            'Basic calculations:
                Do
                    'Call random numbers (here:  the Visual Basic utility; perhaps
                    '  better to use Paul's code), and convert to a -1 to 1 range.
                        Let v1 = 2 * Rnd - 1
                        Let v2 = 2 * Rnd - 1
                        Let r = v1 * v1 + v2 * v2
                    'The following condition guarantees that the output will have
                    '  summed squares less than one.  This happens about 70% of the
                    '  time, so you're pretty much guaranteed to get out fairly soon.
                        If r < 1 And r > 0 Then Exit Do
                Loop
            'The following yields one normal deviate.  Save it for next time.
                Let fac = Sqr(-2 * Log(r) / r)
                'Boersmian original:
                    'Let StoredValue = 2 * v1 * fac
                'Now:
                    Let StoredValue = v1 * fac
            'This computes the other normal deviate, which can be used fresh.
                'Boersmian original:
                '    Let Gaussian = 2 * v2 * fac
                'Now:
                    Let Gaussian = v2 * fac
            'Use the flag to note that you don't have to compute a new one next time.
                Let blnValuedAlreadyStored = True
        Else
            'Return the thriftily stored value.
                Let Gaussian = StoredValue
            'Indicate with the flag that you've used it up and must compute anew next time.
                Let blnValuedAlreadyStored = False
        End If        'Should I compute a new value?
        
        'In demigaussian NNG, all perturbations of weight are positive.
            If gDemiGaussianNHG Then
                If Gaussian < 0 Then
                    Let Gaussian = -1 * Gaussian
                End If
            End If


End Function


Sub ShuffleDataPresentationArray()

    'The DataPresentationArray() must be randomized, to achieve equal numbers
    '   but pseudo-random presentation order.
    
        Dim SwappantFormIndex As Long, SwappantRivalIndex As Long, LearningDatumIndex As Long
        Dim SwappeeIndex As Long
    
    'Loop through each member of the array, and swap it with a random other member.
    
        For LearningDatumIndex = 1 To mTotalFrequency
            'Pick a swappee.
                Let SwappeeIndex = Int(Rnd() * mTotalFrequency) + 1
            'Swap.
                Let SwappantFormIndex = mDataPresentationArray(LearningDatumIndex).FormIndex
                Let SwappantRivalIndex = mDataPresentationArray(LearningDatumIndex).RivalIndex
                Let mDataPresentationArray(LearningDatumIndex).FormIndex = mDataPresentationArray(SwappeeIndex).FormIndex
                Let mDataPresentationArray(LearningDatumIndex).RivalIndex = mDataPresentationArray(SwappeeIndex).RivalIndex
                Let mDataPresentationArray(SwappeeIndex).FormIndex = SwappantFormIndex
                Let mDataPresentationArray(SwappeeIndex).RivalIndex = SwappantRivalIndex
        Next LearningDatumIndex
    
    Exit Sub
    
    'Debug:
    '    Dim DebugFile As Long
    '    Let DebugFile = FreeFile
    '    Open gInputFilePath + "\debug.txt" For Append As #DebugFile
    '    Dim FrequencyIndex As Long
    '    For FrequencyIndex = 1 To mTotalFrequency
    '        Print #DebugFile, FrequencyIndex; vbTab; mDataPresentationArray(FrequencyIndex).FormIndex;
    '        Print #DebugFile, vbTab; mDataPresentationArray(FrequencyIndex).RivalIndex
    '    Next FrequencyIndex
    '    Close #DebugFile

End Sub


Sub NoisyHarmonicGrammarReportProgress(CycleIndex&, PlastMark As Single, PlastFaith As Single)

   'Print out the results so far.

      Dim SlotFiller() As Long
      Dim LocalWeight() As Single
      ReDim SlotFiller(mNumberOfConstraints)
      ReDim LocalWeight(mNumberOfConstraints)
      
      Dim ConstraintIndex As Long
      Dim i As Long, j As Long
         
   'Print the results so far.

        pctProgressWindow.Cls
        
        'Cycle number:
            pctProgressWindow.Print "Completed learning cycle #"; CycleIndex&; "/"; mReportedNumberOfDataPresentations - 1
        
        'Plasticity:
            If mblnUseCustomLearningSchedule = True Then
                pctProgressWindow.Print
                pctProgressWindow.Print "Plasticity for faithfulness is currently:  ";
                pctProgressWindow.Print ThreeDecPlaces(PlastFaith)
                pctProgressWindow.Print "Plasticity for markedness is currently:    ";
                pctProgressWindow.Print ThreeDecPlaces(PlastMark)
            Else
                pctProgressWindow.Print "Plasticity is currently ";
                'Either value will do, when they are locked together--BH.
                pctProgressWindow.Print ThreeDecPlaces(PlastFaith); "."
            End If
            
        'weights:
            pctProgressWindow.Print "Current weights:"
            For ConstraintIndex = 1 To mNumberOfConstraints
                pctProgressWindow.Print FillStringTo(ThreeDecPlaces(mWeight(ConstraintIndex)), 11);
                'pctProgressWindow.Print ThreeDecPlaces(mWeight(ConstraintIndex));
                pctProgressWindow.Print "   "; mConstraintName(ConstraintIndex)
            Next ConstraintIndex

End Sub


Function NHGTestGrammar(Weight() As Single, ShouldIPrint As Boolean, CyclesToTest As Long) As Single
   
   'Variables to keep track of how many of each type are getting generated.
        Dim NumberGenerated() As Long
        ReDim mPercentageGenerated(mNumberOfForms, mMaximumNumberOfRivals)
        Dim PercentageInInput() As Single
        ReDim PercentageInInput(mNumberOfForms, mMaximumNumberOfRivals)
        Dim LocalSum As Single
        Dim LocalWinner As Long

   'Indices
        Dim TrialIndex As Long
        Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
        Dim ProgressCounter As Long
        Dim i As Long, j As Long
       
    'Likelihood, to report
        Dim CandidateLikelihood As Single
        Dim DummyCounter As Integer 'KZ: keeps track of how many loops it's been
                                    'since checked to see if user wants to cancel.
       
    'For log likelihood
        Dim LogLikelihoodPackage As Module1.gLikelihoodCalculation
       
        Dim ErrorSoFar As Single
      
      ReDim NumberGenerated(mNumberOfForms, mMaximumNumberOfRivals)

        If ShouldIPrint = True Then
            pctProgressWindow.Cls
            pctProgressWindow.Print "Testing the grammar..."
        End If

   'Just in case you are doing batch processing, initialize the count of
   '  forms generated.
      For FormIndex = 1 To mNumberOfForms
      
            'I think the 0 is an error -- we have already folded in the winner as a rival.
                 'For RivalIndex = 0 To mNumberOfRivals(FormIndex)
         'For RivalIndex = 0 To mNumberOfRivals(FormIndex)
         For RivalIndex = 1 To mNumberOfRivals(FormIndex)
            Let NumberGenerated(FormIndex, RivalIndex) = 0
         Next RivalIndex
      Next FormIndex

'------------------------------------------------------------------------------

   'Loop through this many trials, for all the forms.

        'Keep track of how far you've gotten.
            Let ProgressCounter = 500
            Let DummyCounter = 0  'KZ
            'KZ: if user has pressed cancel, stop.
                If mblnProcessing = False Then Exit Function

     'Main loop, through trials.
          For TrialIndex = 1 To CyclesToTest
             'Report progress, during the big run:
                 If ShouldIPrint = True Then
                    If TrialIndex > ProgressCounter Then
                        pctProgressWindow.Cls
                        pctProgressWindow.Print "Completed"; TrialIndex; "test trials /"; CyclesToTest
                        Let ProgressCounter = ProgressCounter + 500
                    End If
                End If
            'Using this grammar, generate an output for each form and keep count.
                For FormIndex = 1 To mNumberOfForms
                    Let LocalWinner = GenerateAForm(FormIndex)
                    'Now that you know the winner for this combination of constraint perturbations and input, record it.
                        'The value negative million means that the system was unable to generate a local winner.
                            If LocalWinner <> mcNoWinner Then
                                Let NumberGenerated(FormIndex, LocalWinner) = NumberGenerated(FormIndex, LocalWinner) + 1
                            End If
                Next FormIndex
             'KZ: is it time to check for a cancel command?
                If DummyCounter >= 500 Then
                    'KZ: passes control back to the operating environment to check if there are any pending events--one such
                    '   event would be a second click of cmdRun, which would cancel the running of the algorithm.
                        DoEvents
                        Let DummyCounter = 0
                Else
                    Let DummyCounter = DummyCounter + 1
                End If
          Next TrialIndex           'Loop:  do this many trials of the grammar.

'---------------------------------------------------------------------------

    'For each input, calculate percentages generated for each rival candidate.
        For FormIndex = 1 To mNumberOfForms
           Let LocalSum = 0
           'Dec. 2025:  I think the 0 is an error; we have already folded in the winner.
                    'For RivalIndex = 0 To mNumberOfRivals(FormIndex)
           For RivalIndex = 1 To mNumberOfRivals(FormIndex)
              Let LocalSum = LocalSum + NumberGenerated(FormIndex, RivalIndex)
           Next RivalIndex

           'Dec. 2025:  I think the 0 is an error; we have already folded in the winner.
                    'For RivalIndex = 0 To mNumberOfRivals(FormIndex)
           For RivalIndex = 1 To mNumberOfRivals(FormIndex)
              If LocalSum = 0 Then
                  Let mPercentageGenerated(FormIndex, RivalIndex) = 0
              Else
                  Let mPercentageGenerated(FormIndex, RivalIndex) = NumberGenerated(FormIndex, RivalIndex) / LocalSum
              End If
           Next RivalIndex
        Next FormIndex

    'Calculate analogous percentages for the learning data.
        For FormIndex = 1 To mNumberOfForms
           Let LocalSum = 0
              'Dec. 2025:  I think the 0 is an error; we have already folded in the winner.
            'For RivalIndex = 0 To mNumberOfRivals(FormIndex)
           For RivalIndex = 1 To mNumberOfRivals(FormIndex)
              Let LocalSum = LocalSum + mFrequency(FormIndex, RivalIndex)
           Next RivalIndex
           If LocalSum > 0 Then
            'Ditto
              For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                 Let PercentageInInput(FormIndex, RivalIndex) = mFrequency(FormIndex, RivalIndex) / LocalSum
              Next RivalIndex
           End If
        Next FormIndex

    'Calculate error
        For FormIndex = 1 To mNumberOfForms
            'Dec. 2025:  I think the 0 is an error; we have already folded in the winner.
                'For RivalIndex = 0 To mNumberOfRivals(FormIndex)

            For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                'Previous version was not "centering" the 1-0, .95-.05 case.  Let's try squaring the error.
                Let ErrorSoFar = ErrorSoFar + (PercentageInInput(FormIndex, RivalIndex) - mPercentageGenerated(FormIndex, RivalIndex)) ^ 2
                'Let ErrorSoFar = ErrorSoFar + Abs(PercentageInInput(FormIndex, RivalIndex) - mPercentageGenerated(FormIndex, RivalIndex))
            Next RivalIndex
          'Count the rivals, which will serve as the denominator for the measurement of error.
                Let mTotalNumberOfRivals = mTotalNumberOfRivals + mNumberOfRivals(FormIndex) + 1
       Next FormIndex
       
    'As a function, this routine returns its total error--so it can be used in learning.
            Let NHGTestGrammar = ErrorSoFar
            
    'Calculate the log likelihood of the data.
        Let LogLikelihoodPackage = Module1.CalculateLogLikelihood(mNumberOfForms, mNumberOfRivals(), mPercentageGenerated(), mFrequency())
        Let mLogLikelihood = LogLikelihoodPackage.LogLikelihood
        Let mZeroPredictionWarning = LogLikelihoodPackage.IncludesAZeroProbability
  

'---------------------------------------------------------------------------

    'Print the results, if that's the purpose for which this routine was called.
        If ShouldIPrint = True Then
            'Display the matchups as raw totals, and as percentages.
                'For good line-up, we need to know the longest form.
                    Dim LongestForm As Long
                    Dim LocalLength
                    Let LongestForm = 0
                    For FormIndex = 1 To mNumberOfForms
                        Let LocalLength = Len(mInputForm(FormIndex))
                        If LocalLength > LongestForm Then
                            Let LongestForm = LocalLength
                        End If
                                   'Dec. 2025:  I think the 0 is an error; we have already folded in the winner.
                    'For RivalIndex = 0 To mNumberOfRivals(FormIndex)
                        For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                            Let LocalLength = 3 + Len(DumbSym(mRival(FormIndex, RivalIndex)))
                            If LocalLength > LongestForm Then
                                Let LongestForm = LocalLength
                            End If
                        Next RivalIndex
                    Next FormIndex
        
                Print #mDocFile,
                Print #mDocFile, "\ks"
                Call PrintLevel1Header(mDocFile, mTmpFile, mHTMFile, "Matchup to Input Frequencies")
                
                'Notify if only Wugs were tested.
                '    If TestWugOnly = True Then
                '        Call PrintPara(mDocFile, mTmpFile, mHTMfile, "Note:  only Wug forms (zero frequency total) were tested.")
                '    End If
                    
                For FormIndex = 1 To mNumberOfForms
                    Dim MyTable() As String
                    ReDim MyTable(5, mNumberOfRivals(FormIndex) + 2)
                    If FormIndex > 1 Then Print #mDocFile, "\ks"
                    Print #mDocFile, "\ts5"
                    Print #mTmpFile,
                    Print #mDocFile, "/"; SymbolTag1; mInputForm(FormIndex); SymbolTag2; "/"; vbTab;
                    Print #mTmpFile, "   /"; DumbSym(mInputForm(FormIndex)); "/ ";
                    Let MyTable(1, 1) = mInputForm(FormIndex)
                    For i = Len(mInputForm(FormIndex)) + 4 To LongestForm
                        Print #mTmpFile, " ";
                    Next i
                        Print #mDocFile, "Input Frequencies"; vbTab; "Generated Frequencies"; vbTab; "Input Number (exact)"; vbTab; "Generated Number"
                        Let MyTable(2, 1) = "Input Frequencies"
                        Let MyTable(3, 1) = "Generated Frequencies"
                        Let MyTable(4, 1) = "Input Number (exact)"
                        Let MyTable(5, 1) = "Generated Number"
                        Print #mTmpFile, "Input Fr. Gen Fr.  Input #     Gen. #"
                               'Dec. 2025:  I think the 0 is an error; we have already folded in the winner.
                    'For RivalIndex = 0 To mNumberOfRivals(FormIndex)
                    For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                            Print #mDocFile, SymbolTag1; mRival(FormIndex, RivalIndex); SymbolTag2; vbTab;
                            Let MyTable(1, RivalIndex + 2) = mRival(FormIndex, RivalIndex)
                            Print #mTmpFile, "   "; DumbSym(mRival(FormIndex, RivalIndex));
                            For i = Len(mRival(FormIndex, RivalIndex)) + 1 To LongestForm + 2
                            Print #mTmpFile, " ";
                            Next i
                            Print #mDocFile, ThreeDecPlaces(PercentageInInput(FormIndex, RivalIndex)); vbTab;
                            Let MyTable(2, RivalIndex + 2) = ThreeDecPlaces(PercentageInInput(FormIndex, RivalIndex))
                            Print #mTmpFile, ThreeDecPlaces(PercentageInInput(FormIndex, RivalIndex)); "   ";
                            Print #mDocFile, ThreeDecPlaces(mPercentageGenerated(FormIndex, RivalIndex)); vbTab;
                            Let MyTable(3, RivalIndex + 2) = ThreeDecPlaces(mPercentageGenerated(FormIndex, RivalIndex))
                            Print #mTmpFile, ThreeDecPlaces(mPercentageGenerated(FormIndex, RivalIndex)); "   ";
                            Print #mDocFile, Trim(Str(mActualFrequencyShare(FormIndex, RivalIndex))); vbTab;
                            Let MyTable(4, RivalIndex + 2) = Trim(Str(mActualFrequencyShare(FormIndex, RivalIndex)))
                            Print #mTmpFile, Int8(mActualFrequencyShare(FormIndex, RivalIndex)); "   ";
                            Print #mDocFile, Trim(Str(NumberGenerated(FormIndex, RivalIndex)))
                            Let MyTable(5, RivalIndex + 2) = Trim(Str(NumberGenerated(FormIndex, RivalIndex)))
                            Print #mTmpFile, Int8(NumberGenerated(FormIndex, RivalIndex))
                            'Finally, the tabbed file entries:
                                Print #mTabbedFile, FormIndex; vbTab;
                                Print #mTabbedFile, mInputForm(FormIndex); vbTab;
                                Print #mTabbedFile, RivalIndex; vbTab;
                                Print #mTabbedFile, mRival(FormIndex, RivalIndex); vbTab;
                                Print #mTabbedFile, mActualFrequencyShare(FormIndex, RivalIndex); vbTab;
                                Print #mTabbedFile, NumberGenerated(FormIndex, RivalIndex)
                  Next RivalIndex
                        Print #mDocFile, "\te\ke"
                        Call s.PrintHTMTable(MyTable, mHTMFile, True, False, True)
               Next FormIndex
               
               'Augment the tabbed output with a row-based format.
                    'This needs to be just an option; removing for now.
                        'Print #mTabbedFile,
                        'Print #mTabbedFile,
                        'Print #mTabbedFile,
                        'For FormIndex = 1 To mNumberOfForms
                        '    Print #mTabbedFile, FormIndex; vbtab; mInputForm(FormIndex); vbtab;
                        '    For RivalIndex = 0 To mNumberOfRivals(FormIndex)
                        '        Print #mTabbedFile, RivalIndex; vbtab;
                        '        Print #mTabbedFile, mRival(FormIndex, RivalIndex); vbtab;
                        '        Print #mTabbedFile, ThreeDecPlaces(PercentageInInput(FormIndex, RivalIndex)); vbtab;
                        '        Print #mTabbedFile, mActualFrequencyShare(FormIndex, RivalIndex); vbtab;
                        '        Print #mTabbedFile, NumberGenerated(FormIndex, RivalIndex); vbtab
                        '    Next RivalIndex
                        '    Print #mTabbedFile,
                        'Next FormIndex
                    
                'Add the content of the file.
                
                    Print #mTabbedFile,
                    Call s.PrintContentOfAnInputFile(True, mTabbedFile, mNumberOfConstraints, mConstraintName(), mAbbrev(), _
                        mNumberOfForms, mInputForm(), mWinner(), mWinnerFrequency(), mWinnerViolations(), mNumberOfRivals(), _
                        mRival(), mRivalFrequency(), mRivalViolations())
                        
               
               Close #mTabbedFile
       End If                           'Should I print?
       

End Function



'=================================PRINTING==================================
'===========================================================================

Sub PrintNHGResults(Weight() As Single, ThingFound As String)

   'Print out the results of a numerical algorithm.
        'The Weight() array can be either GLA ranking values or harmonic grammar weights;
        '   and ThingFound (must be grammatically plural) verbally identifies which one it is.
        'Within this code, for historical reasons, the variable is called Weight().

      Dim SlotFiller() As Long
      ReDim SlotFiller(mNumberOfConstraints)
      Dim LocalWeight() As Single
      ReDim LocalWeight(mNumberOfConstraints)

      Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long, InnerConstraintIndex As Long
      Dim ThresholdIndex As Long
      Dim SpaceIndex As Long
      Dim i As Long, j As Long
      
      Dim Swappant As Single, SwapInt As Long
      
      Dim Difference As Single
      
      Dim OldValsFile As Long
      
   'Sort the constraints by their weights.

      For ConstraintIndex = 1 To mNumberOfConstraints
         Let SlotFiller(ConstraintIndex) = ConstraintIndex
         Let LocalWeight(ConstraintIndex) = Weight(ConstraintIndex)
      Next ConstraintIndex

    'December 2025.  Experience suggests this is not useful; the analyst has their own preferred order.
      'For i = 1 To mNumberOfConstraints
      '   For j = 1 To i - 1
      '      If LocalWeight(j) < LocalWeight(i) Then
      '          Let Swappant = LocalWeight(i)
      '          Let LocalWeight(i) = LocalWeight(j)
      '          Let LocalWeight(j) = Swappant
      '          Let SwapInt = SlotFiller(i)
      '          Let SlotFiller(i) = SlotFiller(j)
      '          Let SlotFiller(j) = SwapInt
      '      End If
      '   Next j
      'Next i

   'Print the results of the algorithm.

      Print #mDocFile, "\ks"
      'ThingFound is either weights or ranking values.
        Call PrintLevel1Header(mDocFile, mTmpFile, mHTMFile, ThingFound + " Found")
      
           Print #mDocFile, "\ts2"
           Print #mDocFile, "Constraints"; vbTab; ThingFound + " Found"
            
           For ConstraintIndex = 1 To mNumberOfConstraints
                Print #mDocFile, SmallCapTag1; mConstraintName(SlotFiller(ConstraintIndex)); SmallCapTag2;
                Print #mDocFile, vbTab;
                Print #mDocFile, ThreeDecPlaces(Weight(SlotFiller(ConstraintIndex)));
                Print #mDocFile, vbTab;
                If gExponentialNHG Then
                    Print #mDocFile, ThreeDecPlaces(Exp(Weight(SlotFiller(ConstraintIndex))))
                End If
              
                Print #mTmpFile, FillStringTo(ThreeDecPlaces(Weight(SlotFiller(ConstraintIndex))), 10);
                If gExponentialNHG Then
                    Print #mTmpFile, FillStringTo(ThreeDecPlaces(Exp(Weight(SlotFiller(ConstraintIndex)))), 10);
                End If
                Print #mTmpFile, "   "; mConstraintName(SlotFiller(ConstraintIndex))
           Next ConstraintIndex
        
           Print #mDocFile, "\te\ke"
      
    'We also want a plain output that can be processed by Excel.
    '   Weights should be reported in straight constraint order, to facilitate
    '   averaging over multiple runs.
    
        'If LCase(Right(gFileName, 6)) = "tabbed" Then
            Let mTabbedFile = FreeFile
            Open gOutputFilePath + gFileName + "TabbedOutput.txt" For Output As #mTabbedFile
            For ConstraintIndex = 1 To mNumberOfConstraints
               Print #mTabbedFile, mAbbrev(ConstraintIndex);
               Print #mTabbedFile, vbTab; Weight(ConstraintIndex);
               If gExponentialNHG Then
                    Print #mTabbedFile, vbTab; Exp(Weight(ConstraintIndex));
               End If
               Print #mTabbedFile,
            Next ConstraintIndex
        'End If
      
   'Print a file to save results if you want to run it further.

        'First, make sure there is a folder for these files, a daughter of the
        '   folder in which the input file is located.
            Call Form1.CreateAFolderForOutputFiles
      
        'Print the file.
            Let OldValsFile = FreeFile
            Open gOutputFilePath + gFileName + "MostRecentWeights.txt" For Output As #OldValsFile
            For ConstraintIndex = 1 To mNumberOfConstraints
               Print #OldValsFile, mAbbrev(ConstraintIndex); vbTab; Weight(ConstraintIndex)
            Next ConstraintIndex

End Sub

Sub PrintAHeader(MyAlgorithmName As String)

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
        Print #mTmpFile, NiceDate; ", "; NiceTime
        Print #mTmpFile,
        Print #mTmpFile, "OTSoft " + gMyVersionNumber + ", release date " + gMyReleaseDate
        
          
    'Print a header and diacritic to trigger a page number
        Print #mDocFile, "\hrNHG Results for "; gFileName; gFileSuffix; Chr$(9); NiceDate; Chr$(9); "\pn"


End Sub


Function ThreeDecPlaces(ALong As Variant) As String

    Dim Buffer As String
    'Format numbers with three decimal places
        Let Buffer = Format(ALong, "##,##0.000")
        Let ThreeDecPlaces = Trim(Buffer)

End Function


Function Int8(ALong As Long) As String

    'Format numbers with eight digits and no places
        Dim i As Long
        Let Int8 = Format(ALong, "########")
        For i = Len(Int8) To 7
            Let Int8 = " " + Int8
        Next i

End Function


Sub DebugRoutine(LowerBound As Long)
            
        Dim RivalIndex As Long, FormIndex As Long, ConstraintIndex As Long
        Dim DebugFile As Long
        
        Let DebugFile = FreeFile
        
        'First, make sure there is a folder for these files, a daughter of the
        '   folder in which the input file is located.
            Call Form1.CreateAFolderForOutputFiles
        
        Open gOutputFilePath + "Debug.txt" For Output As #DebugFile
        
        For FormIndex = 1 To mNumberOfForms
            Print #DebugFile, "Input #"; Trim(Str(FormIndex)); "    "; mInputForm(FormIndex)
            Print #DebugFile, "      Winner:  "; mWinner(FormIndex);
            Print #DebugFile, "   Frequency:  "; mWinnerFrequency(FormIndex)
            Print #DebugFile, "         ";
            For ConstraintIndex = 1 To mNumberOfConstraints
                Print #DebugFile, mWinnerViolations(FormIndex, ConstraintIndex); " ";
            Next ConstraintIndex
            Print #DebugFile,
            For RivalIndex = LowerBound To mNumberOfRivals(FormIndex)
                Print #DebugFile, RivalIndex; "   Rival:  "; mRival(FormIndex, RivalIndex); _
                    "   Frequency = "; mFrequency(FormIndex, RivalIndex)
                Print #DebugFile, "         ";
                For ConstraintIndex = 1 To mNumberOfConstraints
                    Print #DebugFile, mRivalViolations(FormIndex, RivalIndex, ConstraintIndex); " ";
                Next ConstraintIndex
                Print #DebugFile,
            Next RivalIndex
        Next FormIndex
        For ConstraintIndex = 1 To mNumberOfConstraints
            Print #DebugFile, ConstraintIndex, mAbbrev(ConstraintIndex), mConstraintName(ConstraintIndex)
        Next ConstraintIndex
            
End Sub


Sub PrepareApproximateTableaux()

    'Make approximate tableaux, by calling the tableau-making machinery of Form1.
    'Constraint names are amplified by their weights.
    
    'We are trying to assemble the following information:

    Dim LocalStratum() As Long      'To store the pseudostrata.
    ReDim LocalStratum(mNumberOfConstraints)
    Dim StratumIndex As Long
    Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
    Dim i As Long, j As Long
    Dim BestWeight As Single  'For converting weights into strata
    Dim BestForStratum As Long      'Ditto
    Dim LocalConstraintName() As String
    ReDim LocalConstraintName(mNumberOfConstraints)
    Dim LocalAbbrev() As String
    ReDim LocalAbbrev(mNumberOfConstraints)
    'Finding winners:
        Dim LocalWinner() As String
        ReDim LocalWinner(mNumberOfForms)
        Dim CurrentBest As Long
    Dim ExitFlag As Boolean
    Dim SlotFiller() As Long
    ReDim SlotFiller(mNumberOfConstraints)
    Dim LocalWeight() As Single
    ReDim LocalWeight(mNumberOfConstraints)
    Dim Swappant As Single
    Dim SwapInt As Long
    Dim LocalWinnerViolations() As Long
    ReDim LocalWinnerViolations(mNumberOfForms, mNumberOfConstraints)
    Dim LocalWinnerFrequency() As Single
    ReDim LocalWinnerFrequency(mNumberOfForms)
    Dim LocalNumberOfRivals() As Long
    ReDim LocalNumberOfRivals(mNumberOfForms)
    Dim PutRivalHere As Long
    Dim LocalRival() As String
    ReDim LocalRival(mNumberOfForms, mMaximumNumberOfRivals)
    Dim LocalRivalFrequency() As Single
    ReDim LocalRivalFrequency(mNumberOfForms, mMaximumNumberOfRivals)
    Dim LocalRivalViolations() As Long
    ReDim LocalRivalViolations(mNumberOfForms, mMaximumNumberOfRivals, mNumberOfConstraints)
    'Find pairwise ranking probabilities:
        Dim ThresholdIndex As Long
        Dim Amplification As String
        Dim Difference As Single
        
   'Sort the constraints by their weights.
        For ConstraintIndex = 1 To mNumberOfConstraints
            Let SlotFiller(ConstraintIndex) = ConstraintIndex
            Let LocalWeight(ConstraintIndex) = mWeight(ConstraintIndex)
        Next ConstraintIndex
        For i = 2 To mNumberOfConstraints
            For j = 1 To i - 1
                If LocalWeight(j) < LocalWeight(i) Then
                    Let Swappant = LocalWeight(i)
                    Let LocalWeight(i) = LocalWeight(j)
                    Let LocalWeight(j) = Swappant
                    Let SwapInt = SlotFiller(i)
                    Let SlotFiller(i) = SlotFiller(j)
                    Let SlotFiller(j) = SwapInt
                End If
            Next j
        Next i
   
   'Convert weights directly into strata.
        'Initialize:
            For ConstraintIndex = 1 To mNumberOfConstraints
                Let LocalStratum(ConstraintIndex) = 0
            Next ConstraintIndex
        'Sort:
            For StratumIndex = 1 To mNumberOfConstraints
                Let BestWeight = -1000000
                Let BestForStratum = 0
                For ConstraintIndex = 1 To mNumberOfConstraints
                    If LocalStratum(ConstraintIndex) = 0 Then
                        If mWeight(ConstraintIndex) > BestWeight Then
                            Let BestForStratum = ConstraintIndex
                            Let BestWeight = mWeight(ConstraintIndex)
                        End If
                    End If
                Next ConstraintIndex
                Let LocalStratum(BestForStratum) = StratumIndex
            Next StratumIndex
   
   'Amplify the strings in mAbbrev() to include the weight.

        For ConstraintIndex = 1 To mNumberOfConstraints
            Let Amplification = " (" + Trim(Str(ThreeDecPlaces(mWeight(SlotFiller(ConstraintIndex))))) + ")"
            Let LocalAbbrev(SlotFiller(ConstraintIndex)) = mAbbrev(SlotFiller(ConstraintIndex)) + Amplification
        Next ConstraintIndex
    
    'The constraint names are not a big deal; just carry them over
        For ConstraintIndex = 1 To mNumberOfConstraints
            Let LocalConstraintName(mNumberOfConstraints) = mConstraintName(mNumberOfConstraints)
        Next ConstraintIndex
               
    'Debug this routine:
    '    Dim DebugFile As Long
    '    Let DebugFile = FreeFile
    '    Open gOutputFilePath + "DebugAbbrevs.txt" For Output As DebugFile
    '    For ConstraintIndex = 1 To mNumberofconstraints
    '        Print #DebugFile, mConstraintName(ConstraintIndex); vbtab; LocalAbbrev(ConstraintIndex)
    '    Next ConstraintIndex
        
   'Find the winners under this GLA grammar, when run nonstochastically,
   '    and install them as such, with violations and frequencies.
        For FormIndex = 1 To mNumberOfForms
            
            'Test wug only if user asked.
            '    If TestWugOnly = True Then
            '        If IsAWugForm(FormIndex) = False Then
            '            GoTo ExitPoint
            '        End If
            '    End If

            Let CurrentBest = 0     'Assume that the zeroth rival is best, then start comparing with 1.
            For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                'Use the SlotFiller() array to get the constraints in descending order
                '   of ranking.
                    For ConstraintIndex = 1 To mNumberOfConstraints
                        If mRivalViolations(FormIndex, RivalIndex, SlotFiller(ConstraintIndex)) < mRivalViolations(FormIndex, CurrentBest, SlotFiller(ConstraintIndex)) Then
                            Let CurrentBest = RivalIndex
                            'It's OT, with strict domination, so exit.
                                Exit For
                        ElseIf mRivalViolations(FormIndex, RivalIndex, SlotFiller(ConstraintIndex)) > mRivalViolations(FormIndex, CurrentBest, SlotFiller(ConstraintIndex)) Then
                            Exit For
                        End If
                    Next ConstraintIndex
            Next RivalIndex
            'You have the winner.  Install it in the arrays that will be used for tableau-printing.
                Let LocalWinner(FormIndex) = mRival(FormIndex, CurrentBest)
                'Note that mWinnerFrequency() and mRivalFrequency() have been amalgamated into mFrequency().
                    Let LocalWinnerFrequency(FormIndex) = mPercentageGenerated(FormIndex, CurrentBest)
                'To help diagnosis, append to the winner its frequency, if nonnull.
                    If LocalWinnerFrequency(FormIndex) > 0 Then
                        Let LocalWinner(FormIndex) = LocalWinner(FormIndex) + " (" + ThreeDecPlaces(LocalWinnerFrequency(FormIndex)) + ")"
                    End If
                Let LocalNumberOfRivals(FormIndex) = mNumberOfRivals(FormIndex)
                For ConstraintIndex = 1 To mNumberOfConstraints
                    Let LocalWinnerViolations(FormIndex, ConstraintIndex) = mRivalViolations(FormIndex, CurrentBest, ConstraintIndex)
                Next ConstraintIndex
            'Now, bunch the rivals in their array, so that position 0 is no longer filled.
                Let PutRivalHere = 1
                    'Dec. 2025:  I think the 0 is an error; we have already folded in the winner.
                    'For RivalIndex = 0 To mNumberOfRivals(FormIndex)
                For RivalIndex = 1 To LocalNumberOfRivals(FormIndex)
                    If RivalIndex <> CurrentBest Then
                        Let LocalRival(FormIndex, PutRivalHere) = mRival(FormIndex, RivalIndex)
                        Let LocalRivalFrequency(FormIndex, PutRivalHere) = mPercentageGenerated(FormIndex, RivalIndex)
                        'To help diagnosis, append to the rival its frequency.
                            If LocalRivalFrequency(FormIndex, PutRivalHere) > 0 Then
                                Let LocalRival(FormIndex, PutRivalHere) = LocalRival(FormIndex, PutRivalHere) + " (" + ThreeDecPlaces(LocalRivalFrequency(FormIndex, PutRivalHere)) + ")"
                            End If
                            For ConstraintIndex = 1 To mNumberOfConstraints
                                Let LocalRivalViolations(FormIndex, PutRivalHere, ConstraintIndex) = mRivalViolations(FormIndex, RivalIndex, ConstraintIndex)
                            Next ConstraintIndex
                        Let PutRivalHere = PutRivalHere + 1
                    End If
                Next RivalIndex
ExitPoint:
        Next FormIndex
        
    'Debug so far:
    '    Dim DebugFile As Long
    '    Let DebugFile = FreeFile
    '    Open gOutputFilePath + "DebugWinners.txt" For Output As DebugFile
    '    For FormIndex = 1 To mNumberOfForms
    '        Print #DebugFile, mInputForm(FormIndex); vbtab; _
    '        "winner #"; vbtab; LocalWinner(FormIndex); _
    '        vbtab; LocalWinner(FormIndex); _
    '            vbtab; "Frequency is "; vbtab; Str(LocalWinnerFrequency(FormIndex)); vbtab; _
    '            "Violations:";
    '        For ConstraintIndex = 1 To mNumberofconstraints
    '            Print #DebugFile, vbtab; LocalWinnerViolations(FormIndex, ConstraintIndex);
     '       Next ConstraintIndex
     ''       Print #DebugFile,
     '       For RivalIndex = 1 To NumberOfRivals(FormIndex)
     '           Print #DebugFile, mInputForm(FormIndex); vbtab; _
      ''          vbtab; vbtab; Rival(FormIndex, RivalIndex); _
      '              vbtab; "Frequency is "; vbtab; LocalRivalFrequency(FormIndex, RivalIndex); vbtab; _
      '              "Violations:";
      '          For ConstraintIndex = 1 To mNumberofconstraints
      '              Print #DebugFile, vbtab; LocalRivalViolations(FormIndex, RivalIndex, ConstraintIndex);
      '          Next ConstraintIndex
      '          Print #DebugFile,
       '     Next RivalIndex
       ''
       ' Next FormIndex
   
   'Print tableaux if the user wants them.
        If mnuIncludeTableaux.Checked = True Then
            Call PrintTableaux.Main(mNumberOfForms, mNumberOfConstraints, LocalConstraintName(), _
                LocalAbbrev(), LocalStratum(), mInputForm(), LocalWinner(), _
                mWinnerFrequency(), LocalWinnerViolations(), _
                mMaximumNumberOfRivals, LocalNumberOfRivals(), LocalRival(), mRivalFrequency(), LocalRivalViolations(), _
                mTmpFile, mDocFile, mHTMFile, _
                "Noisy Harmonic Grammar", False, 0, False, False)
        End If
    
    
End Sub

Sub PrintFinalDetails()
    
    Dim ConstraintIndex As Long
    
    'For sorting:
        Dim SlotFiller() As Long
        Dim LocalWeight() As Single
        ReDim SlotFiller(mNumberOfConstraints)
        ReDim LocalWeight(mNumberOfConstraints)
        Dim i As Long, j As Long
        Dim Swappant As Single, SwapInt As Long
    
    Dim LearningStageIndex As Long      'for learning schedule table
    Dim Buffer As String                'to help with spacing
 
    'Insert the Hasse diagram if appropriate.
        Call Form1.InsertHasseDiagramIntoOutputFile(mDocFile, mHTMFile)
    
    'Identify disasters.
        If mInfinityWarning Then
            Call s.p(mDocFile, mTmpFile)
            Call s.p(mDocFile, mTmpFile, "   ", "Caution: at some point in setting the weights, the program went into an endless loop of harmony ties and force-exited this loop.")
            Call s.p(mDocFile, mTmpFile, "   ", "The results of this event on learning are unpredictable.")
            Call s.p(mDocFile, mTmpFile)
        End If
        
        If mTieWarning Then
            If gExponentialNHG Then
                Call s.p(mDocFile, mTmpFile, "   ", "Caution: at some point a candidate yielded two tied winners.")
                Call s.p(mDocFile, mTmpFile, "   ", "A possible cause of this in Exponential HNG is that a weight has gone so low that its exponentiated version is represented as zero by the programming language;")
                Call s.p(mDocFile, mTmpFile, "   ", "though there can be other causes.")
                Call s.p(mDocFile, mTmpFile, "   ", "In any event, the result of this run should not be considered reliable.")
            Else
                Call s.p(mDocFile, mTmpFile, "   ", "Caution: at some point a candidate yielded two tied winners,")
                Call s.p(mDocFile, mTmpFile, "   ", "so the result of this run should not be considered reliable.")
            End If
            Call s.p(mDocFile, mTmpFile)
        End If
    
    
    'Note the accuracy achieved, and the time needed.
       Print #mDocFile,
        'First, a header:
            Call PrintLevel1Header(mDocFile, mTmpFile, mHTMFile, "Testing the Grammar:  Details")
        'Results:
            Call PrintPara(mDocFile, mTmpFile, mHTMFile, "The grammar was tested for " + Trim(Str(gCyclesToTest)) + " cycles.")
            Call PrintPara(mDocFile, mTmpFile, mHTMFile, "Average error per candidate:  " + ThreeDecPlaces(100 * mErrorTerm / mTotalNumberOfRivals) + " percent")
            Call PrintPara(mDocFile, mTmpFile, mHTMFile, "Log likelihood of the data: " + ThreeDecPlaces(mLogLikelihood))
            If mZeroPredictionWarning Then
                Call PrintPara(mDocFile, mTmpFile, mHTMFile, "Caution:  at least one candidate with positive was assigned zero probability; since zero has no log this was approximated as .001.")
            End If
            Call PrintPara(mDocFile, mTmpFile, mHTMFile, "Learning time:  " + ThreeDecPlaces(mTimeMarker / 60) + " minutes")
    
    'Print the details of the learning simulation.
        Print #mDocFile,
        Print #mTmpFile,
        Call PrintLevel1Header(mDocFile, mTmpFile, mHTMFile, "Parameter Values Used")
    
    'Print the choices made.
    
            Print #mDocFile,
            Print #mTmpFile,
        
        
        'Plainly identifiable versions should be labeled as such.
        
           'Classical NHG.
                If chkNegativeWeightsOK.Value <> vbChecked Then
                    If chkLateNoise.Value <> vbChecked Then
                        If chkNoiseAppliesToTableauCells.Value <> vbChecked Then
                            If chkNoiseIsAddedAfterMultiplication.Value <> vbChecked Then
                                If gExponentialNHG = False Then
                                    Print #mDocFile, "   The algorithm used was the Classical version of Noisy Harmonic Grammar (Boersma and Pater 2013)."
                                    Print #mTmpFile, "   The algorithm used was the Classical version of Noisy Harmonic Grammar (Boersma and Pater 2013)."
                                End If
                            End If
                        End If
                    End If
                End If
        
        'Exponential version.
            If gExponentialNHG Then
            Print #mDocFile,
            Print #mTmpFile,
                Print #mDocFile, "   In calculating winners, Harmony values were exponentiated (Exponential NHG)."
                Print #mTmpFile, "   In calculating winners, Harmony values were exponentiated (Exponential NHG)."
            End If
            
        'Exponential version.
            If gDemiGaussianNHG Then
            Print #mDocFile,
            Print #mTmpFile,
                Print #mDocFile, "   Demigaussian NHG: in calculating winners, weight values were sampled only from the non-negative side of the Gaussian distribution."
                Print #mTmpFile, "   Demigaussian NHG: in calculating winners, weight values were sampled only from the non-negative side of the Gaussian distribution."
            End If
            
        'Noise
            Print #mDocFile,
            Print #mTmpFile,
            Print #mDocFile, "   Noise was set to have a standard deviation of " + Trim(Str(mNoise)) + "."
            Print #mTmpFile, "   Noise was set to have a standard deviation of " + Trim(Str(mNoise)) + "."
        
        'Negative weights.
            If gExponentialNHG = False Then
                Print #mDocFile,
                Print #mTmpFile,
                If chkNegativeWeightsOK.Value = vbChecked Then
                    Print #mDocFile, "   Negative weights were permitted."
                    Print #mDocFile, "   The number of times that a negative weight was recorded during learning and testing was "; Trim(Str(mNumberOfTimesWeightsWentBelowZero)); " out of "; _
                        Trim(Str(mNumberOfChancesForWeightsToGoBelowZero)); " chances."
                    Print #mTmpFile, "   Negative weights were permitted."
                    Print #mTmpFile, "   The number of times that a negative weight was recorded during learning and testing was "; Trim(Str(mNumberOfTimesWeightsWentBelowZero)); " out of "; _
                        Trim(Str(mNumberOfChancesForWeightsToGoBelowZero)); " chances."
                Else
                    Print #mDocFile, "   Negative weights were not permitted."
                    Print #mDocFile, "   The number of times that a negative weight was clipped during learning and testing was "; Trim(Str(mNumberOfTimesWeightsWentBelowZero)); " out of "; _
                        Trim(Str(mNumberOfChancesForWeightsToGoBelowZero)); " chances."
                    Print #mTmpFile, "   Negative weights were not permitted."
                    Print #mTmpFile, "   The number of times that a negative weight was clipped during learning and testing was "; Trim(Str(mNumberOfTimesWeightsWentBelowZero)); " out of "; _
                        Trim(Str(mNumberOfChancesForWeightsToGoBelowZero)); " chances."
                End If
            End If
            
        'Ties
            Print #mDocFile,
            Print #mTmpFile,
            Print #mDocFile, "   The number of ties between winner and rival was  "; Trim(Str(mNumberOfTies)); " out of "; _
                Trim(Str(mNumberOfChancesForTie)); " chances."
            Print #mTmpFile, "   The number of ties between winner and rival was "; Trim(Str(mNumberOfTies)); " out of "; _
                ; Trim(Str(mNumberOfChancesForTie)); " chances."
            If gResolveTiesBySkipping Then
                Print #mDocFile, "   Ties were resolved by dropping the trial in question."
                Print #mTmpFile, "   Ties were resolved by dropping the trial in question."
            Else
                Print #mDocFile, "   Ties were resolved by making a random choice."
                Print #mTmpFile, "   Ties were resolved by making a random choice."
            End If
            
        'Where noise applied.
            'Noise applied late, to candidates.
                If chkLateNoise.Value = vbChecked Then
                    Print #mDocFile,
                    Print #mTmpFile,
                    Print #mDocFile, "   Noise was applied to candidates, following all of the harmony computations."
                    Print #mTmpFile, "   Noise was applied to candidates, following all of the harmony computations."
                Else
                    'The 2 x 2 array of "earlier" noise options.
                        'Cell specific noise.
                            Print #mDocFile,
                            Print #mTmpFile,
                            If chkNoiseAppliesToTableauCells.Value = vbChecked Then
                                Print #mDocFile, "   Noise was applied to individual tableau cells."
                                Print #mTmpFile, "   Noise was applied to individual tableau cells."
                            Else
                                Print #mDocFile, "   Noise was applied to constraints, hence the same for all candidates."
                                Print #mTmpFile, "   Noise was applied to constraints, hence the same for all candidates."
                            End If
                       'Late noise.
                            Print #mDocFile,
                            Print #mTmpFile,
                            If chkNoiseIsAddedAfterMultiplication.Value = vbChecked Then
                                Print #mDocFile, "   Noise was applied after weights multiplied by violations."
                                Print #mTmpFile, "   Noise was applied after weights multiplied by violations."
                            Else
                                Print #mDocFile, "   Noise was applied before weights multiplied by violations."
                                Print #mTmpFile, "   Noise was applied before weights multiplied by violations."
                            End If
                End If
    
        'Initial weights
            'First, a header:
                Print #mDocFile, "\h2Initial weights"
                Print #mTmpFile,
            'We really only want a table if it's necessary."
                Select Case InitialRankingChoice
                    Case AllSame
                        'Trivial, no table.
                            Print #mDocFile,
                            Print #mDocFile, "All constraints started out at the default value of 0."
                            Print #mTmpFile, "   Initial Weights"
                            Print #mTmpFile,
                            Print #mTmpFile, "      All constraints started out at the default value of 0."
                    Case Else
                        'Nontrivial, so a table.
                            'First, sort the constraints by their initial weights.
                                For ConstraintIndex = 1 To mNumberOfConstraints
                                    Let SlotFiller(ConstraintIndex) = ConstraintIndex
                                    Let LocalWeight(ConstraintIndex) = mInitialWeight(ConstraintIndex)
                                Next ConstraintIndex
                                For i = 1 To mNumberOfConstraints
                                    For j = 1 To i - 1
                                        If LocalWeight(j) < LocalWeight(i) Then
                                            Let Swappant = LocalWeight(i)
                                            Let LocalWeight(i) = LocalWeight(j)
                                            Let LocalWeight(j) = Swappant
                                            Let SwapInt = SlotFiller(i)
                                            Let SlotFiller(i) = SlotFiller(j)
                                            Let SlotFiller(j) = SwapInt
                                        End If
                                    Next j
                                Next i
                            'Now you can print the table.
                                Print #mDocFile, "\ts3"
                                Print #mTmpFile, "   Initial Weights (final weights shown in parentheses)"
                                Print #mTmpFile,
                                Print #mDocFile, "Constraint"; vbTab; "Initial Ranking"; vbTab; "Final Ranking"
                                For ConstraintIndex = 1 To mNumberOfConstraints
                                    Print #mDocFile, SmallCapTag1; mConstraintName(SlotFiller(ConstraintIndex)); SmallCapTag2; vbTab;
                                    Print #mDocFile, ThreeDecPlaces(mInitialWeight(SlotFiller(ConstraintIndex))); vbTab;
                                    Print #mDocFile, ThreeDecPlaces(mWeight(SlotFiller(ConstraintIndex)))
                                    Print #mTmpFile, FillStringTo(ThreeDecPlaces(mInitialWeight(SlotFiller(ConstraintIndex))), 13); ThreeDecPlaces(mInitialWeight(SlotFiller(ConstraintIndex))); "   ";
                                    Print #mTmpFile, FillStringTo(ThreeDecPlaces(mWeight(SlotFiller(ConstraintIndex))), 6); "("; ThreeDecPlaces(mWeight(SlotFiller(ConstraintIndex))); ")";
                                    Print #mTmpFile, "   "; mConstraintName(SlotFiller(ConstraintIndex))
                                Next ConstraintIndex
                                Print #mDocFile, "\te"
                End Select      'Did we want a fancy table for nontrivial initial weights?
                Print #mDocFile, "\ke"
            
        'Learning schedule
            'First, a header:
                Print #mDocFile, "\ks"
                Print #mDocFile, "\h2Schedule for Plasticity"
                Print #mTmpFile,
                Print #mTmpFile, "   Schedule for Plasticity"
                Print #mTmpFile,
            'Then a table:
                Print #mDocFile, "\ts6"
                Print #mDocFile, "Stage"; vbTab; "Trials"; vbTab; "PlastMark"; vbTab; "PlastFaith"
                Print #mTmpFile, "      Stage   Trials   PlastMark  PlastFaith "
                For LearningStageIndex = 1 To mNumberOfLearningStages
                    Print #mDocFile, Trim(Val(LearningStageIndex)); vbTab;
                    Print #mDocFile, Trim(Val(mTrialsPerLearningStage(LearningStageIndex))); vbTab;
                    Print #mDocFile, ThreeDecPlaces(CustomPlastMark(LearningStageIndex)); vbTab;
                    Print #mDocFile, ThreeDecPlaces(CustomPlastFaith(LearningStageIndex))
                    Print #mTmpFile, "     ";
                    Let Buffer = Trim(Val(LearningStageIndex))
                        Print #mTmpFile, FillStringTo(Buffer, 9);
                    Let Buffer = Trim(Val(mTrialsPerLearningStage(LearningStageIndex)))
                        Print #mTmpFile, FillStringTo(Buffer, 11);
                    Let Buffer = ThreeDecPlaces(CustomPlastMark(LearningStageIndex))
                        Print #mTmpFile, FillStringTo(Buffer, 11);
                    Let Buffer = ThreeDecPlaces(CustomPlastFaith(LearningStageIndex))
                        Print #mTmpFile, FillStringTo(Buffer, 11)
                Next LearningStageIndex
                Print #mDocFile, "\te"
                Print #mDocFile, "There were a total of "; Trim(Str(mReportedNumberOfDataPresentations)); " learning trials."
                Print #mTmpFile,
                Print #mTmpFile, "      There were a total of "; Trim(Str(mReportedNumberOfDataPresentations)); " learning trials."
                If mnuExactProportions.Checked = True Then
                    Print #mDocFile,
                    Print #mDocFile, "Data were presented non-stochastically, in exact proportions to their frequencies in the input file."
                    Print #mTmpFile,
                    Print #mTmpFile, "      Data were presented non-stochastically, in exact proportions to their "
                    Print #mTmpFile, "      frequencies in the input file."
                End If
                Print #mDocFile, "\ke"
               
End Sub




Sub DebugMe(MyNumberOfForms As Long, myInputForm() As String, myMaximumNumberOfRivals As Long, myNumberOfRivals() As Long, MyRival() As String, MyRivalViolations() As Long, myFrequency() As Single, myNumberOfConstraints As Long, myConstraintnames() As String)

    Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
    Dim DebugFile As Long
    
    Let DebugFile = FreeFile
    Open gOutputFilePath + "/DebuggingNHGDecember2025.txt" For Output As DebugFile
    
    Print #DebugFile, vbTab;
    For ConstraintIndex = 1 To myNumberOfConstraints
        Print #DebugFile, vbTab; myConstraintnames(ConstraintIndex);
    Next ConstraintIndex
    Print #DebugFile,
    For FormIndex = 1 To mNumberOfForms
        Print #DebugFile, myInputForm(FormIndex)
        'Dec. 2025:  I think the 0 is an error; we have already folded in the winner.
                    'For RivalIndex = 0 To mNumberOfRivals(FormIndex)
        For RivalIndex = 1 To myNumberOfRivals(FormIndex)
            Print #DebugFile, MyRival(FormIndex, RivalIndex); myFrequency(FormIndex, RivalIndex);
                For ConstraintIndex = 1 To myNumberOfConstraints
                    Print #DebugFile, vbTab; MyRivalViolations(FormIndex, RivalIndex, ConstraintIndex);
                Next ConstraintIndex
                Print #DebugFile,
        Next RivalIndex
    Next FormIndex

    Close DebugFile
    
End Sub

