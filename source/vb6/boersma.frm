VERSION 5.00
Begin VB.Form GLA 
   Caption         =   "Gradual Learning Algorithm"
   ClientHeight    =   7065
   ClientLeft      =   3060
   ClientTop       =   1815
   ClientWidth     =   9045
   Icon            =   "boersma.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   9045
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Choose framework"
      Height          =   1095
      Left            =   600
      TabIndex        =   17
      Top             =   360
      Width           =   4095
      Begin VB.OptionButton optMaxEnt 
         Caption         =   "Compute weights for MaxEnt"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Value           =   -1  'True
         Width           =   3495
      End
      Begin VB.OptionButton optStochasticOT 
         Caption         =   "Compute ranking values for Stochastic OT"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.TextBox txtValueThatImplementsAPrioriRankings 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3720
      TabIndex        =   14
      Text            =   "20"
      Top             =   3240
      Width           =   975
   End
   Begin VB.PictureBox pctProgressWindow 
      Height          =   6735
      Left            =   4920
      ScaleHeight     =   6675
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
      Top             =   6240
      Width           =   1695
   End
   Begin VB.TextBox txtTimesToTestGrammar 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3720
      TabIndex        =   5
      Text            =   "10000"
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox txtLowerPlasticity 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3720
      TabIndex        =   4
      Text            =   ".001"
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox txtUpperPlasticity 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3720
      TabIndex        =   3
      Text            =   "2"
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtNumberOfCycles 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3720
      TabIndex        =   1
      Text            =   "1000000"
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Run"
      Default         =   -1  'True
      Height          =   615
      Left            =   3000
      TabIndex        =   0
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label lblExact 
      Alignment       =   1  'Right Justify
      Caption         =   "Learning data presented to GLA in exact proportions"
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   4440
      Width           =   4575
   End
   Begin VB.Label lblConstraintsRankedAPrioriMustDifferBy 
      Caption         =   "Constraints ranked a priori must differ  by"
      Height          =   255
      Left            =   600
      TabIndex        =   15
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Label lblAPrioriRankings 
      Alignment       =   1  'Right Justify
      Caption         =   "A priori rankings in effect"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   5520
      Width           =   4455
   End
   Begin VB.Label lblCustomRank 
      Alignment       =   1  'Right Justify
      Caption         =   "Custom initial rankings in effect"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   5160
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label lblCustomLearningSchedule 
      Alignment       =   1  'Right Justify
      Caption         =   "Custom learning schedule in effect"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   4800
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
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label lblFinalMark 
      Alignment       =   1  'Right Justify
      Caption         =   "Final  plasticity"
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label lblInitMark 
      Alignment       =   1  'Right Justify
      Caption         =   "Initial  plasticity"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Number of times to go through forms"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Menu mnuInitialRankings 
      Caption         =   "&Initial ranking values/weights"
      Begin VB.Menu mnuUseDefaultInitialRankingValues 
         Caption         =   "Use default initial values (all same)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuUseSeparateMarkFaithInitialRankings 
         Caption         =   "Use separate initial values for Markedness and Faithfulness"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuUseFullyCustomizedInitialRankingValues 
         Caption         =   "Use fully customized initial values"
      End
      Begin VB.Menu mnuUsePreviousResultsAsInitialRankingValues 
         Caption         =   "Use results of previous run as initial values"
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSpecifySeparateMarkFaithInitialRankings 
         Caption         =   "Specify &separate initial values for Markedness and Faithfulness"
      End
      Begin VB.Menu mnuFullyCustomizedInitialRankingValues 
         Caption         =   "Edit file for fully customized initial values"
      End
   End
   Begin VB.Menu mnuLearningSchedule 
      Caption         =   "&Learning schedule"
      Begin VB.Menu mnuMagri 
         Caption         =   "Use the &Magri update rule (stochastic OT)"
      End
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
      Caption         =   "&A Priori Rankings"
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
   Begin VB.Menu mnuMaxEntHeader 
      Caption         =   "&MaxEnt"
      Begin VB.Menu mnuEditFileForMaxentParameters 
         Caption         =   "Edit file with MaxEnt learning parameters"
      End
      Begin VB.Menu mnuGaussianPrior 
         Caption         =   "Run MaxEnt with Gaussian prior"
      End
      Begin VB.Menu mnuBatchMaxEnt 
         Caption         =   "Run the batch version of MaxEnt"
      End
      Begin VB.Menu mnuNegativeWeightsOK 
         Caption         =   "Allow weights to be negative"
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
         Caption         =   "Print file with history of weights/ranking values"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFullHistory 
         Caption         =   "Print same but more annotations"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuCandidateProbabilityHistory 
         Caption         =   "Print history of candidate probabilities (MaxEnt only)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuExactProportions 
         Caption         =   "Present data to GLA in exact proportions"
      End
      Begin VB.Menu mnuTestWugOnly 
         Caption         =   "Test wug forms only"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mMultipleRuns 
         Caption         =   "Multiple runs"
         Begin VB.Menu mnu10Runs 
            Caption         =   "10 runs"
         End
         Begin VB.Menu mnu100Runs 
            Caption         =   "100 runs"
         End
         Begin VB.Menu mnu1000Runs 
            Caption         =   "1000 runs"
         End
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
      Begin VB.Menu mnuHelpSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAboutOTSoft 
         Caption         =   "About OTSoft"
      End
   End
End
Attribute VB_Name = "GLA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=================================GLA.FRM=====================================
'======Rank a set of constraints given a set of input data with violations====
'==============Use the Gradual Learning Algorithm of Paul Boersma=============

'January 2026:  Add the additional capacity to weight MaxEnt constraints gradually,
'               using the method given by Gerhard Jaeger.
   
   Option Explicit

    'Localized versions of general OTSoft variables.  These get filled by vaues taken from Form1
        
        'Which framework to work with (MaxEnt or StochasticOT).
            Dim mMyFramework As String
        
        'About the learning data:
            Dim mInputForm() As String
            Dim mWinner() As String
            Dim mNumberOfForms As Long              'I.e. number of inputs.
            Dim mWinnerFrequency() As Single        'This gets folded in to a general mFrequency() variable.
            Dim mWinnerViolations() As Long         'Ditto for violations
            Dim mRival() As String                  '"Rival" is used to mean "candidate".
            Dim mRivalFrequency() As Single         'This gets folded in to a general mFrequency() variable.
            Dim mRivalViolations() As Long          'Three dimensions:  input, rival, constraint.
            Dim mNumberOfRivals() As Long           'One dimention:  a rival-count for each input.
            Dim mMaximumNumberOfRivals As Long
        
        'About constraints:
            Dim mNumberOfConstraints As Long
            Dim mConstraintName() As String
            Dim mAbbrev() As String                 'Used to fit nicely into tableaux.
            Dim mFaithfulness() As Boolean
            
        'File numbers, for reading and writing access.
            Dim mTmpFile As Long                    'This is the one I use most of the time.
            Dim mDocFile As Long                    'A Word macro, hard to use these days, can turn this into nice-looking output.
            Dim mHTMFile As Long
            Dim mTabbedFile As Long
   
   'Parameters of the model.
      Dim mRankingValue() As Double             'Variable name chosen long ago; it now designates *weights* as well, when running the program for MaxEnt."
      Dim mInitialRankingValue() As Double      'User specified, for UG principles.
      Dim mActive() As Boolean                  'Monitor if a constraint does any good
      Dim mSigma() As Double                    'For biased learning in MaxEnt.
      Dim mMu() As Double                       'For biased learning in MaxEnt.
      
   'Variables about the learning data:
      Dim mFrequency() As Single                'The frequency of some particular candidate ("rival") for some particular form; two-dimensional array.
      Dim mFrequencyShare() As Double           'Frequency share in the input file.
      Dim mActualFrequencyShare() As Long       'Frequency share as it actually emerged in stochastic learning
      Dim mActualFrequencyPerInput() As Long    'Summing the latter over all rival candidates.
      
    'KZ: some variables for custom plasticity:
        Dim mCustomPlastFaith() As Double
        Dim mCustomPlastMark() As Double
        Dim mNumberOfLearningStages As Long
      
    'KZ: allows random selection of exemplar:
      Dim mFrequencyInterval() As Double
      
      'Save the number of data presentations so that, when a custom schedule
      ' is use, it can be reported to the user without saving it in OTSoftRememberUserChoices.txt
        Dim mReportedNumberOfDataPresentations  As Long
        
    'For reporting history, we want to track the learning trials (even when nothing is done).
        Dim mLearningTrial As Long
      
    'Learning schedule parameters.
        Dim mTrialsPerLearningStage() As Long
        Dim mblnUseCustomLearningSchedule As Boolean
        
    'Testing the grammar:
        Dim mPercentageGenerated() As Single            'In MaxEnt, this is simply the model probability, which is exact and needn't be sampled.
        Dim mNumberGenerated() As Single
        Dim mTotalNumberOfRivals As Long
        Dim mErrorTerm As Double

    'A priori rankings:
        Dim mUseAPrioriRankings As Boolean
        Dim mNumericalEquivalentOfStrictRanking As Double
      
   'Variables for designating files:
      Dim mSimpleHistoryFile As Long
      Dim mFullHistoryFile As Long
      Dim mCandidateHistoryFile As Long

   'Final details:
        Dim mTimeMarker As Single
        
    'KZ: keeps track of whether or not user has cancelled learning
        Dim mblnProcessing As Boolean
    
    'KZ: in case user cancels algorithm but then runs it again without exiting the form, the output files need
    '   to be reopened.
        Dim blnOutputFilesOpen As Boolean
    
    'Variables to report probability of rankings.
        Dim mThreshold(491) As Double
        Dim mProbability(491) As String
                                                    
    'Variables for presenting data in exact frequencies.
        Private Type LearningDatum
            FormIndex As Long
            RivalIndex As Long
        End Type
        Dim mDataPresentationArray() As LearningDatum
        Dim mTotalFrequency As Long
        
'==================================INTERFACE ITEMS==========================================

Sub Main(NumberOfForms As Long, InputForm() As String, _
    Winner() As String, WinnerFrequency() As Single, WinnerViolations() As Long, _
    NumberOfRivals() As Long, Rival() As String, RivalFrequency() As Single, RivalViolations() As Long, _
    NumberOfConstraints As Long, ConstraintName() As String, Abbrev() As String, _
    TmpFile As Long, DocFile As Long, HTMFile As Long, MyFramework As String)

    Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
    
    Let blnOutputFilesOpen = True 'KZ: the files were opened in Main.
    
    'Localize input parameters as module level variables.
        Let mMyFramework = MyFramework
            'Also, put this on the interface.
                If MyFramework = "StochasticOT" Then
                    Let optStochasticOT.Value = True
                Else
                    Let optMaxEnt.Value = True
                End If
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
       
        
    'Localize the output file numbers
        Let mTmpFile = TmpFile
        Let mDocFile = DocFile
        Let mHTMFile = HTMFile
    
    'Put a caption on the form.
        If mMyFramework = "StochasticOT" Then
            Let GLA.Caption = "OTSoft " + gMyVersionNumber + " - GLA-StochasticOT - " + gFileName + gFileSuffix
        ElseIf mMyFramework = "MaxEnt" Then
            Let GLA.Caption = "OTSoft " + gMyVersionNumber + " - GLA-MaxEnt - " + gFileName + gFileSuffix
        Else
            MsgBox "Program error #34971.  Please report this problem to Bruce Hayes at bhayes@humnet.ucla.edu.  Sorry for the trouble."
        End If
    
    'Put on the interface some memorized values for the most crucial parameters:
        Let txtNumberOfCycles.Text = Trim(Str(gNumberOfDataPresentations))
        Let txtUpperPlasticity.Text = Trim(Str(gCoarsestPlastMark))
        Let txtLowerPlasticity.Text = Trim(Str(gFinestPlastMark))
        Let txtTimesToTestGrammar.Text = Trim(Str(gCyclesToTest))
        Let mnuNegativeWeightsOK.Checked = gNegativeWeightsOK
        Let mnuMagri.Checked = gMagriUpdateRuleInEffect
        
        
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
        
    'If a file with initial rankings exists, give the user the option of using it.
        If Dir(gOutputFilePath + gFileName + "ModelParameters.txt") <> "" Then
            'There is a custom file, so use these values.
                Let mnuUseFullyCustomizedInitialRankingValues.Visible = True
                Let mnuUseFullyCustomizedInitialRankingValues.Checked = True
                Let lblCustomRank.Visible = True
                Let lblCustomRank.Caption = "Using customized initial ranking values from file"
                Let InitialRankingChoice = FullyCustomized
                Let mnuUseDefaultInitialRankingValues.Checked = False
        Else
            'Just use the ordinary values--all 100 in Stochastic OT, all 0 in MaxEnt.
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
        
    'Mark in the relevant menu item whether the user wants to use exact input proportions.
        Select Case gExactProportionsForGLAEtc
            Case True
                Let mnuExactProportions.Checked = True
                Let lblExact.Visible = True
            Case False
                Let mnuExactProportions.Checked = False
                Let lblExact.Visible = False
        End Select
        
    'Mark in the relevant menu item whether the user wants to employ the Magri update rule.
        Select Case gMagriUpdateRuleInEffect
            Case True
                Let mnuMagri.Checked = True
            Case False
                Let mnuMagri.Checked = False
        End Select
        
    'We're ready to go, so let the user see the form.
        GLA.Show
    
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




'---------------------------------Initial Rankings Menu-------------------------------------------

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
        Let lblCustomRank.Caption = "Separate initial rankings for Markedness and Faithfulness"
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
    Let lblCustomRank.Caption = "Using customized initial ranking values from file"
    Let lblCustomRank.Visible = True
    
    'Suppress the other choices.
        Let mnuUsePreviousResultsAsInitialRankingValues.Checked = False
        Let mnuUseSeparateMarkFaithInitialRankings.Checked = False
        Let mnuUseDefaultInitialRankingValues.Checked = False
    
End Sub

Private Sub mnuUsePreviousResultsAsInitialRankingValues_Click()
    
    'Check to see if there really is a file.
        If Dir(gOutputFilePath + gFileName + "MostRecentRankingValues.txt") <> "" Then
            Let mnuUsePreviousResultsAsInitialRankingValues.Checked = True
            Let InitialRankingChoice = ValuesFromPreviousRun
            Let lblCustomRank.Visible = True
            Let lblCustomRank.Caption = "Using result of previous run as initial ranking values"
            'Deactivate other choices.
                Let mnuUseDefaultInitialRankingValues.Checked = False
                Let mnuUseSeparateMarkFaithInitialRankings.Checked = False
                Let mnuUseFullyCustomizedInitialRankingValues.Checked = False
        Else
            MsgBox "Sorry, I can't find the file that stores the ranking values from the previous run.", vbExclamation
        End If
    
End Sub

Private Sub mnuFullyCustomizedInitialRankingValues_Click()

    'Control production of file with fully customized initial ranking values.
        
    'Make the file by calling a function.  If it returns True, then that means
    '   it succeeded, so let the user know it will be used.
        If MakeFileForFullyCustomizedModelParameters("StochasticOT") = True Then
            Let mnuUseFullyCustomizedInitialRankingValues.Visible = True
            Let mnuUseFullyCustomizedInitialRankingValues.Checked = True
            'And let OTSoft know too:
                Let InitialRankingChoice = FullyCustomized
                Let lblCustomRank.Caption = "Using customized initial ranking values from " + gFileName + "ModelParameters.txt"
                Let lblCustomRank.Visible = True
            'Decheck the other choices.
                Let mnuUseDefaultInitialRankingValues.Checked = False
                Let mnuUsePreviousResultsAsInitialRankingValues.Checked = False
                Let mnuUseSeparateMarkFaithInitialRankings.Checked = False
        End If
        
End Sub

Private Sub mnuEditFileForMaxentParameters_Click()

    'Make the file by calling a function.  If it returns True, then that means
    '   it succeeded, so let the user know it will be used.
        If MakeFileForFullyCustomizedModelParameters("Maxent") = True Then
            Let mnuUseFullyCustomizedInitialRankingValues.Visible = True
            Let mnuUseFullyCustomizedInitialRankingValues.Checked = True
            'And let OTSoft know too:
                Let InitialRankingChoice = FullyCustomized
                Let lblCustomRank.Caption = "Using customized learning parameters from file"
                Let lblCustomRank.Visible = True
            'Decheck the other choices.
                Let mnuUseDefaultInitialRankingValues.Checked = False
                Let mnuUsePreviousResultsAsInitialRankingValues.Checked = False
                Let mnuUseSeparateMarkFaithInitialRankings.Checked = False
        End If


End Sub


Private Sub mnuSpecifySeparateMarkFaithInitialRankings_Click()

    'KZ: allow user to pick separate values for markedness vs. faithfulness
    '  in the initial ranking values of constraints.
 
    frmInitialRankings.Show
    
    If blnCustomRankCreated = True Then
        'KZ: if there already are custom ranks in effect, allow user to edit them
        Let frmInitialRankings.txtMark.Text = gCustomRankMark
        Let frmInitialRankings.txtFaith.Text = gCustomRankFaith
    End If  'KZ: (otherwise, defaults are loaded)
    
    Let frmInitialRankings.Visible = True
    
End Sub

Function MakeFileForFullyCustomizedModelParameters(MyFramework As String)

    'Open and edit a file for customized initial ranking values.
    
        Dim ModelParametersFile As Long
        Dim ConstraintIndex As Long
        
        'First, make sure there is a folder for this file, a daughter of the
        '   folder in which the input file is located.
            Call Form1.CreateAFolderForOutputFiles
        
        If Dir(gOutputFilePath + gFileName + "ModelParameters.txt") = "" Then
            'You have to make the file anew.
                Let ModelParametersFile = FreeFile
                Open gOutputFilePath + gFileName + "ModelParameters.txt" For Output As ModelParametersFile
                'Print a header
                    Print #ModelParametersFile, "Constraint"; vbTab; "Initial value";
                    If MyFramework = "MaxEnt" Then
                        Print #ModelParametersFile, vbTab; "Mu"; vbTab; "Sigma";
                    End If
                    Print #ModelParametersFile,
                'Print the constraint names (abbreviated) with a suitable default.
                    For ConstraintIndex = 1 To mNumberOfConstraints
                        If optStochasticOT Then
                            Print #ModelParametersFile, mAbbrev(ConstraintIndex); vbTab; "100"
                        Else
                            Print #ModelParametersFile, mAbbrev(ConstraintIndex); vbTab; "0"
                        End If
                    Next ConstraintIndex
                Close ModelParametersFile
            'If you make it, surely you want to use it.
                Call mnuUseFullyCustomizedInitialRankingValues_Click
        End If

        'Now you know it exists, one way or the other.  Edit it.
            If Dir(gExcelLocation) <> "" Then
                'Shell to Excel.
                    Dim Dummy As Long
                    Let Dummy = _
                        Shell(gExcelLocation + " " + Chr(34) + gOutputFilePath + gFileName + "ModelParameters.txt" + Chr(34), vbNormalFocus)
            Else
                'Whatever Windows says.
                Call UseWindowsPrograms.TryShellExecute(gOutputFilePath + gFileName + "ModelParameters.txt")
            End If
            
        'Report success.
            Let MakeFileForFullyCustomizedModelParameters = True
            
        Exit Function

CheckError:

    MsgBox "Program error:  I was unable to edit a file with initial ranking values.  Please report this bug to Bruce Hayes " + _
        "(bhayes@humnet.ucla.edu), specifying error #11198.", vbCritical
    Let MakeFileForFullyCustomizedModelParameters = False
    
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

Private Sub mnuMagri_Click()
    'Toggle the user's choice.
        If mnuMagri.Checked = False Then
            Let mnuMagri.Checked = True
            Let gMagriUpdateRuleInEffect = True
        Else
            Let mnuMagri.Checked = False
            Let gMagriUpdateRuleInEffect = False
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
            Print #CusLSched, "Trials"; vbTab; "PlastMark"; vbTab; "PlastFaith"; vbTab; "NoiseMark"; vbTab; "NoiseFaith"
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
                "Click Yes to edit OTSoftAuxiliarySoftwareLocations.txt (OTSoft will exit), No to edit your custom learning schedule file with whatever program your computer currently uses to open a .txt file.", vbYesNo + vbExclamation)
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
                    "A priori rankings" + Chr(34) + " menu.", vbExclamation
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


'---------------------------------------MaxEnt Menu-------------------------------------------


Private Sub mnuNegativeWeightsOK_Click()
    'Toggle the user's choice.
        If mnuNegativeWeightsOK.Checked = False Then
            Let mnuNegativeWeightsOK.Checked = True
        Else
            Let mnuNegativeWeightsOK.Checked = False
        End If
End Sub


Private Sub mnuBatchMaxEnt_Click()

    'This is the other, basically deprecated algorithm -- if you want batch, why not Excel or R?
            Let gAlgorithmName = "Maximum Entropy"
            Unload Me
            Call MyMaxEnt.Main(mNumberOfForms, mInputForm(), mWinner(), mWinnerFrequency(), mWinnerViolations(), _
                mNumberOfRivals(), mRival(), mRivalFrequency(), mRivalViolations(), _
                mNumberOfConstraints, mConstraintName(), mAbbrev, _
                mTmpFile, mDocFile, mHTMFile)
                Let Form1.cmdRank.Enabled = True
                Let Form1.cmdFacType.Enabled = True
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
Private Sub mnuCandidateProbabilityHistory_Click()
    If mnuCandidateProbabilityHistory.Checked = False Then
        Let mnuCandidateProbabilityHistory.Checked = True
    Else
        Let mnuCandidateProbabilityHistory.Checked = False
    End If
End Sub
Private Sub mnuGaussianPrior_Click()
    If mnuGaussianPrior.Checked = False Then
        Let mnuGaussianPrior.Checked = True
    Else
        Let mnuGaussianPrior.Checked = False
    End If
End Sub

Sub mnuExactProportions_Click()

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


'------------------------------Rest of GLA Interface--------------------------------
Private Sub cmdRun_Click()
    
    'Make sure the user has picked a framework.
        If optStochasticOT.Value = False And optMaxEnt.Value = False Then
            MsgBox "Please choose a framework first."
            Exit Sub
        End If
    
    Call RunGLA

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


Private Sub mnu10Runs_Click()
    Call MultipleRuns(10)
End Sub
Private Sub mnu100Runs_Click()
    Call MultipleRuns(100)
End Sub
Private Sub mnu1000Runs_Click()
    Call MultipleRuns(1000)
End Sub

Sub MultipleRuns(RunCount As Long)
    
    Dim RunNumber As Long
    Dim StoreFileName As String
    Dim RunIndex As Long, ConstraintIndex As Long, FormIndex As Long, RivalIndex As Long
    
    'Purge the current copy, if any, of the CollateRuns.txt file.
        Open gInputFilePath + "\CollateRuns.txt" For Output As #111
        Close #111
    'Do as many runs as requested.
        For RunIndex = 1 To RunCount
            'Remember the old file name.
                Let StoreFileName = gFileName
            'Create a new file name with an index prepended.
                Let gFileName = Trim(Str(RunIndex)) + gFileName
            
            'Since you got here via GLA interface choices, it's ok just to click the button.
                Call Form1.Rank
            'And then, in essence, click the button on the GLA interface, which calls this:
                Call RunGLA
            'print    'Temp:  multiple runs
                Open gInputFilePath + "\CollateRuns.txt" For Append As #111
                For ConstraintIndex = 1 To mNumberOfConstraints
                  Print #111, "G"; vbTab; Trim(Str(RunIndex));
                  Print #111, vbTab; mConstraintName(ConstraintIndex);
                  Print #111, vbTab; mRankingValue(ConstraintIndex)
                Next ConstraintIndex
                For FormIndex = 1 To mNumberOfForms
                    For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                        Print #111, "O"; vbTab;
                        Print #111, Trim(Str(RunIndex)); vbTab;
                        Print #111, FormIndex; vbTab;
                        Print #111, mInputForm(FormIndex); vbTab;
                        Print #111, RivalIndex; vbTab;
                        Print #111, mRival(FormIndex, RivalIndex); vbTab;
                        If RivalIndex = 0 Then
                            Print #111, mWinnerFrequency(FormIndex); vbTab;
                        Else
                            Print #111, mRivalFrequency(FormIndex, RivalIndex); vbTab;
                        End If
                        'Print #111, mActualFrequencyShare(FormIndex, RivalIndex); vbtab;
                        'Print #111, mNumberGenerated(FormIndex, RivalIndex); vbtab;
                        'Print #111, PercentageInInput(FormIndex, RivalIndex); vbtab;
                        Print #111, mPercentageGenerated(FormIndex, RivalIndex)
                      Next RivalIndex
                Next FormIndex
            
            Let gFileName = StoreFileName
            Close #111
            
        Next RunIndex
    

End Sub


'================================MAIN CALCULATION ROUTINE====================================

Private Sub RunGLA()

    'This is the primary routine for this form, which calls all other routines needed for running the GLA.
    
    
    'If the user wants mu and sigma, make sure that the file gets read.
        If mnuGaussianPrior.Checked Then
            Let InitialRankingChoice = FullyCustomized
        End If
    
    'Since the user has clicked a Rank button, (s)he probably wants the settings saved.
        Call Form1.SaveUserChoices
  
    'Dimension arrays needed by the algorithm, according to the size of the problem:
        'Properties of the learning data:
            ReDim mFrequencyShare(mNumberOfForms, mMaximumNumberOfRivals)
            ReDim mActualFrequencyShare(mNumberOfForms, mMaximumNumberOfRivals)
            ReDim mActualFrequencyPerInput(mNumberOfForms)
            'KZ: third dimension gives the lower(inclusive) and upper(exclusive)
            '   bounds of the interval in [0,1] assigned to that form.
                ReDim mFrequencyInterval(mNumberOfForms, mMaximumNumberOfRivals, 1)
                ReDim mFrequency(mNumberOfForms, mMaximumNumberOfRivals)
        
        'Properties of constraints:
            ReDim mRankingValue(mNumberOfConstraints)
            ReDim mInitialRankingValue(mNumberOfConstraints)
            ReDim mActive(mNumberOfConstraints)
            ReDim mSigma(mNumberOfConstraints)
            ReDim mMu(mNumberOfConstraints)
            ReDim mFaithfulness(mNumberOfConstraints) 'KZ
        
    'KZ: the variable mblnProcessing keeps track of whether or not the algorithm is running:
    If mblnProcessing = True Then
        
        'KZ: when the algorithm is running, clicking this button again stops it
            Let mblnProcessing = False 'KZ: return button to "not-running" state and shut down the learner
        
    Else    'KZ: otherwise, run the algorithm
    
        'KZ: let the user know this is now the cancel button:
            cmdRun.Caption = "Cancel"
            mblnProcessing = True
        
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
                'The HTML file:
                    Let mHTMFile = FreeFile
                    Open gOutputFilePath + "ResultsFor" + gFileName + ".htm" For Output As #mHTMFile
                    Call PrintTableaux.InitiateHTML(mHTMFile)
            End If
        
        Call ObtainInformationFromMainWindow
        
        'Clear the progress window.
            pctProgressWindow.Cls
        'Start timing.
            Let mTimeMarker = Timer
        'Report progress.
            pctProgressWindow.Print "Learning..."
        
        'GLAPreliminaries:
            'If all goes will in the preliminary operations, GLAPreliminaries() will return True, and you can continue.
                If GLAPreliminaries() = False Then
                    'Something went wrong in the preliminaries.  Go back to ur-state.
                        Let mblnProcessing = False
                    'Reactivate the various buttons that start things off, and give up.
                        Let Form1.cmdRank.Enabled = True
                        Let Form1.cmdFacType.Enabled = True
                        Let cmdRun.Caption = "&Run GLA"
                    Exit Sub
                End If
            
        Call GLACore    'KZ: GLACore checks periodically to see if the button
                        'has been clicked again (to cancel). If user cancels,
                        'mblnProcessing gets set to False.
                        
        'KZ: don't bother with this part if user has cancelled:
            If mblnProcessing = True Then
                'Different parameter name for the two theories.
                    If optMaxEnt Then
                        Call PrintGLAResults(mRankingValue(), "Weights")
                    Else
                        Call PrintGLAResults(mRankingValue(), "Ranking Values")
                    End If
                'If you're using a priori rankings, look them up and implement them  as initial values.
                    If mnuAPrioriRankings.Checked Then
                        Call Form1.PrintOutTheAprioriRankings(mTmpFile, mDocFile, mHTMFile)
                    End If
                'GLATestGrammar returns a value, namely the degree of error.
                '   This permits it to be used for hill-climbing learning.
                '   The second parameter is True if one wants a printed report.
                '   KZ: GLATestGrammar also checks for cancellation.
                    Let mErrorTerm = GLATestGrammar(mRankingValue(), True, gCyclesToTest)
                
                If optMaxEnt = False Then
                    Call PrepareApproximateTableaux     'This also prints a table
                End If
                                                    '  converting ranking values to probability.
                Call PrintFinalDetails
            End If
        
        'Close output files.
            Close #mTmpFile
            Close #mDocFile
            Print #mHTMFile, "</BODY>"
            Close #mHTMFile

            Let blnOutputFilesOpen = False 'KZ
            
        'KZ: get rid of "cancel" caption:
            Let cmdRun.Caption = "&Run GLA"
        
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
                Close #mCandidateHistoryFile
        
        End If
        
    End If  'Is mblnProcessing true?  (i.e. Rank or Cancel)

End Sub


Sub ObtainInformationFromMainWindow()
  
    'Other parameters are obtained from the interface of the main OTSoft window.
    '   This routine only grabs the parameters from the GLA window, and therefore
    '   has to run after the user has clicked Run.
    
    'Ask the GLA interface how many trials are wanted.
        Let gCyclesToTest = Val(GLA.txtTimesToTestGrammar.Text)
         
    'Determine if apriori rankings are in effect
        If mnuDoGLAWithAPrioriRankings.Checked = True Then
            Let mUseAPrioriRankings = True
        Else
            Let mUseAPrioriRankings = False
        End If

End Sub



Function GLAPreliminaries() As Boolean

   'Execute preliminary actions needed by the Gradual Learning Algorithm.
   
    'Variables of convenience for reading the file with initial rankings.
        Dim MyLine As String, ModelParametersFile As Long
        Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
    
    'To assign plasticity, initial rankings, and noise differently to Markedness
    '   and Faithfulness constraints, you need to look up their status.
        For ConstraintIndex = 1 To mNumberOfConstraints
            Let mFaithfulness(ConstraintIndex) = Form1.FaithfulnessConstraint(mConstraintName(ConstraintIndex))
        Next ConstraintIndex
        
    'Note:  important to run Randomize just once; else a subtle bug.
        Randomize
        
        If optStochasticOT Then
            Call PrintAHeader("Gradual Learning Algorithm -- Stochastic OT")
        Else
            Call PrintAHeader("Gradual Learning Algorithm -- MaxEnt")
        End If
        
    'This is a very tricky thing; creates bugs.  I *think* it is ok now (1/2/26)
        Call InstallWinnerAsRival
        
        'The finder of identical violations happens to be located in another module:
            If HuntForDuplicateViolations(mNumberOfForms, mInputForm(), mNumberOfRivals, mRival(), mRivalViolations(), mDocFile, mTmpFile, mHTMFile) = True Then
                Call AmalgamateCandidates(mNumberOfForms, mMaximumNumberOfRivals, mNumberOfRivals(), mRival(), mRivalViolations(), mFrequency(), mNumberOfConstraints)
            End If
        
        Call FrequencyThresholds
        If InitialRankingValues = False Then
            Let GLAPreliminaries = False
            Exit Function
        End If

    'If you're presenting data in exact proportions, form a chart for this purpose.
        If mnuExactProportions.Checked = True Then
            Call FormDataPresentationArray
        End If
      
    'Set up files to trace the history of ranking/weighting.
    '   Not done if its just running more tests on an existing grammar.
        If gNumberOfDataPresentations > 0 Or mReportedNumberOfDataPresentations > 0 Then
            
            If mnuGenerateHistory.Checked = True Then
                Call OpenSimpleHistoryFile
            End If
            If mnuFullHistory.Checked = True Then
                Call OpenFullHistoryFile
            End If
            'This one must be ordered after the setting of initial ranking values, since it starts with the assessment of the initial-value
            '   grammar.
            'Moreover, it is only possible, for now, with MaxEnt.
                If mnuCandidateProbabilityHistory.Checked = True Then
                    If optStochasticOT Then
                        MsgBox "I'm sorry, but for computational reasons it is difficult to track the probability of candidates in Stochastic OT. Click again to resume."
                        Let mnuCandidateProbabilityHistory.Checked = False
                    Else
                        Call OpenCandidateHistoryFile
                    End If
                End If
            
        End If
    
    'If you're using a priori rankings, look them up and implement them initial values.
        If mUseAPrioriRankings = True Then
            'Vet the value that implements an a priori ranking numerically.
                If GoodDecimal(txtValueThatImplementsAPrioriRankings.Text) = False Then
                    MsgBox "Please enter a valid number in the box labeled " + _
                        Chr(34) + "Constraints ranked a priori must differ by" + Chr(34), vbExclamation
                    Let GLAPreliminaries = False
                    Exit Function
                End If
            Let mNumericalEquivalentOfStrictRanking = Val(txtValueThatImplementsAPrioriRankings)
            'ReadAPrioriRankingsAsTable is a boolean function, which returns False if
            '   it failed to do its job.
            If APrioriRankings.ReadAPrioriRankingsAsTable(mNumberOfConstraints, mAbbrev) = True Then
                'This must happen after the history files are opened, since this routine records
                '   itself when history is being taken.
                    Call AdjustAPrioriRankings_Up
            Else
                Let Form1.mnuConstrainAlgorithmsByAPrioriRankings.Checked = False
            End If
            
        End If
        
    'Determine the schedule for the learning parameters.  If an error, this will return as False.
        If DetermineLearningSchedule() = False Then
            Let GLAPreliminaries = False
            Exit Function
        End If

    'Record the initial ranking values, as part of notating history of the run.
        Call RecordInitialRankingValues
      
    'All is well, so set value as True and exit.
        Let GLAPreliminaries = True

End Function


Function InitialRankingValues() As Boolean

    'Set the ranking value of every constraint at Boersma's arbitrary value of 100.
    '   KZ: or, if user wants, use custom values for faithfulness and markedness
    '   or from file--useful for running more.
    
    'In the case of MaxEnt, we are using 0, with the same proviso.
    
    Dim ModelParametersFile As Long
    Dim MyLine As String
    Dim ConstraintIndex As Long
    
    'We assume success unless a problem is found.
        Let InitialRankingValues = True
    
    Select Case InitialRankingChoice
        Case AllSame
            If optMaxEnt Then
                'MaxEnt usually starts at zero.
                    For ConstraintIndex = 1 To mNumberOfConstraints
                        Let mRankingValue(ConstraintIndex) = 0
                    Next ConstraintIndex
            Else
                'The classical Boersmian value of 100.
                    For ConstraintIndex = 1 To mNumberOfConstraints
                        Let mRankingValue(ConstraintIndex) = 100
                    Next ConstraintIndex
            End If
        Case MarkednessFaithfulness
            'Invoke user's choice for markedness and Faithfulness.
                For ConstraintIndex = 1 To mNumberOfConstraints
                    If mFaithfulness(ConstraintIndex) = True Then
                        Let mRankingValue(ConstraintIndex) = gCustomRankFaith
                    Else
                        Let mRankingValue(ConstraintIndex) = gCustomRankMark
                    End If
                Next ConstraintIndex
        Case FullyCustomized, ValuesFromPreviousRun
            'Read the file with customized choices, or choices saved from last run.
                Let ModelParametersFile = FreeFile
                'Custom values or old values?  Open appropriate file.
                    Select Case InitialRankingChoice
                        Case FullyCustomized
                            'Don't ever try to open a nonexistent file.
                                If Dir(gOutputFilePath + gFileName + "ModelParameters.txt") <> "" Then
                                    Open gOutputFilePath + gFileName + "ModelParameters.txt" For Input As ModelParametersFile
                                Else
                                    MsgBox "Sorry, I can't find the file containing your initial rankings.  It is supposed to be at " + _
                                        gOutputFilePath + gFileName + "ModelParameters.txt.  Click OK to continue.", vbExclamation
                                    Let InitialRankingValues = False
                                    Exit Function
                                End If
                        Case ValuesFromPreviousRun
                            'Don't ever try to open a nonexistent file.
                                If Dir(gOutputFilePath + gFileName + "MostRecentRankingValues.txt") <> "" Then
                                    Open gOutputFilePath + gFileName + "MostRecentRankingValues.txt" For Input As ModelParametersFile
                                Else
                                    MsgBox "Sorry, I can't find the file containing your most recent rankings.  It is supposed to be at " + _
                                        gOutputFilePath + gFileName + "MostRecentRankingValues.txt.  Click OK to continue.", vbExclamation
                                    Let InitialRankingValues = False
                                    Exit Function
                                End If
                    End Select
                'Read the values off the file.
                    'Skip the header, if there is one.
                        Line Input #ModelParametersFile, MyLine
                        Let MyLine = Trim(MyLine)
                        If LCase(Left(MyLine, 10)) <> "constraint" Then
                            'There is probably no header.
                                GoTo NoHeaderPoint
                        End If
                    For ConstraintIndex = 1 To mNumberOfConstraints
                        'Skip the header line and read a content line.
                            Line Input #ModelParametersFile, MyLine
NoHeaderPoint:
                        'The initial ranking value.
                            Let MyLine = s.Residue(MyLine)
                            Let mRankingValue(ConstraintIndex) = Val(Trim(s.Chomp(MyLine)))
                        'We will also read a mu or sigma even if they are not there; harmlessly produces a zero.
                            Let MyLine = s.Residue(MyLine)
                            Let mMu(ConstraintIndex) = Val(Trim(s.Chomp(MyLine)))
                            Let MyLine = s.Residue(MyLine)
                            Let mSigma(ConstraintIndex) = Val(Trim(s.Chomp(MyLine)))
                    Next ConstraintIndex
                Close ModelParametersFile
                
    End Select
    
    
    'Ad hoc desparate stupidity.  But this results in matching the weights and probabilities derived in Excel.  What gives?
    '     For ConstraintIndex = 1 To mNumberOfConstraints
    '        Let mSigma(ConstraintIndex) = mSigma(ConstraintIndex) * Sqr(2)
    '     Next ConstraintIndex
    
    'Save the initial ranking values for future reporting.
        For ConstraintIndex = 1 To mNumberOfConstraints
            Let mInitialRankingValue(ConstraintIndex) = mRankingValue(ConstraintIndex)
        Next ConstraintIndex
    
End Function


Sub OpenSimpleHistoryFile()

    'History of ranking values, for an Excel file.

    Dim ConstraintIndex As Long
    
    'First, make sure there is a folder for these files, a daughter of the
    '   folder in which the input file is located.
        Call Form1.CreateAFolderForOutputFiles
        
    'Open the file:
        Let mSimpleHistoryFile = FreeFile
        Open gOutputFilePath + gFileName + "History" + ".txt" For Output As #mSimpleHistoryFile
        
    'Print a header:
        'The trial number.
            Print #mSimpleHistoryFile, "Trial";
        'The constraint names.
            For ConstraintIndex = 1 To mNumberOfConstraints
                Print #mSimpleHistoryFile, Chr$(9); mAbbrev(ConstraintIndex);
            Next ConstraintIndex
        'End of line
            Print #mSimpleHistoryFile,

End Sub

Sub OpenFullHistoryFile()

    'Complete detailed history of how ranking was done, including the input and the two compared candidates.
    
       On Error GoTo OpenFullHistoryFileErrorPoint
    
        Dim ConstraintIndex
    
    'Open file, and print a header:

        'First, make sure there is a folder for these files, a daughter of the
        '   folder in which the input file is located.
            Call Form1.CreateAFolderForOutputFiles
        
        'Open the file:
            Let mFullHistoryFile = FreeFile
            Open gOutputFilePath + gFileName + "FullHistory.txt" For Output As #mFullHistoryFile
        'Print a header:
            Print #mFullHistoryFile, "Trial #"; vbTab; "Input"; vbTab; "Generated"; vbTab; "Heard"; vbTab;
            'Header for each constraint:  delta, then new value.
                For ConstraintIndex = 1 To mNumberOfConstraints
                    Print #mFullHistoryFile, mAbbrev(ConstraintIndex); vbTab;
                Next ConstraintIndex
            'End of line
                Print #mFullHistoryFile,
                
        Exit Sub
                
OpenFullHistoryFileErrorPoint:

    Select Case Err.Number
        Case 70
            MsgBox "Error.  I conjecture that the file " + gOutputFilePath + gFileName + "FullHistory" + ".txt" + _
                " is already open.  Please close this file, then run OTSoft again.", vbExclamation
            End
        Case Else
            MsgBox "Program error.  For help please contact me at bhayes@humnet.ucla.edu, enclosing your input file and specifying error #14897.", vbCritical
    End Select

End Sub

Sub OpenCandidateHistoryFile()

    'Trace the probability of the candidates over the course of learning.
    'Here, we open the file, provide a header, and record the values from the initial weights.
    
       On Error GoTo OpenCandidateHistoryFileErrorPoint
    
       Dim FormIndex As Long, RivalIndex As Long
    
      'First, make sure there is a folder for these files, a daughter of the folder in which the input file is located.
          Call Form1.CreateAFolderForOutputFiles
    
      'Open the file:
          Let mCandidateHistoryFile = FreeFile
          Open gOutputFilePath + gFileName + "HistoryOfCandidateProbabilities.txt" For Output As #mCandidateHistoryFile
      'Print a header:
          Print #mCandidateHistoryFile, "Trial #";
          'This is awkward, but every candidate gets concatenated with its input.  Probably better than two rows?
          For FormIndex = 1 To mNumberOfForms
              For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                  Print #mCandidateHistoryFile, vbTab; mInputForm(FormIndex) + " -> " + mRival(FormIndex, RivalIndex);
              Next RivalIndex
          Next FormIndex
          Print #mCandidateHistoryFile,
          
       'This must be redimmed early, for on-line reporting.
           ReDim mPercentageGenerated(mNumberOfForms, mMaximumNumberOfRivals)
    
       'Print the predictions of the starting point null grammar, with zero-weighted constrainst
            'First, indicate that this is the initial weightings.
                Print #mCandidateHistoryFile, "(initial)";
            'Grab the predictions for the initial weights.
                Call GenerateMaxEntPredictions
            'Print them.
                For FormIndex = 1 To mNumberOfForms
                    For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                        Print #mCandidateHistoryFile, vbTab; mPercentageGenerated(FormIndex, RivalIndex);
                    Next RivalIndex
                Next FormIndex
                Print #mCandidateHistoryFile,
            
       Exit Sub
                
OpenCandidateHistoryFileErrorPoint:

    Select Case Err.Number
        Case 70
            MsgBox "Error.  I conjecture that the file " + gOutputFilePath + gFileName + "HistoryOfCandidateProbabilities.txt" + _
                " is already open.  Please close this file, then run OTSoft again.", vbExclamation
            End
        Case Else
            MsgBox Err.Description
            MsgBox "Program error.  For help please contact me at bhayes@humnet.ucla.edu, enclosing your input file and specifying error #14888.", vbCritical
    End Select

End Sub



Sub FrequencyThresholds()

   'Forms must be input to the GLA in frequencies that match their real-world counts.
   '  We can do this by keep track of proportions.
   '  This is done on a candidate-by-candidate basis.
   
    Dim FormIndex As Long, RivalIndex As Long
    Dim TotalFrequencies As Double
    Dim PreviousUpperBoundOfInterval As Double

   'Loop through all forms and rivals, and find the sum of the frequencies of all candidates.
        Let TotalFrequencies = 0
        For FormIndex = 1 To mNumberOfForms
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

Sub AmalgamateCandidates(NumberOfForms As Long, MaximumNumberOfRivals As Long, NumberOfRivals() As Long, Rivals() As String, RivalViolations() As Long, Frequency() As Single, _
    NumberOfConstraints As Long)

    'If you've reached this point, there are candidates that have the very same violations.  Deal with this by making them the same candidate; the candidate is given a name consisting
    '   of its ingredients separated by &.
        
    'Variables:
        'Store the amalgamated arrays.
            Dim TempRivalViolations() As Long
            Dim TempRivals() As String
            Dim TempFrequency() As Double
            Dim TempNumberOfRivals() As Long
        'Flag:
            Dim SameViolationsFlag As Boolean
        'Indices:
            Dim FormIndex As Long, RivalIndex As Long, CheckMeRivalIndex As Long, ConstraintIndex As Long
        
    'Redimension the arrays.
        ReDim TempRivalViolations(NumberOfForms, MaximumNumberOfRivals, NumberOfConstraints)
        ReDim TempRivals(NumberOfForms, MaximumNumberOfRivals)
        ReDim TempFrequency(NumberOfForms, MaximumNumberOfRivals)
        ReDim TempNumberOfRivals(NumberOfForms)
    
   For FormIndex = 1 To NumberOfForms
        'The first rival qualifies no matter what.
            Let TempNumberOfRivals(FormIndex) = 0
            Let TempRivals(FormIndex, 0) = Rivals(FormIndex, 0)
            Let TempFrequency(FormIndex, 0) = Frequency(FormIndex, 0)
            For ConstraintIndex = 1 To NumberOfConstraints
                Let TempRivalViolations(FormIndex, 0, ConstraintIndex) = RivalViolations(FormIndex, 0, ConstraintIndex)
            Next ConstraintIndex
        
        'Filter the remaining rivals for non-duplicatehood, one by one.
          For RivalIndex = 1 To NumberOfRivals(FormIndex)
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
                                'This is not a new rival.  Make an amalgamated candidate.  You have the violations already, so just construct the name and augment the frequency.
                                    Let TempRivals(FormIndex, CheckMeRivalIndex) = TempRivals(FormIndex, CheckMeRivalIndex) + " & " + Rivals(FormIndex, RivalIndex)
                                    Let TempFrequency(FormIndex, CheckMeRivalIndex) = TempFrequency(FormIndex, CheckMeRivalIndex) + Frequency(FormIndex, RivalIndex)
                                    'Don't prove it to be a duplicate over and over, just once is ok.
                                        GoTo ResumePoint
                              End If
                     Next CheckMeRivalIndex
                'If you've gotten this far, then this is a new rival.  Install it.
                    Let TempNumberOfRivals(FormIndex) = TempNumberOfRivals(FormIndex) + 1
                    Let TempRivals(FormIndex, TempNumberOfRivals(FormIndex)) = Trim(Rivals(FormIndex, RivalIndex))
                    Let TempFrequency(FormIndex, TempNumberOfRivals(FormIndex)) = Frequency(FormIndex, RivalIndex)
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
                Let mFrequency(FormIndex, RivalIndex) = TempFrequency(FormIndex, RivalIndex)
                For ConstraintIndex = 1 To NumberOfConstraints
                    Let mRivalViolations(FormIndex, RivalIndex, ConstraintIndex) = TempRivalViolations(FormIndex, RivalIndex, ConstraintIndex)
                Next ConstraintIndex
            Next RivalIndex
        Next FormIndex


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
        For RivalIndex = 1 To mNumberOfRivals(FormIndex)
            'Check that this frequency is an integer.  Report -1 and exit if not.
                If mFrequency(FormIndex, RivalIndex) <> Int(mFrequency(FormIndex, RivalIndex)) Then
                    MsgBox "Caution:  you've requested exact frequency matching, but OTSoft can do this only if frequencies are all whole numbers.  Please fix your input file first if you want to use this method.  When you click ok, OTSoft will proceed with inexact frequency matching.", vbExclamation
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
        Dim DebugFile As Long
        Let DebugFile = FreeFile
        Open gInputFilePath + "debugGLADataPrestationArray.txt" For Output As #DebugFile
        For FrequencyIndex = 1 To mTotalFrequency
            Print #DebugFile, FrequencyIndex; vbTab; mDataPresentationArray(FrequencyIndex).FormIndex;
            Print #DebugFile, vbTab; mDataPresentationArray(FrequencyIndex).RivalIndex
        Next FrequencyIndex
        Close #DebugFile

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
                        "Click OK to return to the GLA screen.", vbExclamation
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
                                gOutputFilePath + gFileName + "CustomLearningSchedule.txt.  Click OK to continue.", vbExclamation
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
                            ReDim Preserve mCustomPlastMark(mNumberOfLearningStages)
                            ReDim Preserve mCustomPlastFaith(mNumberOfLearningStages)
                            ReDim Preserve CustomNoiseMark(mNumberOfLearningStages)
                            ReDim Preserve CustomNoiseFaith(mNumberOfLearningStages)
                        'Chomp up the line, recording the values.  Warn the user and
                        '   exit if there are any blanks.
                            Let Buffer = s.Chomp(MyLine)
                                If Trim(Buffer) = "" Then GoTo BlankCell
                                Let mTrialsPerLearningStage(mNumberOfLearningStages) = Val(Buffer)
                                Let MyLine = s.Residue(MyLine)
                            Let Buffer = s.Chomp(MyLine)
                                If Trim(Buffer) = "" Then GoTo BlankCell
                                Let mCustomPlastMark(mNumberOfLearningStages) = Val(Buffer)
                                Let MyLine = s.Residue(MyLine)
                            Let Buffer = s.Chomp(MyLine)
                                If Trim(Buffer) = "" Then GoTo BlankCell
                                Let mCustomPlastFaith(mNumberOfLearningStages) = Val(Buffer)
                                Let MyLine = s.Residue(MyLine)
                            Let Buffer = s.Chomp(MyLine)
                                If Trim(Buffer) = "" Then GoTo BlankCell
                                Let CustomNoiseMark(mNumberOfLearningStages) = Val(s.Chomp(MyLine))
                            Let Buffer = s.Residue(MyLine)
                                If Trim(Buffer) = "" Then GoTo BlankCell
                                Let CustomNoiseFaith(mNumberOfLearningStages) = Val(s.Residue(MyLine))
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
            ReDim Preserve mCustomPlastMark(4)
            ReDim Preserve mCustomPlastFaith(4)
            ReDim Preserve CustomNoiseMark(4)
            ReDim Preserve CustomNoiseFaith(4)
            
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
                    'Let mCustomPlastMark(1) = Val(txtUpperPlasticity)  'High noise, initial boost.
                    Let mCustomPlastMark(1) = Val(txtUpperPlasticity)  'Low noise, recover ranking for ties
                    Let mCustomPlastMark(4) = Val(txtLowerPlasticity)
                'Interpolate other two geometrically.
                    Let mCustomPlastMark(2) = (mCustomPlastMark(1) * mCustomPlastMark(1) * mCustomPlastMark(4)) ^ (1 / 3)
                    Let mCustomPlastMark(3) = (mCustomPlastMark(1) * mCustomPlastMark(4) * mCustomPlastMark(4)) ^ (1 / 3)
            'Noise:
                'Let CustomNoiseMark(1) = 2     'Initial boost
                For LearningStageIndex = 1 To 4
                    Let CustomNoiseMark(LearningStageIndex) = 2
                Next LearningStageIndex
            'Assign remaining material, which is invariant across the five stages.
                For LearningStageIndex = 1 To 4
                    'All stages have same number of trials.
                        Let mTrialsPerLearningStage(LearningStageIndex) = MyTrialsPerLearningStage
                    'Faithfulness plasticity same as Markedness.
                        Let mCustomPlastFaith(LearningStageIndex) = mCustomPlastMark(LearningStageIndex)
                    'Faithfulness noise same as Markedness.
                        Let CustomNoiseFaith(LearningStageIndex) = CustomNoiseMark(LearningStageIndex)
                Next LearningStageIndex
                
            'Assign the right value to the old variables, permitting them to be remembered.
                Let gCoarsestPlastMark = mCustomPlastMark(1)
                Let gFinestPlastMark = mCustomPlastMark(4)

            
    End If      'Learning schedule from file or from interface?
    
    'All is well, so return True.
        Let DetermineLearningSchedule = True
        Exit Function
        
BlankCell:
    'Go here if there's a problem with the learning schedule file.
        MsgBox "Sorry, your input file for the learning schedule contains one or more blank cells.  Please fix before continuing.", vbExclamation
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
        'Debug:
            If mWinnerFrequency(FormIndex) > 0 Then Stop
        Let mRival(FormIndex, 0) = mWinner(FormIndex)
        For ConstraintIndex = 1 To mNumberOfConstraints
            Let mRivalViolations(FormIndex, 0, ConstraintIndex) = mWinnerViolations(FormIndex, ConstraintIndex)
        Next ConstraintIndex
        For RivalIndex = 0 To mNumberOfRivals(FormIndex)
            Let mFrequency(FormIndex, RivalIndex) = mRivalFrequency(FormIndex, RivalIndex)
        Next RivalIndex
    Next FormIndex
    
End Sub


Sub AdjustAPrioriRankings_Up()

    'The idea is to make every pair of constraints in an
    '   a priori relation at least 20 units apart.
    
    'This routine has no danger of running forever, because the a
    '   priori rankings have been pre-vetted.
    
    'It is called initially, and also whenever a constraint is raised.
    
    Dim KeepGoing As Boolean    'If you change anything, keep looking to detect further necessary changes.
    Dim Margin As Double        'The difference in ranking values of two constraints:  enough?
    Dim OuterConstraintIndex As Long, InnerConstraintIndex As Long
    Dim HistoryConstraintIndex As Long
    
    Do
        Let KeepGoing = False
        For OuterConstraintIndex = 1 To mNumberOfConstraints
            For InnerConstraintIndex = 1 To mNumberOfConstraints
                If gAPrioriRankingsTable(OuterConstraintIndex, InnerConstraintIndex) = True Then
                    Let Margin = mRankingValue(OuterConstraintIndex) - mRankingValue(InnerConstraintIndex)
                    'Amazingly, the addition of x to y in Visual Basic does not necessarily
                    '   make x + y greater than y by a margin of x.
                    'Let's try to fix this, by saying not x, but ever so slightly less
                    '   than x.  Sheesh.
                    If Margin < mNumericalEquivalentOfStrictRanking - 0.0001 Then
                        Let mRankingValue(OuterConstraintIndex) = mRankingValue(InnerConstraintIndex) + _
                            mNumericalEquivalentOfStrictRanking
                        Let KeepGoing = True
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
    '        Print #f, mRankingValue(OuterConstraintIndex)
    '    Next OuterConstraintIndex
    '    Close f

End Sub

Sub AdjustAPrioriRankings_Down()

    'The idea is to make every pair of constraints in an a priori relation at least 20 units apart.
    'This routine is called whenever a constraint is lowered.
    
    Dim Margin As Double
    Dim KeepGoing As Boolean
    Dim OuterConstraintIndex As Long, InnerConstraintIndex As Long
    Dim HistoryConstraintIndex As Long
    
    Do
        Let KeepGoing = False
        For OuterConstraintIndex = 1 To mNumberOfConstraints
            For InnerConstraintIndex = 1 To mNumberOfConstraints
                If gAPrioriRankingsTable(OuterConstraintIndex, InnerConstraintIndex) = True Then
                    Let Margin = mRankingValue(OuterConstraintIndex) - mRankingValue(InnerConstraintIndex)
                    If Margin < mNumericalEquivalentOfStrictRanking Then
                        Let mRankingValue(InnerConstraintIndex) = mRankingValue(OuterConstraintIndex) - _
                            mNumericalEquivalentOfStrictRanking
                        Let KeepGoing = True
                    End If
                End If
            Next InnerConstraintIndex
        Next OuterConstraintIndex
        If KeepGoing = False Then Exit Do
    Loop
    
    'Debug this routine:
    '    Dim f As Long
    '    Let f = FreeFile
    '    Open gOutputFilePath + "DebugAPrioriRankingValuesFor" + gFileName + ".txt" For Output As f
    '    For OuterConstraintIndex = 1 To mNumberofconstraints
    '        Print #f, mabbrev(OuterConstraintIndex); vbtab;
    '        Print #f, mRankingValue(OuterConstraintIndex)
    '    Next OuterConstraintIndex
    '    Close f

End Sub

Sub RecordInitialRankingValues()

    'If you're reporting progress, record the initial ranking values.
    
    Dim ConstraintIndex As Long
    
    If mnuFullHistory.Checked = True Then
        Print #mFullHistoryFile, "(Initial)"; vbTab; vbTab; vbTab; vbTab;
        For ConstraintIndex = 1 To mNumberOfConstraints
            Print #mFullHistoryFile, mRankingValue(ConstraintIndex); vbTab;
        Next ConstraintIndex
        Print #mFullHistoryFile,
    End If

End Sub

Sub GLACore()

    'The Gradual Learning Algorithm of Paul Boersma.
    'Depending on user choice, it can be applied to MaxEnt constraint weights, following Jaeger (2004)
    'The latter gives us MaxEnt "on line", learning item-by-item.
   
    'For both, the key actions are determined by picking an empirical form at random (from the observed distribution) and
    '   generating a form at random (from the current grammar's distribution, whether this be Stochastic OT or MaxEnt).
    
    'For selected empirical form:
        Dim SelectedExemplarForm As Long
        Dim SelectedObservedRival As Long
    
    'Learning variables for Stochastic OT:
        Dim LocalRankingValue() As Double
        Dim LocalSlotFiller() As Long
        Dim TotalTried() As Long
      
        ReDim LocalRankingValue(mNumberOfConstraints)
        ReDim LocalSlotFiller(mNumberOfConstraints)
        ReDim SlotFiller(mNumberOfConstraints)
        ReDim TotalTried(mNumberOfForms, mMaximumNumberOfRivals)
      
        Dim LocalWinner As Long
        Dim WorstSoFar As Double

    'The degree of change made in response to mismatched:
        Dim PlastMark As Double
        Dim PlastFaith As Double
        Dim NoiseMark As Double  'KZ: default is 2, but can be customized by user
        Dim NoiseFaith As Double 'KZ: default is 2, but can be customized by user

    'Indices etc.:
        Dim CycleIndex As Long
        Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long, OtherConstraintIndex As Long
        Dim HistoryConstraintIndex As Long
        Dim i As Long, j As Long
        Dim Swappant As Double, SwapInt As Long
      
        Dim DummyCounter As Long     'KZ: every time gets to 200 it checks if the user pressed cancel.
        Dim ProgressCounter As Long
                                    
    'Variables for selecting learning data at random.
        Dim RandomNumber As Double        'KZ
        Dim blnExemplarChosen As Boolean  'KZ
      
    'Variable for reporting what you did in detail.
        Dim Increment As Double             'For remembering plasticity used.
    
    'Variables for time
        Dim TimeSinceLastReport As Single
        Let TimeSinceLastReport = Timer
        
    'Variables for where we are in the run.
        Dim LearningStageIndex As Long
        Dim TotalCycles As Long             'For monitoring progress.
        
    'For exact frequency presentation:
        Dim NumberSelectedFromArray As Long
        
    'This must be redimmed early, for on-line reporting.
       ReDim mPercentageGenerated(mNumberOfForms, mMaximumNumberOfRivals)
    
    'Debug at starting point:
    
        GoTo DebugExitPoint:
        Dim DebugFile As Long
        Let DebugFile = FreeFile
        Open gOutputFilePath + "\DebugGLAAtStartOfCore.txt" For Output As #DebugFile
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
                    Print #DebugFile, RivalIndex; "   Rival:  "; mRival(FormIndex, RivalIndex); _
                        "   Frequency = "; mFrequency(FormIndex, RivalIndex);
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
    
    'Initialize counters.
        
        Let DummyCounter = 0        'Interrupts execution for various purposes
                                    '   (progress report, cancellation)
        Let TotalCycles = 0         'Reports total progress to user.
        'prevent crashes by installing default
            If gReportingFrequency = 0 Then Let gReportingFrequency = 200
        
    'Get a progress report on the screen with the initial rankings.
        'You need to fake some parameters for the first display.
            Let PlastMark = mCustomPlastMark(1)
            Let PlastFaith = mCustomPlastFaith(1)
            Let NoiseMark = CustomNoiseMark(1)
            Let NoiseFaith = CustomNoiseFaith(1)
        'Now it's safe to report progress.
            Call GLAReportProgress(0, PlastMark, PlastFaith, NoiseMark, NoiseFaith)

    'We will keep track of what learning trial we are on.  Just in case of multiple runs, initialize this.
        Let mLearningTrial = 0
        
    'Do as many learning stages as there are.
        
        For LearningStageIndex = 1 To mNumberOfLearningStages
        
            'Establish the plasticities and noises that will hold for
            '   this learning stage.  Use local variables that might be
            '    bit faster than looking them up repeatedly in an array.
                Let PlastMark = mCustomPlastMark(LearningStageIndex)
                Let PlastFaith = mCustomPlastFaith(LearningStageIndex)
                Let NoiseMark = CustomNoiseMark(LearningStageIndex)
                Let NoiseFaith = CustomNoiseFaith(LearningStageIndex)
                    
            'Go through the cycles of this learning stage.
            For CycleIndex = 1 To mTrialsPerLearningStage(LearningStageIndex)
            
                'Note what learning trial we are on.
                    Let mLearningTrial = mLearningTrial + 1
                
                'mblnProcessing (from KZ) lets you interrupt the GLA by clicking
                '   the Run button--relabeled Cancel during a run.
                    If mblnProcessing = False Then
                        GoTo ExitPoint
                    Else
                
                    'We're not canceled, so execute the current learning cycle.

                    'If doing Stochastic OT, form grammars by sorting the ranking values:
                        If optStochasticOT Then
                            GoTo DetermineSelectionPointsAndSort
                        End If
DetermineSelectionPointsAndSortReturnPoint:
                    
                    'Select an exemplar for learning with, respecting the frequencies.
                    '   The code in question returns two values:  SelectedExemplarForm and SelectedObservedRival

                        GoTo SelectExemplars
SelectExemplarsReturnPoint:
                        
                    'Generate a form for the same input using the current grammar, either Stochastic OT or MaxEnt.
                        If optStochasticOT Then
                            GoTo GenerateAFormStochasticOT
                        Else
                            'MaxEnt.
                            'We provide the weights, violations, and input form, and return a sampled rival candidate,
                            '   to be matched again the sample observed form.
                                Let LocalWinner = MaxEntSampledCandidate(SelectedExemplarForm)
                        End If
GenerateAFormReturnPoint:
             
                    'Compare the selected exemplar with the generated form, and adjust the ranking values/weeights of the constraints accordingly.
                    'The key comparison will be in the context of SelectedExamplarRival and LocalWinner.
                    'I believe the same code will work for both Stochastic OT and MaxEnt (cf. Jaeger), though perhaps the plasticities
                    '    need to be different.
                    
                    'The comparison comes in two flavors, depending on whether a priorirankings are in effect.
                       Call RankingValueAdjustment(LocalWinner, SelectedExemplarForm, SelectedObservedRival, PlastMark, PlastFaith)

                    'That completes a learning trial, so now update the history files.
                    
                    'The smaller one:
                        If mnuGenerateHistory.Checked = True Then
                            'Current trial number
                                Print #mSimpleHistoryFile, Trim(Str(mLearningTrial));
                            'Current ranking values:
                                For ConstraintIndex = 1 To mNumberOfConstraints
                                   Print #mSimpleHistoryFile, Chr$(9); FourDecPlaces(mRankingValue(ConstraintIndex));
                                Next ConstraintIndex
                                Print #mSimpleHistoryFile,
                        End If
                        
                    'The larger, more detailed one.
                        If mnuFullHistory.Checked = True Then
                            'Current trial number
                                Print #mFullHistoryFile, Trim(Str(mLearningTrial));
                            'Data involved:
                                Print #mFullHistoryFile, vbTab; mInputForm(SelectedExemplarForm);
                                Print #mFullHistoryFile, vbTab; mRival(SelectedExemplarForm, LocalWinner);
                                Print #mFullHistoryFile, vbTab; mRival(SelectedExemplarForm, SelectedObservedRival);
                            'New parameter values:
                                For ConstraintIndex = 1 To mNumberOfConstraints
                                   Print #mFullHistoryFile, Chr$(9); FourDecPlaces(mRankingValue(ConstraintIndex));
                                Next ConstraintIndex
                                Print #mFullHistoryFile,
                        End If
                        
                    'The one that tracks candidate probabilities.
                        If mnuCandidateProbabilityHistory.Checked = True Then
                            'Calculate the current one
                                Call GenerateMaxEntPredictions
                            'Print the current trial number
                                Print #mCandidateHistoryFile, Trim(Str(mLearningTrial));
                            'Probabilities:
                                'Print them.
                                    For FormIndex = 1 To mNumberOfForms
                                        For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                                                Print #mCandidateHistoryFile, vbTab; mPercentageGenerated(FormIndex, RivalIndex);
                                        Next RivalIndex
                                    Next FormIndex
                                    Print #mCandidateHistoryFile,
                        End If
                            
                    'Keep track of total cycles for reporting progress.
                        Let TotalCycles = TotalCycles + gReportingFrequency
                    
                    'Report progress if 1 second has elapsed.
                        If Timer - TimeSinceLastReport > 1 Then
                            Call GLAReportProgress(mLearningTrial, PlastMark, PlastFaith, NoiseMark, NoiseFaith)
                            Let TimeSinceLastReport = Timer
                        End If

                    'Check if it's time to do tasks that need done every so often.
                        If DummyCounter < gReportingFrequency Then
                            'No.  Just increment the counter.
                                Let DummyCounter = DummyCounter + 1
                        Else
                            'Yes.  Do these tasks.
                            'Check for cancellation by user.
                                'KZ: checks after every gReportingFrequency cycles.
                                DoEvents    'KZ: passes control back to the operating environment
                                            'to check if there are any pending events--one such
                                            'event would be a second click of cmdRun, which would
                                            'cancel the running of the algorithm.
                            

                            'Reset the DummyCounter to start the next interval of gReportingFrequency cycles.
                                Let DummyCounter = 0
                                
                        End If      'Has it been gReportingFrequency cycles?
                End If              'Is mblnProcessing False, so we can keep going?
            
            Next CycleIndex         'Go on to the next cycle of this learning stage.
        Next LearningStageIndex     'Go on to the next learning stage.

ExitPoint:                          'Go here on failure, so you can check the Timer.

    'Determine how long learning took.
        Let mTimeMarker = Timer - mTimeMarker
      
    'This is the end of the main part of GLACore; the rest are subroutines.
   
   Exit Sub

Stop
'-----------------------------------------------------------------------------
DetermineSelectionPointsAndSort:
    
    'Determine a set of selection points, using the ranking values and the Gaussian
    '    distribution.  Then sort them to form an evanescent grammar.

        'Initialize the SlotFiller() array.
            For ConstraintIndex = 1 To mNumberOfConstraints
                Let LocalSlotFiller(ConstraintIndex) = ConstraintIndex
            Next ConstraintIndex

        'Go through all constraints, and assign each a local ranking value,
        '  according to their probability distributions.
            For ConstraintIndex = 1 To mNumberOfConstraints
                Let LocalRankingValue(ConstraintIndex) = mRankingValue(ConstraintIndex) + Gaussian
            Next ConstraintIndex

        'Go through these ranking values, and form a local grammar by sorting.
            For i = 2 To mNumberOfConstraints
                For j = 1 To i - 1
                    If LocalRankingValue(j) < LocalRankingValue(i) Then
                       Let Swappant = LocalRankingValue(i)
                        Let LocalRankingValue(i) = LocalRankingValue(j)
                        Let LocalRankingValue(j) = Swappant
                        Let SwapInt = LocalSlotFiller(i)
                        Let LocalSlotFiller(i) = LocalSlotFiller(j)
                        Let LocalSlotFiller(j) = SwapInt
                    End If
                Next j
            Next i

        'Install the resulting grammar in the batch of grammars.
            For i = 1 To mNumberOfConstraints
                Let SlotFiller(i) = LocalSlotFiller(i)
            Next i

    GoTo DetermineSelectionPointsAndSortReturnPoint

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
                Let SelectedObservedRival = mDataPresentationArray(NumberSelectedFromArray).RivalIndex
            'Keep a record of the learning frequencies
                Let mActualFrequencyShare(mDataPresentationArray(NumberSelectedFromArray).FormIndex, mDataPresentationArray(NumberSelectedFromArray).RivalIndex) = mActualFrequencyShare(mDataPresentationArray(NumberSelectedFromArray).FormIndex, mDataPresentationArray(NumberSelectedFromArray).RivalIndex) + 1
                Let mActualFrequencyPerInput(mDataPresentationArray(NumberSelectedFromArray).FormIndex) = mActualFrequencyPerInput(mDataPresentationArray(NumberSelectedFromArray).FormIndex) + 1

            'Debug:
                Print #DebugFile, "Form:  "; mInputForm(SelectedExemplarForm); vbTab; "Rival:  "; mRival(SelectedExemplarForm, SelectedObservedRival)
            
    Else

        'Stochastic method:  Go through the possible outputs, and select one at
        '    random, according to the frequencies.
           
             Let blnExemplarChosen = False
             'Randomize     'Argh!--this produced not-quite-right results for a long period.
                            'Commented out 12/31/04.
             Let RandomNumber = Rnd()
             For FormIndex = 1 To mNumberOfForms
                 For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                     'Don't bother with zero items.
                     If mFrequencyShare(FormIndex, RivalIndex) > 0 Then
                         If RandomNumber >= mFrequencyInterval(FormIndex, RivalIndex, 0) _
                         And RandomNumber < mFrequencyInterval(FormIndex, RivalIndex, 1) Then
                             Let SelectedExemplarForm = FormIndex
                             Let SelectedObservedRival = RivalIndex
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
'--------------------------------------------------------------------------------------------
GenerateAFormStochasticOT:

   'Go through the candidates, keeping a local best, until you have a final winner.

    'Start with the first candidate (1) as default winner.
        Let LocalWinner = 1
        
    'Ponder each rival in turn.
        For RivalIndex = 2 To mNumberOfRivals(SelectedExemplarForm)
            If RivalIndex <> LocalWinner Then      'Don't bother to compare with self.
                For ConstraintIndex = 1 To mNumberOfConstraints
                    'Select case:  violations of current king of hill, minus violations of
                    '  currently contending rival.
                        Select Case mRivalViolations(SelectedExemplarForm, LocalWinner, SlotFiller(ConstraintIndex)) - mRivalViolations(SelectedExemplarForm, RivalIndex, SlotFiller(ConstraintIndex))
                            Case Is > 0
                                'This kills the current king of the hill.
                                    Let LocalWinner = RivalIndex
                                'Since it's OT, you don't need to look at the rest of the constraints.
                                    Exit For
                            Case Is < 0
                                'This kills the contender. Since it's OT, you don't need to look at the rest of the constraints.
                                    Exit For
                        End Select
                Next ConstraintIndex    'Consider all constraints.
            End If                      'Don't bother to compare with self.
        Next RivalIndex               'Ponder each rival.

   GoTo GenerateAFormReturnPoint

Stop

            
End Sub


Function MaxEntSampledCandidate(myInputForm As Long) As Long

    'We are given a particular input form. Compute the probability distribution over rivals, and sample
    '   from it to pick a local winner.  (This will be compared with the sample empirical winner, to determine reweighting.)
    
    'We need the constraint violations, which are in the module level array mRivalViolations(mNumberOfForms, mMaximumNumberOfRivals, mNumberOfConstraints)
    'We need the current weights, which are in mRankingValues(mNumberOfConstraints)
    'We need the input form being worked with, which is given by MyInputForm.
    
    'We will return the index of a rival candidate, the one that got sampled.
    
    Dim NumberOfRivals As Long
    Dim Harmony() As Single, eHarmony() As Single, Probability() As Single, Z As Single
    Dim MyRandomNumber As Single, AccumulatedProbability As Single
    Dim RivalIndex As Long, ConstraintIndex As Long
    
    'Localize the number of rivals for convenience.
        Let NumberOfRivals = mNumberOfRivals(myInputForm)
    
    'Redimension the local arrays for MaxEnt.
        ReDim Harmony(NumberOfRivals)
        ReDim eHarmony(NumberOfRivals)
        ReDim Probability(NumberOfRivals)
        
    'Compute the Harmony of each candidate.  Note that since we are using old Stochastic OT code, weights are called "ranking values".
        For RivalIndex = 1 To NumberOfRivals
            For ConstraintIndex = 1 To mNumberOfConstraints
                Let Harmony(RivalIndex) = Harmony(RivalIndex) + mRivalViolations(myInputForm, RivalIndex, ConstraintIndex) * mRankingValue(ConstraintIndex)
            Next ConstraintIndex
        Next RivalIndex
    
    'Convert Harmony to eHarmony.
        For RivalIndex = 1 To NumberOfRivals
            Let eHarmony(RivalIndex) = Exp(-1 * Harmony(RivalIndex))
        Next RivalIndex
    
    'Compute Z, by totaling eHarmony.
        For RivalIndex = 1 To NumberOfRivals
            Let Z = Z + eHarmony(RivalIndex)
        Next RivalIndex
        If Z = 0 Then Stop
        
    'Compute probability, i.e. share of Z.
        For RivalIndex = 1 To NumberOfRivals
            Let Probability(RivalIndex) = eHarmony(RivalIndex) / Z
        Next RivalIndex
    
    'Pick a random number between zero and one.
        Let MyRandomNumber = Rnd()
        
    'Go through the candidates and stop at the one where accumulated probability exceeds the random number.
    '   This serves to sample a candidate according to the frequencies predicted by the grammar.
        For RivalIndex = 1 To NumberOfRivals
            Let AccumulatedProbability = AccumulatedProbability + Probability(RivalIndex)
            If AccumulatedProbability >= MyRandomNumber Then
                Let MaxEntSampledCandidate = RivalIndex
                Exit Function
            End If
        Next RivalIndex
        
    'If you get this far, it is an error, so stop.
        Stop

End Function


Sub RankingValueAdjustment(MyLocalWinnerInGrammar As Long, MySelectedExemplarForm As Long, MySelectedObservedRival As Long, MyPlastMark As Double, _
    MyPlastFaith As Double)

   'Compare the sample empirical form with sampled generated form, and adjust ranking values/weights accordingly.
   
   'This covers both MaxEnt and Stochastic OT, which are slightly different.
   
   'There are two non-vanilla options that might be invoked:
   '    The Magri update rule (now perhaps of less value, in light of Magri and Storme's retraction paper)
   '    Biased learning, with sigma and mu.  The math was kindly provided to me in an email (1/2/26) from Gerhard Jaeger.

   'The weight/ranking value change, made at the end.
        Dim RankingValueChange As Single
   
   'The calculations needed for the Magri update rule.
       Dim PromotionAmount As Single
       Dim ConstraintIndex As Long
       
   'For clear implementation of Jaeger's formula, dividing into likelihood and prior.
       Dim LikelihoodBasedChange As Single
       Dim PriorBasedChange As Single

   'Do this only if the two are different.
      If MyLocalWinnerInGrammar = MySelectedObservedRival Then Exit Sub
        
    'If we are using the Magri update rule, we need to compute the promotion amount for this particular set of input, winner, rival.
        If mnuMagri.Checked = True Then
            Let PromotionAmount = MagriPromotionAmount(MySelectedExemplarForm, MyLocalWinnerInGrammar, MySelectedObservedRival)
        End If
            
    'We are ready to compute the adjustments for each ranking value/weight
        
         For ConstraintIndex = 1 To mNumberOfConstraints
         
            If optMaxEnt Then
            'If mnuGaussianPrior.Checked Then
            
                'We will do biased learning separately, to implement precisely what Jaeger specifies.
                    If mFaithfulness(ConstraintIndex) = True Then
                        'Likelihood-based change, which comes ultimately from the O - E theorem for gradient of the log likelihood.
                        '   Specifically:  Plasticity times the difference in violations:  observed minus predicted candidate
                            Let LikelihoodBasedChange = MyPlastFaith * (mRivalViolations(MySelectedExemplarForm, MySelectedObservedRival, ConstraintIndex) - mRivalViolations(MySelectedExemplarForm, MyLocalWinnerInGrammar, ConstraintIndex))
                        'Prior-based change, based on partial derivative of the prior.  The "divided by 2" is out of the blue -- why is it needed?
                            If mnuGaussianPrior.Checked Then
                                Let PriorBasedChange = MyPlastFaith * (mRankingValue(ConstraintIndex) - mMu(ConstraintIndex)) / (mSigma(ConstraintIndex) ^ 2) / 2
                            Else
                                Let PriorBasedChange = 0
                            End If
                    Else
                        'Likelihood-based change, which comes ultimate from the O - E theorem for gradient of the prior.
                        '   Specifically:  Plasticity times the difference in violations:  observed minus predicted candidate
                            Let LikelihoodBasedChange = MyPlastMark * (mRivalViolations(MySelectedExemplarForm, MySelectedObservedRival, ConstraintIndex) - mRivalViolations(MySelectedExemplarForm, MyLocalWinnerInGrammar, ConstraintIndex))
                        'Prior-based change, based on partial derivative of the prior.
                            If mnuGaussianPrior.Checked Then
                                Let PriorBasedChange = MyPlastMark * (mRankingValue(ConstraintIndex) - mMu(ConstraintIndex)) / (mSigma(ConstraintIndex) ^ 2) / 2
                            Else
                                Let PriorBasedChange = 0
                            End If
                    End If
                    'The change will be the difference of these two.
                        Let RankingValueChange = LikelihoodBasedChange + PriorBasedChange
                        
                    'Ok, now we are ready to make the change.
                        Let mRankingValue(ConstraintIndex) = mRankingValue(ConstraintIndex) - RankingValueChange
            
                    'Unless the user is happy with negative weights, enforce positivity.
                        If optMaxEnt = True Then
                            If Not mnuNegativeWeightsOK.Checked Then
                                If mRankingValue(ConstraintIndex) < 0 Then
                                    Let mRankingValue(ConstraintIndex) = 0
                                End If
                            End If
                        End If
            
            Else
            
                'This is the old GLA code, to be used for Stochastic OT.
                    Select Case mRivalViolations(MySelectedExemplarForm, MyLocalWinnerInGrammar, ConstraintIndex) _
                       - mRivalViolations(MySelectedExemplarForm, MySelectedObservedRival, ConstraintIndex)
                       
                       Case Is > 0
                          'The (wrong) LocalWinner violates more.  To improve the grammar, *strengthen* this constraint, to punish the wrong local winner.
        
                        'KZ: plasticity adjustment depends on whether it's a faithfulness constraint or a markedness constraint:
                            If mFaithfulness(ConstraintIndex) = True Then
                                Let RankingValueChange = MyPlastFaith
                                'Use the Magri update rule if it is in effect.
                                    If mnuMagri.Checked = True Then
                                        Let RankingValueChange = RankingValueChange * PromotionAmount
                                    End If
                            Else
                                Let RankingValueChange = MyPlastMark
                                    If mnuMagri.Checked = True Then
                                        Let RankingValueChange = RankingValueChange * PromotionAmount
                                    End If
                            End If
                            
                            If mnuDoGLAWithAPrioriRankings.Checked Then
                                'Enforce the apriori rankings as needed:
                                    Call AdjustAPrioriRankings_Up
                            End If
        
                       Case Is < 0
                            'The (wrong) LocalWinner violates less.  To improve the grammar,
                            '  *weaken* this constraint, so it will not punish the correct form so much.
                            'The Magri update rule imposes no special regime for demotion, so no code here.
        
                            'KZ: plasticity adjustment depends on whether it's a faithfulness constraint or a markedness constraint:
                                If mFaithfulness(ConstraintIndex) = True Then
                                    Let RankingValueChange = -1 * MyPlastFaith
                                Else
                                    Let RankingValueChange = -1 * MyPlastMark
                                End If
        
                    End Select
                    
                    'Ok, now we are ready to make the change.
                        Let mRankingValue(ConstraintIndex) = mRankingValue(ConstraintIndex) + RankingValueChange
                    
                    'Enforce the apriori rankings as needed:
                        If mnuDoGLAWithAPrioriRankings.Checked Then
                            Call AdjustAPrioriRankings_Down
                        End If
                    
            End If      'Are we in MaxEnt or Stochastic OT?
         
         Next ConstraintIndex           'Adjust ranking value or weight of every constraint.
         
End Sub


Function MagriPromotionAmount(MySelectedExemplarForm As Long, MyLocalWinnerInGrammar As Long, MySelectedObservedRival As Long)
            
    'Compute the promotion amount based on the forms you are working with.
    
    Dim ConstraintIndex As Long, NumberOfConstraintsPromoted As Long, NumberOfConstraintsDemoted As Long
    
    'Count the winner-preferring constraints and the loser-preferring constraints.
        For ConstraintIndex = 1 To mNumberOfConstraints
            Select Case mRivalViolations(MySelectedExemplarForm, MyLocalWinnerInGrammar, ConstraintIndex) - mRivalViolations(MySelectedExemplarForm, MySelectedObservedRival, ConstraintIndex)
                Case Is > 0
                    Let NumberOfConstraintsPromoted = NumberOfConstraintsPromoted + 1
                Case Is < 0
                    Let NumberOfConstraintsDemoted = NumberOfConstraintsDemoted + 1
            End Select
        Next ConstraintIndex
    'Compute the promotion amount following Magri's formula.  We use 1 here to add to the denominator, but other values would be possible.
        Let MagriPromotionAmount = NumberOfConstraintsDemoted / (NumberOfConstraintsPromoted + 1)

End Function

Function Gaussian() As Double

    'This algorithm for producing random values from the Gaussian
    '   distribution give you two values at once.  The static variables
    '   below remember the second value, and that you have one available.

        Dim fac As Double, r As Double, v1 As Double, v2 As Double
        Static blnValuedAlreadyStored As Boolean
        Static StoredValue As Double
        
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
                Let StoredValue = 2 * v1 * fac
            'This computes the other normal deviate, which can be used fresh.
                Let Gaussian = 2 * v2 * fac
            'Use the flag to note that you don't have to compute a new one next time.
                Let blnValuedAlreadyStored = True
        Else
            'Return the thriftily stored value.
                Let Gaussian = StoredValue
            'Indicate with the flag that you've used it up and must compute anew next time.
                Let blnValuedAlreadyStored = False
        End If        'Should I compute a new value?

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
    '        Print #DebugFile, FrequencyIndex; vbtab; mDataPresentationArray(FrequencyIndex).FormIndex;
    '        Print #DebugFile, vbtab; mDataPresentationArray(FrequencyIndex).RivalIndex
    '    Next FrequencyIndex
    '    Close #DebugFile

End Sub


Sub GLAReportProgress(CycleIndex&, PlastMark As Double, PlastFaith As Double, NoiseMark As Double, NoiseFaith As Double)

   'Print out the results so far.

      Dim SlotFiller() As Long
      Dim LocalRankingValue() As Double
      ReDim SlotFiller(mNumberOfConstraints)
      ReDim LocalRankingValue(mNumberOfConstraints)
      
      Dim ConstraintIndex As Long
      Dim i As Long, j As Long
      Dim Swappant As Double
      Dim SwapInt As Long
         
   'Sort the constraints by their ranking values.

      For ConstraintIndex = 1 To mNumberOfConstraints
         Let SlotFiller(ConstraintIndex) = ConstraintIndex
         Let LocalRankingValue(ConstraintIndex) = mRankingValue(ConstraintIndex)
      Next ConstraintIndex

      'Hmm, I find I'd prefer not sorting.  Perhaps make it an option.
      GoTo NoSortPoint
      For i = 2 To mNumberOfConstraints
         For j = 1 To i - 1
            If LocalRankingValue(j) < LocalRankingValue(i) Then
               Let Swappant = LocalRankingValue(i)
               Let LocalRankingValue(i) = LocalRankingValue(j)
               Let LocalRankingValue(j) = Swappant
               Let SwapInt = SlotFiller(i)
               Let SlotFiller(i) = SlotFiller(j)
               Let SlotFiller(j) = SwapInt
            End If
         Next j
      Next i
NoSortPoint:

   'Print the results so far.

        pctProgressWindow.Cls
        
        'Cycle number:
            pctProgressWindow.Print "Completed learning cycle #"; CycleIndex&; "/"; mReportedNumberOfDataPresentations
        
        'Plasticity:
            If mblnUseCustomLearningSchedule = True Then
                pctProgressWindow.Print
                pctProgressWindow.Print "Plasticity for faithfulness is currently:  ";
                pctProgressWindow.Print FourDecPlaces(PlastFaith)
                pctProgressWindow.Print "Plasticity for markedness is currently:    ";
                pctProgressWindow.Print FourDecPlaces(PlastMark)
                If optStochasticOT Then
                    pctProgressWindow.Print "Noise for faithfulness is currently:       ";
                    pctProgressWindow.Print FourDecPlaces(NoiseFaith)
                    pctProgressWindow.Print "Noise for markedness is currently:         ";
                    pctProgressWindow.Print FourDecPlaces(NoiseMark)
                End If
                pctProgressWindow.Print
            Else
                pctProgressWindow.Print "Plasticity is currently ";
                'Either value will do, when they are locked together--BH.
                pctProgressWindow.Print FourDecPlaces(PlastFaith); "."
                'No noise in MaxEnt, don't report:
                    If optStochasticOT Then
                        pctProgressWindow.Print "Noise is currently ";
                        pctProgressWindow.Print FourDecPlaces(NoiseFaith); "."
                    End If
            End If
            
        'Ranking values/weights:
            If optStochasticOT Then
                pctProgressWindow.Print "Current ranking values:"
            Else
                pctProgressWindow.Print "Current weights:"
            End If
            For ConstraintIndex = 1 To mNumberOfConstraints
                pctProgressWindow.Print FillStringTo(ThreeDecPlaces(mRankingValue(SlotFiller(ConstraintIndex))), 11);
                'pctProgressWindow.Print ThreeDecPlaces(mRankingValue(SlotFiller(ConstraintIndex)));
                pctProgressWindow.Print "   "; mConstraintName(SlotFiller(ConstraintIndex))
            Next ConstraintIndex

End Sub


Function GLATestGrammar(RankingValue() As Double, ShouldIPrint As Boolean, CyclesToTest As Long) As Double

   'Variable to keep track of how many of each type are getting generated.

       Dim PercentageInInput() As Double
       ReDim PercentageInInput(mNumberOfForms, mMaximumNumberOfRivals)
       Dim LocalSum As Double
       Dim LocalWinner As Long

   'Variables to generate forms:

       Dim LocalRankingValue()  As Double
       ReDim LocalRankingValue(mNumberOfConstraints)
       Dim SlotFiller() As Long
       ReDim SlotFiller(mNumberOfConstraints)
   
   'Indices

       Dim TrialIndex As Long
       Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
       Dim ProgressCounter As Long
       Dim i As Long, j As Long
       
    'Etc.
        Dim Swappant As Double
        Dim SwapInt As Long
        
        Dim DummyCounter As Integer 'KZ: keeps track of how many loops it's been
                                    'since checked to see if user wants to cancel.
       
        Dim ErrorSoFar As Double
      
      ReDim mNumberGenerated(mNumberOfForms, mMaximumNumberOfRivals)

    
    'We will need this number even if we use MaxEnt:
        For FormIndex = 1 To mNumberOfForms
           Let mTotalNumberOfRivals = mTotalNumberOfRivals + mNumberOfRivals(FormIndex) + 1
        Next FormIndex
    
    'And likewise this:  calculate percentages for the input.
        For FormIndex = 1 To mNumberOfForms
           Let LocalSum = 0
           For RivalIndex = 1 To mNumberOfRivals(FormIndex)
              Let LocalSum = LocalSum + mFrequency(FormIndex, RivalIndex)
           Next RivalIndex
           If LocalSum > 0 Then
              For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                 Let PercentageInInput(FormIndex, RivalIndex) = mFrequency(FormIndex, RivalIndex) / LocalSum
              Next RivalIndex
           End If
        Next FormIndex
    
    'Do a branch-off for the rather different calculations needed to test the MaxEnt grammar.
        If optMaxEnt = True Then
            Call GenerateMaxEntPredictions
            GoTo PrintStart
        End If
   
        If ShouldIPrint = True Then
            pctProgressWindow.Cls
            pctProgressWindow.Print "Testing the grammar..."
        End If



   'Just in case you are doing batch processing, initialize the count of forms generated.

      For FormIndex = 1 To mNumberOfForms
         For RivalIndex = 1 To mNumberOfRivals(FormIndex)
            Let mNumberGenerated(FormIndex, RivalIndex) = 0
         Next RivalIndex
      Next FormIndex

'------------------------------------------------------------------------------

   'Loop through this many trials, for all the forms.

        'Keep track of how far you've gotten.
            Let ProgressCounter = 500
            Let DummyCounter = 0  'KZ
            If mblnProcessing = True Then  'KZ: if user has pressed cancel, stop.

     'Main loop, through trials.
      For TrialIndex = 1 To CyclesToTest

         'Report progress, during the big run:
             If ShouldIPrint = True Then
                If TrialIndex = ProgressCounter Then
                    pctProgressWindow.Cls
                    pctProgressWindow.Print "Completed"; TrialIndex; "test trials /"; CyclesToTest
                    Let ProgressCounter = ProgressCounter + 500
                End If
            End If
         
         GoTo FormAGrammarAgain
FormAGrammarAgainReturnPoint:

        'Using this grammar, generate an output for each form and keep count.
            For FormIndex = 1 To mNumberOfForms
                
                GoTo GenerateAFormForTest
GenerateAFormForTestReturnPoint:

                If LocalWinner = 0 Then Stop
                
                'Keep count.
                    Let mNumberGenerated(FormIndex, LocalWinner) = mNumberGenerated(FormIndex, LocalWinner) + 1
ExitPoint:
            Next FormIndex
         
         'KZ: is it time to check for a cancel command?
            If DummyCounter >= 500 Then  'KZ: checks after every 500 trials.
                DoEvents 'KZ: passes control back to the operating environment
                        'to check if there are any pending events--one such
                        'event would be a second click of cmdRun, which would
                        'cancel the running of the algorithm.
                Let DummyCounter = 0
            Else
                Let DummyCounter = DummyCounter + 1
            End If
            
      Next TrialIndex           'Loop:  do this many trials of the grammar.

'---------------------------------------------------------------------------

    'Calculate percentages generated of the rival forms.
        For FormIndex = 1 To mNumberOfForms
            'Test wug only if user asked.
            '    If TestWugOnly = True Then
            '        If IsAWugForm(FormIndex) = False Then
            '            GoTo ExitPoint2
            '        End If
            '    End If
           Let LocalSum = 0
           For RivalIndex = 1 To mNumberOfRivals(FormIndex)
              Let LocalSum = LocalSum + mNumberGenerated(FormIndex, RivalIndex)
           Next RivalIndex
           For RivalIndex = 1 To mNumberOfRivals(FormIndex)
              Let mPercentageGenerated(FormIndex, RivalIndex) = mNumberGenerated(FormIndex, RivalIndex) / LocalSum
           Next RivalIndex
ExitPoint2:
        Next FormIndex

    'Calculate error
        For FormIndex = 1 To mNumberOfForms
            For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                'Previous version was not "centering" the 1-0, .95-.05 case.  Let's try squaring the error.
                Let ErrorSoFar = ErrorSoFar + (PercentageInInput(FormIndex, RivalIndex) - mPercentageGenerated(FormIndex, RivalIndex)) ^ 2
                'Let ErrorSoFar = ErrorSoFar + Abs(PercentageInInput(FormIndex, RivalIndex) - mPercentageGenerated(FormIndex, RivalIndex))
            Next RivalIndex
       Next FormIndex
      'As a function, this routine returns its total error--so it can be used in learning.
            Let GLATestGrammar = ErrorSoFar

'---------------------------------------------------------------------------

PrintStart:
    'Now, print, if that's the purpose for which you were called.
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
                        For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                            Let LocalLength = 3 + Len(DumbSym(mRival(FormIndex, RivalIndex)))
                            If LocalLength > LongestForm Then
                                Let LongestForm = LocalLength
                            End If
                        Next RivalIndex
                    Next FormIndex
        
                Print #mDocFile,
                Print #mDocFile, "\ks"
                Print #mTmpFile,
                
                Call PrintLevel1Header(mDocFile, mTmpFile, mHTMFile, "Matchup to Input Frequencies")
                
                'Print column headers for tabbed output.
                    Print #mTabbedFile, "Input#"; vbTab; "Input"; vbTab; "Cand#"; vbTab; "Cand"; vbTab; "Freq. from input file"; vbTab; "Learning tokens"; vbTab; "Output of grammar test"; vbTab; "Target proportion"; vbTab; "Predicted proportion"
                
                'Notify if only Wugs were tested.
                '    If TestWugOnly = True Then
                '        Print #mTmpFile,
                '        Print #mTmpFile, "   Note:  only Wug forms (zero frequency total) were tested."
                '        Print #mDocFile,
                '        Print #mDocFile, "Note:  only Wug forms (zero frequency total) were tested."
                '        Print #mDocFile,
                '    End If
                    
                For FormIndex = 1 To mNumberOfForms
                    'Restart a table for html.
                        Dim MyTable() As String
                        ReDim MyTable(5, mNumberOfRivals(FormIndex) + 2)
                    'Test wug only if user asked.
                    '    If TestWugOnly = True Then
                    '        If IsAWugForm(FormIndex) = False Then
                    '            GoTo ExitPoint5
                    '        End If
                    '    End If
                    'Header stuff:
                        If FormIndex > 1 Then
                            Print #mDocFile, "\ks"
                            Print #mTmpFile,
                        End If
                        Print #mDocFile, "\ts5"
                        Print #mDocFile, "/"; SymbolTag1; mInputForm(FormIndex); SymbolTag2; "/"; vbTab;
                        Print #mTmpFile, "   /"; DumbSym(mInputForm(FormIndex)); "/ ";
                        Let MyTable(1, 1) = "/" + DumbSym(mInputForm(FormIndex)) + "/"
                        For i = Len(mInputForm(FormIndex)) + 4 To LongestForm
                            Print #mTmpFile, " ";
                        Next i
                        'This needs to be different for Stochastic OT and Maxent
                            If optStochasticOT Then
                                Print #mDocFile, "Input Frequencies"; vbTab; "Generated Frequencies"; vbTab; "Input Number (exact)"; vbTab; "Generated Number"
                                Print #mTmpFile, "Input Fr. Gen Fr.  Input #     Gen. #"
                                Let MyTable(2, 1) = "Input Frequencies"
                                Let MyTable(3, 1) = "Generated Frequencies"
                                Let MyTable(4, 1) = "Input Number (exact)"
                                Let MyTable(5, 1) = "Generated Number"
                            Else
                                Print #mDocFile, "Input Frequencies"; vbTab; "Output probabilities"; vbTab; "Input Number (exact)"
                                Print #mTmpFile, "Input Fr.  P     Input #"
                                Let MyTable(2, 1) = "Input Frequencies"
                                Let MyTable(3, 1) = "Probability"
                                Let MyTable(4, 1) = "Input Number (exact)"
                            End If
                    
                    For RivalIndex = 1 To mNumberOfRivals(FormIndex)
        
                            'The rival:
                                Print #mDocFile, SymbolTag1; mRival(FormIndex, RivalIndex); SymbolTag2; vbTab;
                                Print #mTmpFile, "   "; DumbSym(mRival(FormIndex, RivalIndex));
                                For i = Len(mRival(FormIndex, RivalIndex)) + 1 To LongestForm + 2
                                    Print #mTmpFile, " ";
                                Next i
                                Let MyTable(1, RivalIndex + 2) = mRival(FormIndex, RivalIndex)
                            
                            'Its numbers
                                Print #mDocFile, FourDecPlaces(PercentageInInput(FormIndex, RivalIndex)); vbTab;
                                Print #mTmpFile, FourDecPlaces(PercentageInInput(FormIndex, RivalIndex)); "   ";
                                Let MyTable(2, RivalIndex + 2) = FourDecPlaces(PercentageInInput(FormIndex, RivalIndex))
                                
                                Print #mDocFile, FourDecPlaces(mPercentageGenerated(FormIndex, RivalIndex)); vbTab;
                                Print #mTmpFile, FourDecPlaces(mPercentageGenerated(FormIndex, RivalIndex)); "   ";
                                Let MyTable(3, RivalIndex + 2) = FourDecPlaces(mPercentageGenerated(FormIndex, RivalIndex))
                                
                                Print #mDocFile, Trim(Str(mActualFrequencyShare(FormIndex, RivalIndex))); vbTab;
                                If optMaxEnt = True Then
                                    Print #mTmpFile, Str(mActualFrequencyShare(FormIndex, RivalIndex)); "   ";
                                Else
                                    Print #mTmpFile, Int8(mActualFrequencyShare(FormIndex, RivalIndex)); "   ";
                                End If
                                Let MyTable(4, RivalIndex + 2) = Trim(Str(mActualFrequencyShare(FormIndex, RivalIndex)))
                                
                                'This is not used for MaxEnt, which generates precise output probabilities.
                                    If optStochasticOT Then
                                        Print #mDocFile, Trim(Str(mNumberGenerated(FormIndex, RivalIndex)))
                                        Print #mTmpFile, Str(mNumberGenerated(FormIndex, RivalIndex)); "   ";
                                        Let MyTable(5, RivalIndex + 2) = Trim(Str(mNumberGenerated(FormIndex, RivalIndex)))
                                    
                                        If mActualFrequencyPerInput(FormIndex) > 0 Then
                                            Print #mTmpFile, "   "; FourDecPlaces(mActualFrequencyShare(FormIndex, RivalIndex) / mActualFrequencyPerInput(FormIndex));
                                        End If
                                    
                                    End If
                                
                            'End of line
                                Print #mTmpFile,
                                Print #mDocFile,
                            
                            'Finally, the tabbed file entries:
                                'If LCase(Right(gFileName, 6)) = "tabbed" Then
                                    Print #mTabbedFile, FormIndex; vbTab;
                                    Print #mTabbedFile, mInputForm(FormIndex); vbTab;
                                    Print #mTabbedFile, RivalIndex; vbTab;
                                    Print #mTabbedFile, mRival(FormIndex, RivalIndex); vbTab;
                                    If RivalIndex = 0 Then
                                        Print #mTabbedFile, mWinnerFrequency(FormIndex); vbTab;
                                    Else
                                        Print #mTabbedFile, mRivalFrequency(FormIndex, RivalIndex); vbTab;
                                    End If
                                    Print #mTabbedFile, mActualFrequencyShare(FormIndex, RivalIndex); vbTab;
                                    Print #mTabbedFile, mNumberGenerated(FormIndex, RivalIndex); vbTab;
                                    Print #mTabbedFile, PercentageInInput(FormIndex, RivalIndex); vbTab;
                                    Print #mTabbedFile, mPercentageGenerated(FormIndex, RivalIndex)

                                'End If
                  Next RivalIndex
                        'This is:  bold top row, no bold left column, centered non-left column cells.
                            Call s.PrintHTMTable(MyTable(), mHTMFile, True, False, True)
                        Print #mDocFile, "\te\ke"
               Next FormIndex
               
               Close #mTabbedFile
       End If                           'Should I print?
       
    Exit Function


Stop
'------------------------------------------------------------------------------
FormAGrammarAgain:

   'Go through all constraints, and assign each a local ranking value,
   '  according to their probability distributions.

      For ConstraintIndex = 1 To mNumberOfConstraints
         Let LocalRankingValue(ConstraintIndex) = RankingValue(ConstraintIndex) + Gaussian
      Next ConstraintIndex

   'Go through these ranking values, and form a local grammar by sorting.

      For ConstraintIndex = 1 To mNumberOfConstraints
         Let SlotFiller(ConstraintIndex) = ConstraintIndex
      Next ConstraintIndex

      For i = 2 To mNumberOfConstraints
         For j = 1 To i - 1
            If LocalRankingValue(j) < LocalRankingValue(i) Then
                Let Swappant = LocalRankingValue(i)
                Let LocalRankingValue(i) = LocalRankingValue(j)
                Let LocalRankingValue(j) = Swappant
                Let SwapInt = SlotFiller(i)
                Let SlotFiller(i) = SlotFiller(j)
                Let SlotFiller(j) = SwapInt
            End If
         Next j
      Next i

   GoTo FormAGrammarAgainReturnPoint

Stop
'------------------------------------------------------------------------------
GenerateAFormForTest:

    'xxx bad. delete or fix.

   'Go through the candidates, keeping a local best, until you have a final winner.

      Let LocalWinner = 1      'Start with the first candidate (1) as default winner.

      For RivalIndex = 1 To mNumberOfRivals(FormIndex)
         If RivalIndex <> LocalWinner Then      'Don't bother to compare with self.
            For ConstraintIndex = 1 To mNumberOfConstraints

               'Select case:  violations of current king of hill, minus violations of
               '  currently contending rival.

                  Select Case mRivalViolations(FormIndex, LocalWinner, SlotFiller(ConstraintIndex)) - mRivalViolations(FormIndex, RivalIndex, SlotFiller(ConstraintIndex))
                     Case Is > 0
                        'This kills the current king of the hill.
                            Let LocalWinner = RivalIndex
                            Exit For
                     Case Is < 0
                        'This kills the contender.
                            Exit For
                  End Select

            Next ConstraintIndex
         End If                   'END IF:  Don't bother to compare with self.
      Next RivalIndex
      
      'Now see which constraints are active under this ranking.
        For RivalIndex = 1 To mNumberOfRivals(FormIndex)
            If RivalIndex <> LocalWinner Then      'Don't bother to compare with self.
                For ConstraintIndex = 1 To mNumberOfConstraints
                    Select Case mRivalViolations(FormIndex, LocalWinner, SlotFiller(ConstraintIndex)) - mRivalViolations(FormIndex, RivalIndex, SlotFiller(ConstraintIndex))
                       Case Is > 0
                          'Hmm... a winner-preferrer.  Shouldn't occur, I think.
                              Stop
                       Case Is < 0
                          'This kills the rival, showing that the constraint is active.
                            Let mActive(SlotFiller(ConstraintIndex)) = True
                            Exit For
                    End Select
                Next ConstraintIndex
            End If                   'END IF:  Don't bother to compare with self.
        Next RivalIndex

        GoTo GenerateAFormForTestReturnPoint


    End If  'end if for mblnProcessing

End Function

Sub GenerateMaxEntPredictions()

    'This is easy; we need not sample, but just apply the MaxEnt formula.
    
    On Error GoTo ErrorPoint
    
    Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
    Dim NumberOfRivals As Long
    Dim Harmony() As Single, eHarmony() As Single
    Dim Z As Single, DataCountForThisInput As Long
    
    For FormIndex = 1 To mNumberOfForms
        'If FormIndex = 2 Then Stop
        'Initialize.
            Let Z = 0
            Let NumberOfRivals = mNumberOfRivals(FormIndex)
            ReDim Harmony(NumberOfRivals)
            ReDim eHarmony(NumberOfRivals)
        'Examine all rivals.
            For RivalIndex = 1 To NumberOfRivals
                'Compute harmony of this rival
                    For ConstraintIndex = 1 To mNumberOfConstraints
                        Let Harmony(RivalIndex) = Harmony(RivalIndex) + mRivalViolations(FormIndex, RivalIndex, ConstraintIndex) * mRankingValue(ConstraintIndex)
                    Next ConstraintIndex
                'From this, eHarmony
                    Let eHarmony(RivalIndex) = Exp(-1 * Harmony(RivalIndex))
                'and Z
                    Let Z = Z + eHarmony(RivalIndex)
                'We also need the raw number of forms for this input.
                    Let DataCountForThisInput = DataCountForThisInput + mRivalFrequency(FormIndex, RivalIndex)
            Next RivalIndex
            'If Z < 0.1 Then MsgBox "Z is " + Str(Z)
        For RivalIndex = 1 To NumberOfRivals
            'Probability and predicted counts
                If Z = 0 Then
                    'Stop
                    Close
                    End
                End If
                Let mPercentageGenerated(FormIndex, RivalIndex) = eHarmony(RivalIndex) / Z
                'Let mNumberGenerated(FormIndex, RivalIndex) = mPercentageGenerated(FormIndex, RivalIndex) * DataCountForThisInput
        Next RivalIndex
    Next FormIndex
    
    Exit Sub
    
ErrorPoint:
    MsgBox "Program crashed; autoexit with saving of files.  Please notify Bruce Hayes at bhayes@humnet.ucla.edu, specifying error number 89993."
    Close
    End

End Sub




'=================================PRINTING==================================
'===========================================================================

Sub PrintGLAResults(RankingValue() As Double, ThingFound As String)

   'Print out the results of a numerical algorithm.
        'The RankingValue() array can be either GLA ranking values or Maximum Entropy weights;
        '   and ThingFound (must be grammatically plural) verbally identifies which one it is.
        'Within this code, for historical reasons, the variable is called RankingValue().

      Dim SlotFiller() As Long
      ReDim SlotFiller(mNumberOfConstraints)
      Dim LocalRankingValue() As Double
      ReDim LocalRankingValue(mNumberOfConstraints)

      Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long, InnerConstraintIndex As Long
      Dim ThresholdIndex As Long
      Dim SpaceIndex As Long
      Dim i As Long, j As Long
      
      Dim Swappant As Double, SwapInt As Long
      
      Dim Difference As Double
      
      Dim OldValsFile As Long
      
   'Sort the constraints by their ranking values.

      For ConstraintIndex = 1 To mNumberOfConstraints
         Let SlotFiller(ConstraintIndex) = ConstraintIndex
         Let LocalRankingValue(ConstraintIndex) = RankingValue(ConstraintIndex)
      Next ConstraintIndex

      'I find I would prefer not to sort; perhaps an option some day.
        GoTo DontSort
        For i = 1 To mNumberOfConstraints
           For j = 1 To i - 1
              If LocalRankingValue(j) < LocalRankingValue(i) Then
                  Let Swappant = LocalRankingValue(i)
                  Let LocalRankingValue(i) = LocalRankingValue(j)
                  Let LocalRankingValue(j) = Swappant
                  Let SwapInt = SlotFiller(i)
                  Let SlotFiller(i) = SlotFiller(j)
                  Let SlotFiller(j) = SwapInt
              End If
           Next j
        Next i
DontSort:

   'Print the results of the algorithm.

      Print #mDocFile, "\ks"
      'ThingFound is either weights or ranking values.
        Call PrintLevel1Header(mDocFile, mTmpFile, mHTMFile, ThingFound + " Found")
      
      Print #mDocFile, "\ts2"
      Dim MyTable() As String
      ReDim MyTable(2, mNumberOfConstraints)
       
      For ConstraintIndex = 1 To mNumberOfConstraints
         Print #mDocFile, SmallCapTag1; mConstraintName(SlotFiller(ConstraintIndex)); SmallCapTag2; vbTab;
         Print #mDocFile, FourDecPlaces(RankingValue(SlotFiller(ConstraintIndex)))
         Let MyTable(1, ConstraintIndex) = mConstraintName(SlotFiller(ConstraintIndex))
         Let MyTable(2, ConstraintIndex) = FourDecPlaces((RankingValue(SlotFiller(ConstraintIndex))))
         Print #mTmpFile, FillStringTo(mAbbrev(SlotFiller(ConstraintIndex)), s.Longest(mAbbrev()) + 2); '; ThreeDecPlaces(RankingValue(SlotFiller(ConstraintIndex)));
         Print #mTmpFile, "   "; FourDecPlaces(RankingValue(SlotFiller(ConstraintIndex)))
      Next ConstraintIndex
   
      Print #mDocFile, "\te\ke"
      Call s.PrintHTMTable(MyTable(), mHTMFile, False, False, True)
      
    'We also want a plain output that can be processed by Excel.
    '   Ranking values should be reported in straight constraint order, to facilitate
    '   averaging over multiple runs.
    
        Let mTabbedFile = FreeFile
        Open gOutputFilePath + gFileName + "TabbedOutput.txt" For Output As #mTabbedFile
        For ConstraintIndex = 1 To mNumberOfConstraints
           Print #mTabbedFile, mAbbrev(ConstraintIndex); vbTab; RankingValue(ConstraintIndex)
        Next ConstraintIndex
        Print #mTabbedFile,
      
   'Print a file to save results if you want to run it further.

        'First, make sure there is a folder for these files, a daughter of the folder in which the input file is located.
            Call Form1.CreateAFolderForOutputFiles
      
        'Print the file.
            Let OldValsFile = FreeFile
            Open gOutputFilePath + gFileName + "MostRecentRankingValues.txt" For Output As #OldValsFile
            For ConstraintIndex = 1 To mNumberOfConstraints
               Print #OldValsFile, mAbbrev(ConstraintIndex); vbTab; RankingValue(ConstraintIndex)
            Next ConstraintIndex

End Sub

Sub PrintAHeader(MyAlgorithmName As String)

    'Print a header for the output file:
        Call PrintTopLevelHeader(mDocFile, mTmpFile, mHTMFile, "Result of Applying " + MyAlgorithmName + " to " + gFileName + gFileSuffix)
        Print #mDocFile, "\cn"; NiceDate; ", "; NiceTime
        Print #mDocFile,
        Print #mDocFile, "\cn"; "OTSoft " + gMyVersionNumber + ", release date " + gMyReleaseDate
        Print #mDocFile,
        Call PrintPara(-1, mTmpFile, mHTMFile, NiceDate + ", " + NiceTime)
        Call PrintPara(-1, mTmpFile, mHTMFile, "OTSoft " + gMyVersionNumber + ", release date " + gMyReleaseDate)
          
    'Identify the framework.
        If optMaxEnt = True Then
            Call PrintPara(mDocFile, mTmpFile, mHTMFile, "The GLA was used to find best-fit MaxEnt weights in gradual fashion, as in Jäger (2004).")
        Else
            Call PrintPara(mDocFile, mTmpFile, mHTMFile, "The GLA was used to find best-fit ranking values for a Stochastic OT grammar (Boersma 1998).")
        End If
    
    'Print a header and diacritic to trigger a page number
        Print #mDocFile, "\hrGLA Results for "; gFileName; gFileSuffix; Chr$(9); NiceDate; Chr$(9); "\pn"


End Sub


Public Function ThreeDecPlaces(ALong As Variant) As String

    Dim Buffer As String
    'Format numbers with three decimal places
        Let Buffer = Format(ALong, "##,##0.000")
        Let ThreeDecPlaces = Trim(Buffer)

End Function

Public Function FourDecPlaces(ALong As Variant) As String

    Dim Buffer As String
    'Format numbers with four decimal places
        Let Buffer = Format(ALong, "##,##0.0000")
        Let FourDecPlaces = Trim(Buffer)

End Function


Function Int8(ALong As Long) As String

    'Format numbers with eight digits and no places
        Dim i As Long
        Let Int8 = Format(ALong, "########")
        For i = Len(Int8) To 7
            Let Int8 = " " + Int8
        Next i

End Function

Sub PrintPairwiseRankingProbabilities(SlotFiller() As Long)
            
    'Each difference of ranking value implies a probability of outranking.
    '   Report this information.
    
    'This also prints the material needed for the Hasse diagram.
    
        Dim ThresholdIndex As Long
        Dim ConstraintIndex As Long, InnerConstraintIndex As Long
        Dim SpaceIndex As Long
        Dim Difference As Double
    
    'Header to report ranking probabilities:
        Print #mDocFile,
        Print #mDocFile, "\ks"
        
        Call PrintLevel1Header(mDocFile, mTmpFile, mHTMFile, "Ranking Value to Ranking Probability Conversion")
        
        Call PrintPara(mDocFile, mTmpFile, mHTMFile, "The computed ranking values imply the pairwise ranking probabilities given below.  PARAIn the table, the probability given is that of the constraint in the row headings outranking the PARAconstraint in the column headings.")
        Print #mDocFile,
        
        'Print #mTmpFile,
        'Print #mTmpFile, "   The computed ranking values correspond to the following pairwise "
        'Print #mTmpFile, "   ranking probabilities:"
        'Print #mTmpFile,
    
    'Set up the Hasse diagram file:
        
        'First, make sure there is a folder for these files, a daughter of the
        '   folder in which the input file is located.
            Call Form1.CreateAFolderForOutputFiles
        
        Let HasseFile = FreeFile
        Open gOutputFilePath + gFileName + "Hasse.txt" For Output As #HasseFile
        Print #HasseFile, "digraph G {"
        
    'Put the names of the constraints into the Hasse diagram file:
        For ConstraintIndex = 1 To mNumberOfConstraints
            Print #HasseFile, "   "; Trim(Str(ConstraintIndex));
            'Note:   chr(34) is a double quote.
            Print #HasseFile, " [label="; Chr(34); mAbbrev(ConstraintIndex); Chr(34);
            Print #HasseFile, ",fontsize = "; Trim(Str(HasseFontSize)); "]"
        Next ConstraintIndex
    
        Call LookUpProbabilities
    
    'Find the probability of each pairwise ranking.
        For ConstraintIndex = 1 To mNumberOfConstraints
            For InnerConstraintIndex = ConstraintIndex + 1 To mNumberOfConstraints
                Let Difference = mRankingValue(SlotFiller(ConstraintIndex)) - mRankingValue(SlotFiller(InnerConstraintIndex))
                'Find the first threshold that this is smaller than.
                '   Let's go for a less detailed story; right now it seems to make a Hasse diagram that is just too busy.
                    For ThresholdIndex = 1 To 443       '481 is max
                        If Difference < mThreshold(ThresholdIndex) Then
                            'Now you know the probability of this ranking difference.
                            'First, inform the user:
                                'Print #mTmpFile, "     "; mProbability(ThresholdIndex);
                                'For SpaceIndex = Len(mProbability(ThresholdIndex)) To 5
                                '    Print #mTmpFile, " ";
                                'Next SpaceIndex
                                'Print #mTmpFile, mAbbrev(SlotFiller(ConstraintIndex)); " >> "; mAbbrev(SlotFiller(InnerConstraintIndex))
                            'Next, update the Hasse diagram file.
                                'The ranking given:
                                    Print #HasseFile, "   "; Trim(Str(SlotFiller(ConstraintIndex)));
                                    Print #HasseFile, " -> "; Trim(Str(SlotFiller(InnerConstraintIndex)));
                                    Print #HasseFile, " [fontsize="; Trim(Str(HasseFontSize - 3));
                                    'The ranking that has a probability of < .95 gets a dotted line.
                                        If Val(mProbability(ThresholdIndex)) < 0.95 Then
                                            Print #HasseFile, ",style=dotted";
                                        End If
                                        Print #HasseFile, ",label="; mProbability(ThresholdIndex);
                                        Print #HasseFile, ",fontsize="; Trim(Str(HasseFontSize - 3)); "]"
                                'The opposite ranking:
                                '    If Val(mProbability(ThresholdIndex)) < 0.95 Then
                                '        Print #HasseFile, "   "; Trim(Str(SlotFiller(InnerConstraintIndex)));
                                '        Print #HasseFile, " -> "; Trim(Str(SlotFiller(ConstraintIndex)));
                                '        Print #HasseFile, " [fontsize="; Trim(Str(HasseFontSize - 3));
                                '        'The ranking that has a probability of < .5 gets a dotted line.
                                '            If Val(mProbability(ThresholdIndex)) > 0.5 Then
                                '                Print #HasseFile, ",style=dotted";
                                '            End If
                                '            Print #HasseFile, ",label="; ThreeDecPlaces(1 - Val(mProbability(ThresholdIndex))); "]"
                                '    End If
                            GoTo ThresholdExitPoint
                        End If
                    Next ThresholdIndex
                'If you've gotten this far, then the probability is large.
                    'Report this to the user.  New fine-grain method stops at .999
                        'Print #mTmpFile, " >.999999  "; mabbrev(SlotFiller(ConstraintIndex)); " >> "; mabbrev(SlotFiller(InnerConstraintIndex))
                        'Print #mTmpFile, "    >.999  "; mAbbrev(SlotFiller(ConstraintIndex)); " >> "; mAbbrev(SlotFiller(InnerConstraintIndex))
                    'Put it as an absolute arc in the Hasse diagram--but only for adjacent-strata
                    '   constraints.
                        If Abs(ConstraintIndex - InnerConstraintIndex) = 1 Then
                            Print #HasseFile, "   "; Trim(Str(SlotFiller(ConstraintIndex)));
                            Print #HasseFile, " -> "; Trim(Str(SlotFiller(InnerConstraintIndex)));
                            Print #HasseFile, " [fontsize="; Trim(Str(HasseFontSize - 3));
                            Print #HasseFile, ",label="; Chr(32); "1"; Chr(32); "]"
                        End If
                    
ThresholdExitPoint:
            Next InnerConstraintIndex
            'Print #mTmpFile,
        Next ConstraintIndex
    
    'Print a pretty tabular version for the Microsoft Word printout and HTML.
    
        Dim Table() As String
        ReDim Table(mNumberOfConstraints, mNumberOfConstraints)
        Print #mDocFile, "\ts" + Trim(Str(mNumberOfConstraints))
        'The constraint names.  Initial tab correctly leaves blank in upper left corner.
            For ConstraintIndex = 2 To mNumberOfConstraints
                Print #mDocFile, Chr$(9); SmallCapTag1; mAbbrev(SlotFiller(ConstraintIndex)); SmallCapTag2;
                Let Table(ConstraintIndex, 1) = mAbbrev(SlotFiller(ConstraintIndex))
            Next ConstraintIndex
            Print #mDocFile,
        'The outranked constraints, and probabilities:
            For ConstraintIndex = 1 To mNumberOfConstraints - 1
                Print #mDocFile, SmallCapTag1; mAbbrev(SlotFiller(ConstraintIndex)); SmallCapTag2;
                Let Table(1, ConstraintIndex + 1) = mAbbrev(SlotFiller(ConstraintIndex))
                For InnerConstraintIndex = 2 To mNumberOfConstraints
                    Print #mDocFile, vbTab;
                    'If it's in the upper diagonal, print a value.  Else shading.
                    If InnerConstraintIndex <= ConstraintIndex Then
                        Print #mDocFile, "\sh";
                        Let Table(InnerConstraintIndex, ConstraintIndex + 1) = "\sh"
                    Else
                        Let Difference = mRankingValue(SlotFiller(ConstraintIndex)) - mRankingValue(SlotFiller(InnerConstraintIndex))
                        'Find the first threshold that this is smaller than.
                            For ThresholdIndex = 1 To 481
                                If Difference < mThreshold(ThresholdIndex) Then
                                    Print #mDocFile, Trim(mProbability(ThresholdIndex));
                                    Let Table(InnerConstraintIndex, ConstraintIndex + 1) = Trim(mProbability(ThresholdIndex))
                                    GoTo ThresholdExitPoint2
                                End If
                            Next ThresholdIndex
                        'If you've gotten this far, then the probability is large.
                            Print #mDocFile, ">.999999";
                            Let Table(InnerConstraintIndex, ConstraintIndex + 1) = ">.999999"
                        End If
ThresholdExitPoint2:
                Next InnerConstraintIndex
                Print #mDocFile,
            Next ConstraintIndex
            Print #mDocFile, "\te\ke"
            Call s.PrintTable(-1, mTmpFile, mHTMFile, Table(), False, False, True)
            
            
    'Finish off the Hasse file and call the routine needed to make a Hasse diagram, using ATT's "dot.exe".
        Print #HasseFile, "}"
        Close #HasseFile
        'Report progress:
            GLA.pctProgressWindow.Cls
            GLA.pctProgressWindow.Print "Creating Hasse diagram..."
            DoEvents
        Call Form1.RunATTDot

End Sub

Sub PrintActiveConstraints(SlotFiller() As Long, Active() As Boolean)
            
    'Print which constraints are active.
    
        Dim ConstraintIndex As Long
    
    'Header to report active constraints:
        Print #mDocFile,
        Print #mDocFile, "\ks"
        Print #mTmpFile,
        
        Call PrintLevel1Header(mDocFile, mTmpFile, mHTMFile, "Active Constraints")
        Call PrintPara(mDocFile, mTmpFile, mHTMFile, "A constraint is active if it causes the winning candidate to defeat a rival PARAin at least one competition.")
        
        Print #mDocFile,
        Print #mDocFile, "\ts2"
        Print #mDocFile, "Constraint"; vbTab; "Status"
        
    'Print them:
        Dim MyTable() As String
        ReDim MyTable(2, mNumberOfConstraints)
        For ConstraintIndex = 1 To mNumberOfConstraints
            Print #mDocFile, SmallCapTag1; mConstraintName(SlotFiller(ConstraintIndex)); SmallCapTag2; vbTab;
            Print #mDocFile, RTrim(ActiveLabel(Active(SlotFiller(ConstraintIndex))))
            Let MyTable(1, ConstraintIndex) = mConstraintName(SlotFiller(ConstraintIndex))
            Let MyTable(2, ConstraintIndex) = RTrim(ActiveLabel(Active(SlotFiller(ConstraintIndex))))
        Next ConstraintIndex
        Call s.PrintTable(-1, mTmpFile, mHTMFile, MyTable(), False, False, False)
            
    'Blank lines:
        Print #mDocFile, "\te\ke"
        Print #mTmpFile,

End Sub

Function ActiveLabel(MyBoolean As Boolean) As String
    If MyBoolean = True Then
        Let ActiveLabel = "Active  "
    Else
        Let ActiveLabel = "Inactive"
    End If
End Function

Sub DebugRoutine(LowerBound As Long)
            
    'Trying this, Jan. 1, 2026
        
        Dim RivalIndex As Long, FormIndex As Long, ConstraintIndex As Long
        Dim DebugFile As Long
        
        Let DebugFile = FreeFile
        
        'First, make sure there is a folder for these files, a daughter of the
        '   folder in which the input file is located.
            Call Form1.CreateAFolderForOutputFiles
        
        Open gOutputFilePath + "/DebugGLA.txt" For Output As #DebugFile
        
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

    'Make approximate tableaux, by calling the tableau-making machinery
    '   of Form1.
    
    'Constraint names are amplified by their ranking values, so users
    '   can guess where the free variation would come through.
    
    'We are trying to assemble the following information:
    
                         'Save a sorted version of the input file if it was so requested.
'                        If mnuSaveAsTxtSortedByRank.Checked = True Then
'                            Call SaveSortedInputFile(LocalConstraintName(), LocalAbbrev(), _
'                                LocalStratum(), _
'                                LocalWinner(), LocalWinnerViolations(), _
'                                LocalNumberOfRivals(), LocalRival(), LocalRivalViolations(), _
'                                LocalWinnerFrequency(), LocalRivalFrequency())
'                        End If

    Dim LocalStratum() As Long      'To store the pseudostrata.
    ReDim LocalStratum(mNumberOfConstraints)
    Dim StratumIndex As Long
    Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
    Dim i As Long, j As Long
    Dim BestRankingValue As Double  'For converting ranking values into strata
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
    Dim LocalRankingValue() As Double
    ReDim LocalRankingValue(mNumberOfConstraints)
    Dim Swappant As Double
    Dim SwapInt As Long
    Dim LocalWinnerViolations() As Long
    ReDim LocalWinnerViolations(mNumberOfForms, mNumberOfConstraints)
    Dim LocalWinnerFrequency() As Double
    ReDim LocalWinnerFrequency(mNumberOfForms)
    Dim LocalNumberOfRivals() As Long
    ReDim LocalNumberOfRivals(mNumberOfForms)
    Dim PutRivalHere As Long
    Dim LocalRival() As String
    ReDim LocalRival(mNumberOfForms, mMaximumNumberOfRivals)
    Dim LocalRivalFrequency() As Double
    ReDim LocalRivalFrequency(mNumberOfForms, mMaximumNumberOfRivals)
    Dim LocalRivalViolations() As Long
    ReDim LocalRivalViolations(mNumberOfForms, mMaximumNumberOfRivals, mNumberOfConstraints)
    'Find pairwise ranking probabilities:
        Dim ThresholdIndex As Long
        Dim Amplification As String
        Dim Difference As Double
        
   'Sort the constraints by their ranking values.
        For ConstraintIndex = 1 To mNumberOfConstraints
            Let SlotFiller(ConstraintIndex) = ConstraintIndex
            Let LocalRankingValue(ConstraintIndex) = mRankingValue(ConstraintIndex)
        Next ConstraintIndex
        For i = 2 To mNumberOfConstraints
            For j = 1 To i - 1
                If LocalRankingValue(j) < LocalRankingValue(i) Then
                    Let Swappant = LocalRankingValue(i)
                    Let LocalRankingValue(i) = LocalRankingValue(j)
                    Let LocalRankingValue(j) = Swappant
                    Let SwapInt = SlotFiller(i)
                    Let SlotFiller(i) = SlotFiller(j)
                    Let SlotFiller(j) = SwapInt
                End If
            Next j
        Next i
   
   'Convert ranking values directly into strata.
        'Initialize:
            For ConstraintIndex = 1 To mNumberOfConstraints
                Let LocalStratum(ConstraintIndex) = 0
            Next ConstraintIndex
        'Sort:
            For StratumIndex = 1 To mNumberOfConstraints
                Let BestRankingValue = -1000000
                Let BestForStratum = 0
                For ConstraintIndex = 1 To mNumberOfConstraints
                    If LocalStratum(ConstraintIndex) = 0 Then
                        If mRankingValue(ConstraintIndex) > BestRankingValue Then
                            Let BestForStratum = ConstraintIndex
                            Let BestRankingValue = mRankingValue(ConstraintIndex)
                        End If
                    End If
                Next ConstraintIndex
                Let LocalStratum(BestForStratum) = StratumIndex
            Next StratumIndex
   
   'Amplify the strings in mabbrev() to include the relative probability that they will
   '    dominate the next constraint down.
   
        Call LookUpProbabilities
        
        For ConstraintIndex = 1 To mNumberOfConstraints - 1
            Let Difference = mRankingValue(SlotFiller(ConstraintIndex)) - mRankingValue(SlotFiller(ConstraintIndex + 1))
            'Find the first threshold that this is smaller than.
                For ThresholdIndex = 1 To 481
                    If Difference < mThreshold(ThresholdIndex) Then
                        Let Amplification = " (" + Trim(mProbability(ThresholdIndex)) + ")"
                        GoTo ThresholdExitPoint2
                    End If
                Next ThresholdIndex
            'If you've gotten this far, then the probability is large.
                Let Amplification = "(1)"
ThresholdExitPoint2:
            Let LocalAbbrev(SlotFiller(ConstraintIndex)) = mAbbrev(SlotFiller(ConstraintIndex)) + Amplification
            'ThreeDecPlaces (mRankingValue(ConstraintIndex)) + ")"
        Next ConstraintIndex
        Let LocalAbbrev(SlotFiller(mNumberOfConstraints)) = mAbbrev(SlotFiller(mNumberOfConstraints))
    
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
        
   'Print tableaux if the user wants them.
        If mnuIncludeTableaux.Checked = True Then
            Call PrintTableaux.Main(mNumberOfForms, mNumberOfConstraints, LocalConstraintName(), _
                LocalAbbrev(), LocalStratum(), mInputForm(), LocalWinner(), _
                mWinnerFrequency(), LocalWinnerViolations(), _
                mMaximumNumberOfRivals, LocalNumberOfRivals(), LocalRival(), mRivalFrequency(), LocalRivalViolations(), _
                mTmpFile, mDocFile, mHTMFile, _
                "Gradual Learning Algorithm", False, 0, False, False)
                
        End If
    
    'Print the probabilities of outranking.
        'If mnuPairwiseRankingProbabilities.Checked = True Then
        'Not applicable to MaxEnt:
            If optStochasticOT Then
                Call PrintPairwiseRankingProbabilities(SlotFiller())
            End If
        'End If
        
    'Print which constraints are active.
            Call PrintActiveConstraints(SlotFiller(), mActive())
    
End Sub

Sub PrintFinalDetails()
    
    Dim ConstraintIndex As Long
    
    'For sorting:
        Dim SlotFiller() As Long
        Dim LocalRankingValue() As Double
        ReDim SlotFiller(mNumberOfConstraints)
        ReDim LocalRankingValue(mNumberOfConstraints)
        Dim i As Long, j As Long
        Dim Swappant As Double, SwapInt As Long
    
    Dim LearningStageIndex As Long      'for learning schedule table
    Dim Buffer As String                'to help with spacing
 
    'For log likelihood.
        Dim LogLikelihoodPackage As gLikelihoodCalculation
        Dim LogLikelihood As Single
        Dim ZeroPredictionWarning As Boolean
    
    'All of the following are only for Stochastic OT:
        If optStochasticOT Then
            'Insert the Hasse diagram if appropriate.
                Call Form1.InsertHasseDiagramIntoOutputFile(mDocFile, mHTMFile)
            
            'Note the accuracy achieved, and the time needed.
               Print #mDocFile,
                'First, a header:
                    Call PrintLevel1Header(mDocFile, mTmpFile, mHTMFile, "Testing the Grammar:  Details")
        
                'Results:
                    Call PrintPara(mDocFile, mTmpFile, mHTMFile, "The grammar was tested for " + Trim(Str(gCyclesToTest)) + " cycles.")
                    Call PrintPara(mDocFile, mTmpFile, mHTMFile, "Average error per candidate:  " + ThreeDecPlaces(100 * mErrorTerm / mTotalNumberOfRivals) + " percent")
                    Call PrintPara(mDocFile, mTmpFile, mHTMFile, "Learning time:  " + ThreeDecPlaces(mTimeMarker / 60) + " minutes")
                    If gNegativeWeightsOK Then
                        Call PrintPara(mDocFile, mTmpFile, mHTMFile, "Negative weights were permitted.")
                    Else
                        Call PrintPara(mDocFile, mTmpFile, mHTMFile, "Negative weights were not permitted.")
                    End If
        End If
        
    'Print the details of the learning simulation.
        Print #mDocFile,
        Print #mTmpFile,
        
        Call PrintLevel1Header(mDocFile, mTmpFile, mHTMFile, "Parameter Values Used by the GLA")
        
    'Initial rankings
        'We really only want a table if it's necessary."
            Select Case InitialRankingChoice
                Case AllSame
                    'Trivial, no table.
                        If optMaxEnt = True Then
                            Call PrintPara(mDocFile, mTmpFile, mHTMFile, "Initial Weights")
                            Call PrintPara(mDocFile, mTmpFile, mHTMFile, "All weights started out at the default value of 0.")
                        Else
                            Call PrintPara(mDocFile, mTmpFile, mHTMFile, "Initial Rankings")
                            Call PrintPara(mDocFile, mTmpFile, mHTMFile, "All constraints started out at the default value of 100.")
                        End If
                Case Else
                    'Nontrivial, so a table.
                        'First, sort the constraints by their initial ranking values.
                            For ConstraintIndex = 1 To mNumberOfConstraints
                                Let SlotFiller(ConstraintIndex) = ConstraintIndex
                                Let LocalRankingValue(ConstraintIndex) = mInitialRankingValue(ConstraintIndex)
                            Next ConstraintIndex
                            For i = 1 To mNumberOfConstraints
                                For j = 1 To i - 1
                                    If LocalRankingValue(j) < LocalRankingValue(i) Then
                                        Let Swappant = LocalRankingValue(i)
                                        Let LocalRankingValue(i) = LocalRankingValue(j)
                                        Let LocalRankingValue(j) = Swappant
                                        Let SwapInt = SlotFiller(i)
                                        Let SlotFiller(i) = SlotFiller(j)
                                        Let SlotFiller(j) = SwapInt
                                    End If
                                Next j
                            Next i
                        'Now you can print the table.
                            Call PrintPara(mDocFile, mTmpFile, mHTMFile, "Initial Rankings (final rankings also shown)")
                            Dim MyTable() As String
                            ReDim MyTable(3, mNumberOfConstraints + 1)
                            Print #mDocFile, "\ts3"
                            'Print #mTmpFile,
                            Print #mDocFile, "Constraint"; vbTab; "Initial Ranking"; vbTab; "Final Ranking"
                            Let MyTable(1, 1) = "Constraint)"
                            Let MyTable(2, 1) = "Initial Ranking"
                            Let MyTable(3, 1) = "Final Ranking"
                            For ConstraintIndex = 1 To mNumberOfConstraints
                                Print #mDocFile, SmallCapTag1; mConstraintName(SlotFiller(ConstraintIndex)); SmallCapTag2; vbTab;
                                Print #mDocFile, ThreeDecPlaces(mInitialRankingValue(SlotFiller(ConstraintIndex))); vbTab;
                                Print #mDocFile, ThreeDecPlaces(mRankingValue(SlotFiller(ConstraintIndex)))
                                Let MyTable(1, ConstraintIndex + 1) = mConstraintName(SlotFiller(ConstraintIndex))
                                Let MyTable(2, ConstraintIndex + 1) = ThreeDecPlaces(mInitialRankingValue(SlotFiller(ConstraintIndex)))
                                Let MyTable(3, ConstraintIndex + 1) = ThreeDecPlaces(mRankingValue(SlotFiller(ConstraintIndex)))
                                'Print #mTmpFile, FillStringTo(ThreeDecPlaces(mInitialRankingValue(SlotFiller(ConstraintIndex))), 13); ThreeDecPlaces(mInitialRankingValue(SlotFiller(ConstraintIndex))); "   ";
                                'Print #mTmpFile, FillStringTo(ThreeDecPlaces(mRankingValue(SlotFiller(ConstraintIndex))), 6); "("; ThreeDecPlaces(mRankingValue(SlotFiller(ConstraintIndex))); ")";
                                'Print #mTmpFile, "   "; mConstraintName(SlotFiller(ConstraintIndex))
                            Next ConstraintIndex
                            Print #mDocFile, "\te"
                            Call s.PrintTable(-1, mTmpFile, mHTMFile, MyTable(), True, False, True)
            End Select      'Did we want a fancy table for nontrivial initial rankings?
        
    'Learning schedule
        'First, a header:
            Call PrintPara(mDocFile, mTmpFile, mHTMFile, "Schedule for GLA Parameters")
        'Then a table:
            ReDim MyTable(6, mNumberOfLearningStages + 1)
            Let MyTable(1, 1) = "Stage"
            Let MyTable(2, 1) = "Trials"
            Let MyTable(3, 1) = "PlastMark"
            Let MyTable(4, 1) = "PlastFaith"
            'MaxEnt doesn't use noise.
                If optStochasticOT Then
                    Let MyTable(5, 1) = "NoiseMark"
                    Let MyTable(6, 1) = "NoiseFaith"
                End If
            
            For LearningStageIndex = 1 To mNumberOfLearningStages
                Let MyTable(1, LearningStageIndex + 1) = Trim(Val(LearningStageIndex))
                Let MyTable(2, LearningStageIndex + 1) = Trim(Val(mTrialsPerLearningStage(LearningStageIndex)))
                Let MyTable(3, LearningStageIndex + 1) = FourDecPlaces(mCustomPlastMark(LearningStageIndex))
                Let MyTable(4, LearningStageIndex + 1) = FourDecPlaces(mCustomPlastFaith(LearningStageIndex))
                If optStochasticOT Then
                    Let MyTable(5, LearningStageIndex + 1) = FourDecPlaces(CustomNoiseMark(LearningStageIndex))
                    Let MyTable(6, LearningStageIndex + 1) = FourDecPlaces(CustomNoiseFaith(LearningStageIndex))
                End If
            Next LearningStageIndex
            
            Call s.PrintTable(mDocFile, mTmpFile, mHTMFile, MyTable(), True, False, True)
            
            Print #mTmpFile,
            Call PrintPara(mDocFile, mTmpFile, mHTMFile, "There were a total of " + Trim(Str(mReportedNumberOfDataPresentations)) + " learning trials.")
            If mnuExactProportions.Checked = True Then
                Print #mDocFile,
                Call PrintPara(mDocFile, mTmpFile, mHTMFile, "Data were presented non-stochastically, in exact proportions to their PARAfrequencies in the input file.")
            End If
            If mnuMagri.Checked = True Then
                Print #mDocFile,
                Call PrintPara(mDocFile, mTmpFile, mHTMFile, "The Magri update rule was employed.")
            End If
            
      'Specify the Gaussian prior if it was used.
            If mnuGaussianPrior.Checked Then
                Call PrintPara(mDocFile, mTmpFile, mHTMFile, "")
                Call PrintPara(mDocFile, mTmpFile, mHTMFile, "A Gaussian prior for MaxEnt learning was in effect.")
                Call PrintPara(mDocFile, mTmpFile, mHTMFile, "")
                Call PrintPara(mDocFile, mTmpFile, mHTMFile, "Constraint" + vbTab + "Mu" + vbTab + "Sigma")
                For ConstraintIndex = 1 To mNumberOfConstraints
                    Print #mTmpFile, mAbbrev(ConstraintIndex); vbTab; mMu(ConstraintIndex); vbTab; mSigma(ConstraintIndex)
                Next ConstraintIndex
            End If

      'Log likelihood of the data.
            'Calculate it.
                'MsgBox Str(mPercentageGenerated(1, 1))
                Let LogLikelihoodPackage = Module1.CalculateLogLikelihood(mNumberOfForms, mNumberOfRivals(), mPercentageGenerated(), mFrequency())
                Let LogLikelihood = LogLikelihoodPackage.LogLikelihood
                Let ZeroPredictionWarning = LogLikelihoodPackage.IncludesAZeroProbability
            'Print it.
                Print #mTmpFile,
                Print #mTmpFile, "Log likelihood of the data:  " + Trim(Str(LogLikelihood))
                Print #mDocFile, "Log likelihood of the data:  " + Trim(Str(LogLikelihood))
                If ZeroPredictionWarning Then
                    Print #mTmpFile, "Caution:  there were attested values assigned zero probability. To enable likelihood calculation, these were assigned the probability 0.001 by convention."
                    Print #mDocFile, "Caution:  there were attested values assigned zero probability. To enable likelihood calculation, these were assigned the probability 0.001 by convention."
                End If


End Sub

Sub LookUpProbabilities()
    
    'Here are the probabilities, listed as code to reduce installation problems.
        Let mThreshold(1) = 0
        Let mProbability(1) = "0.5"
        Let mThreshold(2) = 0.01
        Let mProbability(2) = "0.501"
        Let mThreshold(3) = 0.02
        Let mProbability(3) = "0.503"
        Let mThreshold(4) = 0.03
        Let mProbability(4) = "0.504"
        Let mThreshold(5) = 0.04
        Let mProbability(5) = "0.506"
        Let mThreshold(6) = 0.05
        Let mProbability(6) = "0.507"
        Let mThreshold(7) = 0.06
        Let mProbability(7) = "0.508"
        Let mThreshold(8) = 0.07
        Let mProbability(8) = "0.51"
        Let mThreshold(9) = 0.08
        Let mProbability(9) = "0.511"
        Let mThreshold(10) = 0.09
        Let mProbability(10) = "0.513"
        Let mThreshold(11) = 0.1
        Let mProbability(11) = "0.514"
        Let mThreshold(12) = 0.11
        Let mProbability(12) = "0.516"
        Let mThreshold(13) = 0.12
        Let mProbability(13) = "0.517"
        Let mThreshold(14) = 0.13
        Let mProbability(14) = "0.518"
        Let mThreshold(15) = 0.14
        Let mProbability(15) = "0.52"
        Let mThreshold(16) = 0.15
        Let mProbability(16) = "0.521"
        Let mThreshold(17) = 0.16
        Let mProbability(17) = "0.523"
        Let mThreshold(18) = 0.17
        Let mProbability(18) = "0.524"
        Let mThreshold(19) = 0.18
        Let mProbability(19) = "0.525"
        Let mThreshold(20) = 0.19
        Let mProbability(20) = "0.527"
        Let mThreshold(21) = 0.2
        Let mProbability(21) = "0.528"
        Let mThreshold(22) = 0.21
        Let mProbability(22) = "0.53"
        Let mThreshold(23) = 0.22
        Let mProbability(23) = "0.531"
        Let mThreshold(24) = 0.23
        Let mProbability(24) = "0.532"
        Let mThreshold(25) = 0.24
        Let mProbability(25) = "0.534"
        Let mThreshold(26) = 0.25
        Let mProbability(26) = "0.535"
        Let mThreshold(27) = 0.26
        Let mProbability(27) = "0.537"
        Let mThreshold(28) = 0.27
        Let mProbability(28) = "0.538"
        Let mThreshold(29) = 0.28
        Let mProbability(29) = "0.539"
        Let mThreshold(30) = 0.29
        Let mProbability(30) = "0.541"
        Let mThreshold(31) = 0.3
        Let mProbability(31) = "0.542"
        Let mThreshold(32) = 0.31
        Let mProbability(32) = "0.544"
        Let mThreshold(33) = 0.32
        Let mProbability(33) = "0.545"
        Let mThreshold(34) = 0.33
        Let mProbability(34) = "0.546"
        Let mThreshold(35) = 0.34
        Let mProbability(35) = "0.548"
        Let mThreshold(36) = 0.35
        Let mProbability(36) = "0.549"
        Let mThreshold(37) = 0.36
        Let mProbability(37) = "0.551"
        Let mThreshold(38) = 0.37
        Let mProbability(38) = "0.552"
        Let mThreshold(39) = 0.38
        Let mProbability(39) = "0.553"
        Let mThreshold(40) = 0.39
        Let mProbability(40) = "0.555"
        Let mThreshold(41) = 0.4
        Let mProbability(41) = "0.556"
        Let mThreshold(42) = 0.41
        Let mProbability(42) = "0.558"
        Let mThreshold(43) = 0.42
        Let mProbability(43) = "0.559"
        Let mThreshold(44) = 0.43
        Let mProbability(44) = "0.56"
        Let mThreshold(45) = 0.44
        Let mProbability(45) = "0.562"
        Let mThreshold(46) = 0.45
        Let mProbability(46) = "0.563"
        Let mThreshold(47) = 0.46
        Let mProbability(47) = "0.565"
        Let mThreshold(48) = 0.47
        Let mProbability(48) = "0.566"
        Let mThreshold(49) = 0.48
        Let mProbability(49) = "0.567"
        Let mThreshold(50) = 0.49
        Let mProbability(50) = "0.569"
        Let mThreshold(51) = 0.5
        Let mProbability(51) = "0.57"
        Let mThreshold(52) = 0.51
        Let mProbability(52) = "0.572"
        Let mThreshold(53) = 0.52
        Let mProbability(53) = "0.573"
        Let mThreshold(54) = 0.53
        Let mProbability(54) = "0.574"
        Let mThreshold(55) = 0.54
        Let mProbability(55) = "0.576"
        Let mThreshold(56) = 0.55
        Let mProbability(56) = "0.577"
        Let mThreshold(57) = 0.56
        Let mProbability(57) = "0.578"
        Let mThreshold(58) = 0.57
        Let mProbability(58) = "0.58"
        Let mThreshold(59) = 0.58
        Let mProbability(59) = "0.581"
        Let mThreshold(60) = 0.59
        Let mProbability(60) = "0.583"
        Let mThreshold(61) = 0.6
        Let mProbability(61) = "0.584"
        Let mThreshold(62) = 0.61
        Let mProbability(62) = "0.585"
        Let mThreshold(63) = 0.62
        Let mProbability(63) = "0.587"
        Let mThreshold(64) = 0.63
        Let mProbability(64) = "0.588"
        Let mThreshold(65) = 0.64
        Let mProbability(65) = "0.59"
        Let mThreshold(66) = 0.65
        Let mProbability(66) = "0.591"
        Let mThreshold(67) = 0.66
        Let mProbability(67) = "0.592"
        Let mThreshold(68) = 0.67
        Let mProbability(68) = "0.594"
        Let mThreshold(69) = 0.68
        Let mProbability(69) = "0.595"
        Let mThreshold(70) = 0.69
        Let mProbability(70) = "0.596"
        Let mThreshold(71) = 0.7
        Let mProbability(71) = "0.598"
        Let mThreshold(72) = 0.71
        Let mProbability(72) = "0.599"
        Let mThreshold(73) = 0.72
        Let mProbability(73) = "0.6"
        Let mThreshold(74) = 0.73
        Let mProbability(74) = "0.602"
        Let mThreshold(75) = 0.74
        Let mProbability(75) = "0.603"
        Let mThreshold(76) = 0.75
        Let mProbability(76) = "0.605"
        Let mThreshold(77) = 0.76
        Let mProbability(77) = "0.606"
        Let mThreshold(78) = 0.77
        Let mProbability(78) = "0.607"
        Let mThreshold(79) = 0.78
        Let mProbability(79) = "0.609"
        Let mThreshold(80) = 0.79
        Let mProbability(80) = "0.61"
        Let mThreshold(81) = 0.8
        Let mProbability(81) = "0.611"
        Let mThreshold(82) = 0.81
        Let mProbability(82) = "0.613"
        Let mThreshold(83) = 0.82
        Let mProbability(83) = "0.614"
        Let mThreshold(84) = 0.83
        Let mProbability(84) = "0.615"
        Let mThreshold(85) = 0.84
        Let mProbability(85) = "0.617"
        Let mThreshold(86) = 0.85
        Let mProbability(86) = "0.618"
        Let mThreshold(87) = 0.86
        Let mProbability(87) = "0.619"
        Let mThreshold(88) = 0.87
        Let mProbability(88) = "0.621"
        Let mThreshold(89) = 0.88
        Let mProbability(89) = "0.622"
        Let mThreshold(90) = 0.89
        Let mProbability(90) = "0.623"
        Let mThreshold(91) = 0.9
        Let mProbability(91) = "0.625"
        Let mThreshold(92) = 0.91
        Let mProbability(92) = "0.626"
        Let mThreshold(93) = 0.92
        Let mProbability(93) = "0.628"
        Let mThreshold(94) = 0.93
        Let mProbability(94) = "0.629"
        Let mThreshold(95) = 0.94
        Let mProbability(95) = "0.63"
        Let mThreshold(96) = 0.95
        Let mProbability(96) = "0.632"
        Let mThreshold(97) = 0.96
        Let mProbability(97) = "0.633"
        Let mThreshold(98) = 0.97
        Let mProbability(98) = "0.634"
        Let mThreshold(99) = 0.98
        Let mProbability(99) = "0.636"
        Let mThreshold(100) = 0.99
        Let mProbability(100) = "0.637"
        Let mThreshold(101) = 1
        Let mProbability(101) = "0.638"
        Let mThreshold(102) = 1.01
        Let mProbability(102) = "0.639"
        Let mThreshold(103) = 1.02
        Let mProbability(103) = "0.641"
        Let mThreshold(104) = 1.03
        Let mProbability(104) = "0.642"
        Let mThreshold(105) = 1.04
        Let mProbability(105) = "0.643"
        Let mThreshold(106) = 1.05
        Let mProbability(106) = "0.645"
        Let mThreshold(107) = 1.06
        Let mProbability(107) = "0.646"
        Let mThreshold(108) = 1.07
        Let mProbability(108) = "0.647"
        Let mThreshold(109) = 1.08
        Let mProbability(109) = "0.649"
        Let mThreshold(110) = 1.09
        Let mProbability(110) = "0.65"
        Let mThreshold(111) = 1.1
        Let mProbability(111) = "0.651"
        Let mThreshold(112) = 1.11
        Let mProbability(112) = "0.653"
        Let mThreshold(113) = 1.12
        Let mProbability(113) = "0.654"
        Let mThreshold(114) = 1.13
        Let mProbability(114) = "0.655"
        Let mThreshold(115) = 1.14
        Let mProbability(115) = "0.657"
        Let mThreshold(116) = 1.15
        Let mProbability(116) = "0.658"
        Let mThreshold(117) = 1.16
        Let mProbability(117) = "0.659"
        Let mThreshold(118) = 1.17
        Let mProbability(118) = "0.66"
        Let mThreshold(119) = 1.18
        Let mProbability(119) = "0.662"
        Let mThreshold(120) = 1.19
        Let mProbability(120) = "0.663"
        Let mThreshold(121) = 1.2
        Let mProbability(121) = "0.664"
        Let mThreshold(122) = 1.21
        Let mProbability(122) = "0.666"
        Let mThreshold(123) = 1.22
        Let mProbability(123) = "0.667"
        Let mThreshold(124) = 1.23
        Let mProbability(124) = "0.668"
        Let mThreshold(125) = 1.24
        Let mProbability(125) = "0.669"
        Let mThreshold(126) = 1.25
        Let mProbability(126) = "0.671"
        Let mThreshold(127) = 1.26
        Let mProbability(127) = "0.672"
        Let mThreshold(128) = 1.27
        Let mProbability(128) = "0.673"
        Let mThreshold(129) = 1.28
        Let mProbability(129) = "0.675"
        Let mThreshold(130) = 1.29
        Let mProbability(130) = "0.676"
        Let mThreshold(131) = 1.3
        Let mProbability(131) = "0.677"
        Let mThreshold(132) = 1.31
        Let mProbability(132) = "0.678"
        Let mThreshold(133) = 1.32
        Let mProbability(133) = "0.68"
        Let mThreshold(134) = 1.33
        Let mProbability(134) = "0.681"
        Let mThreshold(135) = 1.34
        Let mProbability(135) = "0.682"
        Let mThreshold(136) = 1.35
        Let mProbability(136) = "0.683"
        Let mThreshold(137) = 1.36
        Let mProbability(137) = "0.685"
        Let mThreshold(138) = 1.37
        Let mProbability(138) = "0.686"
        Let mThreshold(139) = 1.38
        Let mProbability(139) = "0.687"
        Let mThreshold(140) = 1.39
        Let mProbability(140) = "0.688"
        Let mThreshold(141) = 1.4
        Let mProbability(141) = "0.69"
        Let mThreshold(142) = 1.41
        Let mProbability(142) = "0.691"
        Let mThreshold(143) = 1.42
        Let mProbability(143) = "0.692"
        Let mThreshold(144) = 1.43
        Let mProbability(144) = "0.693"
        Let mThreshold(145) = 1.44
        Let mProbability(145) = "0.695"
        Let mThreshold(146) = 1.45
        Let mProbability(146) = "0.696"
        Let mThreshold(147) = 1.46
        Let mProbability(147) = "0.697"
        Let mThreshold(148) = 1.47
        Let mProbability(148) = "0.698"
        Let mThreshold(149) = 1.48
        Let mProbability(149) = "0.7"
        Let mThreshold(150) = 1.49
        Let mProbability(150) = "0.701"
        Let mThreshold(151) = 1.5
        Let mProbability(151) = "0.702"
        Let mThreshold(152) = 1.51
        Let mProbability(152) = "0.703"
        Let mThreshold(153) = 1.52
        Let mProbability(153) = "0.705"
        Let mThreshold(154) = 1.53
        Let mProbability(154) = "0.706"
        Let mThreshold(155) = 1.54
        Let mProbability(155) = "0.707"
        Let mThreshold(156) = 1.55
        Let mProbability(156) = "0.708"
        Let mThreshold(157) = 1.56
        Let mProbability(157) = "0.709"
        Let mThreshold(158) = 1.57
        Let mProbability(158) = "0.711"
        Let mThreshold(159) = 1.58
        Let mProbability(159) = "0.712"
        Let mThreshold(160) = 1.59
        Let mProbability(160) = "0.713"
        Let mThreshold(161) = 1.6
        Let mProbability(161) = "0.714"
        Let mThreshold(162) = 1.61
        Let mProbability(162) = "0.715"
        Let mThreshold(163) = 1.62
        Let mProbability(163) = "0.717"
        Let mThreshold(164) = 1.63
        Let mProbability(164) = "0.718"
        Let mThreshold(165) = 1.64
        Let mProbability(165) = "0.719"
        Let mThreshold(166) = 1.65
        Let mProbability(166) = "0.72"
        Let mThreshold(167) = 1.66
        Let mProbability(167) = "0.721"
        Let mThreshold(168) = 1.67
        Let mProbability(168) = "0.723"
        Let mThreshold(169) = 1.68
        Let mProbability(169) = "0.724"
        Let mThreshold(170) = 1.69
        Let mProbability(170) = "0.725"
        Let mThreshold(171) = 1.7
        Let mProbability(171) = "0.726"
        Let mThreshold(172) = 1.71
        Let mProbability(172) = "0.727"
        Let mThreshold(173) = 1.72
        Let mProbability(173) = "0.728"
        Let mThreshold(174) = 1.73
        Let mProbability(174) = "0.73"
        Let mThreshold(175) = 1.74
        Let mProbability(175) = "0.731"
        Let mThreshold(176) = 1.75
        Let mProbability(176) = "0.732"
        Let mThreshold(177) = 1.76
        Let mProbability(177) = "0.733"
        Let mThreshold(178) = 1.77
        Let mProbability(178) = "0.734"
        Let mThreshold(179) = 1.78
        Let mProbability(179) = "0.735"
        Let mThreshold(180) = 1.79
        Let mProbability(180) = "0.737"
        Let mThreshold(181) = 1.8
        Let mProbability(181) = "0.738"
        Let mThreshold(182) = 1.81
        Let mProbability(182) = "0.739"
        Let mThreshold(183) = 1.82
        Let mProbability(183) = "0.74"
        Let mThreshold(184) = 1.83
        Let mProbability(184) = "0.741"
        Let mThreshold(185) = 1.84
        Let mProbability(185) = "0.742"
        Let mThreshold(186) = 1.85
        Let mProbability(186) = "0.743"
        Let mThreshold(187) = 1.86
        Let mProbability(187) = "0.745"
        Let mThreshold(188) = 1.87
        Let mProbability(188) = "0.746"
        Let mThreshold(189) = 1.88
        Let mProbability(189) = "0.747"
        Let mThreshold(190) = 1.89
        Let mProbability(190) = "0.748"
        Let mThreshold(191) = 1.9
        Let mProbability(191) = "0.749"
        Let mThreshold(192) = 1.91
        Let mProbability(192) = "0.75"
        Let mThreshold(193) = 1.92
        Let mProbability(193) = "0.751"
        Let mThreshold(194) = 1.93
        Let mProbability(194) = "0.752"
        Let mThreshold(195) = 1.94
        Let mProbability(195) = "0.754"
        Let mThreshold(196) = 1.95
        Let mProbability(196) = "0.755"
        Let mThreshold(197) = 1.96
        Let mProbability(197) = "0.756"
        Let mThreshold(198) = 1.97
        Let mProbability(198) = "0.757"
        Let mThreshold(199) = 1.98
        Let mProbability(199) = "0.758"
        Let mThreshold(200) = 1.99
        Let mProbability(200) = "0.759"
        Let mThreshold(201) = 2
        Let mProbability(201) = "0.76"
        Let mThreshold(202) = 2.01
        Let mProbability(202) = "0.761"
        Let mThreshold(203) = 2.02
        Let mProbability(203) = "0.762"
        Let mThreshold(204) = 2.03
        Let mProbability(204) = "0.764"
        Let mThreshold(205) = 2.04
        Let mProbability(205) = "0.765"
        Let mThreshold(206) = 2.05
        Let mProbability(206) = "0.766"
        Let mThreshold(207) = 2.06
        Let mProbability(207) = "0.767"
        Let mThreshold(208) = 2.07
        Let mProbability(208) = "0.768"
        Let mThreshold(209) = 2.08
        Let mProbability(209) = "0.769"
        Let mThreshold(210) = 2.09
        Let mProbability(210) = "0.77"
        Let mThreshold(211) = 2.1
        Let mProbability(211) = "0.771"
        Let mThreshold(212) = 2.11
        Let mProbability(212) = "0.772"
        Let mThreshold(213) = 2.12
        Let mProbability(213) = "0.773"
        Let mThreshold(214) = 2.13
        Let mProbability(214) = "0.774"
        Let mThreshold(215) = 2.14
        Let mProbability(215) = "0.775"
        Let mThreshold(216) = 2.15
        Let mProbability(216) = "0.776"
        Let mThreshold(217) = 2.16
        Let mProbability(217) = "0.777"
        Let mThreshold(218) = 2.17
        Let mProbability(218) = "0.779"
        Let mThreshold(219) = 2.18
        Let mProbability(219) = "0.78"
        Let mThreshold(220) = 2.19
        Let mProbability(220) = "0.781"
        Let mThreshold(221) = 2.2
        Let mProbability(221) = "0.782"
        Let mThreshold(222) = 2.21
        Let mProbability(222) = "0.783"
        Let mThreshold(223) = 2.22
        Let mProbability(223) = "0.784"
        Let mThreshold(224) = 2.23
        Let mProbability(224) = "0.785"
        Let mThreshold(225) = 2.24
        Let mProbability(225) = "0.786"
        Let mThreshold(226) = 2.25
        Let mProbability(226) = "0.787"
        Let mThreshold(227) = 2.26
        Let mProbability(227) = "0.788"
        Let mThreshold(228) = 2.27
        Let mProbability(228) = "0.789"
        Let mThreshold(229) = 2.28
        Let mProbability(229) = "0.79"
        Let mThreshold(230) = 2.29
        Let mProbability(230) = "0.791"
        Let mThreshold(231) = 2.3
        Let mProbability(231) = "0.792"
        Let mThreshold(232) = 2.31
        Let mProbability(232) = "0.793"
        Let mThreshold(233) = 2.32
        Let mProbability(233) = "0.794"
        Let mThreshold(234) = 2.33
        Let mProbability(234) = "0.795"
        Let mThreshold(235) = 2.34
        Let mProbability(235) = "0.796"
        Let mThreshold(236) = 2.35
        Let mProbability(236) = "0.797"
        Let mThreshold(237) = 2.36
        Let mProbability(237) = "0.798"
        Let mThreshold(238) = 2.37
        Let mProbability(238) = "0.799"
        Let mThreshold(239) = 2.38
        Let mProbability(239) = "0.8"
        Let mThreshold(240) = 2.39
        Let mProbability(240) = "0.801"
        Let mThreshold(241) = 2.4
        Let mProbability(241) = "0.802"
        Let mThreshold(242) = 2.41
        Let mProbability(242) = "0.803"
        Let mThreshold(243) = 2.42
        Let mProbability(243) = "0.804"
        Let mThreshold(244) = 2.43
        Let mProbability(244) = "0.805"
        Let mThreshold(245) = 2.44
        Let mProbability(245) = "0.806"
        Let mThreshold(246) = 2.45
        Let mProbability(246) = "0.807"
        Let mThreshold(247) = 2.46
        Let mProbability(247) = "0.808"
        Let mThreshold(248) = 2.47
        Let mProbability(248) = "0.809"
        Let mThreshold(249) = 2.48
        Let mProbability(249) = "0.81"
        Let mThreshold(250) = 2.49
        Let mProbability(250) = "0.811"
        Let mThreshold(251) = 2.5
        Let mProbability(251) = "0.812"
        Let mThreshold(252) = 2.51
        Let mProbability(252) = "0.813"
        Let mThreshold(253) = 2.52
        Let mProbability(253) = "0.814"
        Let mThreshold(254) = 2.54
        Let mProbability(254) = "0.815"
        Let mThreshold(255) = 2.55
        Let mProbability(255) = "0.816"
        Let mThreshold(256) = 2.56
        Let mProbability(256) = "0.817"
        Let mThreshold(257) = 2.57
        Let mProbability(257) = "0.818"
        Let mThreshold(258) = 2.58
        Let mProbability(258) = "0.819"
        Let mThreshold(259) = 2.59
        Let mProbability(259) = "0.82"
        Let mThreshold(260) = 2.6
        Let mProbability(260) = "0.821"
        Let mThreshold(261) = 2.61
        Let mProbability(261) = "0.822"
        Let mThreshold(262) = 2.62
        Let mProbability(262) = "0.823"
        Let mThreshold(263) = 2.63
        Let mProbability(263) = "0.824"
        Let mThreshold(264) = 2.64
        Let mProbability(264) = "0.825"
        Let mThreshold(265) = 2.65
        Let mProbability(265) = "0.826"
        Let mThreshold(266) = 2.66
        Let mProbability(266) = "0.827"
        Let mThreshold(267) = 2.68
        Let mProbability(267) = "0.828"
        Let mThreshold(268) = 2.69
        Let mProbability(268) = "0.829"
        Let mThreshold(269) = 2.7
        Let mProbability(269) = "0.83"
        Let mThreshold(270) = 2.71
        Let mProbability(270) = "0.831"
        Let mThreshold(271) = 2.72
        Let mProbability(271) = "0.832"
        Let mThreshold(272) = 2.73
        Let mProbability(272) = "0.833"
        Let mThreshold(273) = 2.74
        Let mProbability(273) = "0.834"
        Let mThreshold(274) = 2.75
        Let mProbability(274) = "0.835"
        Let mThreshold(275) = 2.77
        Let mProbability(275) = "0.836"
        Let mThreshold(276) = 2.78
        Let mProbability(276) = "0.837"
        Let mThreshold(277) = 2.79
        Let mProbability(277) = "0.838"
        Let mThreshold(278) = 2.8
        Let mProbability(278) = "0.839"
        Let mThreshold(279) = 2.81
        Let mProbability(279) = "0.84"
        Let mThreshold(280) = 2.82
        Let mProbability(280) = "0.841"
        Let mThreshold(281) = 2.84
        Let mProbability(281) = "0.842"
        Let mThreshold(282) = 2.85
        Let mProbability(282) = "0.843"
        Let mThreshold(283) = 2.86
        Let mProbability(283) = "0.844"
        Let mThreshold(284) = 2.87
        Let mProbability(284) = "0.845"
        Let mThreshold(285) = 2.88
        Let mProbability(285) = "0.846"
        Let mThreshold(286) = 2.89
        Let mProbability(286) = "0.847"
        Let mThreshold(287) = 2.91
        Let mProbability(287) = "0.848"
        Let mThreshold(288) = 2.92
        Let mProbability(288) = "0.849"
        Let mThreshold(289) = 2.93
        Let mProbability(289) = "0.85"
        Let mThreshold(290) = 2.94
        Let mProbability(290) = "0.851"
        Let mThreshold(291) = 2.95
        Let mProbability(291) = "0.852"
        Let mThreshold(292) = 2.97
        Let mProbability(292) = "0.853"
        Let mThreshold(293) = 2.98
        Let mProbability(293) = "0.854"
        Let mThreshold(294) = 2.99
        Let mProbability(294) = "0.855"
        Let mThreshold(295) = 3
        Let mProbability(295) = "0.856"
        Let mThreshold(296) = 3.02
        Let mProbability(296) = "0.857"
        Let mThreshold(297) = 3.03
        Let mProbability(297) = "0.858"
        Let mThreshold(298) = 3.04
        Let mProbability(298) = "0.859"
        Let mThreshold(299) = 3.05
        Let mProbability(299) = "0.86"
        Let mThreshold(300) = 3.07
        Let mProbability(300) = "0.861"
        Let mThreshold(301) = 3.08
        Let mProbability(301) = "0.862"
        Let mThreshold(302) = 3.09
        Let mProbability(302) = "0.863"
        Let mThreshold(303) = 3.11
        Let mProbability(303) = "0.864"
        Let mThreshold(304) = 3.12
        Let mProbability(304) = "0.865"
        Let mThreshold(305) = 3.13
        Let mProbability(305) = "0.866"
        Let mThreshold(306) = 3.14
        Let mProbability(306) = "0.867"
        Let mThreshold(307) = 3.16
        Let mProbability(307) = "0.868"
        Let mThreshold(308) = 3.17
        Let mProbability(308) = "0.869"
        Let mThreshold(309) = 3.18
        Let mProbability(309) = "0.87"
        Let mThreshold(310) = 3.2
        Let mProbability(310) = "0.871"
        Let mThreshold(311) = 3.21
        Let mProbability(311) = "0.872"
        Let mThreshold(312) = 3.22
        Let mProbability(312) = "0.873"
        Let mThreshold(313) = 3.24
        Let mProbability(313) = "0.874"
        Let mThreshold(314) = 3.25
        Let mProbability(314) = "0.875"
        Let mThreshold(315) = 3.27
        Let mProbability(315) = "0.876"
        Let mThreshold(316) = 3.28
        Let mProbability(316) = "0.877"
        Let mThreshold(317) = 3.29
        Let mProbability(317) = "0.878"
        Let mThreshold(318) = 3.31
        Let mProbability(318) = "0.879"
        Let mThreshold(319) = 3.32
        Let mProbability(319) = "0.88"
        Let mThreshold(320) = 3.34
        Let mProbability(320) = "0.881"
        Let mThreshold(321) = 3.35
        Let mProbability(321) = "0.882"
        Let mThreshold(322) = 3.36
        Let mProbability(322) = "0.883"
        Let mThreshold(323) = 3.38
        Let mProbability(323) = "0.884"
        Let mThreshold(324) = 3.39
        Let mProbability(324) = "0.885"
        Let mThreshold(325) = 3.41
        Let mProbability(325) = "0.886"
        Let mThreshold(326) = 3.42
        Let mProbability(326) = "0.887"
        Let mThreshold(327) = 3.44
        Let mProbability(327) = "0.888"
        Let mThreshold(328) = 3.45
        Let mProbability(328) = "0.889"
        Let mThreshold(329) = 3.47
        Let mProbability(329) = "0.89"
        Let mThreshold(330) = 3.48
        Let mProbability(330) = "0.891"
        Let mThreshold(331) = 3.5
        Let mProbability(331) = "0.892"
        Let mThreshold(332) = 3.51
        Let mProbability(332) = "0.893"
        Let mThreshold(333) = 3.53
        Let mProbability(333) = "0.894"
        Let mThreshold(334) = 3.54
        Let mProbability(334) = "0.895"
        Let mThreshold(335) = 3.56
        Let mProbability(335) = "0.896"
        Let mThreshold(336) = 3.57
        Let mProbability(336) = "0.897"
        Let mThreshold(337) = 3.59
        Let mProbability(337) = "0.898"
        Let mThreshold(338) = 3.61
        Let mProbability(338) = "0.899"
        Let mThreshold(339) = 3.62
        Let mProbability(339) = "0.9"
        Let mThreshold(340) = 3.64
        Let mProbability(340) = "0.901"
        Let mThreshold(341) = 3.65
        Let mProbability(341) = "0.902"
        Let mThreshold(342) = 3.67
        Let mProbability(342) = "0.903"
        Let mThreshold(343) = 3.69
        Let mProbability(343) = "0.904"
        Let mThreshold(344) = 3.7
        Let mProbability(344) = "0.905"
        Let mThreshold(345) = 3.72
        Let mProbability(345) = "0.906"
        Let mThreshold(346) = 3.74
        Let mProbability(346) = "0.907"
        Let mThreshold(347) = 3.75
        Let mProbability(347) = "0.908"
        Let mThreshold(348) = 3.77
        Let mProbability(348) = "0.909"
        Let mThreshold(349) = 3.79
        Let mProbability(349) = "0.91"
        Let mThreshold(350) = 3.81
        Let mProbability(350) = "0.911"
        Let mThreshold(351) = 3.82
        Let mProbability(351) = "0.912"
        Let mThreshold(352) = 3.84
        Let mProbability(352) = "0.913"
        Let mThreshold(353) = 3.86
        Let mProbability(353) = "0.914"
        Let mThreshold(354) = 3.88
        Let mProbability(354) = "0.915"
        Let mThreshold(355) = 3.9
        Let mProbability(355) = "0.916"
        Let mThreshold(356) = 3.91
        Let mProbability(356) = "0.917"
        Let mThreshold(357) = 3.93
        Let mProbability(357) = "0.918"
        Let mThreshold(358) = 3.95
        Let mProbability(358) = "0.919"
        Let mThreshold(359) = 3.97
        Let mProbability(359) = "0.92"
        Let mThreshold(360) = 3.99
        Let mProbability(360) = "0.921"
        Let mThreshold(361) = 4.01
        Let mProbability(361) = "0.922"
        Let mThreshold(362) = 4.03
        Let mProbability(362) = "0.923"
        Let mThreshold(363) = 4.05
        Let mProbability(363) = "0.924"
        Let mThreshold(364) = 4.07
        Let mProbability(364) = "0.925"
        Let mThreshold(365) = 4.09
        Let mProbability(365) = "0.926"
        Let mThreshold(366) = 4.11
        Let mProbability(366) = "0.927"
        Let mThreshold(367) = 4.13
        Let mProbability(367) = "0.928"
        Let mThreshold(368) = 4.15
        Let mProbability(368) = "0.929"
        Let mThreshold(369) = 4.17
        Let mProbability(369) = "0.93"
        Let mThreshold(370) = 4.19
        Let mProbability(370) = "0.931"
        Let mThreshold(371) = 4.21
        Let mProbability(371) = "0.932"
        Let mThreshold(372) = 4.23
        Let mProbability(372) = "0.933"
        Let mThreshold(373) = 4.25
        Let mProbability(373) = "0.934"
        Let mThreshold(374) = 4.28
        Let mProbability(374) = "0.935"
        Let mThreshold(375) = 4.3
        Let mProbability(375) = "0.936"
        Let mThreshold(376) = 4.32
        Let mProbability(376) = "0.937"
        Let mThreshold(377) = 4.34
        Let mProbability(377) = "0.938"
        Let mThreshold(378) = 4.37
        Let mProbability(378) = "0.939"
        Let mThreshold(379) = 4.39
        Let mProbability(379) = "0.94"
        Let mThreshold(380) = 4.41
        Let mProbability(380) = "0.941"
        Let mThreshold(381) = 4.44
        Let mProbability(381) = "0.942"
        Let mThreshold(382) = 4.46
        Let mProbability(382) = "0.943"
        Let mThreshold(383) = 4.49
        Let mProbability(383) = "0.944"
        Let mThreshold(384) = 4.51
        Let mProbability(384) = "0.945"
        Let mThreshold(385) = 4.54
        Let mProbability(385) = "0.946"
        Let mThreshold(386) = 4.56
        Let mProbability(386) = "0.947"
        Let mThreshold(387) = 4.59
        Let mProbability(387) = "0.948"
        Let mThreshold(388) = 4.62
        Let mProbability(388) = "0.949"
        Let mThreshold(389) = 4.64
        Let mProbability(389) = "0.95"
        Let mThreshold(390) = 4.67
        Let mProbability(390) = "0.951"
        Let mThreshold(391) = 4.7
        Let mProbability(391) = "0.952"
        Let mThreshold(392) = 4.73
        Let mProbability(392) = "0.953"
        Let mThreshold(393) = 4.76
        Let mProbability(393) = "0.954"
        Let mThreshold(394) = 4.79
        Let mProbability(394) = "0.955"
        Let mThreshold(395) = 4.82
        Let mProbability(395) = "0.956"
        Let mThreshold(396) = 4.85
        Let mProbability(396) = "0.957"
        Let mThreshold(397) = 4.88
        Let mProbability(397) = "0.958"
        Let mThreshold(398) = 4.91
        Let mProbability(398) = "0.959"
        Let mThreshold(399) = 4.94
        Let mProbability(399) = "0.96"
        Let mThreshold(400) = 4.97
        Let mProbability(400) = "0.961"
        Let mThreshold(401) = 5.01
        Let mProbability(401) = "0.962"
        Let mThreshold(402) = 5.04
        Let mProbability(402) = "0.963"
        Let mThreshold(403) = 5.08
        Let mProbability(403) = "0.964"
        Let mThreshold(404) = 5.11
        Let mProbability(404) = "0.965"
        Let mThreshold(405) = 5.15
        Let mProbability(405) = "0.966"
        Let mThreshold(406) = 5.19
        Let mProbability(406) = "0.967"
        Let mThreshold(407) = 5.22
        Let mProbability(407) = "0.968"
        Let mThreshold(408) = 5.26
        Let mProbability(408) = "0.969"
        Let mThreshold(409) = 5.3
        Let mProbability(409) = "0.97"
        Let mThreshold(410) = 5.35
        Let mProbability(410) = "0.971"
        Let mThreshold(411) = 5.39
        Let mProbability(411) = "0.972"
        Let mThreshold(412) = 5.43
        Let mProbability(412) = "0.973"
        Let mThreshold(413) = 5.48
        Let mProbability(413) = "0.974"
        Let mThreshold(414) = 5.52
        Let mProbability(414) = "0.975"
        Let mThreshold(415) = 5.57
        Let mProbability(415) = "0.976"
        Let mThreshold(416) = 5.62
        Let mProbability(416) = "0.977"
        Let mThreshold(417) = 5.68
        Let mProbability(417) = "0.978"
        Let mThreshold(418) = 5.73
        Let mProbability(418) = "0.979"
        Let mThreshold(419) = 5.78
        Let mProbability(419) = "0.98"
        Let mThreshold(420) = 5.84
        Let mProbability(420) = "0.981"
        Let mThreshold(421) = 5.9
        Let mProbability(421) = "0.982"
        Let mThreshold(422) = 5.97
        Let mProbability(422) = "0.983"
        Let mThreshold(423) = 6.04
        Let mProbability(423) = "0.984"
        Let mThreshold(424) = 6.11
        Let mProbability(424) = "0.985"
        Let mThreshold(425) = 6.18
        Let mProbability(425) = "0.986"
        Let mThreshold(426) = 6.26
        Let mProbability(426) = "0.987"
        Let mThreshold(427) = 6.34
        Let mProbability(427) = "0.988"
        Let mThreshold(428) = 6.44
        Let mProbability(428) = "0.989"
        Let mThreshold(429) = 6.53
        Let mProbability(429) = "0.99"
        Let mThreshold(430) = 6.64
        Let mProbability(430) = "0.991"
        Let mThreshold(431) = 6.76
        Let mProbability(431) = "0.992"
        Let mThreshold(432) = 6.88
        Let mProbability(432) = "0.993"
        Let mThreshold(433) = 7.03
        Let mProbability(433) = "0.994"
        Let mThreshold(434) = 7.2
        Let mProbability(434) = "0.995"
        Let mThreshold(435) = 7.39
        Let mProbability(435) = "0.996"
        Let mThreshold(436) = 7.63
        Let mProbability(436) = "0.997"
        Let mThreshold(437) = 7.94
        Let mProbability(437) = "0.998"
        Let mThreshold(438) = 8.4
        Let mProbability(438) = "0.9985"
        Let mThreshold(439) = 8.43
        Let mProbability(439) = "0.9986"
        Let mThreshold(440) = 8.49
        Let mProbability(440) = "0.9987"
        Let mThreshold(441) = 8.56
        Let mProbability(441) = "0.9988"
        Let mThreshold(442) = 8.63
        Let mProbability(442) = "0.9989"
        Let mThreshold(443) = 8.7
        Let mProbability(443) = "0.999"
        Let mThreshold(444) = 8.79
        Let mProbability(444) = "0.9991"
        Let mThreshold(445) = 8.88
        Let mProbability(445) = "0.9992"
        Let mThreshold(446) = 8.98
        Let mProbability(446) = "0.9993"
        Let mThreshold(447) = 9.1
        Let mProbability(447) = "0.9994"
        Let mThreshold(448) = 9.24
        Let mProbability(448) = "0.9995"
        Let mThreshold(449) = 9.4
        Let mProbability(449) = "0.9996"
        Let mThreshold(450) = 9.59
        Let mProbability(450) = "0.9997"
        Let mThreshold(451) = 9.85
        Let mProbability(451) = "0.9998"
        Let mThreshold(452) = 10.23
        Let mProbability(452) = "0.99985"
        Let mThreshold(453) = 10.26
        Let mProbability(453) = "0.99986"
        Let mThreshold(454) = 10.31
        Let mProbability(454) = "0.99987"
        Let mThreshold(455) = 10.36
        Let mProbability(455) = "0.99988"
        Let mThreshold(456) = 10.42
        Let mProbability(456) = "0.99989"
        Let mThreshold(457) = 10.49
        Let mProbability(457) = "0.9999"
        Let mThreshold(458) = 10.56
        Let mProbability(458) = "0.99991"
        Let mThreshold(459) = 10.64
        Let mProbability(459) = "0.99992"
        Let mThreshold(460) = 10.73
        Let mProbability(460) = "0.99993"
        Let mThreshold(461) = 10.83
        Let mProbability(461) = "0.99994"
        Let mThreshold(462) = 10.94
        Let mProbability(462) = "0.99995"
        Let mThreshold(463) = 11.08
        Let mProbability(463) = "0.99996"
        Let mThreshold(464) = 11.25
        Let mProbability(464) = "0.99997"
        Let mThreshold(465) = 11.48
        Let mProbability(465) = "0.99998"
        Let mThreshold(466) = 11.81
        Let mProbability(466) = "0.999985"
        Let mThreshold(467) = 11.83
        Let mProbability(467) = "0.999986"
        Let mThreshold(468) = 11.88
        Let mProbability(468) = "0.999987"
        Let mThreshold(469) = 11.93
        Let mProbability(469) = "0.999988"
        Let mThreshold(470) = 11.98
        Let mProbability(470) = "0.999989"
        Let mThreshold(471) = 12.04
        Let mProbability(471) = "0.99999"
        Let mThreshold(472) = 12.1
        Let mProbability(472) = "0.999991"
        Let mThreshold(473) = 12.17
        Let mProbability(473) = "0.999992"
        Let mThreshold(474) = 12.25
        Let mProbability(474) = "0.999993"
        Let mThreshold(475) = 12.34
        Let mProbability(475) = "0.999994"
        Let mThreshold(476) = 12.44
        Let mProbability(476) = "0.999995"
        Let mThreshold(477) = 12.56
        Let mProbability(477) = "0.999996"
        Let mThreshold(478) = 12.72
        Let mProbability(478) = "0.999997"
        Let mThreshold(479) = 12.92
        Let mProbability(479) = "0.999998"
        Let mThreshold(480) = 13.22
        Let mProbability(480) = "0.999999"
        Let mThreshold(481) = 13.23
        Let mProbability(481) = ">.999999"

End Sub

Private Sub mnuWhatIsMagri_Click()
    MsgBox "The Magri update rule, invented by Giorgio Magri, promotes winner-preferring constraints less if there are many of them.  You can read about it in " + _
        "Giorgio Magri (2012) Convergence of error-driven ranking algorithms. Phonology, 29.2: 213-269, 2012."
End Sub


