VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "OTSoft 2.3"
   ClientHeight    =   8445
   ClientLeft      =   2880
   ClientTop       =   1485
   ClientWidth     =   10785
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   8445
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Choose framework"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   360
      TabIndex        =   17
      Top             =   2520
      Width           =   3735
      Begin VB.OptionButton optConstraintDemotion 
         Caption         =   "Classical OT"
         Height          =   435
         Left            =   120
         TabIndex        =   25
         ToolTipText     =   "See What is it?, on the right, for details of this algorithm."
         Top             =   240
         Width           =   2415
      End
      Begin VB.OptionButton optGLA 
         Caption         =   "Stochastic OT"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         ToolTipText     =   "See What is it?, on the right, for details of this algorithm."
         Top             =   1440
         Width           =   2295
      End
      Begin VB.CommandButton cmdIdentifyCD 
         Caption         =   "What is it?"
         Height          =   195
         Left            =   2640
         TabIndex        =   23
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdIdentifyGLA 
         Caption         =   "What is it?"
         Height          =   195
         Left            =   2640
         TabIndex        =   22
         Top             =   1440
         Width           =   975
      End
      Begin VB.OptionButton optMaximumEntropy 
         Caption         =   "Maximum Entropy"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   "See What is it?, on the right, for details of this algorithm."
         Top             =   720
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.CommandButton cmdIdentifyMaximumEntropy 
         Caption         =   "What is it?"
         Height          =   195
         Left            =   2640
         TabIndex        =   20
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton optNoisyHarmonicGrammar 
         Caption         =   "Noisy Harmonic Grammar"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton cmdIdentifyNHG 
         Caption         =   "What is it?"
         Height          =   195
         Left            =   2640
         TabIndex        =   18
         Top             =   1080
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdRank 
      Caption         =   "Compute ranking"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      TabIndex        =   0
      ToolTipText     =   "Rank constraints using the algorithm selected below"
      Top             =   1080
      Width           =   3735
   End
   Begin VB.CommandButton cmdFacType 
      Caption         =   "Compute factorial typology"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   8040
      TabIndex        =   15
      ToolTipText     =   "Compute the factorial typology of your constraint set"
      Top             =   240
      Width           =   2655
   End
   Begin VB.Frame frmArguments 
      Caption         =   "Ranking argumentation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4560
      TabIndex        =   10
      Top             =   1800
      Width           =   3255
      Begin VB.CheckBox chkMiniTableaux 
         Caption         =   "Include illustrative minitableaux"
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   1320
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CheckBox chkDetailedArguments 
         Caption         =   "Show details of argumentation"
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   960
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CheckBox chkMostInformativeBasis 
         Caption         =   "Use Most Informative Basis"
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   600
         Width           =   2655
      End
      Begin VB.CheckBox chkArguerOn 
         Caption         =   "Include ranking arguments"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Value           =   1  'Checked
         Width           =   2775
      End
   End
   Begin VB.CheckBox chkDiagnosticTableaux 
      Caption         =   "Diagnostics if ranking fails"
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      ToolTipText     =   "OTSoft will print out various diagnostics to help you figure out what is wrong with your constraints or entered violations"
      Top             =   3840
      Value           =   1  'Checked
      Width           =   2655
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options for crowded tableaux"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   8040
      TabIndex        =   5
      Top             =   2160
      Width           =   2655
      Begin VB.OptionButton optNeverSwitchAxes 
         Caption         =   "Never switch axes"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   2175
      End
      Begin VB.OptionButton optSwitchSomeAxes 
         Caption         =   "Switch axes where needed"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   2295
      End
      Begin VB.OptionButton optSwitchAll 
         Caption         =   "Switch axes for all tableaux "
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   2295
      End
   End
   Begin VB.TextBox txtViewOutput 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   7080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   675
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Exit this program"
      Top             =   7320
      Width           =   10575
   End
   Begin VB.CommandButton cmdViewResults 
      Caption         =   "&View Results (see View menu for options)"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "View the results; see View menu above for details."
      Top             =   6600
      Width           =   10575
   End
   Begin VB.Label Label1 
      Caption         =   "The ""run"" button. Pick a framework then click here:"
      Height          =   495
      Left            =   480
      TabIndex        =   26
      Top             =   480
      Width           =   3495
   End
   Begin VB.Line Line4 
      X1              =   10680
      X2              =   10680
      Y1              =   6480
      Y2              =   4920
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   10680
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   120
      Y1              =   4920
      Y2              =   6480
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   10680
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "To work with a different file, drag it into this box:"
      Height          =   255
      Left            =   4680
      TabIndex        =   16
      Top             =   4560
      Width           =   3855
   End
   Begin VB.Shape Shape2 
      Height          =   4575
      Left            =   120
      Top             =   240
      Width           =   4215
   End
   Begin VB.Label lblProgressWindow 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "(This window will show progress while the program runs.)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4560
      TabIndex        =   3
      Top             =   240
      Width           =   3255
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuReload 
         Caption         =   "&Reload new version of xxx"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu mnuSaveAsIn 
         Caption         =   "Save as .&in file"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSaveAsTxt 
         Caption         =   "Save as .&txt file"
      End
      Begin VB.Menu mnuSaveAsPraat 
         Caption         =   "Save as &Praat file"
      End
      Begin VB.Menu mnuSaveAsR 
         Caption         =   "Save as &R file (logistic regression)"
      End
      Begin VB.Menu mnuExitWithoutSaving 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuSepFile 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenRecent 
         Caption         =   "Open &recent"
         Begin VB.Menu mnuOpenRecent1 
            Caption         =   "1"
         End
         Begin VB.Menu mnuOpenRecent2 
            Caption         =   "2"
         End
         Begin VB.Menu mnuOpenRecent3 
            Caption         =   "3"
         End
         Begin VB.Menu mnuOpenRecent4 
            Caption         =   "4"
         End
         Begin VB.Menu mnuOpenRecent5 
            Caption         =   "5"
         End
         Begin VB.Menu mnuOpenRecent6 
            Caption         =   "6"
         End
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCurrentFile 
         Caption         =   "&Edit current file"
      End
   End
   Begin VB.Menu mnuWordProcessorChoice 
      Caption         =   "&View"
      Begin VB.Menu mnuViewHere 
         Caption         =   "&View result in OTSoft"
      End
      Begin VB.Menu mnuViewCuston 
         Caption         =   "View with your &word processor"
      End
      Begin VB.Menu mnuViewAsWebPage 
         Caption         =   "View result as web & page"
      End
      Begin VB.Menu mnuPrepareForPrinting 
         Caption         =   "Prepare for &printing (MS Word)"
      End
      Begin VB.Menu mnuSepView 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewHasseDiagramII 
         Caption         =   "View &Hasse diagram of rankings"
      End
      Begin VB.Menu mnuShowHowRankingWasDone 
         Caption         =   "Show &how ranking was done"
      End
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "&Print"
      Begin VB.Menu mnuPrintMenuPrepareForPrinting 
         Caption         =   "Prepare for quality printing (Microsoft Word)"
      End
      Begin VB.Menu mnuDraftPrint 
         Caption         =   "Draft print (Microsoft Word not needed)"
      End
   End
   Begin VB.Menu mnuFactoricalTypology 
      Caption         =   "&Factorial Typology"
      Begin VB.Menu mnuIncludeRankingInFTResults 
         Caption         =   "Include &rankings in results"
      End
      Begin VB.Menu mnuIncludeTableaux 
         Caption         =   "Include tableaux in results"
      End
      Begin VB.Menu mnuFTSumFile 
         Caption         =   "Generate compact factorial typology summary file"
      End
      Begin VB.Menu mnuCompactFTFile 
         Caption         =   "Compact file collapsing neutralized outputs"
      End
      Begin VB.Menu mnuSepFacType 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewCompactFactorialTypologySummaryFile 
         Caption         =   "View compact factorial typology summary file"
      End
      Begin VB.Menu mnuViewCompactFileCollapsingNeutralizedOutputs 
         Caption         =   "View compact file collapsing neutralized outputs"
      End
   End
   Begin VB.Menu mnuAPrioriRankings 
      Caption         =   "&A Priori Rankings"
      Begin VB.Menu mnuConstrainAlgorithmsByAPrioriRankings 
         Caption         =   "Rank constraints constrained by a priori rankings"
      End
      Begin VB.Menu mnuSepAPriori 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTemplateForAprioriRankings 
         Caption         =   "Make or edit a file containing a priori rankings"
      End
      Begin VB.Menu mnuSaveAPrioriRankings 
         Caption         =   "Use strata obtained in ranking to construct a priori ranking file"
      End
   End
   Begin VB.Menu mnuHasse 
      Caption         =   "&Hasse"
      Begin VB.Menu mnuViewHasseDiagram 
         Caption         =   "View Hasse diagram"
      End
      Begin VB.Menu mnuSepHasse 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditHasseTextFile 
         Caption         =   "Edit the source file underlying Hasse diagram"
      End
      Begin VB.Menu mnuReplotHasse 
         Caption         =   "Replot Hasse diagram from altered source file"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuSaveAsTxtSortedByRank 
         Caption         =   "On ranking, save copy of input file, &sorted by rank"
      End
      Begin VB.Menu mnuSmallCaps 
         Caption         =   "Print constraint names in small caps"
      End
      Begin VB.Menu mnuDeleteTmpFiles 
         Caption         =   "Delete temporary files on exit"
      End
      Begin VB.Menu mnuLowFaithfulness 
         Caption         =   "Use the Low Faithfulness version of RCD"
      End
      Begin VB.Menu mnuBiasedConstraintDemotion 
         Caption         =   "Use Biased Constraint Demotion"
      End
      Begin VB.Menu mnuSpecificBCD 
         Caption         =   "BCD favors specific Faithfulness constraints"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSortCandidatesByHarmony 
         Caption         =   "Sort candidates in tableaux by harmony"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuRestoreDefaultSettings 
         Caption         =   "Restore &default settings"
      End
      Begin VB.Menu mnuEditAuxiliary 
         Caption         =   "Edit OTSoftAuxiliarySoftwareLocations.txt"
      End
   End
   Begin VB.Menu mnuHTML 
      Caption         =   "&HTML"
      Begin VB.Menu mnuHTMLOptions 
         Caption         =   "&Options for HTML output"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Hel&p"
      Begin VB.Menu mnuOpenHelpAsPDF 
         Caption         =   "View manual as Adobe PDF file"
      End
      Begin VB.Menu mnuSepHelp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAboutOTSoft 
         Caption         =   "&About OTSoft"
      End
   End
   Begin VB.Menu mnuReturnToMainMenu 
      Caption         =   "&Return to main menu"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================OTSOFT 2.7====================================
'============Software package for Optimality Theory and related frameworks============
'=============================Written by Bruce Hayes, UCLA============================
'===================with contributions by Kie Zuraw and Bruce Tesar===================

   Option Explicit

      Const Yes As Long = 1
      Const No As Long = 0
          
    'Learning data and constraints.  These are module level variables, which get passed
    '   to peripheral modules.
        Dim mMaximumNumberOfForms As Long
        Dim mNumberOfForms As Long
        Dim mInputForm() As String
        Dim mWinner() As String
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
    
    'Files to print to.
        Dim mTmpFile As Long       'Quickie, casual formatting.
        Dim mDocFile As Long       'For conversion to pretty output in Word.
        Dim mHTMFile As Long       'HMTL display.
    
    'To hold the output of Constraint Demotion.
        Private Type DiscreteRankingResult
            Converged As Boolean
            NumberOfStrata As Long
            Stratum(1000) As Long
        End Type
        Dim TSResult As DiscreteRankingResult
        
    'Variables for inputting the data file:
        Dim mUserName As String
        Dim mRecentlyOpenedFiles(6) As String
  
    'Factorial typology variables:
        Public FactorialTypologyAlreadyRunOnThisFile As Boolean
        Public RunningFactorialTypology As Boolean
        
    'Variables for keeping Faithfulness low:
        Dim Subset() As Boolean
        Dim Subsetted() As Boolean   'Excluded from current stratum by the existence of
                                     '  a more specific Faithfulness constraint.

    'A priori rankings (only for archaic input file type; else global)
        Dim mAPrioriRankingsList() As Long
        Dim mNumberOfAPrioriRankings As Long
        Dim mMaximumNumberOfAPrioriRankings As Long
    
    'Flags:
        Dim SaveStrataAsAPrioriRankings As Boolean  'Make a file showing learned strata as a priori rankings.
        Dim ReplotFirst As Boolean                  'Replot a Hasse diagram before displaying it.
        Public mMoreThanOneWinner As Boolean            'Is this a file with more than one winner?
        Dim mGlobalTie As Boolean                    'Is this a file with more than one winner bearing the maximum frequency value?
        Dim mPraatFileFlag As Boolean                'Don't do sorting if it's a Praat file you're trying to make.  This should be more general.

    'File number of self-progress monitoring file.
        Dim ShowMe As Long
        
    'Global flags for remembering if Prince and Tesar's BCD has made an arbitrary
    '  decision on ties.
        Dim UpperTieFlag As Boolean
        Dim TieFlag As Boolean
    
    'Choice of viewing apparatus
        Dim EditorChoice As String
        

'=================================INTERFACE ITEMS==================================
'==================================================================================

Private Sub Form_Load()

    'Do all things that have to be done, and done early.
    
        Dim RunningOTSoftForFirstTime As Boolean
    
    'Accommodate people with tiny screens.
        If Me.Width > Screen.Width - 400 Then Let Me.Width = Screen.Width - 400
        If Me.Height > Screen.Height - 400 Then Let Me.Height = Screen.Height - 400
        
    'Center the form on this user's particular screen.
        Let Me.Left = (Screen.Width - Width) / 2
        Let Me.Top = (Screen.Height - Height) / 2
        
    'Establish where you can read saved user information by invoking the boolean
    '   function FindSavePlace.  The value returned will also tell you if
    '   this is the first time that this copy of OTSoft has been run.
        If FindSafePlace = False Then
            Let RunningOTSoftForFirstTime = True
        Else
            Let RunningOTSoftForFirstTime = False
        End If
   
    'Start the whole business by learning the current file name, so you can label
    '   the buttons.
        'OTSoftRememberUserChoices.txt has the last file name worked on.
            Call ReadOTSoftIni
        'Learn where are the Windows programs you must interact with.
            Call ReadSoftwareLocations
        'If the user opened OTSoft by right clicking on a file name, use that file name.
            Call LetWindowsDictateTheFile("")
        'But if it's the very first run of the program, you need to let them start with
        '   a safe file.
            If RunningOTSoftForFirstTime = True Then
                Call GiveUserAFileToRun
                'And let the user override if they wish, from now on:
                Let RunningOTSoftForFirstTime = False
            End If
        
    'Put the file name on the command buttons, alerting user to how to find a file.
        Let cmdRank.Caption = "Rank " + gFileName + gFileSuffix
        Let cmdFacType.Caption = "Factorial typology for " + gFileName + gFileSuffix
        Let Me.Caption = "OTSoft " + gMyVersionNumber + " - " + gInputFilePath + gFileName + gFileSuffix
        'And on the Edit menu.
            Let mnuEditCurrentFile.Caption = "Edit the input file " + gFileName + gFileSuffix
                       
    'Update the Open Recent menu items.
        Call UpdateOpenRecentMenu
        
    'Label the View Results button appropriately
        Let cmdViewResults.Caption = "View Results"
        
    'Record that the program has not yet been run.
    '    Let gHasTheProgramBeenRun = False
    
    'Record that factorial typology has not yet been calculated, and is
    '   not currently being calculated.
        Let FactorialTypologyAlreadyRunOnThisFile = False
        Let RunningFactorialTypology = False
        
    'Interface memory.
        Let TestWugOnly = False
        
    'Establish the form of web tables (can't do this as constant, for some reason.)
       ' Let gHTMLTableSpecs = "<TABLE BORDER CELLSPACING=" + Chr(34) + "1" + Chr(34) + "CELLPADDING = " + Chr(34) + "; 1; " + Chr(34) + " > "
        
        
End Sub
      
Function FindSafePlace() As Boolean

    'In the cruel modern world, you often can't write to App.Path.
    '   We'll look around the hard disk for other possibilities.
    '   If there was a safe refuge set up earlier, we will know
    '   because it has OTSoftRememberUserChoices.txt in it.
    
    'This routine returns True if OTSoftRememberUserChoices.txt is found in the safe place,
    '   False if not.

    On Error GoTo CheckError
    
    Dim DirString As String     'Remember whether you found OTSoftRememberUserChoices.txt; null if no.

    If Dir("c:\windows\temp\*.*") <> "" Then
        Let gSafePlaceToWriteTo = "c:\windows\temp"
        Let DirString = Dir("c:\windows\temp\OTSoftRememberUserChoices.txt")
    ElseIf Dir("c:\winnt\temp\*.*") <> "" Then
        Let gSafePlaceToWriteTo = "c:\winnt\temp"
        Let DirString = Dir("c:\winnt\temp\OTSoftRememberUserChoices.txt")
    ElseIf Dir("d:\*.*") <> "" Then
        Let gSafePlaceToWriteTo = "d:"
        Let DirString = Dir("d:\OTSoftRememberUserChoices.txt")
    ElseIf Dir("c:\*.*") <> "" Then
        Let gSafePlaceToWriteTo = "c:"
        Let DirString = Dir("c:\OTSoftRememberUserChoices.txt")
    Else
        Let gSafePlaceToWriteTo = App.Path               'A desperation move.
        Let DirString = Dir(App.Path + "\OTSoftRememberUserChoices.txt")
    End If
    
    If DirString = "" Then
        Let FindSafePlace = False
    Else
        Let FindSafePlace = True
    End If
    
    'Debug:  what if it's first time?
    '    Let gSafePlaceToWriteTo = App.Path
    '    Let FindSafePlace = False
    
    Exit Function
    
CheckError:
    'Probably, you're doomed, but you can at least try App.path to see if it works.
        Let gSafePlaceToWriteTo = App.Path
        Let FindSafePlace = False
        
End Function

Sub GiveUserAFileToRun()
    
    'If it's the first time the user has used OTSoft, we need to put
    '    file up for running.  But it had better not be in App.path, since
    '    in many systems this can't be written to.
    
    On Error GoTo CheckError
    
    FileCopy App.Path + "\TinyIllustrativeFile.xls", gSafePlaceToWriteTo + "\TinyIllustrativeFile.xls"
    FileCopy App.Path + "\TinyIllustrativeFile.txt", gSafePlaceToWriteTo + "\TinyIllustrativeFile.txt"
    
    Let gFileName = "TinyIllustrativeFile"
    Let gFileSuffix = ".txt"
    Let gInputFilePath = gSafePlaceToWriteTo + "\"
    Let gOutputFilePath = gInputFilePath + "FilesFor" + gFileName + "\"
    
CheckError:
    
    Exit Sub
    
    
End Sub


Sub mnuAboutOTSoft_Click()
    frmAboutOTSoft.Show
End Sub





'-----------------------------------The menus in order---------------------------------------------

'--------------------------------------The File Menu-----------------------------------------------

Private Sub mnuOpen_Click()

    On Error GoTo ErrHandler
    
    Dim MyPref As Integer
    
    'Since you're doing something new, clear the progress window.
        Let lblProgressWindow.Caption = ""
    
    'The Common Dialog is now suspended, due to causing terrible problems in installation and compatibility.
        'Alert the user and ask if they want to quit.
            Let MyPref = _
                MsgBox("To open a new file in OTSoft 2.6, drag the file onto the program itself, or into the space designated on the program interface." + vbCr + vbLf + "Click Ok to continue, Cancel to exit OTSoft.", vbOKCancel)
        'Act accordingly.
            Select Case MyPref
                Case vbOK
                    Exit Sub
                Case vbCancel
                    Close
                    End
            End Select
        'Close up.
    
ErrHandler:
    'User pressed Cancel button
        Exit Sub
            
End Sub


Private Sub mnuSaveAs_Click()

    'Converted to user warning due to CommonDialog problem.
    
    On Error GoTo ErrHandler
    
    MsgBox "This menu item is deactivated, sorry.  However, a file of the form FileNameBackup.txt is automatically installed in the output folder for your input file, whenever OTSoft is run, " + _
        "and you can move this file wherever you like.  Click OK to continue."
    Exit Sub
        
ErrHandler:
    'User pressed Cancel button
        Exit Sub

End Sub

Private Sub mnuSaveAsPraat_Click()
    Call SaveAs.SaveAsPraat(gFileName, False, mWinner(), mNumberOfConstraints, mConstraintName(), mNumberOfForms, mInputForm(), mNumberOfRivals(), mWinnerViolations(), mWinnerFrequency(), mRival(), mRivalViolations(), _
        mRivalFrequency())
End Sub

Private Sub mnuSaveAsR_Click()
    Call SaveAs.SaveAsR(gFileName, False, mWinner(), mNumberOfConstraints, mConstraintName(), mNumberOfForms, mInputForm(), mNumberOfRivals(), mWinnerViolations(), mWinnerFrequency(), mRival(), mRivalViolations(), _
        mRivalFrequency())
End Sub

Private Sub mnuSaveAsTxt_Click()
    Call SaveAsTxt(gFileName, False, mWinner())
End Sub

Private Sub mnuExitWithoutSaving_Click()
    Call cmdExit_Click
End Sub


Private Sub mnuReload_Click()
    
    'You've got the right file name, just reload it, since user has changed it.
    
    'Make sure you've got the normal interface.
        Call mnuReturnToMainMenu_Click
    
    'Open this file.
        Close
        Call DigestTheInputFile(gInputFilePath, gFileName, gFileSuffix)  'KZ: DigestTheInputFile is now a boolean function

    'Reset the flag indicating that this file hasn't been processed yet.
        Let gHasTheProgramBeenRun = False

End Sub

'The Open Recent menu items.  If user clicks, then let them be the file name.

    'File names need to be parsed into path, filename, and suffix.  The routine
    '   LetWindowsDictateTheFile(), written earlier, does the work of parsing
    '   these strings.

Private Sub mnuOpenRecent1_Click()
    Call LetWindowsDictateTheFile(mRecentlyOpenedFiles(1))
    Let cmdFacType.Caption = "Factorial typology for " + gFileName + gFileSuffix
    Let cmdRank.Caption = "Rank " + gFileName + gFileSuffix
    Let Me.Caption = "OTSoft " + gMyVersionNumber + " - " + gInputFilePath + gFileName + gFileSuffix
    Call RefreshRecentlyOpenedFiles(gInputFilePath + gFileName + gFileSuffix)
End Sub
Private Sub mnuOpenRecent2_Click()
    Call LetWindowsDictateTheFile(mRecentlyOpenedFiles(2))
    Let cmdFacType.Caption = "Factorial typology for " + gFileName + gFileSuffix
    Let cmdRank.Caption = "Rank " + gFileName + gFileSuffix
    Let Me.Caption = "OTSoft " + gMyVersionNumber + " - " + gInputFilePath + gFileName + gFileSuffix
    Call RefreshRecentlyOpenedFiles(gInputFilePath + gFileName + gFileSuffix)
End Sub
Private Sub mnuOpenRecent3_Click()
    Call LetWindowsDictateTheFile(mRecentlyOpenedFiles(3))
    Let cmdFacType.Caption = "Factorial typology for " + gFileName + gFileSuffix
    Let cmdRank.Caption = "Rank " + gFileName + gFileSuffix
    Let Me.Caption = "OTSoft " + gMyVersionNumber + " - " + gInputFilePath + gFileName + gFileSuffix
    Call RefreshRecentlyOpenedFiles(gInputFilePath + gFileName + gFileSuffix)
End Sub
Private Sub mnuOpenRecent4_Click()
    Call LetWindowsDictateTheFile(mRecentlyOpenedFiles(4))
    Let cmdFacType.Caption = "Factorial typology for " + gFileName + gFileSuffix
    Let cmdRank.Caption = "Rank " + gFileName + gFileSuffix
    Let Me.Caption = "OTSoft " + gMyVersionNumber + " - " + gInputFilePath + gFileName + gFileSuffix
    Call RefreshRecentlyOpenedFiles(gInputFilePath + gFileName + gFileSuffix)
End Sub
Private Sub mnuOpenRecent5_Click()
    Call LetWindowsDictateTheFile(mRecentlyOpenedFiles(5))
    Let cmdFacType.Caption = "Factorial typology for " + gFileName + gFileSuffix
    Let cmdRank.Caption = "Rank " + gFileName + gFileSuffix
    Let Me.Caption = "OTSoft " + gMyVersionNumber + " - " + gInputFilePath + gFileName + gFileSuffix
    Call RefreshRecentlyOpenedFiles(gInputFilePath + gFileName + gFileSuffix)
End Sub
Private Sub mnuOpenRecent6_Click()
    Call LetWindowsDictateTheFile(mRecentlyOpenedFiles(6))
    Let cmdFacType.Caption = "Factorial typology for " + gFileName + gFileSuffix
    Let cmdRank.Caption = "Rank " + gFileName + gFileSuffix
    Let Me.Caption = "OTSoft " + gMyVersionNumber + " - " + gInputFilePath + gFileName + gFileSuffix
    Call RefreshRecentlyOpenedFiles(gInputFilePath + gFileName + gFileSuffix)
End Sub


'-----------------------------------The Edit Menu------------------------------------------


Private Sub mnuEditCurrentFile_Click()
    
    'Call up the editor for this kind of file.
    
    On Error GoTo CheckError
    
    'First, inform the reader if the file to be edited no longer exists.
        If Dir(gInputFilePath + gFileName + gFileSuffix) = "" Then
            MsgBox "Sorry, I can't find the file you're trying to edit, which is:" + _
                Chr(10) + Chr(10) + _
                gInputFilePath + gFileName + gFileSuffix + _
                Chr(10) + Chr(10) + _
                "You can try to locate this file yourself by clicking on the Work With Different File button.", vbExclamation
            Exit Sub
        End If
        
    'Since hardly anybody's Windows computer is set up to edit a file labeled ".in", let's
    '   arrange this:
        If gFileSuffix = ".in" Then
            MsgBox "Sorry, but I'm unable to call up a Windows editor for this file.  Please exit OTSoft and edit the file on your own, then restart.", vbExclamation
            End
        Else
            'The other files can be edited as Windows is set up to edit them.
            Select Case UseWindowsPrograms.TryShellExecute(gInputFilePath + gFileName + gFileSuffix)
                Case gShellExecuteWasSuccessful
                    'Do nothing
                Case Else
                    MsgBox "Sorry, I can't find your input file, which is supposed to be located at" + _
                        Chr(10) + Chr(10) + _
                        gInputFilePath + gFileName + gFileSuffix + _
                        Chr(10) + Chr(10) + _
                        "You may have better luck trying to open this file outside of OTSoft." + _
                        Chr(10) + Chr(10) + _
                        "Click OK to return to the main OTSoft screen.", vbExclamation
                    Exit Sub
            End Select          'How did ShellExecute go?
        End If
    
    'So, at least it exists.  Now, is it already open in Excel?  Try it, with errors
    '   detectable just in case.
    '    AppActivate "Microsoft Excel - " + gInputFilePath + gFileName + gFileSuffix
        
    'Since you're editing the file, the user might be interested in reloading
    '   an updated version.
        Let mnuReload.Caption = "Reload " + gFileName + gFileSuffix
        Let mnuReload.Visible = True
        
        Exit Sub
        
CheckError:
    Select Case Err.Number
        Case 5
            'Excel wasn't already open.  So try opening Excel first.
            Call OpenExcelForEditing
        Case Else
            MsgBox "Program error.  For help please contact bhayes@humnet.ucla, specifying error #95972, and including a copy of your input file.  " + _
                "For now, I suggest you edit your input file, which is located at:" + _
                Chr(10) + Chr(10) + _
                gInputFilePath + gFileName + gFileSuffix + _
                Chr(10) + Chr(10) + _
                "by exiting OTSoft first and opening the file using whatever software you normally use.", vbCritical
            Exit Sub
    End Select
        
End Sub
Private Sub OpenExcelForEditing()

    'Open Excel to edit the user's input file.
    
    On Error GoTo CheckError
    
    'We know from the routine that called this one that the input file exists.
    '   Now, try to open it in Excel.  Use Shell if the user's specified copy of a spreadsheet
    '   program can be found; otherwise use ShellExecute to let Windows look for whatever
    '   is opening .xls files on this computer.
    
        If Dir(gExcelLocation) <> "" Then
            Dim Dummy As Long
            'Note:  chr(34), the quotation mark, must surround all file names in a
            '   Shell command, else spaces will foil it.
                Let Dummy = Shell(gExcelLocation + " " + _
                    Chr(34) + gInputFilePath + gFileName + gFileSuffix + Chr(34), _
                    vbNormalFocus)
        Else
            Select Case UseWindowsPrograms.TryShellExecute(gInputFilePath + gFileName + gFileSuffix)
                Case gShellExecuteWasSuccessful
                    'Do nothing
                Case Else
                    MsgBox "Sorry, I can't find your input file, which is supposed to be located at" + _
                        Chr(10) + Chr(10) + _
                        gInputFilePath + gFileName + gFileSuffix + _
                        Chr(10) + Chr(10) + _
                        "You may have better luck trying to open this file outside of OTSoft." + _
                        Chr(10) + Chr(10) + _
                        "Click OK to return to the main OTSoft screen.", vbExclamation
                    Exit Sub
            End Select          'How did ShellExecute go?
        End If                  'Could the user-specified spreadsheet be found?
        
    'Since you're editing the file, the user might be interested in reloading
    '   an updated version.
        Let mnuReload.Caption = "Reload " + gFileName + gFileSuffix
        Let mnuReload.Visible = True
        
        Exit Sub
        
CheckError:

    MsgBox "Program error.  For help please contact bhayes@humnet.ucla, specifying error #14962, and including a copy of your input file." + _
        Chr(10) + Chr(10) + _
        "For now, I suggest you edit your input file by exiting OTSoft and accessing it directly with your own software." + _
        Chr(10) + Chr(10) + _
        "Click OK to return to the main OTSoft screen.", vbCritical
        
End Sub


'-----------------------------------The View Menu-------------------------------------------

Private Sub mnuViewCuston_Click()
    Let EditorChoice = "User specified"
    Let cmdViewResults.Caption = "View Results"
    Call cmdViewResults_Click
End Sub

Private Sub mnuViewHere_Click()
    Call mnuReturnToMainMenu_Click
    Let EditorChoice = "InternalViewer"
    Let cmdViewResults.Caption = "View Results"
    Call cmdViewResults_Click
End Sub
Private Sub mnuViewAsWebpage_Click()
    Let EditorChoice = "WebPage"
    Let cmdViewResults.Caption = "View Results"
    Call cmdViewResults_Click
End Sub

Private Sub mnuPrepareForPrinting_Click()
    Let EditorChoice = "PrepareForPrinting"
    Let cmdViewResults.Caption = "View Results"
    Call cmdViewResults_Click
End Sub

Private Sub mnuViewHasseDiagramII_Click()
    'Deliberate redundancy.
        Call mnuViewHasseDiagram_Click
End Sub

Private Sub mnuShowHowRankingWasDone_Click()

    'Put up in the display window the file "HowIRanked" + gFileName  + ".txt", showing how ranking was done.
    '   Or, for GLA, the file "FILENAMEFullHistory.xls", in Excel.
    
    On Error GoTo CheckError
    
    Let cmdRank.Enabled = True
    Let cmdFacType.Enabled = True

    If optGLA.Value = True Then
    
        'GLA.  Open "FILENAMEFullHistory.txt" in Excel, if possible.
        
            'Make sure the file exists.
                If Dir(gOutputFilePath + gFileName + "FullHistory.txt") = "" Then
                    MsgBox "I can't find the file needed to display how ranking worked.  Be sure that before you run the GLA, you selected Print File With History of All Actions from the Options menu of the GLA screen." + _
                        Chr(10) + Chr(10) + _
                        "Click OK to return to the main OTSoft screen.", vbExclamation
                    Exit Sub
                End If
                
            'Make sure Excel exists.
                'First:  is it correctly encoded in OTSoftAuxiliarySoftwareLocations.txt?
                If Dir(gExcelLocation) <> "" Then
                    'Should be fine.  Go ahead and try to open.
                        'Note, however:  all file names in Shell commands must be
                        '   surrounded by chr(34), the quotation mark, else spaces
                        '   foil the search.
                        Shell gExcelLocation + " " + _
                            Chr(34) + gOutputFilePath + gFileName + "FullHistory.txt" + Chr(34), _
                            vbNormalFocus
                        Exit Sub
                Else
                    'Not in OTSoftAuxiliarySoftwareLocations.txt.  Try letting Windows find it.
                    Select Case UseWindowsPrograms.TryShellExecute(gOutputFilePath + gFileName + "FullHistory.txt")
                        Case gShellExecuteWasSuccessful
                            'Things are fine; just leave this routine.
                                Exit Sub
                        Case Else
                            'It's no good.  Report an error.
                            MsgBox "In order to show how ranking was done, I need to have a copy of Excel, or some other spreadsheet program, or at least some program that can display tab-separated text. " + _
                                "I'm currently looking for this program in this location:" + _
                                Chr(10) + Chr(10) + _
                                gExcelLocation + _
                                Chr(10) + Chr(10) + _
                                "Please open the file" + _
                                Chr(10) + Chr(10) + _
                                App.Path + "\OTSoftAuxiliarySoftwareLocations.txt" + _
                                Chr(10) + Chr(10) + _
                                "and type in the location of your spreadsheet program on the relevant line.  Then rerun OTSoft." + _
                                Chr(10) + Chr(10) + _
                                "You may also want to try opening the relevant file outside of OTSoft.  It is located at:" + _
                                Chr(10) + Chr(10) + _
                                gOutputFilePath + gFileName + "FullHistory.txt" + _
                                Chr(10) + Chr(10) + _
                                "Click OK to return to the main OTSoft screen.", vbExclamation
                                'Old:  gSafePlaceToWriteTo + "\OTSoftAuxiliarySoftwareLocations.txt"
                            Exit Sub
                    End Select          'How did ShellExecute go?
                End If
                
    Else
    
        'Discrete algorithms.  Open the pure-text file "HowIRanked" + gFileName + ".txt"
        '   and display it internally.
        
            Dim OutFile As Long
            Let OutFile = FreeFile
            Let lblProgressWindow.Caption = ""
           
            'Put the output file in the text window:
            '   I don't think this ever gets bigger than the window can handle.
                
            'The code for filling the OTSoft-internal viewer window seems to be not working
            '   on some new versions of Windows, particularly Windows XP.  So, let's make
            '   the potentially-fatal code that fills the view window into a Boolean
            '   function, returning False if it fails.
            
            'Try to put the output file in the text window:
            
            'Check that the file exists before opening it.
                If Dir(gOutputFilePath + "HowIRanked" + gFileName + ".txt") <> "" Then
                    Open gOutputFilePath + "HowIRanked" + gFileName + ".txt" For Input As #OutFile
                Else
                    MsgBox "Sorry, I can't find your file " + _
                    gOutputFilePath + "HowIRanked" + gFileName + ".txt.  Click OK to continue.", vbExclamation
                    Exit Sub
                End If

            If FillTheOutputWindow(OutFile) = False Then
                'The problem did indeed happen.  So try a backup route.
                    Select Case UseWindowsPrograms.TryShellExecute(gOutputFilePath + "HowIRanked" + gFileName + ".txt")
                        Case gShellExecuteWasSuccessful
                            'Do nothing
                        Case Else
                            MsgBox "Sorry, I'm having trouble displaying your how-ranking-happened file.  You may want to open it from outside OTSoft; it is located at:" + _
                                Chr(10) + Chr(10) + _
                                gOutputFilePath + "HowIRanked" + gFileName + ".txt" + _
                                Chr(10) + Chr(10) + _
                                "Click OK to return to the main OTSoft screen.", vbExclamation
                            Exit Sub
                    End Select
            End If
                
            Close #OutFile
            Exit Sub
            
    End If
    
CheckError:

    Select Case Err.Number
        Case 53
            If optGLA.Value = True Then
            'Can't find Excel:
                MsgBox "In order to show history of ranking, I need to have a copy of Excel, or some other spreadsheet program, or at least some program that can display tab-separated text. " + _
                    "I'm currently looking for this program in this location:" + _
                    Chr(10) + Chr(10) + _
                    gExcelLocation + _
                    Chr(10) + Chr(10) + _
                    "Please open the file" + _
                    App.Path + "\OTSoftAuxiliarySoftwareLocations.txt" + _
                    Chr(10) + Chr(10) + _
                    "and type in the location of your spreadsheet program on the relevant line.  Then rerun OTSoft.", vbExclamation
                    'Old:  gSafePlaceToWriteTo + "\OTSoftAuxiliarySoftwareLocations.txt"
            Else
                'Generic:
                    MsgBox "I'm having trouble showing this material.  To see how ranking proceeded, please use a word processor to open the file " + _
                    Chr(10) + Chr(10) + _
                    gOutputFilePath + "HowIRanked" + gFileName + ".txt" + ".", vbExclamation
            End If
        Case Else
            If optGLA.Value = True Then
                'Can't find Excel file:
                    MsgBox "I'm having trouble showing this material.  To see how ranking proceeded, please use a word processor to open the file " + _
                    Chr(10) + Chr(10) + _
                    gOutputFilePath + gFileName + "FullHistory.txt" + ".", , vbExclamation
            Else
                'Can't find the text file for categorical rankings.
                    MsgBox "I'm having trouble showing this material.  To see how ranking proceeded, please use a word processor to open the file " + _
                    Chr(10) + Chr(10) + _
                    gOutputFilePath + "HowIRanked" + gFileName + ".txt" + ".", vbExclamation
            End If
            Call mnuReturnToMainMenu_Click
        
    End Select
      
End Sub


'-----------------------------------The Print Menu------------------------------------------

Private Sub mnuPrintMenuPrepareForPrinting_Click()
    
    'This is the same menu item as under the File menu, only located under Print.
    Call mnuPrepareForPrinting_Click

End Sub

Private Sub mnuDraftPrint_Click()
        
    'It's perilous having this immediately do its job, because people often click menu
    '   items by mistake.  So instead, just bring up the menu.
        frmPrinting.Show
    
End Sub


'-----------------------------The Factorial Typology Menu------------------------------------

Private Sub mnuIncludeTableaux_Click()
    If mnuIncludeTableaux.Checked = False Then
        Let mnuIncludeTableaux.Checked = True
        'If you include tableaux, you should also include the rankings.
            Let mnuIncludeRankingInFTResults.Checked = True
    Else
        'But it's fine to have rankings but no tableaux.
        Let mnuIncludeTableaux.Checked = False
    End If
End Sub

Private Sub mnuIncludeRankingInFTResults_Click()
    If mnuIncludeRankingInFTResults.Checked = False Then
        Let mnuIncludeRankingInFTResults.Checked = True
    Else
        'If you don't include rankings, it makes no sense to include tableaux either.
            Let mnuIncludeRankingInFTResults.Checked = False
            Let mnuIncludeTableaux.Checked = False
    End If
End Sub

Private Sub mnuFTSumFile_Click()
    If mnuFTSumFile.Checked = False Then
        Let mnuFTSumFile.Checked = True
        Let mnuViewCompactFactorialTypologySummaryFile.Visible = True
        Let mnuSepFacType.Visible = True
    Else
        Let mnuFTSumFile.Checked = False
        Let mnuViewCompactFactorialTypologySummaryFile.Visible = False
        Let mnuSepFacType.Visible = False
    End If
End Sub

Private Sub mnuCompactFTFile_Click()
    If mnuCompactFTFile.Checked = False Then
        Let mnuCompactFTFile.Checked = True
        Let mnuViewCompactFileCollapsingNeutralizedOutputs.Visible = True
        Let mnuSepFacType.Visible = True
    Else
        Let mnuCompactFTFile.Checked = False
        Let mnuViewCompactFileCollapsingNeutralizedOutputs.Visible = False
        Let mnuSepFacType.Visible = False
    End If
End Sub

Private Sub mnuViewCompactFactorialTypologySummaryFile_Click()
    On Error Resume Next
    'Open a file.
        Dim f As Long
        Let f = FreeFile
        TryShellExecute (gOutputFilePath + gFileName + "FTSum" + ".txt")
End Sub

Private Sub mnuViewCompactFileCollapsingNeutralizedOutputs_Click()
    On Error Resume Next
    'Open a file.
        Dim f As Long
        Let f = FreeFile
        TryShellExecute (gOutputFilePath + gFileName + "CompactSum" + ".txt")
End Sub


'-------------------------------The A Priori Rankings Menu----------------------------------------

Private Sub mnuConstrainAlgorithmsByAPrioriRankings_Click()
    If mnuConstrainAlgorithmsByAPrioriRankings.Checked = False Then
        'Check to see that there is a file.
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
        Let mnuConstrainAlgorithmsByAPrioriRankings.Checked = True
        Let lblProgressWindow.Caption = "A priori rankings in effect"
    Else
        Let mnuConstrainAlgorithmsByAPrioriRankings.Checked = False
    End If
End Sub

Private Sub mnuTemplateForAPrioriRankings_Click()
    Call APrioriRankings.PrintOutTemplateForAPrioriRankings(mNumberOfConstraints, mAbbrev())
End Sub

Private Sub mnuSaveAPrioriRankings_Click()
    Let SaveStrataAsAPrioriRankings = True
End Sub


'------------------------------------The Hasse Menu----------------------------------------

Private Sub mnuViewHasseDiagram_Click()

    On Error GoTo CheckError
    
    Let cmdRank.Enabled = True
    Let cmdFacType.Enabled = True
    
    'Avoid this error:  user edits the Hasse, file, then tries to look at
    '   it without replotting.
        If ReplotFirst = True Then
            MsgBox "You've just edited the Hasse input file.  Please select the option Replot Hasse Diagram from the Hasse menu before inspecting your Hasse diagram." + _
            Chr(10) + Chr(10) + _
            "You can view the old Hasse diagram anyway if you like simply by clicking View Hasse diagram again.", vbExclamation
            Let ReplotFirst = False
            Exit Sub
        End If

    'First, recheck if the ATT GraphViz software is installed.
        If Dir(gDotExeLocation) = "" Then
            Select Case MsgBox("OTSoft did not generate a graphics file for the Hasse diagram." + _
                Chr(10) + Chr(10) + _
                "I conjecture that this because OTSoft was unable to access the necessary ATT GraphViz software." + _
                Chr(10) + Chr(10) + _
                "If you have already installed GraphViz, open the file" + _
                Chr(10) + Chr(10) + _
                App.Path + "\OTSoftAuxiliarySoftwareLocations.txt." + _
                Chr(10) + Chr(10) + _
                "and alter it to tell OTSoft where the crucial program of GraphViz, dot.exe, resides on your computer.  " + _
                Chr(10) + Chr(10) + _
                "If you have not installed GraphViz , you can get it for free by downloading it from " + gATTWebSite + "." + _
                Chr(10) + Chr(10) + _
                "Note that you can still view your results without the Hasse diagram." + _
                Chr(10) + Chr(10) + _
                "If you're connected to the Web, click Yes to try to download GraphViz, else click No to return to the main OTSoft screen.", vbYesNo + vbExclamation)
                'Old:  gSafePlaceToWriteTo + "\OTSoftAuxiliarySoftwareLocations.txt
                Case vbYes
                    Call UseWindowsPrograms.OpenWebBrowser(gATTWebSite)
            End Select
                Exit Sub
        End If
    
    'If the user didn't select the arguer, you won't have a new Hasse diagram.
        If HasseDiagramCreated = False Then
            'The issue is whether the user wants to just look at the old Hasse diagram.
            '   Check to see if there is one available.
                If Dir(gOutputFilePath + gFileName + "Hasse.gif") = "" Then
                    MsgBox "You haven't run a suitable ranking algorithm yet, so I can't show you a Hasse diagram." + _
                        Chr(10) + Chr(10) + _
                        "To get a Hasse diagram, either " + _
                        Chr(10) + Chr(10) + _
                        "   Run the Gradual Learning Algorithm on your file, or" + _
                        Chr(10) + _
                        "   Run any of the other ranking algorithms, checking Include Ranking Arguments first.", vbExclamation
                    Exit Sub
                Else
                    'There is a Hasse diagram, but it's old.
                        MsgBox "You haven't run a suitable ranking algorithm yet, so the Hasse diagram I'm about to show, if any, is one already on your disk drive." + _
                            Chr(10) + Chr(10) + _
                            "If you want a fresh Hasse diagram, either " + _
                            Chr(10) + Chr(10) + _
                            "   Run the Gradual Learning Algorithm, requesting Pairwise Ranking Probabilities on the options menu, or" + _
                            Chr(10) + _
                            "   Run any of the other ranking algorithms, checking Include Ranking Arguments first.", vbExclamation
                End If
        End If
        
    'Load the form.
        frmHasse.Show

    Exit Sub

CheckError:

    MsgBox "Program error.  For help please contact bhayes@humnet.ucla, specifying error #45972, and including a copy of your input file.  Click ok to continue without a Hasse diagram.", vbCritical
    Exit Sub

End Sub


Private Sub mnuEditHasseTextFile_Click()
    
    'Call up the text editor to edit the current Hasse file.
    
    On Error GoTo CheckError
    
    'Unload the Hasse window.
        Unload frmHasse
    
    'Avoid the danger of missing files.
        'First, no output file:
            If Dir(gOutputFilePath + gFileName + "hasse.txt") = "" Then
                MsgBox "I can't find the text file " + _
                    Chr(10) + Chr(10) + _
                    gOutputFilePath + gFileName + "Hasse.txt" + _
                    Chr(10) + Chr(10) + _
                    "that underlies your Hasse diagram.  You may have to rank your constraints again in order to obtain this file.", vbExclamation
                    Exit Sub
            End If
        'Second, no word processor.
            If Dir(gUsersWordProcessor) <> "" Then
                'Safe to use the user's word processor.
                '   chr(34), the quote mark, must be used so that Shell will find the path
                '       when the path contains blanks.
                    Dim Dummy As Long
                    Let Dummy = Shell(gUsersWordProcessor + " " + _
                        Chr(34) + gOutputFilePath + gFileName + "Hasse.txt" + Chr(34), _
                        vbNormalFocus)
                    Let ReplotFirst = True
                    Exit Sub
            Else
                'Can't find word processor.  Try letting Windows find it.
                    Select Case UseWindowsPrograms.TryShellExecute(gOutputFilePath + gFileName + "Hasse.txt")
                        Case gShellExecuteWasSuccessful
                            Let ReplotFirst = True
                        Case Else
                            MsgBox "I can't show you the source file for your Hasse diagram because I can't find your word processor." + Chr(10) + Chr(10) + _
                                "Please open the file " + Chr(10) + Chr(10) + _
                                    App.Path + "\OTSoftAuxiliarySoftwareLocations.txt" + Chr(10) + Chr(10) + _
                                    "and type in the location of your word processor.", vbExclamation
                                    'Old:  gSafePlaceToWriteTo + "\OTSoftAuxiliarySoftwareLocations.txt"
                            Exit Sub
                    End Select          'How did ShellExecute go?
            End If                      'Can you find the user-designated word processor?
        
        Exit Sub
        
CheckError:

    MsgBox "Program error.  Please notify Bruce Hayes at bhayes@humnet.ucla.  Kindly include your input file and specify error number 59888.  OTSoft will continue (without editing your Hasse source file) when you click OK." + Chr(10) + Chr(10) + _
        "Note that you can still edit your Hasse source file, located at" + Chr(10) + Chr(10) + _
        gOutputFilePath + gFileName + "Hasse.txt" + Chr(10) + Chr(10) + _
        "by exiting OTSoft and using your word processor in the ordinary way.", vbCritical
    Exit Sub

End Sub

Sub mnuReplotHasse_Click()
    
    'Basically, just call the program, but if something goes wrong, diagnose.
        
    On Error GoTo CheckError
        
        If Dir(gDotExeLocation) = "" Then
            MsgBox "I can't replot your Hasse diagram because I can't find the ATT GraphViz software needed." + _
                Chr(10) + Chr(10) + _
                "Please make sure this software is installed (you can download it from " + gATTWebSite + _
                "), then open the file" + _
                Chr(10) + Chr(10) + _
                App.Path + "\OTSoftAuxiliarySoftwareLocations.txt" + _
                Chr(10) + Chr(10) + _
                "and type in the location of the GraphViz program called dot.exe.", vbExclamation
                'Old:  gSafePlaceToWriteTo + "\OTSoftAuxiliarySoftwareLocations.txt"
            Exit Sub
        End If
        
        Call RunATTDot
        
        'We need to let ATTDot finish, and even its internal code doesn't seem to guarantee this.
        
            Dim MyTimer As Long
            Let MyTimer = Timer
            Do
                If Dir(gOutputFilePath + gFileName + "hasse.gif") = "" Then Exit Do
                If Timer - MyTimer > 30 Then
                    MsgBox "Sorry, I can't plot a Hasse diagram.", vbExclamation
                    'Let it be known that you have not created a Hasse diagram.
                        Let HasseDiagramCreated = False
                    Exit Sub
                End If
                DoEvents
            Loop

        Let HasseDiagramCreated = True
        Let ReplotFirst = False
        
        Exit Sub
        
CheckError:

    MsgBox "Program error.  For help contact Bruce Hayes at bhayes@humnet.ucla.edu, including a copy of your input file, and specifying error #44488.  When you click OK, OTSoft will return to the main menu.", vbCritical
    Exit Sub
        
End Sub


'------------------------------------The Options Menu-----------------------------------------

Private Sub mnuSaveAsTxtSortedByRank_Click()
    'Toggle option for saving sorted input file.
        If mnuSaveAsTxtSortedByRank.Checked = True Then
            Let mnuSaveAsTxtSortedByRank.Checked = False
        Else
            Let mnuSaveAsTxtSortedByRank.Checked = True
        End If
End Sub

Private Sub mnuSmallCaps_Click()
    If mnuSmallCaps.Checked = False Then
        Let mnuSmallCaps.Checked = True
        Let SmallCapTag1 = "\cs"
        Let SmallCapTag2 = "\ce"
    Else
        Let mnuSmallCaps.Checked = False
        Let SmallCapTag1 = ""
        Let SmallCapTag2 = ""
    End If
End Sub

Private Sub mnuLowFaithfulness_Click()
    'Toggle option for used of the Low Faithfulness version of RCD.
        If mnuLowFaithfulness.Checked = True Then
            Let mnuLowFaithfulness.Checked = False
        Else
            Let mnuLowFaithfulness.Checked = True
            'Only one can be used.
                Let mnuBiasedConstraintDemotion.Checked = False
        End If

End Sub

Private Sub mnuBiasedConstraintDemotion_Click()
    'Toggle option for used of the Biased version of RCD.
        If mnuBiasedConstraintDemotion.Checked = True Then
            Let mnuBiasedConstraintDemotion.Checked = False
        Else
            Let mnuBiasedConstraintDemotion.Checked = True
            'Only one can be used.
                Let mnuLowFaithfulness.Checked = False
        End If

End Sub


Private Sub mnuSpecificBCD_Click()
    'Toggle option for requiring BCD to give priority to specific constraints.
        If mnuSpecificBCD.Checked = True Then
            Let mnuSpecificBCD.Checked = False
        Else
            Let mnuSpecificBCD.Checked = True
        End If
End Sub

Private Sub mnuDeleteTmpFiles_Click()
    If mnuDeleteTmpFiles.Checked = False Then
        Let mnuDeleteTmpFiles.Checked = True
    Else
        Let mnuDeleteTmpFiles.Checked = False
    End If
End Sub


Private Sub mnuSortCandidatesByHarmony_Click()
    If mnuSortCandidatesByHarmony.Checked = True Then
        Let mnuSortCandidatesByHarmony.Checked = False
    Else
        Let mnuSortCandidatesByHarmony.Checked = True
    End If
End Sub

Private Sub mnuRestoreDefaultSettings_Click()
    
    On Error GoTo CheckError
    
    Select Case MsgBox("Are you sure you want to restore default settings?", vbYesNo + vbExclamation)
        Case vbYes
            FileCopy App.Path + "\DefaultSettings.ini ", gSafePlaceToWriteTo + "\OTSoftRememberUserChoices.txt"
            Call ReadOTSoftIni
            Call Form_Load
            Exit Sub
        Case Else
            Exit Sub
    End Select
    
CheckError:

    MsgBox "I was unable to restore the default settings.  Try going into the OTSoft" + _
        " folder and copying the contents of DefaultSettings.ini into OTSoftRememberUserChoices.txt.  " + _
        "Sorry for the inconvenience.", vbExclamation
    Exit Sub
    
End Sub


Private Sub mnuEditAuxiliary_Click()
    
    'Call up the text editor to edit OTSoftAuxiliarySoftwareLocations.txt.
    
    On Error GoTo CheckError
    
    'Avoid the danger of missing files.
        'First, no file:
            If Dir(App.Path + "\OTSoftAuxiliarySoftwareLocations.txt") = "" Then
                MsgBox "I can't find the file " + _
                    Chr(10) + Chr(10) + _
                    App.Path + "\OTSoftAuxiliarySoftwareLocations.txt" + _
                    Chr(10) + Chr(10) + _
                    "Please find this file yourself in the OTSoft program folder (it is part of the download package) and use software like Notepad to edit it.", vbExclamation
                    Exit Sub
            End If
        'Second, no word processor.
            If Dir(gUsersWordProcessor) <> "" Then
                'Safe to use the user's word processor.
                    Dim Dummy As Long
                    Let Dummy = Shell(gUsersWordProcessor + " " + _
                        App.Path + "\OTSoftAuxiliarySoftwareLocations.txt", vbNormalFocus)
                    Let ReplotFirst = True
                    Exit Sub
            Else
                'Can't find word processor.  Try letting Windows find it.
                    Select Case UseWindowsPrograms.TryShellExecute(App.Path + "\OTSoftAuxiliarySoftwareLocations.txt")
                        Case gShellExecuteWasSuccessful
                            Let ReplotFirst = True
                        Case Else
                            MsgBox "I can't edit " + App.Path + "\OTSoftAuxiliarySoftwareLocations.txt because I can't find your word processor. You will have to edit the file manually."
                            Exit Sub
                    End Select          'How did ShellExecute go?
            End If                      'Can you find the user-designated word processor?
        
        Exit Sub
        
CheckError:

    MsgBox "Program error.  Please notify Bruce Hayes at bhayes@humnet.ucla.edu"
    Exit Sub

End Sub


'-------------------------------------The HTML Menu-----------------------------------------
Private Sub mnuHTMLOptions_Click()
    Call frmHTMLOptions.Main
End Sub

'-------------------------------------The Help Menu-----------------------------------------

Sub mnuOpenHelpAsPDF_Click()
    
    'We have no idea where the user's PDF reader is.  So use ShellExecute to find the
    '   reader, as whatever this copy of Windows has set up to read .pdf files.
    
    On Error GoTo CheckError
    
    Dim Dummy As Long
    Let Dummy = UseWindowsPrograms.TryShellExecute(App.Path + "\OTSoftManual_" + gMyVersionNumber + ".pdf")
    Select Case Dummy
        Case gShellExecuteWasSuccessful
            'Do nothing
        Case SE_ERR_FNF
            MsgBox "Sorry, I can't find" + _
                Chr(10) + Chr(10) + _
                App.Path + "OTSoftManual.pdf," + _
                Chr(10) + Chr(10) + _
                "which is the PDF version of the manual." + _
                Chr(10) + Chr(10) + _
                "You may have better luck trying to open this file outside of OTSoft." + _
                Chr(10) + Chr(10) + _
                "Click OK to return to the main OTSoft screen.", vbExclamation
        Case SE_NOASSOC, 5  'Five is a Windows error.
             Select Case MsgBox("I can't find your copy of the Adobe Acrobat Reader, needed to read the PDF version of the manual." + _
                Chr(10) + Chr(10) + _
                "If you don't have the Acrobat Reader, download it from" + _
                Chr(10) + Chr(10) + _
                "http://www.adobe.com/products/acrobat/readermain.html" + _
                Chr(10) + Chr(10) + _
                "and try again." + _
                Chr(10) + Chr(10) + _
                "If you do have the Reader, then there is a program error.  Please contact Bruce Hayes at bhayes@humnet.ucla.edu, specifying error #99457." + _
                Chr(10) + Chr(10) + _
                "Click Yes if you would like to try downloading the Adobe Acrobat Reader, No to return to the main OTSoft screen.", vbYesNo + vbExclamation)
                    Case vbYes
                        Call UseWindowsPrograms.OpenWebBrowser("http://www.adobe.com/products/acrobat/readermain.html")
                    Case vbNo
                        Exit Sub
            End Select
                
        Case Else
            MsgBox "Program error.  I can't open the PDF version of the manual.  Please contact Bruce Hayes at bhayes@humnet.ucla.edu, specifying error #95457." + _
                Chr(10) + Chr(10) + _
                "Click OK to return to the main OTSoft screen.", vbExclamation
    End Select
    
    Exit Sub
    
CheckError:
    
    MsgBox "Program error.  For help contact Bruce Hayes at bhayes@humnet.ucla.edu, enclosing a copy of your input file and specifying error #99376." + _
        Chr(10) + Chr(10) + _
        "You can consult the Help file on your own if you like; use the free downloadable PDF reader (http://www.adobe.com/products/acrobat/readermain.html) to open this file:" + _
        Chr(10) + Chr(10) + _
        "   " + App.Path + "OTSoftManual.pdf" + _
        Chr(10) + Chr(10) + _
        "Click OK to return to the main OTSoft screen.", vbExclamation
    Exit Sub
    
End Sub

Sub mnuViewManual_Click()
    
    'Load the manual for inspection using the user's word processor.
        
    On Error GoTo CheckError
    
    Dim Dummy As Long, MsgStr As String
        
    'Make sure files exist before using them.
    
        'The help file itself in .doc form:
        
        If Dir(App.Path + "\OTSoftManual.doc") = "" Then
            'Not there.  Report the problem.
                Let MsgStr = "Currently, Help consists of showing you the manual (which has contents and links)."
                Let MsgStr = MsgStr + Chr(10) + Chr(10) + "But I can't find the manual in its proper place, i.e. "
                Let MsgStr = MsgStr + Chr(10) + Chr(10) + App.Path + "\OTSoftManual.doc"
                Let MsgStr = MsgStr + Chr(10) + Chr(10) + "Please check the folder in which you installed OTSoft; if all else fails, you might try reinstalling OTSoft to fix this problem."
                MsgBox MsgStr, vbExclamation
                Exit Sub
        Else
            'OTSoftManual.doc is available.
            '   Try opening it with the file specified in OTSoftAuxiliarySoftwareLocations.txt.
            '   Note:  chr(34), the quote mark, is needed to get Shell to work around
            '       blanks in the path name.
            If Dir(gUsersWordProcessor) <> "" Then
                Let Dummy = Shell(gUsersWordProcessor + " " + _
                    Chr(34) + App.Path + "\OTSoftManual.doc" + Chr(34), _
                    vbNormalFocus)
                Exit Sub
            Else
                'Couldn't find user's word processor in specified location.
                '   But it still might be possible to find it with ShellExecute.
                    Select Case UseWindowsPrograms.TryShellExecute(App.Path + "\OTSoftManual.doc")
                        Case gShellExecuteWasSuccessful
                            'Do nothing
                        Case SE_ERR_FNF
                            MsgBox "Sorry, I can't find" + _
                                Chr(10) + Chr(10) + _
                                App.Path + "\OTSoftManual.doc," + _
                                Chr(10) + Chr(10) + _
                                "which is the Word version of the manual." + _
                                Chr(10) + Chr(10) + _
                                "You may have better luck trying to open this file outside of OTSoft." + _
                                Chr(10) + Chr(10) + _
                                "Click OK to return to the main OTSoft screen.", vbExclamation
                        Case Else
                            Let MsgStr = "Currently, Help consists of using your word processor to show you the manual."
                            Let MsgStr = MsgStr + Chr(10) + Chr(10) + "But I can't find your word processor for this purpose.  Please open this file:"
                            'Let MsgStr = MsgStr + Chr(10) + Chr(10) + gSafePlaceToWriteTo + "\OTSoftAuxiliarySoftwareLocations.txt"
                            Let MsgStr = MsgStr + Chr(10) + Chr(10) + App.Path + "\OTSoftAuxiliarySoftwareLocations.txt"
                            Let MsgStr = MsgStr + Chr(10) + Chr(10) + "and type in the name and location of your word processor.  Then try Help again."
                            MsgBox MsgStr, vbExclamation
                            Exit Sub
                    End Select          'How did ShellExecute go?
            End If                      'Was word processor specified in OTSoftAuxiliarySoftwareLocations.txt available?
        End If                          'Is OTSoftManual.doc available?
    
    Exit Sub
    
CheckError:

    'Given the prechecks that happened, diagnosis will be hard if execution reaches
    '   this point.  Punt.
    MsgBox "Sorry, I can't display Help.  For assistance contact Bruce Hayes at bhayes@humnet.ucla.edu, specifying error #14773.  " + _
        Chr(10) + Chr(10) + _
        "You may be able to read the Help file simply by opening it in the following location:" + _
        Chr(10) + Chr(10) + _
        App.Path + "\OTSoftManual.doc" + _
        Chr(10) + Chr(10) + _
        "Click OK to return to the main OTSoft menu.", vbExclamation
    
End Sub


'-----------------------------The Return to Main Menu Menu-----------------------------------------

Private Sub mnuReturnToMainMenu_Click()
    
    'Back to normal screen size.  This needs to be fixed to restore resizing capacity. xxx
        Let Form1.WindowState = 0
        Let Width = 10905
        Let Height = 8790
    
    'Center the form on this user's particular screen.
        Let Left = (Screen.Width - Width) / 2
        Let Top = (Screen.Height - Height) / 2

    'Show all the normal controls, rather than the text window/Hasse diagram.
        Let txtViewOutput.Visible = False
        Unload frmHasse
        
        Let frmArguments.Visible = True
        Let Frame2.Visible = True
        Let Frame3.Visible = True
        Let cmdRank.Visible = True
        Let cmdViewResults.Visible = True
        Let cmdFacType.Visible = True
        Let cmdExit.Visible = True
        Let chkDiagnosticTableaux.Visible = True
        
    'Restore hidden menu items.
        Let mnuFactoricalTypology.Visible = True
        Let mnuAPrioriRankings.Visible = True

    
    'Suppress unwanted menu item.
        Let mnuReturnToMainMenu.Visible = False

End Sub

'-----------------------Command Buttons, Option Buttons, Check Boxes-----------------------

'These are in the order seen on the interface itself, going down in columns.

Private Sub cmdRank_Click()
    Call Rank
End Sub

'Relabel the main button by its function.

Private Sub optConstraintDemotion_Click()
    Let cmdRank.Caption = "Rank constraints for " + gFileName + gFileSuffix
End Sub
Private Sub optGLA_Click()
   Let cmdRank.Caption = "Compute ranking values for " + gFileName + gFileSuffix
End Sub
Private Sub optMaximumEntropy_Click()
    Let cmdRank.Caption = "Compute weights for " + gFileName + gFileSuffix
End Sub
Private Sub optNoisyHarmonicGrammar_Click()
    Let cmdRank.Caption = "Compute weights for " + gFileName + gFileSuffix
End Sub

'Next five:  Identify the algorithms for beginners
'   And let them read on the Web to learn about them.
Sub cmdIdentifyCD_Click()
    MsgBox ("This choice covers Classical Optimality Theory, invented 1993 by Tesar and Smolensky. The ranking algorithm used is Recursive Constraint Demotion, invented 1993 by Tesar and Smolensky. This finds a feasible ranking for your constraints if one exists, detects redundant constraints, and can provide help in diagnosing failed constraint sets.References:  Tesar and Smolensky, http://ruccs.rutgers.edu/roa.html #2, 155, 156, etc.")
End Sub
Sub cmdIdentifyGLA_Click()
    MsgBox ("Stochastic OT assigns ranking values to constraints, perturbs ranking values at random, sorts by ranking value, then picks winner as in Classical OT. References can be downloaded from:" + _
        Chr(10) + Chr(10) + _
        "https://brucehayes.org/papers/BoersmaAndHayes2001GLA.pdf")
End Sub
Private Sub cmdIdentifyMaximumEntropy_Click()
    Select Case MsgBox("Maximum Entropy OT is a stochastic evolved version of OT. It assigns a weight to every constraint, and outputs a probability for each candidate.  Reference:  Goldwater and Johnson, https://brucehayes.org/otsoft/pdf/goldwaterjohnson03.pdf" + _
        Chr(10) + Chr(10) + _
        "If you're connected to the Web, click Yes to read a tutorial article by Goldwater and Johnson, else No to return to the main OTSoft screen.", vbYesNo + vbInformation)
        Case vbYes
            Call UseWindowsPrograms.OpenWebBrowser("https://brucehayes.org/otsoft/pdf/goldwaterjohnson03.pdf")
    End Select
End Sub

Private Sub cmdIdentifyNHG_Click()
    Select Case MsgBox("Noisy Harmonic Grammar is a stochastic evolved version of OT. It assigns a weight to every constraint, and outputs a probability for each candidate.  Reference:  Boersma and Pater, https://brucehayes.org/otsoft/pdf/boersma-pater-2013.pdf" + _
        Chr(10) + Chr(10) + _
        "If you're connected to the Web, click Yes to read this article, else No to return to the main OTSoft screen.", vbYesNo + vbInformation)
        Case vbYes
            Call UseWindowsPrograms.OpenWebBrowser("https://brucehayes.org/otsoft/pdf/boersma-pater-2013.pdf")
    End Select
End Sub

Private Sub chkArguerOn_Click()
    'If the user checks the arguer button, then make all the other possibilities visible too.
        Select Case chkArguerOn.Value
            Case 1
                Let chkDetailedArguments.Visible = True
                Let chkMiniTableaux.Visible = True
                Let chkMostInformativeBasis.Visible = True
                'Use the magic formula (copied from form1_resize()) to get the right height,
                '   depending on the height of the whole form.
                Let frmArguments.Height = 2415 * (Form1.Height - 600) / (7905 - 600)
                
            Case 0
                Let chkDetailedArguments.Visible = False
                Let chkMiniTableaux.Visible = False
                Let chkMostInformativeBasis.Visible = False
                Let frmArguments.Height = 1000 * (Form1.Height - 600) / (7905 - 600)
        End Select
End Sub

Private Sub cmdFacType_Click()
        Dim ReportErrorFileName As String
       
    'On Error GoTo CheckError
            
    'Avoid confusion by deactivating the Rank and Factorial Typology buttons.
        Let Form1.cmdRank.Enabled = False
        Let Form1.cmdFacType.Enabled = False
    
    'If the user has gotten this far, (s)he probably wants the settings saved.
        Call SaveUserChoices
    
    'Make sure there is a folder for output files, a daughter of the
    '   folder in which the input file is located.
        Call CreateAFolderForOutputFiles
    
    'Open the output files.  Know in advance what file this will be,
    '  in case it's already open, then you can report the error to
    '  the user in a useful way.
    
        Let ReportErrorFileName = gOutputFilePath + gFileName + "DraftOutput.txt"
        Let mTmpFile = FreeFile
        Open gOutputFilePath + gFileName + "DraftOutput.txt" For Output As #mTmpFile
        'Initialize the header numbers, in case this isn't the first run.
            Let gLevel1HeadingNumber = 0
        
        Let ReportErrorFileName = gOutputFilePath + gFileName + "QualityOutput.txt"
        Let mDocFile = FreeFile
        Open gOutputFilePath + gFileName + "QualityOutput.txt" For Output As #mDocFile
    
        'The HTML output:
            Let ReportErrorFileName = "ResultsFor" + gFileName + ".htm"
            Let mHTMFile = FreeFile
            Open gOutputFilePath + "ResultsFor" + gFileName + ".htm" For Output As #mHTMFile
            Call PrintTableaux.InitiateHTML(mHTMFile)

      Let lblProgressWindow = "Working..."
      DoEvents

    'Preliminary actions:
   
        'Remember that you're doing factorial typology.  This is used to
        '   avoid the construction of a million Hasse diagrams, also
        '   to invoke normal tableaux even if GLA is selected on the Ranking
        '   side of the interface.
            Let RunningFactorialTypology = True
   
        'Digest the input file
            If gHaveIOpenedTheFile = False Then
                If DigestTheInputFile(gInputFilePath, gFileName, gFileSuffix) = False Then
                    'Crucial that the user not be stranded by inability to click on Rank or FacType.
                        Let cmdRank.Enabled = True
                        Let cmdFacType.Enabled = True
                    Exit Sub    'KZ: false = input file couldn't be opened
                End If
            End If
            Let gHaveIOpenedTheFile = True
        
   'Execute the factorial typology algorithm:
        Call FactorialTypology.Main(mNumberOfForms, mNumberOfConstraints, mInputForm(), _
            mWinner(), mWinnerFrequency(), mWinnerViolations(), _
            mNumberOfRivals(), mRival(), mRivalFrequency(), mRivalViolations(), mConstraintName(), mAbbrev(), _
            mTmpFile, mDocFile, mHTMFile)
   
   'Close the output files.
        Close #mTmpFile
        Close #mDocFile
        Print #mHTMFile, "</BODY>"
        Close #mHTMFile
   
   'Remember that you've run factorial typology, so the any subsequent run will know.
        Let Form1.FactorialTypologyAlreadyRunOnThisFile = True
   
    'Announce completion, and facilitate viewing.
        Let Form1.lblProgressWindow.Caption = "I'm done."
    
    'Avoid confusion by reactivate the Rank and Factorial Typology buttons.
        Let Form1.cmdRank.Enabled = True
        Let Form1.cmdFacType.Enabled = True

      Close
              
   'Guide user to the View Results button
      Let Form1.cmdViewResults.Font.Size = 10
      Let Form1.cmdViewResults.FontBold = True
      
   'Get ready to View Results
        Form1.cmdViewResults.SetFocus
        Let gHasTheProgramBeenRun = True
        Let Form1.RunningFactorialTypology = False
      
      Exit Sub
      
CheckError:
    Select Case Err.Number  ' Evaluate error number.
        Case 70 ' "File already open" error.
            MsgBox "Error.  Probably what is happening is this:  I'm trying to open the file " + _
                ReportErrorFileName + " for purposes of storing my results, but a file of this name is already open.  I suggest you try to find this file, close it, then click OK.", vbExclamation
            Resume
        Case Else
            MsgBox "Program error, specifically:  " + Err.Description + ".   I would appreciate your letting me know the about the problem.  Email me at bhayes@humnet.ucla.edu, specifying error #84575, and including a copy of your input file.", vbCritical
    End Select
      

End Sub


Sub cmdViewResults_Click()

    On Error GoTo CheckError
    
    Dim OutFile As Long
    Let OutFile = FreeFile
    
    'Enable the basic command buttons--crucial that the user not be
    '   stranded by inability to click on Rank or FacType.
        Let cmdRank.Enabled = True
        Let cmdFacType.Enabled = True
    
    'Guard against catastrophes:
        If gHasTheProgramBeenRun = False Then
            'The user might be wanting to look at an old output file.  Check if one is available.
                If Dir(gOutputFilePath + gFileName + "DraftOutput.txt") = "" Then
                    MsgBox "I can't display an output.  I conjecture that this is because you haven't run the program yet.  Click on Rank or Factorial Typology first, then try View again.", vbExclamation
                    Exit Sub
                Else
                    'There is an old file available--does the user want to look at it?
                        Select Case MsgBox("You haven't run the program yet.  Click Yes to see the results of the last time OTSoft was run, else click No and click Rank or Factorial Typology.", vbYesNo + vbExclamation)
                            Case vbYes
                                'Do nothing here; just keep going.
                            Case vbNo
                                Exit Sub
                        End Select
                End If
        End If
        
    'It's unlikely that the program could run, but not produce an output file.
    '   But just in case, be prepared to warn the user.
        If Dir(gOutputFilePath + gFileName + "DraftOutput.txt") = "" Then
            MsgBox "I can't find the output file generated by OTSoft.  It is supposed to be located at:" + _
                Chr(10) + Chr(10) + _
                gOutputFilePath + gFileName + "DraftOutput.txt" + _
                Chr(10) + Chr(10) + _
                "I suggest you rerun OTSoft and try again.  If this doesn't work, then there is a program error; please contact Bruce Hayes at bhayes@humnet.ucla.edu, specifying error #41126.", vbExclamation
                Exit Sub
            Exit Sub
        End If              'Is there an output file?
        
    Let lblProgressWindow.Caption = ""
    
    'If you've gotten here, all is well with respect to there being a desired
    '   output file.  Now, view this file with the user's choice of method.
    
RestartPoint:                       'Restart if you failed to fill the output screen.
    Select Case EditorChoice
        Case "User specified"
            'Check if the word processor specified in OTSoftAuxiliarySoftwareLocations.txt can be found:
            If Dir(gUsersWordProcessor) <> "" Then
                'Yes.  So use it.
                '   Note:  chr(34) = quote needed to get Shell to deal with blanks in
                '       path names.
                    Shell gUsersWordProcessor + " " + _
                        Chr(34) + gOutputFilePath + gFileName + "DraftOutput.txt" + Chr(34), _
                        vbNormalFocus
            Else
                'No.  So use the Shell execute method, probably the safest.
                Let EditorChoice = "Shell execute"
                GoTo RestartPoint
            End If
        Case "Shell execute"
                'Try to open with whatever Windows wants to use.
                    Select Case UseWindowsPrograms.TryShellExecute(gOutputFilePath + gFileName + "DraftOutput.txt")
                        Case gShellExecuteWasSuccessful
                            'Do nothing
                        Case Else
                            MsgBox "Sorry, I'm having trouble displaying your output file.  You may want to open it from outside OTSoft; it is located at:" + _
                                Chr(10) + Chr(10) + _
                                gOutputFilePath + gFileName + "DraftOutput.txt" + _
                                Chr(10) + Chr(10) + _
                                "Click OK to return to the main OTSoft screen.", vbExclamation
                            Exit Sub
                    End Select
        Case "PrepareForPrinting"
            'Note:  chr(34), quote, needed to get Shell to work with pathnames that have
            '   blanks
            Shell gUsersWordProcessor + " " + _
                Chr(34) + gOutputFilePath + gFileName + "QualityOutput.txt" + Chr(34), _
                vbNormalFocus
        Case "InternalViewer"
            'This seems to be not working on some new versions of Windows, particularly
            '   Windows XP.  So, let's make the potentially-fatal code that fills the
            '   view window into a Boolean function, returning False if it fails.
                    'Try to put the output file in the text window:
                    'First, check that the file exists.
                        If Dir(gOutputFilePath + gFileName + "DraftOutput.txt") <> "" Then
                            Open gOutputFilePath + gFileName + "DraftOutput.txt" For Input As #OutFile
                        Else
                            MsgBox "Sorry, I can't find the file " + gOutputFilePath + gFileName + "DraftOutput.txt.  Click OK to continue.", vbExclamation
                            Exit Sub
                        End If
                        
                    If FillTheOutputWindow(OutFile) = False Then
                        UseWindowsPrograms.TryShellExecute (gOutputFilePath + gFileName + "DraftOutput.txt")
                        Close #OutFile
                        'GoTo RestartPoint
                    End If
        Case "WebPage"
            'First, check that the file exists.
                If Dir(gOutputFilePath + "ResultsFor" + gFileName + ".htm") <> "" Then
                    Select Case UseWindowsPrograms.TryShellExecute(Chr(34) + gOutputFilePath + "ResultsFor" + gFileName + ".htm" + Chr(34))
                        Case gShellExecuteWasSuccessful
                            'Do nothing
                        Case Else
                            MsgBox "Sorry, I'm having trouble displaying your output file.  You may want to open it from outside OTSoft; it is located at:" + _
                                Chr(10) + Chr(10) + _
                                gOutputFilePath + "ResultsFor" + gFileName + ".htm" + _
                                Chr(10) + Chr(10) + _
                                "Click OK to return to the main OTSoft screen.", vbExclamation
                            Exit Sub
                    End Select
                    'Call UseWindowsPrograms.TryShellExecute(Chr(34) + gOutputFilePath + "ResultsFor" + gFileName + ".htm" + Chr(34))
                    'Call Shell("C:\Program Files\Mozilla Firefox\firefox.exe " + Chr(34) + gOutputFilePath + "ResultsFor" + gFileName + ".htm" + Chr(34), vbNormalFocus)
                Else
                    MsgBox "Sorry, I can't find the file " + gOutputFilePath + "ResultsFor" + gFileName + ".htm.  Click OK to continue.", vbExclamation
                    Exit Sub
                End If
            
    End Select
   Close
   
   'Resize the type on the button.
        Let cmdViewResults.Font.Size = 8
        Let cmdViewResults.FontBold = False
      
    Exit Sub
    
CheckError:

    Select Case Err.Number  ' Evaluate error number.
        Case 7      'Not enough memory.
            Select Case _
                MsgBox("Your output file is too big to be displayed in OTSoft.  Click Yes to display it with your word processor; No to exit OTSoft." + _
                Chr(10) + Chr(10) + _
                "If you choose to view the output file some other way, note that it is located in the following place:" + _
                Chr(10) + Chr(10) + _
                gOutputFilePath + gFileName + "DraftOutput.txt", vbYesNo + vbExclamation)
                Case vbYes
                    Call mnuReturnToMainMenu_Click
                    Call mnuViewCuston_Click
                    Exit Sub
                Case vbNo
                    Call cmdExit_Click
            End Select
        Case 53
            MsgBox "Error.  I conjecture that your computer doesn't have a word processor in the location " + _
                gUsersWordProcessor + ".  " + _
                Chr(10) + Chr(10) + _
                "You can fix this by using Notepad to edit OTSoftAuxiliarySoftwareLocations.txt, which is located at " + _
                App.Path + "\OTSoftAuxiliarySoftwareLocations.txt.  " + _
                "Change the line that follows " + _
                "Path and name for custom word processor:, indicating where your word processor really is.  " + _
                Chr(10) + Chr(10) + _
                "For now, it you want to look at your results, exit OTSoft, and " + _
                "use a word processor to open either " + gOutputFilePath + gFileName + "DraftOutput.txt or " + _
                gOutputFilePath + gFileName + "QualityOutput.txt.", vbExclamation
        Case Else
            MsgBox "Program error.  Please contact Bruce Hayes at bhayes@humnet.ucla, including a copy of your input file and specifying error #32228.", vbCritical
    End Select
      
End Sub


Function FillTheOutputWindow(OutFile As Long) As Boolean
    
    Dim FillTheOutputWindowText As String
    Dim LinesInFile As Long, MyLine As String
    
    'To put the content of a file into the OTSoft-internal output window, we
    '   need to use a function that crashes in some versions of Windows.  Fail-
    '   safe this by returning False if the loading process fails.  Then other
    '   modes of display can be used.
    
    On Error GoTo CheckError
    
    'First, count the number of lines in the output file and abort if too many to fit.
        Dim QuickRead As Long
        Let QuickRead = FreeFile
        Open gOutputFilePath + gFileName + "DraftOutput.txt" For Input As #QuickRead
            Let LinesInFile = 0
            Do While Not EOF(QuickRead)
                Line Input #QuickRead, MyLine
                Let LinesInFile = LinesInFile + 1
                If LinesInFile > 2000 Then
                    Select Case _
                        MsgBox("Your output file is too big to be displayed in OTSoft.  Click Yes to display it with your word processor; No to exit OTSoft." + _
                        Chr(10) + Chr(10) + _
                        "If you choose to view the output file some other way, note that it is located in the following place:" + _
                        Chr(10) + Chr(10) + _
                        gOutputFilePath + gFileName + "DraftOutput.txt", vbYesNo + vbExclamation)
                        Case vbYes
                            'This returns the value false, leading the higher code to try something else.
                                Exit Function
                        Case vbNo
                            'Giving up in despair.
                                Call cmdExit_Click
                    End Select
                End If
            Loop
        Close #QuickRead
        
    'The following code, which fills the window, is by Taesun Moon.
        Do Until EOF(1)
            Let FillTheOutputWindowText = FillTheOutputWindowText + Input(1, OutFile)
        Loop

        Let txtViewOutput.Text = FillTheOutputWindowText
    
    Call PrepareTheScreenToShowAFile
    Let FillTheOutputWindow = True
    Exit Function
    
CheckError:
    Let FillTheOutputWindow = False
    'Restore the screen to normal form.
    Call mnuReturnToMainMenu_Click
    Exit Function

End Function

Sub cmdExit_Click()
    
    'Save the .ini file.
        Call SaveUserChoices
    'Update the recently opened files list.
        Call RefreshRecentlyOpenedFiles(gInputFilePath + gFileName + gFileSuffix)
    'Delete temporary files if requested
        If mnuDeleteTmpFiles.Checked = True Then Call DeleteTmpFiles
    'Close all files.
        Close
    'Unload all forms
        Dim frm As Form
        For Each frm In Forms
          Unload frm
        Next
        
    End
    
End Sub


'-----------------Routines for Altering the Appearance of the Main Window----------------

Sub PrepareTheScreenToShowAFile()
    
    'Make the viewing area as big as it can be.
        Let Left = 0
        Let Top = 0
        Let Width = Screen.Width - 100      'old:  whole width.  Now:  just a bit so it can be resized.
        Let Height = Screen.Height - 400    '- 250    'Not quite the whole screen:  leave
                                            'room for task bar.
        
    
    'Size the text output window.
        'KZ: I made it a bit shorter, because the bottom was cut off.
        '  BH:  Also note that for big file, one wants to make the horizontal
        '       scroll bar accessible.
        Let txtViewOutput.Height = Form1.Height - 900   '600
        Let txtViewOutput.Width = Form1.Width - 200
        Let txtViewOutput.Top = 0
        Let txtViewOutput.Left = 100
    
    'Show the text window, and suppress all the controls.
        Let txtViewOutput.Visible = True
        Let frmArguments.Visible = False
        Let Frame2.Visible = False
        Let Frame3.Visible = False
        Let cmdRank.Visible = False
        Let cmdFacType.Visible = False
        Let chkDiagnosticTableaux.Visible = False
        
    'Add a menu item permitting user to return:
        Let mnuReturnToMainMenu.Visible = True

End Sub

Sub Form_Resize()

    'Keep everything proportional when all is resized.
    
    'Turn this routine off for debugging.
        Exit Sub
    
    On Error Resume Next
    
        Dim HeightRatio As Single
        Dim WidthRatio As Single
        Dim MinRatio As Single      'Whichever is less--for font size.
        
        Let HeightRatio = (Form1.Height - 600) / (7905 - 600)
        Let WidthRatio = Form1.Width / 10905
        If HeightRatio < WidthRatio Then
            Let MinRatio = HeightRatio
        Else
            Let MinRatio = WidthRatio
        End If
        
        Let cmdRank.Height = 1455 * HeightRatio
        Let cmdRank.Left = 360 * WidthRatio
        Let cmdRank.Top = 360 * HeightRatio
        Let cmdRank.Width = 3735 * WidthRatio
        Let cmdRank.Font.Size = Int(10 * MinRatio)
        
        Let cmdFacType.Height = 1695 * HeightRatio
        Let cmdFacType.Left = 8040 * WidthRatio
        Let cmdFacType.Top = 240 * HeightRatio
        Let cmdFacType.Width = 2655 * WidthRatio
        Let cmdFacType.Font.Size = Int(10 * MinRatio)
        
        Let cmdViewResults.Height = 675 * HeightRatio
        Let cmdViewResults.Left = 180 * WidthRatio
        'Let cmdViewResults.Top = 5520 * HeightRatio
        Let cmdViewResults.Top = 5610 * HeightRatio
        Let cmdViewResults.Width = 10455 * WidthRatio
        Let cmdViewResults.Font.Size = 8 * MinRatio
        
        Let cmdExit.Height = 675 * HeightRatio
        Let cmdExit.Left = 180 * WidthRatio
        'Let cmdExit.Top = 6480 * HeightRatio
        Let cmdExit.Top = 6410 * HeightRatio
        Let cmdExit.Width = 10455 * WidthRatio
        Let cmdExit.Font.Size = 8 * MinRatio
        
        'Ranking arguments:
        
            If MinRatio < 0.9 Then
                Let frmArguments.Caption = "Arguments"
            Else
                Let frmArguments.Caption = "Ranking Argumentation"
            End If
            Let frmArguments.Left = 4560 * WidthRatio
            Let frmArguments.Width = 3255 * WidthRatio
            Let frmArguments.Top = 1800 * HeightRatio
            'This box is sometimes different sizes.
                Select Case chkArguerOn.Value
                    Case 1
                        Let frmArguments.Height = 1800 * HeightRatio
                        'Let frmArguments.Height = 2415 * HeightRatio
                    Case 0
                        Let frmArguments.Height = 1000 * HeightRatio
                End Select
            Let frmArguments.Font.Size = 8 * MinRatio
            
                If MinRatio < 0.9 Then
                    Let chkMiniTableaux.Caption = "Minitableaux"
                Else
                    Let chkMiniTableaux.Caption = "Include Illustrative Minitableaux"
                End If
                Let chkMiniTableaux.Height = 375 * HeightRatio
                Let chkMiniTableaux.Left = 360 * WidthRatio
                Let chkMiniTableaux.Top = 1320 * HeightRatio
                'Let chkMiniTableaux.Top = 1800 * HeightRatio
                Let chkMiniTableaux.Width = 2655 * WidthRatio
                Let chkMiniTableaux.Font.Size = 8 * MinRatio
            
                If MinRatio < 0.9 Then
                    Let chkDetailedArguments.Caption = "Details"
                Else
                    Let chkDetailedArguments.Caption = "Show details of argumentation"
                End If
                Let chkDetailedArguments.Height = 375 * HeightRatio
                Let chkDetailedArguments.Left = 360 * WidthRatio
                Let chkDetailedArguments.Top = 960 * HeightRatio
                'Let chkDetailedArguments.Top = 1320 * HeightRatio
                Let chkDetailedArguments.Width = 2655 * WidthRatio
                Let chkDetailedArguments.Font.Size = 8 * MinRatio
            
                If MinRatio < 0.9 Then
                    Let chkMostInformativeBasis.Caption = "Bundle"
                Else
                    Let chkMostInformativeBasis.Caption = "Prefer few, bundled arguments"
                End If
                Let chkMostInformativeBasis.Height = 375 * HeightRatio
                Let chkMostInformativeBasis.Left = 360 * WidthRatio
                Let chkMostInformativeBasis.Top = 640 * HeightRatio
                Let chkMostInformativeBasis.Width = 2655 * WidthRatio
                Let chkMostInformativeBasis.Font.Size = 8 * MinRatio
            
                If MinRatio < 0.9 Then
                    Let chkArguerOn.Caption = "Include"
                Else
                    Let chkArguerOn.Caption = "Include ranking arguments"
                End If
                Let chkArguerOn.Height = 375 * HeightRatio
                Let chkArguerOn.Left = 360 * WidthRatio
                Let chkArguerOn.Top = 320 * HeightRatio
                Let chkArguerOn.Width = 2295 * WidthRatio
                Let chkArguerOn.Font.Size = 8 * MinRatio
        
    'Diagnostic tableaux:
        
        If MinRatio < 0.85 Then
            Let chkDiagnosticTableaux.Caption = "Diagnostics"
        Else
            Let chkDiagnosticTableaux.Caption = "Diagnostics if ranking fails"
        End If
        Let chkDiagnosticTableaux.Height = 375 * HeightRatio
        Let chkDiagnosticTableaux.Left = 4920 * WidthRatio
        Let chkDiagnosticTableaux.Top = 3840 * HeightRatio
        Let chkDiagnosticTableaux.Width = 2175 * WidthRatio
        Let chkDiagnosticTableaux.Font.Size = 8 * MinRatio
        
    'Options for crowded tableaux:
        
        If MinRatio < 0.95 Then
            Let Frame2.Caption = "Crowded tableaux"
        Else
            Let Frame2.Caption = "Options for crowded tableaux"
        End If
        Let Frame2.Height = 2055 * HeightRatio
        Let Frame2.Top = 2160 * HeightRatio
        Let Frame2.Left = 8040 * WidthRatio
        Let Frame2.Width = 2655 * WidthRatio
        Let Frame2.Font.Size = 8 * MinRatio
        
            If MinRatio < 0.7 Then
                Let optNeverSwitchAxes.Caption = "Never"
            Else
                Let optNeverSwitchAxes.Caption = "Never switch axes"
            End If
            Let optNeverSwitchAxes.Height = 375 * HeightRatio
            Let optNeverSwitchAxes.Left = 120 * WidthRatio
            Let optNeverSwitchAxes.Top = 1560 * HeightRatio
            Let optNeverSwitchAxes.Width = 2295 * WidthRatio
            Let optNeverSwitchAxes.Font.Size = 8 * MinRatio
        
            If MinRatio < 0.85 Then
                Let optSwitchSomeAxes.Caption = "If needed"
            Else
                Let optSwitchSomeAxes.Caption = "Switch axes where needed"
            End If
            Let optSwitchSomeAxes.Height = 375 * HeightRatio
            Let optSwitchSomeAxes.Left = 120 * WidthRatio
            Let optSwitchSomeAxes.Top = 960 * HeightRatio
            Let optSwitchSomeAxes.Width = 2295 * WidthRatio
            Let optSwitchSomeAxes.Font.Size = 8 * MinRatio
        
            If MinRatio < 0.85 Then
                Let optSwitchAll.Caption = "All"
            Else
                Let optSwitchAll.Caption = "Switch axes for all tableaux"
            End If
            Let optSwitchAll.Height = 375 * HeightRatio
            Let optSwitchAll.Left = 120 * WidthRatio
            Let optSwitchAll.Top = 360 * HeightRatio
            Let optSwitchAll.Width = 2295 * WidthRatio
            Let optSwitchAll.Font.Size = 8 * MinRatio
        
    'Choose ranking algorithm:
        
        If MinRatio < 0.7 Then
            Let Frame3.Caption = "Algorithm"
        Else
            Let Frame3.Caption = "Choose Ranking Algorithm"
        End If
        'Let Frame3.Height = 1815 * HeightRatio
        'Let Frame3.Height = 2175 * HeightRatio
        Let Frame3.Height = 2515 * HeightRatio
        'Let Frame3.Height = 2875 * HeightRatio
        Let Frame3.Top = 1920 * HeightRatio
        Let Frame3.Left = 360 * WidthRatio
        Let Frame3.Width = 3735 * WidthRatio
        Let Frame3.Font.Size = 8 * MinRatio
        
            'The little buttons that identify algorithms are of low priority,
            '   and should disappear under crowded conditions.
            
            If MinRatio < 0.85 Or txtViewOutput.Visible = True Then
                Let cmdIdentifyGLA.Visible = False
                Let cmdIdentifyCD.Visible = False
                Let cmdIdentifyMaximumEntropy.Visible = False
                Let cmdIdentifyNHG.Visible = False
            Else
                Let cmdIdentifyGLA.Visible = True
                Let cmdIdentifyCD.Visible = True
                Let cmdIdentifyMaximumEntropy.Visible = True
                Let cmdIdentifyNHG.Visible = True
            End If
            
            Let cmdIdentifyCD.Height = 195 * HeightRatio
            Let cmdIdentifyCD.Left = 2640 * WidthRatio
            Let cmdIdentifyCD.Top = 360 * HeightRatio
            Let cmdIdentifyCD.Width = 975 * WidthRatio
            Let cmdIdentifyCD.Font.Size = 8 * MinRatio
        
            Let cmdIdentifyGLA.Height = 195 * HeightRatio
            Let cmdIdentifyGLA.Left = 2640 * WidthRatio
            Let cmdIdentifyGLA.Top = 720 * HeightRatio
            Let cmdIdentifyGLA.Width = 975 * WidthRatio
            Let cmdIdentifyGLA.Font.Size = 8 * MinRatio
        
            Let cmdIdentifyMaximumEntropy.Height = 195 * HeightRatio
            Let cmdIdentifyMaximumEntropy.Left = 2640 * WidthRatio
            Let cmdIdentifyMaximumEntropy.Top = 1800 * HeightRatio
            Let cmdIdentifyMaximumEntropy.Width = 975 * WidthRatio
            Let cmdIdentifyMaximumEntropy.Font.Size = 8 * MinRatio
        
            Let cmdIdentifyNHG.Height = 195 * HeightRatio
            Let cmdIdentifyNHG.Left = 2640 * WidthRatio
            Let cmdIdentifyNHG.Top = 2160 * HeightRatio
            Let cmdIdentifyNHG.Width = 975 * WidthRatio
            Let cmdIdentifyNHG.Font.Size = 8 * MinRatio
        
        
            If MinRatio < 0.9 Then
                Let optConstraintDemotion.Caption = "CD"
            Else
                Let optConstraintDemotion.Caption = "Constraint Demotion"
            End If
            Let optConstraintDemotion.Height = 195 * HeightRatio
            Let optConstraintDemotion.Left = 120 * WidthRatio
            Let optConstraintDemotion.Top = 360 * HeightRatio
            Let optConstraintDemotion.Width = 2600 * WidthRatio
            Let optConstraintDemotion.Font.Size = 8 * MinRatio
        
            If MinRatio < 0.9 Then
                Let optGLA.Caption = "GLA"
            Else
                Let optGLA.Caption = "Gradual Learning Algorithm"
            End If
            Let optGLA.Height = 195 * HeightRatio
            Let optGLA.Left = 120 * WidthRatio
            Let optGLA.Top = 720 * HeightRatio
            Let optGLA.Width = 2600 * WidthRatio
            Let optGLA.Font.Size = 8 * MinRatio
            
            If MinRatio < 0.9 Then
                Let optMaximumEntropy.Caption = "MaxEnt"
            Else
                Let optMaximumEntropy.Caption = "Maximum Entropy"
            End If
            Let optMaximumEntropy.Height = 195 * HeightRatio
            Let optMaximumEntropy.Left = 120 * WidthRatio
            Let optMaximumEntropy.Top = 1800 * HeightRatio
            Let optMaximumEntropy.Width = 2600 * WidthRatio
            Let optMaximumEntropy.Font.Size = 8 * MinRatio
        
            If MinRatio < 0.9 Then
                Let optNoisyHarmonicGrammar.Caption = "NHG"
            Else
                Let optNoisyHarmonicGrammar.Caption = "Noisy Harmonic Grammar"
            End If
            Let optNoisyHarmonicGrammar.Height = 195 * HeightRatio
            Let optNoisyHarmonicGrammar.Left = 120 * WidthRatio
            Let optNoisyHarmonicGrammar.Top = 2160 * HeightRatio
            Let optNoisyHarmonicGrammar.Width = 2600 * WidthRatio
            Let optNoisyHarmonicGrammar.Font.Size = 8 * MinRatio
        
        
        
        'Let Shape2.Height = 4095 * HeightRatio
        Let Shape2.Height = 4455 * HeightRatio
        Let Shape2.Left = 120 * WidthRatio
        Let Shape2.Top = 240 * HeightRatio
        Let Shape2.Width = 4215 * WidthRatio
        
        Let lblProgressWindow.Height = 1335 * HeightRatio
        Let lblProgressWindow.Left = 4560 * WidthRatio
        Let lblProgressWindow.Top = 240 * HeightRatio
        Let lblProgressWindow.Width = 3255 * WidthRatio
        Let lblProgressWindow.Font.Size = 12 * MinRatio
            
    'Not visible at the start:
        Let txtViewOutput.Height = Form1.Height - 200 '6300
        Let txtViewOutput.Width = Form1.Width - 200  '9500

    'Strangely, the rank button refuses to be default.  Try code override.
        Let cmdFacType.Default = False
        Let cmdRank.Default = True
        DoEvents

End Sub

Sub Form1_Unload()
    Call cmdExit_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'If someone clicks the little X in the upper right hand corner of the interface,
    '  we want the same actions to happen as if they click on the Exit button.
        Call cmdExit_Click
    
End Sub

'===============================DEALING WITH FILES=================================
'==================================================================================

Sub ReadOTSoftIni()

   'Get information from the OTSoftRememberUserChoices.txt file.
      
        On Error GoTo CheckError
        
        Dim CheckThisLine As String
        Dim IniFile As Long
        Let IniFile = FreeFile
        
        'If the OTSoftRememberUserChoices.txt in gSafePlaceToWriteTo is bad, copy over the one in app.path.
            If Dir(gSafePlaceToWriteTo & "\OTSoftRememberUserChoices.txt") = "" Then
                If Dir(App.Path + "\OTSoftRememberUserChoices.txt") <> "" Then
                    FileCopy App.Path + "\OTSoftRememberUserChoices.txt", gSafePlaceToWriteTo + "\OTSoftRememberUserChoices.txt"
                Else
                    'What if it isn't even in App.path?
                        MsgBox "Caution:  a file, OTSoftRememberUserChoices.txt, that normally comes with OTSoft " + _
                            "is missing from your OTSoft installation.  OTSoft will attempt to proceed without this file." + _
                            vbLf + vbLf + _
                            "If you have trouble, it may be necessary to reinstall the program." + _
                            "When you click OK, OTSoft will continue on this provisional basis.  Try selecting File, Open, in order to proceed.", vbExclamation
                        Exit Sub
                End If
            End If
            
        'Now you can safely (?) open it.
            Open gSafePlaceToWriteTo + "\OTSoftRememberUserChoices.txt" For Input As #IniFile
        
        Do While Not EOF(IniFile)
            Line Input #IniFile, CheckThisLine
            Select Case Trim(LCase(CheckThisLine))
            'Select Case Trim(CheckThisLine)
                'Last file read:
                    Case "last file read:  name  (e.g. french)"
                        Line Input #IniFile, gFileName
                    Case "last file read:  suffix (.txt/.in)"
                        Line Input #IniFile, gFileSuffix
                    Case "last file read:  path (e.g. c:\program files\otsoft\)"
                        Line Input #IniFile, gInputFilePath
                        'We need to establish the output file path early, so we can locate apriori rankings.
                            Let gOutputFilePath = gInputFilePath + "FilesFor" + gFileName + "\"
                Case "last ranking algorithm used (constraint demotion/low faith/gla):"
                    Line Input #IniFile, CheckThisLine
                    Select Case Trim(LCase(CheckThisLine))
                        Case "constraint demotion"
                            Let optConstraintDemotion.Value = True
                        Case "maximum entropy"
                            Let optMaximumEntropy.Value = True
                        Case "noisy harmonic grammar"
                            Let optNoisyHarmonicGrammar.Value = True
                    End Select
                Case "diagnostics if ranking fails (yes/no):"
                    Line Input #IniFile, CheckThisLine
                    Select Case Trim(LCase(CheckThisLine))
                        Case "yes"
                            Let chkDiagnosticTableaux.Value = 1
                        Case "no"
                            Let chkDiagnosticTableaux.Value = 0
                    End Select
                'Ranking argumentation:
                    Case "ranking argumentation(yes/no):"
                        Line Input #IniFile, CheckThisLine
                        Select Case Trim(LCase(CheckThisLine))
                            Case "yes"
                                Let chkArguerOn.Value = 1
                            Case "no"
                                Let chkArguerOn.Value = 0
                        End Select
                    Case "use most informative basis (yes/no):"
                        Line Input #IniFile, CheckThisLine
                        Select Case Trim(LCase(CheckThisLine))
                            Case "yes"
                                Let chkMostInformativeBasis.Value = 1
                            Case "no"
                                Let chkMostInformativeBasis.Value = 0
                        End Select
                    Case "show details of argumentation (yes/no):"
                        Line Input #IniFile, CheckThisLine
                        Select Case Trim(LCase(CheckThisLine))
                            Case "yes"
                                Let chkDetailedArguments.Value = 1
                            Case "no"
                                Let chkDetailedArguments.Value = 0
                        End Select
                Case "include illustrative minitableaux (yes/no):"
                    Line Input #IniFile, CheckThisLine
                    Select Case Trim(LCase(CheckThisLine))
                        Case "yes"
                            Let chkMiniTableaux.Value = 1
                        Case "no"
                            Let chkMiniTableaux.Value = 0
                    End Select
                'Factorial typology:
                    Case "include rankings in factorial typology printout (yes/no):"
                        Line Input #IniFile, CheckThisLine
                        Select Case Trim(LCase(CheckThisLine))
                            Case "yes"
                                Let mnuIncludeRankingInFTResults.Checked = True
                            Case "no"
                                Let mnuIncludeRankingInFTResults.Checked = False
                        End Select
                    Case "include tableaux in factorial typology printout (yes/no):"
                        Line Input #IniFile, CheckThisLine
                        Select Case Trim(LCase(CheckThisLine))
                            Case "yes"
                                Let mnuIncludeTableaux.Checked = True
                            Case "no"
                                Let mnuIncludeTableaux.Checked = False
                        End Select
                'Tableaux options:
                    Case "switch axes for all crowded tableaux (all/where needed/never):"
                        Line Input #IniFile, CheckThisLine
                        Select Case Trim(LCase(CheckThisLine))
                            Case "all"
                                Let optSwitchAll.Value = True
                            Case "where needed"
                                Let optSwitchSomeAxes.Value = True
                            Case "never"
                                Let optNeverSwitchAxes.Value = True
                        End Select
                    Case "sort candidates in tableaux by harmony"
                        Line Input #IniFile, CheckThisLine
                        Select Case Trim(LCase(CheckThisLine))
                            Case "yes"
                                Let mnuSortCandidatesByHarmony.Checked = True
                            Case "no"
                                Let mnuSortCandidatesByHarmony.Checked = False
                            Case Else
                                'Perhaps the default should be sorted?
                                    Let mnuSortCandidatesByHarmony.Checked = vbChecked
                        End Select
                        
                Case "word processor for examining results (pcwrite/internal/custom):"
                    Line Input #IniFile, CheckThisLine
                    Select Case Trim(LCase(CheckThisLine))
                        Case "pcwrite"
                            Let EditorChoice = "PCWrite"
                        Case "internal"
                            Let EditorChoice = "InternalViewer"
                        Case "custom"
                            Let EditorChoice = "User specified"
                        Case "webpage"
                            Let EditorChoice = "WebPage"
                    End Select
                    
                'This needed only for non-Excel input; else deduced from the Excel spreadsheet.
                Case "font for printing candidates (normal/IPA):"
                    Line Input #IniFile, CheckThisLine
                    Select Case CheckThisLine
                        Case "normal"
                            Let SymbolTag1 = ""
                        Case "IPA"
                            Let SymbolTag1 = "\ss"
                    End Select
                    
                Case "delete .tmp files on exit (yes/no):"
                    Line Input #IniFile, CheckThisLine
                    Select Case Trim(LCase(CheckThisLine))
                        Case "yes"
                            Let mnuDeleteTmpFiles.Checked = True
                        Case "no"
                            Let mnuDeleteTmpFiles.Checked = False
                    End Select
                Case "include summary file for factorial typology (yes/no):"
                    Line Input #IniFile, CheckThisLine
                    Select Case Trim(LCase(CheckThisLine))
                        Case "yes"
                            Let mnuFTSumFile.Checked = True
                            Let mnuViewCompactFactorialTypologySummaryFile.Visible = True
                            Let mnuSepFacType.Visible = True
                        Case "no"
                            Let mnuFTSumFile.Checked = False
                            Let mnuViewCompactFactorialTypologySummaryFile.Visible = False
                            Let mnuSepFacType.Visible = False
                    End Select
                Case "compact file for factorial typology, collapsing neutralized outputs (yes/no):"
                    Line Input #IniFile, CheckThisLine
                    Select Case Trim(LCase(CheckThisLine))
                        Case "yes"
                            Let mnuCompactFTFile.Checked = True
                            Let mnuViewCompactFileCollapsingNeutralizedOutputs.Visible = True
                            Let mnuSepFacType.Visible = True
                        Case "no"
                            Let mnuCompactFTFile.Checked = False
                            Let mnuViewCompactFileCollapsingNeutralizedOutputs.Visible = False
                            Let mnuSepFacType.Visible = False
                    End Select
                Case "constraint names in small caps (yes/no):"
                    Line Input #IniFile, CheckThisLine
                    Select Case Trim(LCase(CheckThisLine))
                        Case "yes"
                            Let mnuSmallCaps.Checked = True
                            Let SmallCapTag1 = "\cs"
                            Let SmallCapTag2 = "\ce"
                        Case "no"
                            Let mnuSmallCaps.Checked = False
                            Let SmallCapTag1 = ""
                            Let SmallCapTag2 = ""
                    End Select
                    'KZ: "font for printing candidates"?
                    
                'GLAMaxent parameters.  More need to be added.
                    Case "times to go through forms"
                        Input #IniFile, gNumberOfDataPresentations
                    Case "initial plasticity"
                        Input #IniFile, gCoarsestPlastMark
                    Case "final plasticity"
                        Input #IniFile, gFinestPlastMark
                    Case "number of times to test grammar"
                        Input #IniFile, gCyclesToTest
                    Case "Allow weights to go negative"
                        Input #IniFile, gNegativeWeightsOK
                        
                    Case "include tableaux in gla output (yes/no)"
                        Input #IniFile, CheckThisLine
                        Select Case Trim(LCase(CheckThisLine))
                            Case "yes"
                                Let IncludeTableauxInGLAOutput = True
                            Case "no"
                                Let IncludeTableauxInGLAOutput = False
                        End Select
                    Case "input forms to gla in exact proportion to their frequency (yes/no)"
                        Input #IniFile, CheckThisLine
                        Select Case Trim(LCase(CheckThisLine))
                            Case "yes"
                                Let gExactProportionsForGLAEtc = True
                            Case "no"
                                Let gExactProportionsForGLAEtc = False
                        End Select
                    Case "use the magri update rule"
                        Input #IniFile, CheckThisLine
                        Select Case Trim(LCase(CheckThisLine))
                            Case "yes"
                                Let gMagriUpdateRuleInEffect = True
                            Case "no"
                                Let gMagriUpdateRuleInEffect = False
                        End Select
                    Case "frequency at which learning state should be reported"
                        Input #IniFile, gReportingFrequency
                        
                    'NHG options
                    Case "for nhg, add noise to candidates"
                        Input #IniFile, CheckThisLine
                        Select Case Trim(LCase(CheckThisLine))
                            Case "yes"
                                Let gNHGLateNoise = True
                            Case "no"
                                Let gNHGLateNoise = False
                        End Select
                    Case "for nhg, negative weights ok"
                        Input #IniFile, CheckThisLine
                        Select Case Trim(LCase(CheckThisLine))
                            Case "yes"
                                Let gNHGNegativeWeightsOK = True
                            Case "no"
                                Let gNHGNegativeWeightsOK = False
                        End Select
                    Case "for nhg, apply noise to cells, not constraints"
                        Input #IniFile, CheckThisLine
                        Select Case Trim(LCase(CheckThisLine))
                            Case "yes"
                                Let gNHGNoiseAppliesToTableauCells = True
                            Case "no"
                                Let gNHGNoiseAppliesToTableauCells = False
                        End Select
                    Case "for nhg, apply noise to cells without violations"
                        Input #IniFile, CheckThisLine
                        Select Case Trim(LCase(CheckThisLine))
                            Case "yes"
                                Let gNHGNoiseForZeroCells = True
                            Case "no"
                                Let gNHGNoiseForZeroCells = False
                        End Select
                    Case "for nhg, apply noise after multiplication by violations"
                        Input #IniFile, CheckThisLine
                        Select Case Trim(LCase(CheckThisLine))
                            Case "yes"
                                Let gNHGNoiseIsAddedAfterMultiplication = True
                            Case "no"
                                Let gNHGNoiseIsAddedAfterMultiplication = False
                        End Select
                    Case "exponential nhg"
                        Input #IniFile, CheckThisLine
                        Select Case Trim(LCase(CheckThisLine))
                            Case "yes"
                                Let gExponentialNHG = True
                            Case "no"
                                Let gExponentialNHG = False
                        End Select
                    Case "demigaussiannhg"
                        Input #IniFile, CheckThisLine
                        Select Case Trim(LCase(CheckThisLine))
                            Case "yes"
                                Let gDemiGaussianNHG = True
                            Case "no"
                                Let gDemiGaussianNHG = False
                        End Select
                    Case "resolve ties by dropping the trial"
                        Input #IniFile, CheckThisLine
                        Select Case Trim(LCase(CheckThisLine))
                            Case "yes"
                                Let gResolveTiesBySkipping = True
                            Case "no"
                                Let gResolveTiesBySkipping = False
                        End Select
                    
                      
                'HTML parameters.
                    Case "shading color for html output (use normal html color codes, i.e. six-digit hexadecimal)"
                        Line Input #IniFile, CheckThisLine
                        Let gShadingColor = Trim(CheckThisLine)
                    
            End Select
        Loop
        
        Close #IniFile
        
        Exit Sub
    
CheckError:
    Select Case Err.Number  ' Evaluate error number.
        Case 53 'Can't find file.
            MsgBox ("I can't find the file OTSoftRememberUserChoices.txt.  Make sure your installation of OTSoft has this file, then click OK to continue."), vbExclamation
            Resume
        Case Else
            MsgBox "Program error.  Please notify me at bhayes@humnet.ucla.edu, enclosing a copy of your input file.  Specify location 49820.", vbCritical
    End Select

End Sub

Public Sub SaveUserChoices()
    
    'Rewrite the file with user choices.
    
     On Error GoTo CheckError
     
     'Close #UserChoiceFile
     Dim UserChoiceFile As Long
     Let UserChoiceFile = FreeFile
     
RestartPoint:
     Open gSafePlaceToWriteTo + "\OTSoftRememberUserChoices.txt" For Output As #UserChoiceFile
        
        'Input file:
            Print #UserChoiceFile, "Last file read:  Name  (e.g. French)"
            Print #UserChoiceFile, gFileName
            
            Print #UserChoiceFile, "Last file read:  suffix (.txt/.in)"
            Print #UserChoiceFile, gFileSuffix
            
            Print #UserChoiceFile, "Last file read:  path (e.g. c:\program files\otsoft\)"
            Print #UserChoiceFile, gInputFilePath
        
        Print #UserChoiceFile, "last ranking algorithm used (constraint demotion/low faith/gla):"
            If optConstraintDemotion.Value = True Then
                Print #UserChoiceFile, "Constraint demotion"
            ElseIf optGLA.Value = True Then
                Print #UserChoiceFile, "GLA"
            ElseIf optMaximumEntropy.Value = True Then
                Print #UserChoiceFile, "Maximum entropy"
            ElseIf optNoisyHarmonicGrammar.Value = True Then
                Print #UserChoiceFile, "noisy harmonic grammar"
            Else
                Print #UserChoiceFile, "BCD"
            End If
        
        Print #UserChoiceFile, "Diagnostics if ranking fails (yes/no):"
            If chkDiagnosticTableaux.Value = 1 Then
                Print #UserChoiceFile, "yes"
            Else
                Print #UserChoiceFile, "no"
            End If
        
        'Ranking argumentation:
            Print #UserChoiceFile, "Ranking argumentation(yes/no):"
                If chkArguerOn.Value = 1 Then
                    Print #UserChoiceFile, "yes"
                Else
                    Print #UserChoiceFile, "no"
                End If
            Print #UserChoiceFile, "Use Most Informative Basis (yes/no):"
                If chkMostInformativeBasis.Value = 1 Then
                    Print #UserChoiceFile, "yes"
                Else
                    Print #UserChoiceFile, "no"
                End If
            Print #UserChoiceFile, "Show details of argumentation (yes/no):"
                If chkDetailedArguments.Value = 1 Then
                    Print #UserChoiceFile, "yes"
                Else
                    Print #UserChoiceFile, "no"
                End If
            Print #UserChoiceFile, "Include illustrative minitableaux (yes/no):"
                If chkMiniTableaux.Value = 1 Then
                    Print #UserChoiceFile, "yes"
                Else
                    Print #UserChoiceFile, "no"
                End If
                
    'For legibility, the file says yes or no when the variable in question is Boolean.
    '   The function TrueFalseToYesNo(), in Module1, performs the conversion.
        
        'Factorial typology options:
            Print #UserChoiceFile, "Include rankings in factorial typology printout (yes/no):"
                Print #UserChoiceFile, TrueFalseToYesNo(mnuIncludeRankingInFTResults.Checked)
            Print #UserChoiceFile, "Include tableaux in factorial typology printout (yes/no):"
                Print #UserChoiceFile, TrueFalseToYesNo(mnuIncludeTableaux.Checked)
                
        'Tableaux options:
            Print #UserChoiceFile, "Switch axes for all crowded tableaux (all/where needed/never):"
                If optSwitchAll.Value = True Then
                    Print #UserChoiceFile, "all"
                ElseIf optSwitchSomeAxes.Value = True Then
                    Print #UserChoiceFile, "where needed"
                Else
                    Print #UserChoiceFile, "never"
                End If
                
            Print #UserChoiceFile, "sort candidates in tableaux by harmony"
                If mnuSortCandidatesByHarmony.Checked = True Then
                    Print #UserChoiceFile, "yes"
                Else
                    Print #UserChoiceFile, "no"
                End If
            
        Print #UserChoiceFile, "Word processor for examining results (pcwrite/internal/custom):"
            If EditorChoice = "PCWrite" Then
                Print #UserChoiceFile, "pcwrite"
            ElseIf EditorChoice = "User specified" Then
                Print #UserChoiceFile, "custom"
            ElseIf EditorChoice = "WebPage" Then
                Print #UserChoiceFile, "WebPage"
            Else
                Print #UserChoiceFile, "internal"
            End If
            
        Print #UserChoiceFile, "Font for printing candidates (normal/IPA):"
            Select Case SymbolTag1
                Case ""
                    Print #UserChoiceFile, "normal"
                Case "\ss"
                    Print #UserChoiceFile, "IPA"
            End Select
            
        Print #UserChoiceFile, "Delete .tmp files on exit (yes/no):"
            Print #UserChoiceFile, TrueFalseToYesNo(mnuDeleteTmpFiles.Checked)
        Print #UserChoiceFile, "Include summary file for factorial typology (yes/no):"
            Print #UserChoiceFile, TrueFalseToYesNo(mnuFTSumFile.Checked)
        Print #UserChoiceFile, "Compact file for factorial typology, collapsing neutralized outputs (yes/no):"
            Print #UserChoiceFile, TrueFalseToYesNo(mnuCompactFTFile.Checked)
        Print #UserChoiceFile, "Constraint names in small caps (yes/no):"
            Print #UserChoiceFile, TrueFalseToYesNo(mnuSmallCaps.Checked)
        
        'GLA options:
            Print #UserChoiceFile, "Times to go through forms"
                Print #UserChoiceFile, gNumberOfDataPresentations
            Print #UserChoiceFile, "Initial plasticity"
                Print #UserChoiceFile, gCoarsestPlastMark
            Print #UserChoiceFile, "Final plasticity"
                Print #UserChoiceFile, gFinestPlastMark
            Print #UserChoiceFile, "Number of times to test grammar"
                Print #UserChoiceFile, gCyclesToTest
            Print #UserChoiceFile, "Allow weights to go negative"
                Print #UserChoiceFile, gNegativeWeightsOK
            
            Call PrintAUserChoiceFileEntry(UserChoiceFile, "Include Tableaux in GLA Output (yes/no)", IncludeTableauxInGLAOutput)
            Call PrintAUserChoiceFileEntry(UserChoiceFile, "Input forms to GLA in exact proportion to their frequency (yes/no)", gExactProportionsForGLAEtc)
            Call PrintAUserChoiceFileEntry(UserChoiceFile, "Use the Magri update rule", gMagriUpdateRuleInEffect)
        
        'Noisy Harmonic Grammar options:
            Call PrintAUserChoiceFileEntry(UserChoiceFile, "For NHG, add noise to candidates", gNHGLateNoise)
            Call PrintAUserChoiceFileEntry(UserChoiceFile, "For NHG, negative weights ok", gNHGNegativeWeightsOK)
            Call PrintAUserChoiceFileEntry(UserChoiceFile, "For NHG, apply noise to cells, not constraints", gNHGNoiseAppliesToTableauCells)
            Call PrintAUserChoiceFileEntry(UserChoiceFile, "For NHG, apply noise to cells without violations", gNHGNoiseForZeroCells)
            Call PrintAUserChoiceFileEntry(UserChoiceFile, "For NHG, apply noise after multiplication by violations", gNHGNoiseIsAddedAfterMultiplication)
            Call PrintAUserChoiceFileEntry(UserChoiceFile, "Exponential NHG", gExponentialNHG)
            Call PrintAUserChoiceFileEntry(UserChoiceFile, "Demigaussian NHG", gDemiGaussianNHG)
            Call PrintAUserChoiceFileEntry(UserChoiceFile, "resolve ties by dropping the trial", gResolveTiesBySkipping)
            
        'Reporting frequency:
            Print #UserChoiceFile, "Frequency at which learning state should be reported"
                Print #UserChoiceFile, gReportingFrequency
                
        'HTML options:
             Print #UserChoiceFile, "shading color for html output (use normal html color codes, i.e. six-digit hexadecimal)"
             Print #UserChoiceFile, gShadingColor
                    

        Close #UserChoiceFile

        Exit Sub
        
CheckError:

    Select Case Err.Number
        Case 70
            MsgBox "Error.  I conjecture that the file " + gSafePlaceToWriteTo + "\OTSoftRememberUserChoices.txt" + _
                " is already open (i.e. with some othe program).  Please close OTSoftRememberUserChoices.txt, then run OTSoft again.", vbCritical
            End
        Case 75
            'It looks like the administrator won't let you write to App.Path.
            '   Try to find a safe haven on the hard disk.
TempLineLabel:
                If Dir("c:/windows/temp") <> "" Then
                    Let gSafePlaceToWriteTo = "c:/windows/temp"
                ElseIf Dir("c:/winnt/temp") <> "" Then
                    Let gSafePlaceToWriteTo = "c:/winnt/temp"
                ElseIf Dir("d:") <> "" Then
                    Let gSafePlaceToWriteTo = "d:"
                Else
                    'In desperation, don't save any user settings.
                    MsgBox "Sorry, I can't save your previous settings.  Please contact Bruce Hayes at bhayes@humnet.ucla.edu, including a copy of your input file and specifying error #40822.  OTSoft will resume without saving your settings.", vbCritical
                    Exit Sub
                End If
            'Rescue the SoftwareLocations and RecentlyOpenedFiles by copying them over.
            '    If Dir(App.Path + "\OTSoftAuxiliarySoftwareLocations.txt") <> "" Then
            '        FileCopy App.Path + "\OTSoftAuxiliarySoftwareLocations.txt", gSafePlaceToWriteTo + "\OTSoftAuxiliarySoftwareLocations.txt"
            '    End If
                If Dir(App.Path + "\OTSoftRecentlyOpenedFiles.txt") <> "" Then
                    FileCopy App.Path + "\OTSoftRecentlyOpenedFiles.txt", gSafePlaceToWriteTo + "RecentlyOpenedFiles.ini"
                End If
                    
            GoTo RestartPoint
        Case Else
            MsgBox "Sorry, I can't save your previous settings.  Please contact Bruce Hayes at bhayes@humnet.ucla.edu, including a copy of your input file and specifying error #40822.  OTSoft will resume without saving your settings.", vbCritical
    End Select
    
    Stop


End Sub

Sub PrintAUserChoiceFileEntry(FileNumber As Long, MyLabel As String, MyValue As Boolean)
    
    'Code a repeatedly-invoked task.
        'Caption
            Print #FileNumber, MyLabel
        'Value
        Select Case MyValue
            Case True
                Print #FileNumber, "yes"
            Case False
                Print #FileNumber, "no"
        End Select

End Sub


Sub ReadSoftwareLocations()

    'Open the file OTSoftAuxiliarySoftwareLocations.txt and learn where Word, Excel, java.exe, dot.exe, and Paint
    '   are.
    
    'This needs to be a separate file from OTSoftRememberUserChoices.txt, partly for user convenience,
    '    and partly because OTSoftRememberUserChoices.txt can have the "defaults" restored, which we
    '    don't want to have affect knowledge of program locations.
      
        On Error GoTo CheckError
        
        Dim CheckThisLine As String
        Dim SWFile As Long
        Let SWFile = FreeFile
        
        Let CheckThisLine = ""
        
        'Ideally, OTSoftAuxiliarySoftwareLocations.txt will be in the safe place folder, and you can
        '  get it from there.  But on the first use, it won't be, so you have to get
        '  from there, and copy it.  PUNTED.  Now, OTSoftAuxiliarySoftwareLocations.txt will reside in app.path, since
        '   it's basically an administrator function anyway.
        '    If Dir(gSafePlaceToWriteTo + "\OTSoftAuxiliarySoftwareLocations.txt") = "" Then
        '        FileCopy App.Path + "\OTSoftAuxiliarySoftwareLocations.txt", gSafePlaceToWriteTo + "\OTSoftAuxiliarySoftwareLocations.txt"
        '    End If
                
        'Now you should have a copy no matter what, and you can open it.
            'Open gSafePlaceToWriteTo + "\OTSoftAuxiliarySoftwareLocations.txt" For Input As #SWFile
            Open App.Path + "\OTSoftAuxiliarySoftwareLocations.txt" For Input As #SWFile
        
        Do While Not EOF(SWFile)
            Line Input #SWFile, CheckThisLine
            Select Case Trim(LCase(CheckThisLine))
                'Word or other word processor:
                    Case "location of your word processor:"
                        Line Input #SWFile, CheckThisLine
                        Let gUsersWordProcessor = Trim(CheckThisLine)
                'Excel:
                    Case "location of excel or other spreadsheet program:"
                        Line Input #SWFile, CheckThisLine
                        Let gExcelLocation = Trim(CheckThisLine)
                'Paint:
                    Case "location of paint or other graphics editing program:"
                        Line Input #SWFile, CheckThisLine
                        Let gPaintLocation = Trim(CheckThisLine)
                'dot.exe:
                    Case "location of dot.exe, part of att graphviz software:"
                        Line Input #SWFile, CheckThisLine
                        Let gDotExeLocation = Trim(CheckThisLine)
            End Select
        Loop
        
        Close #SWFile
        
        Exit Sub
    
CheckError:
    Select Case Err.Number  ' Evaluate error number.
        Case 53 'Can't find file.
            'Try finding it in app.path.
                If Dir(App.Path + "\OTSoftAuxiliarySoftwareLocations.txt") <> "" Then
                    Open App.Path + "\OTSoftAuxiliarySoftwareLocations.txt" For Input As #SWFile
                Else
                    MsgBox ("I can't find the file" + Chr(10) + Chr(10) + _
                        App.Path + "\OTSoftAuxiliarySoftwareLocations.txt." + Chr(10) + Chr(10) + _
                        "Make sure your installation of OTSoft has this file, then restart."), vbCritical
                    End
                End If
        Case Else
            MsgBox "Program error.  Please notify me at bhayes@humnet.ucla.edu, enclosing a copy of your input file.  Specify location 49820.", vbCritical
    End Select
 
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
  
    'Easy way to open a file:  drag it onto the interface.
    
      Dim FileString As String, i As Long
    
    'Got this code from the internet, and I don't know what all the arguments are.  Don't mess it up.
        With Data
            Let FileString = Data.Files.Item(1)
        End With
  
    'MsgBox FileString
    Call LetWindowsDictateTheFile(FileString)
    
    Exit Sub

End Sub



Sub LetWindowsDictateTheFile(MyFileName As String)

    'If the user started OTSoft by clicking on a file, then the default values for
    '   gFileName, gFileSuffix, and gInputFilePath should be set by using the Command
    '   function, which communicates what Windows tells OTSoft.
    
    'Command basically tells you what would be on the command line, if the program
    '  had been started with the command line option.  It omits the name of the
    '  program itself, however.
    
    'If there are errors, user can at least click the "Pick a different file" button.
    
    'Debug:
    'Dim D As Long
    'Let D = FreeFile
    'Open "c:\2004Debug.txt" For Output As #D
            
    Dim MyCommand As String

    On Error GoTo CheckError
    
    'There are two ways to use this routine.  If we're getting the file name from
    '   windows, then MyFileName will be null.  If we called this from an Open Recent
    '   menu item, then MyFileName will be the name from that menu item.  Either
    '   way, the parsing carried out here is necessary.
    
        If MyFileName = "" Then
            'Print #D, "Letting MyCommend = Command; Command = "; Command
            Let MyCommand = Command
        Else
            'Print #D, "Letting MyCommend = MyFileName; MyFileName = "; MyFileName
            Let MyCommand = MyFileName
        End If
        
    'Debugging; pretend you got this from Windows, ok?
    '    Let MyCommand = "C:\AO\OTSoft\Input\McCarthyProblem\CombinedSet\Trap.xls"
    
    If MyCommand <> "" Then
    
        'Peel off any quotes that surround MyCommand
            If Left(MyCommand, 1) = Chr(34) Then
                Let MyCommand = Mid(MyCommand, 2)
                'Print #D, "Peeling off a left quote from MyCommend, which is now"; MyCommand
            End If
            If Right(MyCommand, 1) = Chr(34) Then
                Let MyCommand = Left(MyCommand, Len(MyCommand) - 1)
                'Print #D, "Peeling off a right quote from MyCommend, which is now"; MyCommand
            End If
    
        'You have to do a parse.  This goes from right to left, yielding
        '   gFileSuffix, gFileName, and gInputFilePath.
            
            Dim MyString As String
            Dim MyStringLength As Long
            Dim PositionIndex As Long
            
            Let MyString = MyCommand
            Let MyStringLength = Len(MyString)
        
        'Parse gFileSuffix:
            For PositionIndex = MyStringLength To 1 Step -1
                If Mid(MyString, PositionIndex, 1) = "." Then
                    Let gFileSuffix = LCase(Mid(MyString, PositionIndex))
                    'Print #D, "I've found gFileSuffix, which is"; gFileSuffix
                    'Keep the residue, trimmed, for further parsing:
                        Let MyString = Left(MyString, PositionIndex - 1)
                        Let MyStringLength = Len(MyString)
                    Exit For
                End If
            Next PositionIndex
        
        'Parse gFileName and gInputFilePath:
            For PositionIndex = MyStringLength To 1 Step -1
                Select Case Mid(MyString, PositionIndex, 1)
                    Case "/", "\"
                        Let gFileName = Mid(MyString, PositionIndex + 1)
                        'Print #D, "Found gFileName, which is "; gFileName
                        Let gInputFilePath = Left(MyString, PositionIndex)
                        'Print #D, "Found gInputFilePath, which is "; gInputFilePath
                        
                        'We need the output file path early, so we can locate apriori rankings.
                            Let gOutputFilePath = gInputFilePath + "FilesFor" + gFileName + "\"
                        Exit For
                End Select
            Next PositionIndex

    End If          'Was MyCommand null?
    
    'Update the interface.
        'Put the file name on the command buttons, alerting user to how to find a file.
            Let cmdRank.Caption = "Rank " + gFileName + gFileSuffix
            Let cmdFacType.Caption = "Factorial typology for " + gFileName + gFileSuffix
            Let Me.Caption = "OTSoft " + gMyVersionNumber + " - " + gInputFilePath + gFileName + gFileSuffix
        'And on the Edit menu.
            Let mnuEditCurrentFile.Caption = "Edit the input file " + gFileName + gFileSuffix
        DoEvents

    Exit Sub
        
CheckError:

    MsgBox "I'm having trouble finding the file you want me to work on.  Please try again using the button that says Work with Different File.", vbExclamation
    Exit Sub

End Sub


Function DigestTheInputFile(InputFilePath As String, FileName As String, FileSuffix As String) As Boolean

   Dim CheckThisLine As String
   Dim LineNumber As Long
   Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
   
   If RedimensionTheArrays = True Then
   
       'KZ: CheckError in RedimensionTheArrays needs to let DigestTheInputFile
       'know that the file wasn't found so that it won't continue and cause
       'a program-terminating error. So I changed RedimensionTheArrays from a
       'sub to a function that returns a
       'Boolean value: true means proceed, false means abort.
       'In the same fashion, DigestTheInputFile needs to pass back the information
       'that the file wasn't opened, so I made it into a boolean function too.
        Let DigestTheInputFile = True   'KZ: the file was successfully opened.
    
        'Whenever you redigest the input file, then factorial typology should behave
        '   as it does on an initial run.
            Let FactorialTypologyAlreadyRunOnThisFile = False
            
        'This has three possibilities:  old Ranker file, tab-delimited text and Excel.  Also, correct impossible cases if possible."
            Select Case FileSuffix
                Case ".in"
                    Call DigestTraditionalFile
                Case ".txt"
                    If DigestTabDelimitedTextFile(FileName, FileSuffix) = False Then
                        Let DigestTheInputFile = False
                        Exit Function
                    End If
                Case ".xlsx"
                    If DigestTrueExcelFile(InputFilePath, FileName, FileSuffix) = False Then
                        Let DigestTheInputFile = False
                        Exit Function
                    End If
                Case ".xls"
                    MsgBox "Sorry, I can't run old Excel files with the file suffix .xls. Please open your file in Excel and re-save it as a modern Excel file, with suffix .xlsx."
                Case Else
                    MsgBox "Sorry, I can't run files with the suffix " + FileSuffix + "."
                    Close
                    End
            End Select

        'When the ranking algorithm is nonstochastic, we want an outright winner, the one with the highest frequency.
            Select Case gAlgorithmName
                Case "Recursive Constraint Demotion", "Biased Constraint Demotion", "Low Faithfulness Constraint Demotion"
                    Call DeduceWinnersFromFrequencies
            End Select
                    
'The following is big trouble; 10/14/21.  Why call structural descriptions if you don't want it?
                    
GoTo AdHocSkip
                    
                    Call StructuralDescriptions.Main(mNumberOfForms, mInputForm(), mWinner(), mWinnerViolations(), _
                        mNumberOfRivals(), mRival(), mRivalViolations(), _
                        mNumberOfConstraints, mConstraintName(), mAbbrev())
                    
                    'Retrieve what it calculated
                        For FormIndex = 1 To mNumberOfForms
                            For ConstraintIndex = 1 To mNumberOfConstraints
                                Let mWinnerViolations(FormIndex, ConstraintIndex) = StructuralDescriptions.mWinnerViolations(FormIndex, ConstraintIndex)
                            Next ConstraintIndex
                            For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                                For ConstraintIndex = 1 To mNumberOfConstraints
                                    Let mRivalViolations(FormIndex, RivalIndex, ConstraintIndex) = StructuralDescriptions.mRivalViolations(FormIndex, RivalIndex, ConstraintIndex)
                                Next ConstraintIndex
                            Next RivalIndex
                        Next FormIndex
            
AdHocSkip:
            
            
        'Let gHaveIOpenedTheFile = True
        Let DigestTheInputFile = True
        
        'Do a few elementary checks.
            If mNumberOfConstraints = 0 Then
                Select Case MsgBox("Sorry, it looks like your input file doesn't have any constraints in it.")
                '" + _
                '    Chr(10) + Chr(10) + _
                '    "Click Yes to edit the file, No to exit OTSoft." + _
                '    Chr(10) + Chr(10) + _
                '    "Note:  the location of your input file is as follows:  " + _
                '    InputFilePath + FileName + FileSuffix, vbYesNo, "OTSoft " + gMyVersionNumber, vbExclamation)
                    Case vbYes
                        Call mnuEditCurrentFile_Click
                    Case vbNo
                        Call cmdExit_Click
                End Select
                Let DigestTheInputFile = False
                Exit Function
            End If
            If mNumberOfForms = 0 Then
                Select Case MsgBox("Sorry, it looks like your input file doesn't have any learning data in it.    " + _
                    Chr(10) + Chr(10) + _
                    "Click Yes to edit the file, No to exit OTSoft." + _
                    Chr(10) + Chr(10) + _
                    "Note:  the location of your input file is as follows:  " + _
                    InputFilePath + FileName + FileSuffix, vbYesNo, "OTSoft " + gMyVersionNumber, vbExclamation)
                    Case vbYes
                        Call mnuEditCurrentFile_Click
                    Case vbNo
                        Call cmdExit_Click
                End Select
                Let DigestTheInputFile = False
                Exit Function
            End If
            
         'Make a backup file
            Call SaveAsTxt(FileName + "Backup", False, mWinner())
            
         'Let the user know you're ready to go.
            Let lblProgressWindow.Caption = "Ready"
            DoEvents
           
    Else    'KZ: input file wasn't opened; user must find it.
        Let DigestTheInputFile = False
    End If

End Function


Function DigestTrueExcelFile(InputFilePath As String, FileName As String, FileSuffix As String) As Boolean
       
    'Digest an input file in Excel format.
    
    'For how I learned to do this, I must consult in my Sandbox a program downloaded from
    'Source:  https://www.vbforums.com/showthread.php?768523-A-simple-program-to-open-an-excel-file-in-vb6-0 and modified there.
    
    'MsgBox "I'm sorry, but OTSoft no longer supports the old .xls Excel file format.  Please submit your file to OTSoft as tab-demited text, which" + _
    '    " you can obtain by saving in this format from Excel. The program will now close."
    'Close
    'End
    
    'On Error GoTo CheckError
    
        Dim FormIndex As Long, RivalIndex As Long, ExtraIndex As Long, ConstraintIndex As Long, i As Long
        Dim RowIndex As Long        'Raw count
        Dim RowNumber As Long       'Actual row, with the two rows at the top taken by constraint names/abbreviations.
        Dim ColumnNumber As Long
        Dim LocalNumberOfRivals As Long
        Dim LastInput As String
        Dim Buffer As String
        Dim FoundOne As Boolean
        Dim AbbreviationRowIsUsed As Boolean
        
    'Variables needed for Excel.
        Dim ApExcel As Excel.Application, MyExcelDocument As Excel.Workbook, Current As Excel.Worksheet, cell
        
    'Create objects for Excel.
        Set ApExcel = CreateObject("Excel.Application")
        'This command also opens the spreadsheet.
            Set MyExcelDocument = ApExcel.Workbooks.Open(InputFilePath + FileName + ".xlsx")
        Set Current = MyExcelDocument.Sheets(1)
   
    'Let the user know you are working on the file:
        lblProgressWindow.Caption = "Processing Input File..."
    
    'Find out what font the forms are in.  This is done by examining the first form.
        'MsgBox "Font is:  " + ApExcel.Range("A3:A3").Font.Name
        Select Case LCase(ApExcel.Range("A3:A3").Font.Name)
            Case "sildoulosipa93", "sildoulos ipa93", "sildoulosipa", "sildoulos ipa"
                Let SymbolTag1 = "\ss"
                Let SymbolTag2 = "\se"
            Case Else                   'for running a new file without closing the program.
                Let SymbolTag1 = ""
                Let SymbolTag2 = ""
        End Select
        
    'Find out how big the main arrays for this program have to be.
                    
        Let ColumnNumber = 3
        Do
            Let ColumnNumber = ColumnNumber + 1
            Let Buffer = Trim(ApExcel.Cells(1, ColumnNumber).Formula)
            Select Case Buffer
                Case ""
                    'Make sure there aren't more columns following.  If so, warn the user to fix it.
                        For ExtraIndex = 1 To 30        'Unlikely the gap is wider than 30.
                            'Let Buffer = Trim(ApExcel.Cells(1, ColumnNumber + ExtraIndex).Formula)
                            If Buffer <> "" Then
                                'Don't keep this workbook object open.
                                    MyExcelDocument.Close
                                    ApExcel.Quit
                                    Set MyExcelDocument = Nothing
                                    Set ApExcel = Nothing           'Note:  the alternative command ApExcel.Quit doesn't seem to work.
                                'Clear progress window.
                                    lblProgressWindow.Caption = ""  'KZ: to clear the "processing..." message.
                                'Report the trouble:  little gap/big gap.
                                    If ExtraIndex = 1 Then
                                        Select Case MsgBox("Sorry, but it looks like there's a problem with the input file.  " + _
                                                "The top row of your file is discontinuous, " + _
                                                "with a blank cell in column " + Chr(ColumnNumber + 64) + _
                                                ".  Since I get confused by such blanks, this needs to be fixed before I can proceed." + Chr(10) + Chr(10) + _
                                                "Click Yes to edit the file, No to exit OTSoft." + _
                                                Chr(10) + Chr(10) + _
                                                "Note:  the location of your input file is as follows:  " + _
                                                InputFilePath + FileName + FileSuffix, vbYesNo, "OTSoft " + gMyVersionNumber, vbExclamation)
                                            Case vbYes
                                                Call mnuEditCurrentFile_Click
                                            Case vbNo
                                                Call cmdExit_Click
                                        End Select
                                    Else
                                        Select Case MsgBox("Sorry, but it looks like there's a problem with the input file.  " + _
                                                "The top row of your file is discontinuous, " + _
                                                "with blank cells in columns " + Chr(ColumnNumber + 64) + " to " + Chr(ColumnNumber + ExtraIndex + 63) + _
                                                ".  Since I get confused by such blanks, this needs to be fixed before I can proceed." + Chr(10) + Chr(10) + _
                                                "Click Yes to edit the file, No to exit OTSoft." + _
                                                Chr(10) + Chr(10) + _
                                                "Note:  the location of your input file is as follows:  " + _
                                                InputFilePath + FileName + FileSuffix, vbYesNo, "OTSoft " + gMyVersionNumber)
                                            Case vbYes
                                                Call mnuEditCurrentFile_Click
                                            Case vbNo
                                                Call cmdExit_Click
                                        End Select
                                    End If
                                'Report failure and get out.
                                    Let DigestTrueExcelFile = False
                                    Exit Function
                            End If                  'Did you detect a gap?
                        Next ExtraIndex             'Keep looking for a gap over 30 columnes.
                    'There are no blanks for 30 cells down the pike.  So it's probably safe to move on.
                        Exit Do
                Case Else
                    Let MaximumNumberOfConstraints = MaximumNumberOfConstraints + 1
            End Select
        Loop
        'MsgBox "Max number of constraints is:  " + Str(MaximumNumberOfConstraints)
        
    'Guess whether the second row is actually constraint violations.
        Dim SecondRow() As String
        ReDim SecondRow(MaximumNumberOfConstraints + 3)
        For i = 1 To MaximumNumberOfConstraints + 3
            'Let SecondRow(i) = Trim(ApExcel.Cells(2, i).Formula)
        Next i
        If SecondRowIsViolations(SecondRow()) Then
            Let AbbreviationRowIsUsed = False
        Else
            Let AbbreviationRowIsUsed = True
        End If
    
    'Check to make sure that the second row matches in length--same number of abbreviations as constraints.
        If AbbreviationRowIsUsed = True Then
            For ColumnNumber = 4 To MaximumNumberOfConstraints + 3
                Let Buffer = Trim(ApExcel.Cells(2, ColumnNumber).Formula)
                If Buffer = "" Then
                    'Don't keep this workbook object open.
                        MyExcelDocument.Close
                        ApExcel.Quit
                        Set MyExcelDocument = Nothing
                        Set ApExcel = Nothing           'Note:  the alternative command ApExcel.Quit doesn't seem to work.
                    'Clear progress window.
                        lblProgressWindow.Caption = ""
                    'Report the trouble and solicit repair
                        Select Case MsgBox("Sorry, but it looks like there's a problem with the input file.  " + _
                                "The first row should contain full constraint names, the second row constraint abbreviations.  " + _
                                "There's a constraint name that appears to lack an abbreviation." + _
                                Chr(10) + Chr(10) + _
                                "Since I need both full constraint names and abbreviations, this needs to be fixed before I can proceed." + Chr(10) + Chr(10) + _
                                "Click Yes to edit the file, No to exit OTSoft." + _
                                Chr(10) + Chr(10) + _
                                "Note:  the location of your input file is as follows:  " + _
                                InputFilePath + FileName + FileSuffix, vbYesNo + vbExclamation, "OTSoft " + gMyVersionNumber)
                            Case vbYes
                                Call mnuEditCurrentFile_Click
                            Case vbNo
                                Call cmdExit_Click
                        End Select
                    'Record failure and get out.
                        Let DigestTrueExcelFile = False
                        Exit Function
                End If
            Next ColumnNumber
        End If

    'Also, that there aren't any abbreviations without corresponding constraint names.
        If AbbreviationRowIsUsed = True Then
            For ColumnNumber = MaximumNumberOfConstraints + 4 To MaximumNumberOfConstraints + 9
                Let Buffer = Trim(ApExcel.Cells(2, ColumnNumber).Formula)
                If Buffer <> "" Then
                     'Don't keep this workbook object open.
                        MyExcelDocument.Close
                        ApExcel.Quit
                        Set MyExcelDocument = Nothing
                        Set ApExcel = Nothing           'Note:  the alternative command ApExcel.Quit doesn't seem to work.
                    'Clear progress window.
                        lblProgressWindow.Caption = ""
                    'Report the trouble and solicit repair.
                        Select Case MsgBox("Sorry, but it looks like there's a problem with the input file.  " + _
                                "The first row should contain full constraint names, the second row constraint abbreviations.  " + _
                                "There's a constraint abbreviation that appears to lack a corresponding name (or perhaps, you've left out one of these two rows)." + _
                                Chr(10) + Chr(10) + _
                                "Since I need both full constraint names and abbreviations, this needs to be fixed before I can proceed." + Chr(10) + Chr(10) + _
                                "Click Yes to edit the file, No to exit OTSoft." + _
                                Chr(10) + Chr(10) + _
                                "Note:  the location of your input file is as follows:  " + _
                                InputFilePath + FileName + FileSuffix, vbYesNo, "OTSoft " + gMyVersionNumber, vbExclamation, vbYesNo)
                            Case vbYes
                                Call mnuEditCurrentFile_Click
                            Case vbNo
                                Call cmdExit_Click
                        End Select
                    'Record failure and get out.
                        Let DigestTrueExcelFile = False
                        Exit Function
                End If
            Next ColumnNumber
        End If
     
    'Calculate mNumberOfForms and mMaximumNumberOfRivals
            
        'We'll start in row 2 if there are no abbreviations, else in row 3.
            If AbbreviationRowIsUsed = True Then
                Let RowNumber = 2
            Else
                Let RowNumber = 1
            End If
           
        'Initialize:
            Let LastInput = ""
            
        'Loop through the file, augmenting
            Do
                Let RowNumber = RowNumber + 1
                'The second column is always filled to the end.
                Select Case Trim(ApExcel.Cells(RowNumber, 2).Formula)
                    Case ""
                        'Probably, you're reached the end of the file.
                        '  But occasionally a foolish user puts in a blank line.
                            If Trim(ApExcel.Cells(RowNumber + 1, 2).Formula) <> "" Then
                                MsgBox "Sorry, but it looks like there's a problem with the input file.  You have a row, number" + Str(RowNumber) + ", that has no candidate in it; i.e. there's a gap.  I can only read files that have no gaps of this sort.  Please fix your input file, then try again.  When you click OK, OTSoft will close down.", vbCritical
                                ApExcel.Workbooks.Close
                                Call cmdExit_Click
                            End If
                        Exit Do
                    Case Else
                        Let Buffer = Trim(ApExcel.Cells(RowNumber, 1).Formula)
                        If Buffer <> "" Then   'BH:  Let's comment this out:  sometimes you *want* identical inputs.  And Buffer <> LastInput
                            'You've found a new input.
                            Let mMaximumNumberOfForms = mMaximumNumberOfForms + 1
                            Let LocalNumberOfRivals = 0
                            Let LastInput = Buffer
                        Else
                            'You're continuing on rivals for the same input
                            Let LocalNumberOfRivals = LocalNumberOfRivals + 1
                            If LocalNumberOfRivals > mMaximumNumberOfRivals Then
                               Let mMaximumNumberOfRivals = LocalNumberOfRivals
                            End If
                        End If
                End Select
            Loop

    'Do the actual redimensioning.
        Call RedimensioningFinalStep
        
    'ACTUAL READING OF THE .XLSX INPUT FILE:
    
    'Initialize the populations of these things, for second uses.
       Call InitializeArrays
       
    'Get constraint names off of first line.
        Let ColumnNumber = 3
        Do
            Let ColumnNumber = ColumnNumber + 1
                Let Buffer = Trim(ApExcel.Cells(1, ColumnNumber).Formula)
                If Buffer <> "" Then
                    Let mNumberOfConstraints = mNumberOfConstraints + 1
                    Let mConstraintName(mNumberOfConstraints) = Buffer
                    'MsgBox "Constraint number is " + Trim(Str(mNumberOfConstraints)) + ".    Name is " + mConstraintName(mNumberOfConstraints)
                    'If there was no abbreviation row, you need to make the constraint names the abbreviation names as well, since the user provided none.
                        If AbbreviationRowIsUsed = False Then
                            Let mAbbrev(mNumberOfConstraints) = mConstraintName(mNumberOfConstraints)
                        End If
                Else
                    Exit Do
                End If
        Loop
    
    'Get constraint abbreviations off of second line.
        If AbbreviationRowIsUsed Then
            For ConstraintIndex = 1 To mNumberOfConstraints
                Let ColumnNumber = ConstraintIndex + 3
                Let mAbbrev(ConstraintIndex) = Trim(ApExcel.Cells(2, ColumnNumber).Formula)
                'MsgBox "Constraint number is " + Trim(Str(mNumberOfConstraints)) + ".    Abbreviation is " + mAbbrev(ConstraintIndex)
            Next ConstraintIndex
        End If
            
    'Get inputs, candidates, and violations from the rest of the lines.
        
        If AbbreviationRowIsUsed Then
            Let RowNumber = 2
        Else
            Let RowNumber = 1
        End If
   
        Do
   
            Let RowNumber = RowNumber + 1
   
            'All legitimate rows have something in the second column
                If Trim(ApExcel.Cells(RowNumber, 2).Formula) = "" Then
                    Exit Do
                Else
                    'Remember how many rows of inputs etc. you've got.
                        Let gTotalNumberOfRows = gTotalNumberOfRows + 1
                    'Check for new input forms and rivals.
                    'Note that the winner is determined later, from frequencies.
                    Let Buffer = Trim(ApExcel.Cells(RowNumber, 1).Formula)
                    If Buffer <> "" Then    'BH:  identical inputs allowed.  See comment above.  And Buffer <> mInputForm(mNumberOfForms) Then
                        Let mNumberOfForms = mNumberOfForms + 1
                        Let mInputForm(mNumberOfForms) = Buffer
                        'MsgBox "Added an input: " + mInputForm(mNumberOfForms)
                        Let mNumberOfRivals(mNumberOfForms) = 1
                        Let mRival(mNumberOfForms, 1) = Trim(ApExcel.Cells(RowNumber, 2).Formula)
                        'MsgBox "Added a rival: " + mRival(mNumberOfForms, 1)
                    Else
                        Let mNumberOfRivals(mNumberOfForms) = mNumberOfRivals(mNumberOfForms) + 1
                        Let mRival(mNumberOfForms, mNumberOfRivals(mNumberOfForms)) = _
                            Trim(ApExcel.Cells(RowNumber, 2).Formula)
                            'MsgBox "Added a rival: " + mRival(mNumberOfForms, mNumberOfRivals(mNumberOfForms))
                    End If
                End If
   
        'Read off frequency.
            'If it's not a number, then give the user a change to edit the file.
            Let Buffer = Trim(ApExcel.Cells(RowNumber, 3).Formula)
            If StructuralDescriptions.OnlyDigitsAndDecimalPoint(Buffer) = False Then
                Let DigestTrueExcelFile = False
                ApExcel.Workbooks.Close
                Select Case MsgBox("Sorry, but there's a problem with your input file.  The third column should contain only numbers (1's for outright winners, 0's for outright losers, other values when there is free variation).  But in the third column of your input file, I find the following non-numerical expression:  " + Chr(10) + Chr(10) + _
                        "    " + Buffer + Chr(10) + Chr(10) + _
                        "Click Yes to edit the file, No to exit OTSoft." + _
                        Chr(10) + Chr(10) + _
                        "Note:  the location of your input file is as follows:  " + _
                        InputFilePath + FileName + FileSuffix, vbYesNo, "OTSoft " + gMyVersionNumber)
                    Case vbYes
                        Call mnuEditCurrentFile_Click
                    Case vbNo
                        Call cmdExit_Click
                End Select
                Exit Function
            Else
                Let mRivalFrequency(mNumberOfForms, mNumberOfRivals(mNumberOfForms)) = _
                    Val(Trim(ApExcel.Cells(RowNumber, 3).Formula))
            End If
               
        'Read off constraint violations.
        '   Note:  the columns may contain alphabetic material, for
        '       encoding constraints.  It is harmless to let this material be mistakenly
        '       read as contraint violations.  The Val() function simply returns zero,
        '       and the bad values will be overwritten by the StructuralDescriptions module
        '       later one.
            For ConstraintIndex = 1 To mNumberOfConstraints
                Let ColumnNumber = ConstraintIndex + 3
                Let mRivalViolations(mNumberOfForms, _
                    mNumberOfRivals(mNumberOfForms), ConstraintIndex) = _
                    Val(Trim(ApExcel.Cells(RowNumber, ColumnNumber).Formula))
            Next ConstraintIndex
        
EscapePoint:        'In case the line was remarked out.
        
        Loop
                    
        'For STRUCTURAL DESCRIPTIONS, it helps to have a different coding of the file.
            ReDim Preserve gRawColumns(mNumberOfConstraints, gTotalNumberOfRows)
            
            'First, read in all the material that lines up with the inputs and candidates:
                For ConstraintIndex = 1 To mNumberOfConstraints
                    'Skip the forms and frequencies:
                        Let ColumnNumber = ConstraintIndex + 3
                    'Read a whole column until you get to a gap:
                        For RowIndex = 1 To gTotalNumberOfRows
                            If AbbreviationRowIsUsed Then
                                Let RowNumber = RowIndex + 2
                            Else
                                Let RowNumber = RowIndex + 1
                            End If
                            Let gRawColumns(ConstraintIndex, RowIndex) = Trim(ApExcel.Cells(RowNumber, ColumnNumber).Formula)
                        Next RowIndex
                Next ConstraintIndex
                
            'It's possible that the material for constraint structural descriptions will go lower
            '  down the sheet than the inputs and candidate did.  So look for more.
            
                Let RowIndex = gTotalNumberOfRows
                Do
                    Let FoundOne = False
                    Let RowIndex = RowIndex + 1
                    Let RowNumber = RowIndex + 2
                    For ConstraintIndex = 1 To mNumberOfConstraints
                        Let ColumnNumber = ConstraintIndex + 3
                        Let Buffer = Trim(ApExcel.Cells(RowNumber, ColumnNumber).Formula)
                        If Buffer <> "" Then
                            'You've found an entry going past the rows of the inputs and candidates.
                            '   If it's the first one on this row, redimension array.
                                    If FoundOne = False Then
                                        Let gTotalNumberOfRows = gTotalNumberOfRows + 1
                                        ReDim Preserve gRawColumns(mNumberOfConstraints, gTotalNumberOfRows)
                                    End If
                                'Record the new item:
                                    Let FoundOne = True
                                    Let gRawColumns(ConstraintIndex, RowIndex) = Buffer
                        End If
                    Next ConstraintIndex
                    'If you didn't find any, keep looking for a while, reporting a gap to the
                    '   reader.
                        If FoundOne = False Then
                            For ExtraIndex = 1 To 10
                                Let RowNumber = (gTotalNumberOfRows + ExtraIndex) + 2
                                For ConstraintIndex = 1 To mNumberOfConstraints
                                    Let ColumnNumber = ConstraintIndex + 3
                                    Let Buffer = Trim(ApExcel.Cells(RowNumber, ColumnNumber).Formula)
                                    If Buffer <> "" Then
                                        'Don't keep this workbook object open.
                                            ApExcel.Workbooks.Close
                                        'Clear progress window.
                                            Let lblProgressWindow.Caption = ""  '
                                        'Report the trouble and solicit repair.
                                            Select Case MsgBox("There's a problem with your input file.  I find that for the constraint" + Chr(10) + Chr(10) + _
                                                    "    " + mConstraintName(ConstraintIndex) + Chr(10) + Chr(10) + _
                                                    "there is a column of structural descriptions that has a gap in it.  " + _
                                                    "Since I can't deal with such gaps, , this needs to be fixed before I can proceed." + Chr(10) + Chr(10) + _
                                                    "Click Yes to edit the file, No to exit OTSoft." + _
                                                    Chr(10) + Chr(10) + _
                                                    "Note:  the location of your input file is as follows:  " + _
                                                    InputFilePath + FileName + FileSuffix, vbYesNo, "OTSoft " + gMyVersionNumber, vbExclamation)
                                                Case vbYes
                                                    Call mnuEditCurrentFile_Click
                                                Case vbNo
                                                    Call cmdExit_Click
                                            End Select
                                        'Record failure and get out.
                                            Let DigestTrueExcelFile = False
                                            Exit Function
                                    End If
                                Next ConstraintIndex
                            Next ExtraIndex
                            Exit Do
                        End If
                Loop
                
        'Debug
        Dim DebugFile As Long
        Let DebugFile = FreeFile
        Open gOutputFilePath + "/DebugJustAfterReadingInput.txt" For Output As #DebugFile
        Dim j As Long
        For i = 1 To mNumberOfForms
            Print #DebugFile, "Input:"; vbTab; mInputForm(i)
            Print #DebugFile, vbTab; "Winner:  ["; vbTab; mWinner(i); "]"; vbTab; "Frequency"; vbTab; mWinnerFrequency(i)
            For j = 1 To mNumberOfRivals(mNumberOfForms)
                Print #DebugFile, vbTab; "RivalIndex:"; vbTab; Trim(Str(j)); vbTab; "Rival:"; vbTab; mRival(i, j); vbTab; "Frequency:"; vbTab; mRivalFrequency(i, j)
            Next j
            Print #DebugFile,
        Next i
        Close #DebugFile
        
                    
                
      '  'You've got all the info needed, so close up.
            MyExcelDocument.Close
            ApExcel.Quit
            Set MyExcelDocument = Nothing
            Set ApExcel = Nothing           'Note:  the alternative command ApExcel.Quit doesn't seem to work.
            Let lblProgressWindow.Caption = "" 'KZ: to clear the "processing..." message.
     
     
        'Debug
        'Dim DebugFile As Long
        'Let DebugFile = FreeFile
        'Open gOutputFilePath + "/DebugJustAfterReadingInput.txt" For Output As #DebugFile
        'Dim j As Long
        'For i = 1 To mNumberOfForms
        '    Print #DebugFile, "Input:"; vbTab; mInputForm(i)
        '    Print #DebugFile, vbTab; "Winner:  ["; vbTab; mWinner(i); "]"; vbTab; "Frequency"; vbTab; mWinnerFrequency(i)
        '    For j = 1 To mNumberOfRivals(mNumberOfForms)
        '        Print #DebugFile, vbTab; "RivalIndex:"; vbTab; Trim(Str(j)); vbTab; "Rival:"; vbTab; mRival(i, j); vbTab; "Frequency:"; vbTab; mRivalFrequency(i, j)
        '    Next j
        '    Print #DebugFile,
        'Next i
        'Close #DebugFile
     
     '   'Report success and quit.
            Let DigestTrueExcelFile = True
            Exit Function
        
CheckError:
    Select Case Err.Number
        Case 440
            'Can't access an Excel spreadsheet, probably because Excel not installed.
            MsgBox "I'm having trouble accessing your Excel spreadsheet.  I suggest you save your Excel file as tab-delimited text, and reopen with OTSoft in that format instead.", vbExclamation
            End
        Case Else
            MsgBox "I'm having trouble opening your input file with Excel." + _
            Chr(10) + Chr(10) + _
            "Here are two of the things that might be going wrong." + _
            Chr(10) + Chr(10) + _
            "1. You might not have Excel on your computer." + _
            Chr(10) + Chr(10) + _
            "2. The file " + App.Path + "\OTSoftAuxiliarySoftwareLocations.txt might not indicate the correct location of your copy of Excel (you can edit this file to fix)." + _
            Chr(10) + Chr(10) + _
            "Here are two things you can do:" + _
            Chr(10) + Chr(10) + _
            "a. Save your input file as tab-delimited text (see Help) and use that as a non-Excel input file." + _
            Chr(10) + Chr(10) + _
            "b. Contact me at bhayes@humnet.ucla.edu, including a copy of your input file, and the error code 15552.", vbExclamation
            Exit Function
    End Select

End Function

Function DigestTabDelimitedTextFile(FileName As String, FileSuffix As String) As Boolean

   'Digest an input file in tab-delimited text format.
   
    'On Error GoTo CheckError
   
    Dim CheckThisLine As String, MyChomp As String, MyResidue As String
    Dim LineNumber As Long, ColumnNumber As Long, RowNumber As Long
    Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long, RowIndex As Long, ColumnIndex As Long, PositionIndex As Long
    
    'Debug file
        Dim D As Long
    
    'Dealing with omitted abbreviation rows:
        Dim SecondRow() As String, SecondRowCount As Long
        Dim AbbreviationRowIsUsed As Boolean
    
    'Open the file.
        Dim InF As Long
        Let InF = FreeFile
        'Check to make sure it exists.
            If Dir(gInputFilePath + FileName + FileSuffix) <> "" Then
                Open gInputFilePath + FileName + FileSuffix For Input As #InF
            Else
                MsgBox "Sorry, I can't find the file " + gInputFilePath + gFileName + FileSuffix + ".  Press OK to continue.", vbExclamation
                Let DigestTabDelimitedTextFile = False
                Exit Function
            End If
            
    'We want to know if the user left out an abbreviation line, so input line 2 and find out.
        Line Input #InF, CheckThisLine
        Line Input #InF, CheckThisLine
        Do
            Let MyChomp = s.Chomp(CheckThisLine)
            Let SecondRowCount = SecondRowCount + 1
            ReDim Preserve SecondRow(SecondRowCount)
            Let SecondRow(SecondRowCount) = MyChomp
            Let CheckThisLine = s.Residue(CheckThisLine)
            If Trim(CheckThisLine) = "" Then Exit Do
        Loop
        If SecondRowIsViolations(SecondRow()) Then
            Let AbbreviationRowIsUsed = False
        Else
            Let AbbreviationRowIsUsed = True
        End If
    
    'Now that you know, reopen it.
        Close #InF
        Open gInputFilePath + FileName + FileSuffix For Input As #InF
    
    'Initialize the populations of these things, for second uses.
        Call InitializeArrays
    
    'Initialize the line number.
        Let LineNumber = 0
    
    'Loop through all the lines.
        Do
BlankRedo:
            If EOF(InF) Then Exit Do
            Line Input #InF, CheckThisLine
            If VacuousLine(CheckThisLine) Then
                GoTo BlankRedo
            End If
                    
            'Augment the line number, and initialize the column number.
                Let LineNumber = LineNumber + 1
                Let ColumnNumber = 0
                
            'Augment gTotalNumberOfRows, which will be used for structural descriptions.  It doesn't include the header rows.
                If LineNumber >= 3 Or (AbbreviationRowIsUsed = False And LineNumber = 2) Then
                    Let gTotalNumberOfRows = gTotalNumberOfRows + 1
                End If
                
            'Go through each line and parse it.  Chomping divides at first tab found.
            
            'Restart point is where you go for the next chomp.
RestartPoint:
                Let MyChomp = s.Chomp(CheckThisLine)
                Let CheckThisLine = s.Residue(CheckThisLine)
                Let ColumnNumber = ColumnNumber + 1
                
            'Process what you just chomped.
                Select Case LineNumber
                    Case 1  'This row for constraint names.
                        'We have no interest in the first two columns,
                        '   which can contain whatever the user wants.
                            If ColumnNumber > 3 Then
                                Let mNumberOfConstraints = mNumberOfConstraints + 1
                                Let mConstraintName(mNumberOfConstraints) = MyChomp
                                'Let this be the abbreviation, in case there is no abbreviation row.
                                    Let mAbbrev(mNumberOfConstraints) = MyChomp
                            End If
                    Case 2 And AbbreviationRowIsUsed    'This row for constraint abbreviations.
                            
                        'Again, we have no interest in the first two columns,
                        '   which can contain whatever the user wants.
                        If ColumnNumber > 3 Then
                            'Are there too many, or two few, abbreviations?
                            If ColumnNumber - 3 > mNumberOfConstraints Or _
                                ColumnNumber - 3 > mNumberOfConstraints And CheckThisLine = "" Then
                                'Clear progress window.
                                    Let lblProgressWindow.Caption = ""  '
                                'Record failure
                                    Let DigestTabDelimitedTextFile = False
                                    Close #InF
                                'Report the trouble and solicit repair.
                                    Select Case MsgBox("The first row should contain full constraint names, the second row constraint abbreviations.  There's a constraint abbreviation that appears to lack a corresponding full name (or perhaps, you've left out one of these two rows)." + _
                                            Chr(10) + Chr(10) + _
                                            "Since I need both full constraint names and abbreviations, this needs to be fixed before I can proceed." + Chr(10) + Chr(10) + _
                                            "Click Yes to edit the file, No to exit OTSoft." + _
                                            Chr(10) + Chr(10) + _
                                            "Note:  the location of your input file is as follows:  " + _
                                            gInputFilePath + gFileName + FileSuffix, vbYesNo, "OTSoft " + gMyVersionNumber, vbExclamation)
                                        Case vbYes
                                            Call mnuEditCurrentFile_Click
                                        Case vbNo
                                            Call cmdExit_Click
                                    End Select
                                'Get out.
                                    Exit Function
                            Else
                                Let mAbbrev(ColumnNumber - 3) = MyChomp
                            End If
                        End If
                    Case Is >= 3 Or (AbbreviationRowIsUsed = False And LineNumber = 2)   'This row for inputs, rivals, frequencies, violations.

                        'Look at this row, column by column.
                        
                        Select Case ColumnNumber
                            Case 1
                                If Trim(MyChomp) <> "" Then
                                    'Note:  this could be identical to predecessor; let this happen.
                                        Let mNumberOfForms = mNumberOfForms + 1
                                        Let mInputForm(mNumberOfForms) = MyChomp
                                        Let mNumberOfRivals(mNumberOfForms) = 0
                                End If
                            Case 2
                                'Keep a temporary MyChomp of rivals, since you don't know the winner yet.
                                    Let mNumberOfRivals(mNumberOfForms) = mNumberOfRivals(mNumberOfForms) + 1
                                    Let mRival(mNumberOfForms, mNumberOfRivals(mNumberOfForms)) = Trim(MyChomp)
                            Case 3
                                'The frequency of this rival.
                                    Let mRivalFrequency(mNumberOfForms, mNumberOfRivals(mNumberOfForms)) = Val(MyChomp)
                            Case Is > mNumberOfConstraints + 3
                                'Too many violations on this row.
                                    'Clear progress window.
                                        Let lblProgressWindow.Caption = ""  '
                                    'Record failure
                                        Let DigestTabDelimitedTextFile = False
                                        Close #InF
                                    'Report the trouble and solicit repair.
                                        Select Case MsgBox("Caution:  in row #" + Trim(LineNumber) + " of your input file, you have more columns of constraint violations than you have actual constraints.")
                                        '+ _
                                        '        Chr(10) + Chr(10) + _
                                        '        "This needs to be fixed before I can proceed." + Chr(10) + Chr(10) + _
                                        '        "Click Yes to edit the file, No to exit OTSoft." + _
                                        '        Chr(10) + Chr(10) + _
                                        '        "Note:  the location of your input file is as follows:  " + _
                                        '        gInputFilePath + gFileName + FileSuffix, vbYesNo, "OTSoft " + gMyVersionNumber, vbExclamation)
                                            Case vbYes
                                                Call mnuEditCurrentFile_Click
                                            Case vbNo
                                                Call cmdExit_Click
                                        End Select
                                    'Get out.
                                        Exit Function
                            Case Else
                                'This is likely a violation count, but might also be a structural description.
                                '  If the latter, it will be handled by the structural description routine later.
                                    If IsNumeric(MyChomp) Then
                                        Let mRivalViolations(mNumberOfForms, mNumberOfRivals(mNumberOfForms), ColumnNumber - 3) = Val(MyChomp)
                                    End If
                                'Fill the array that will later be interpreted by the structural
                                '  description code.
                                    Let ConstraintIndex = ColumnNumber - 3
                                    Let RowIndex = LineNumber - 2
                                    'Let gRawColumns(ConstraintIndex, RowIndex) = Val(MyChomp)
                        End Select
                End Select              'What line in the file?
                
            'You've digested this chomp.  If there's more to the line, devour it.
                If CheckThisLine <> "" Then GoTo RestartPoint
                
            'At this point, you've digested a whole line.  Perform some checks.
            
                'I. Commonly people forget an abbreviation line.  Detect this and fill
                '   it in for them.  xxx not yet done.
                    
                'II. An incomplete line should be filled in.
                    If ColumnNumber < mNumberOfConstraints + 3 Then
                        Select Case LineNumber
                            Case 2
                            'Fill in abbreviations by copying constraint names.
                                For ColumnIndex = ColumnNumber - 3 + 1 To mNumberOfConstraints
                                    Let mAbbrev(ColumnIndex) = mConstraintName(ColumnIndex)
                                Next ColumnIndex
                            Case Is >= 3
                            'Fill in rival violations.
                                For ColumnIndex = ColumnNumber - 3 + 1 To mNumberOfConstraints
                                    Let mRivalViolations(mNumberOfForms, mNumberOfRivals(mNumberOfForms), ColumnIndex) = 0
                                Next ColumnIndex
                        End Select
                    End If
            
        Loop                                'Loop through all the lines.
    
    Close #InF
    
    'This is causing trouble; turn it off for now.  (10/14/21)


'GoTo TempGoToPoint
    
    
    
        'For STRUCTURAL DESCRIPTIONS, it helps to have a different coding of the file.
            ReDim Preserve gRawColumns(mNumberOfConstraints, gTotalNumberOfRows)
            
            'Reopen the file.
                Open gInputFilePath + FileName + FileSuffix For Input As #InF
            
            'Skip the header lines
                If AbbreviationRowIsUsed Then
                    Line Input #InF, CheckThisLine
                    Line Input #InF, CheckThisLine
                Else
                    Line Input #InF, CheckThisLine
                End If
            
            'First, read in all the material that lines up with the inputs and candidates:
                For RowIndex = 1 To gTotalNumberOfRows
                    If AbbreviationRowIsUsed Then
                        Let RowNumber = RowIndex + 2
                    Else
                        Let RowNumber = RowIndex + 1
                    End If
                    Line Input #InF, CheckThisLine
                    'Peel off the input, candidate, and frequency, which are not different.
                        Let CheckThisLine = s.Residue(s.Residue(s.Residue(CheckThisLine)))
                    'Grab the cell contents that (may or may not) embody formalized constraints.
                        For ConstraintIndex = 1 To mNumberOfConstraints
                            'Skip the forms and frequencies:
                                Let ColumnNumber = ConstraintIndex + 3
                            'Read a whole column until you get to a gap:
                                Let gRawColumns(ConstraintIndex, RowIndex) = s.Chomp(CheckThisLine)
                                Let CheckThisLine = s.Residue(CheckThisLine)
                        Next ConstraintIndex
                Next RowIndex
                
            'It's possible that the material for constraint structural descriptions will go lower
            '  down the sheet than the inputs and candidate did.  So look for more.
            
            
                Dim FoundOne As Boolean, Buffer As String, ExtraIndex As Long, InputFilePath As String
                
                Let RowIndex = gTotalNumberOfRows
                Do
                    Let FoundOne = False
                    Let RowIndex = RowIndex + 1
                    Let RowNumber = RowIndex + 2
                    For ConstraintIndex = 1 To mNumberOfConstraints
                        Let ColumnNumber = ConstraintIndex + 3
                        'Let Buffer = Trim(ApExcel.Cells(RowNumber, ColumnNumber).Formula)
                        If Buffer <> "" Then
                            'You've found an entry going past the rows of the inputs and candidates.
                            '   If it's the first one on this row, redimension array.
                                    If FoundOne = False Then
                                        Let gTotalNumberOfRows = gTotalNumberOfRows + 1
                                        ReDim Preserve gRawColumns(mNumberOfConstraints, gTotalNumberOfRows)
                                    End If
                                'Record the new item:
                                    Let FoundOne = True
                                    Let gRawColumns(ConstraintIndex, RowIndex) = Buffer
                        End If
                    Next ConstraintIndex
                    'If you didn't find any, keep looking for a while, reporting a gap to the
                    '   reader.
                        If FoundOne = False Then
                            For ExtraIndex = 1 To 10
                                Let RowNumber = (gTotalNumberOfRows + ExtraIndex) + 2
                                For ConstraintIndex = 1 To mNumberOfConstraints
                                    Let ColumnNumber = ConstraintIndex + 3
                                    'Let Buffer = Trim(ApExcel.Cells(RowNumber, ColumnNumber).Formula)
                                    If Buffer <> "" Then
                                        'Don't keep this workbook object open.
                                            'ApExcel.Workbooks.Close
                                        'Clear progress window.
                                            Let lblProgressWindow.Caption = ""  '
                                        'Report the trouble and solicit repair.
                                            Select Case MsgBox("There's a problem with your input file.  I find that for the constraint" + Chr(10) + Chr(10) + _
                                                    "    " + mConstraintName(ConstraintIndex) + Chr(10) + Chr(10) + _
                                                    "there is a column of structural descriptions that has a gap in it.  " + _
                                                    "Since I can't deal with such gaps, , this needs to be fixed before I can proceed." + Chr(10) + Chr(10) + _
                                                    "Click Yes to edit the file, No to exit OTSoft." + _
                                                    Chr(10) + Chr(10) + _
                                                    "Note:  the location of your input file is as follows:  " + _
                                                    InputFilePath + FileName + FileSuffix, vbYesNo, "OTSoft " + gMyVersionNumber, vbExclamation)
                                                Case vbYes
                                                    Call mnuEditCurrentFile_Click
                                                Case vbNo
                                                    Call cmdExit_Click
                                            End Select
                                        'Record failure and get out.
                                            'Let DigestTrueExcelFile = False
                                            Exit Function
                                    End If
                                Next ConstraintIndex
                            Next ExtraIndex
                            Exit Do
                        End If
                Loop

            Close #InF
    
TempGoToPoint:
    
    'Record success.
        Let DigestTabDelimitedTextFile = True
       
    'Debug:  show what you got.
        If True Then
            Let D = FreeFile
            Call CreateAFolderForOutputFiles
            Open gOutputFilePath + "debug.txt" For Output As #D
            'Constraint row:
                Print #D, Chr(9); Chr(9);
                For ConstraintIndex = 1 To mNumberOfConstraints
                    Print #D, Chr(9); mConstraintName(ConstraintIndex);
                Next ConstraintIndex
                Print #D,
            'Abbreviation row:
                Print #D, Chr(9); Chr(9);
                For ConstraintIndex = 1 To mNumberOfConstraints
                    Print #D, Chr(9); mAbbrev(ConstraintIndex);
                Next ConstraintIndex
                Print #D,
            'Inputs, candidates, violations
                For FormIndex = 1 To mNumberOfForms
                    For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                        'Inputs and candidates:
                            If RivalIndex = 1 Then
                                Print #D, mInputForm(FormIndex);
                            End If
                            Print #D, Chr(9); mRival(FormIndex, RivalIndex);
                        'Frequencies:
                            If mRivalFrequency(FormIndex, RivalIndex) = 0 Then
                                Print #D, Chr(9);
                            Else
                                Print #D, Chr(9); mRivalFrequency(FormIndex, RivalIndex);
                            End If
                        For ConstraintIndex = 1 To mNumberOfConstraints
                            'If mRivalViolations(FormIndex, RivalIndex, ConstraintIndex) = 0 Then
                            '    Print #d, Chr(9);
                            'Else
                                Print #D, Chr(9); mRivalViolations(FormIndex, RivalIndex, ConstraintIndex);
                            'End If
                        Next ConstraintIndex
                        Print #D,
                    Next RivalIndex
                Next FormIndex
            Close #D
            'Stop
       End If
       
      
    'Debug for raw text content:
        If False Then
            Let D = FreeFile
            Call CreateAFolderForOutputFiles
            Open gOutputFilePath + "debug.txt" For Output As #D
            For ConstraintIndex = 1 To mNumberOfConstraints
                Print #D, mConstraintName(ConstraintIndex)
                For RowIndex = 1 To gTotalNumberOfRows
                    Print #D, "    "; RowIndex; Chr(9); gRawColumns(ConstraintIndex, RowIndex)
                Next RowIndex
                Print #D,
            Next ConstraintIndex
            Close
        End If
        
    'Only error-checking code follows, so exit the function.
        Exit Function
    
CheckError:

    Select Case Err.Number
        Case 70
            MsgBox "Error.  I conjecture that the file " + gOutputFilePath + "debug.txt" + _
                " is already open (i.e. with some othe program).  Please close this file, then run OTSoft again.", vbExclamation
            End
        Case Else
            MsgBox "Program error.  Please contact Bruce Hayes at bhayes@humnet.ucla.edu, including a copy of your input file and specifying error #43922.", vbCritical
    End Select
            End

    MsgBox "I can't find this input file.  Please look for it using the Work With Different File button.", vbCritical

End Function

Function SecondRowIsViolations(SecondRow() As String) As Boolean

    'Sometimes users don't put in an abbreviation row.  Accommodate them.
    'Returns true if the second row looks like the beginning of the inputs and candidates.
    
        Dim i As Long
        
    'There should be an input in cell 1.
        If SecondRow(1) = "" Then
            Let SecondRowIsViolations = False
            Exit Function
        End If
        
    'There should be a candidate in cell 2.
        If SecondRow(2) = "" Then
            Let SecondRowIsViolations = False
            Exit Function
        End If
        
    'There should be integers in cells 4 through end.
    '    For i = 4 To UBound(SecondRow())
    '        If SecondRow(i) = "" Or s.IsAnInteger(SecondRow(i)) Then
    '            'Do nothing, looks ok so far.
    '        Else
    '            Let SecondRowIsViolations = False
    '            Exit Function
    '        End If
    '    Next i
    
    'Passed all tests, so return True
        Let SecondRowIsViolations = True
    
End Function

Function VacuousLine(MyString As String) As Boolean
    'Return true if a line contains at most blanks and tabs.
        Dim i As Long, TargetSegment As String
        
    'MsgBox MyString
        For i = 1 To Len(MyString)
            'MsgBox (Len(MyString))
            Let TargetSegment = Mid(MyString, i, 1)
            'MsgBox "Here is character #" + Trim(Str(i)) + " surrounded by X's:  X" + TargetSegment + "X.  Its ASCII code is " + Trim(Str(Asc(TargetSegment)))
            If Asc(TargetSegment) >= 32 Then
                If Asc(TargetSegment) <= 255 Then
                    Let VacuousLine = False
                    Exit Function
                End If
            End If
            
            Select Case Mid(MyString, i)
                Case " "
                    'do nothing
                Case Chr(9)
                    'do nothing
                Case Else
                    'do nothing
            End Select
        Next i
        Let VacuousLine = True
End Function


Sub InitializeArrays()
    
    'Sometimes the program is run again without exiting.  Make output reliable by initializing everything.
    
    Dim FormIndex As Long

    Let mNumberOfForms = 0
    For FormIndex = 1 To UBound(mNumberOfRivals())
        Let mNumberOfRivals(FormIndex) = 0
    Next FormIndex
    Let mNumberOfConstraints = 0
    Let gTotalNumberOfRows = 0

End Sub

Sub DeduceWinnersFromFrequencies()

    'Note that this applies only when we use Recursive Constraint Demotion.
    
    'Find the winner, making use of frequency, and install the winner arrays.
    '   Return True if successful (i.e. uncancelled).
    
        Dim LocalWinner As Long     'Current best in a highest-frequency search among candidates for a given input.
        Dim WinningValue As Single  'Highest frequency for any candidate for a given input.
        Dim LocalTie As Boolean     'Any input for which there are two winners?
        
        Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
        
        Let LocalTie = False
        Let mGlobalTie = False
        
        'Variables to monitor the presence of more than one winner.
            Let mMoreThanOneWinner = False
            Dim NumberOfWinnersPerInput As Long
        
        For FormIndex = 1 To mNumberOfForms
            'We'll be counting winners, so initialize.
                Let NumberOfWinnersPerInput = 0
            
            'Prepare to locate the highest-frequency winner.
                Let LocalWinner = 1
                Let WinningValue = -1
            
            'Inspect all candidates.
                For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                    Select Case mRivalFrequency(FormIndex, RivalIndex)
                        Case Is > WinningValue
                            'Don't do this for stochastic algorithms or for making a Praat file.  This should also be a user option.
                                If optMaximumEntropy.Value = True Or mPraatFileFlag Or optGLA.Value = True Or optNoisyHarmonicGrammar.Value = True Then
                                    Let LocalWinner = 1
                                    Let WinningValue = mRivalFrequency(FormIndex, 1)
                                    Let LocalTie = False
                                Else
                                    Let LocalWinner = RivalIndex
                                    Let WinningValue = mRivalFrequency(FormIndex, RivalIndex)
                                    Let LocalTie = False
                                End If
                        Case WinningValue
                            Let LocalTie = True
                    End Select
                    'Count winners and issue a warning if needed.
                        If mRivalFrequency(FormIndex, RivalIndex) > 0 Then
                            Let NumberOfWinnersPerInput = NumberOfWinnersPerInput + 1
                                'Reporting here is not needed: this is done all at once in its own routine.
                                        'Select Case gAlgorithmName
                                        '    Case "Recursive Constraint Demotion", "Biased Constraint Demotion", "Low Faithfulness Constraint Demotion"
                                        '        If NumberOfWinnersPerInput > 1 Then
                                        '            MsgBox "Caution:  the input " + mInputForm(FormIndex) + " has more than one winners I am using the first one. If there really are two winners, consider using a stochastic algorithm like MaxEnt or Noisy Harmonic Grammar."
                                        '        End If
                                        'End Select
                        End If
                        
                Next RivalIndex
                
            'If LocalTie got set to True, and never got set back to False, you have
            '   a true tie.  Record this.
                If LocalTie = True Then
                    Let mGlobalTie = True
                End If
            
            'Install as the winner the one with the highest frequency, with ties resolved in favor of
            '  the first one.
                Let mWinner(FormIndex) = mRival(FormIndex, LocalWinner)
                Let mWinnerFrequency(FormIndex) = mRivalFrequency(FormIndex, LocalWinner)
                For ConstraintIndex = 1 To mNumberOfConstraints
                    Let mWinnerViolations(FormIndex, ConstraintIndex) = mRivalViolations(FormIndex, _
                        LocalWinner, ConstraintIndex)
                Next ConstraintIndex
            
            'Now there's one fewer rival.
                Let mNumberOfRivals(FormIndex) = mNumberOfRivals(FormIndex) - 1
            
            'Move up the rivals to occupy a contiguous stretch of their array.
                For RivalIndex = LocalWinner To mNumberOfRivals(FormIndex)
                    Let mRival(FormIndex, RivalIndex) = mRival(FormIndex, RivalIndex + 1)
                    Let mRivalFrequency(FormIndex, RivalIndex) = mRivalFrequency(FormIndex, RivalIndex + 1)
                    For ConstraintIndex = 1 To mNumberOfConstraints
                        Let mRivalViolations(FormIndex, RivalIndex, ConstraintIndex) = _
                            mRivalViolations(FormIndex, RivalIndex + 1, ConstraintIndex)
                    Next ConstraintIndex
                Next RivalIndex
            
            'Keep track if there are any forms in which there is more than one winner.
                If NumberOfWinnersPerInput > 1 Then
                    Let mMoreThanOneWinner = True
                End If
        
        Next FormIndex          'Examine all inputs.
       
        
            
End Sub

Sub DigestTraditionalFile()
    
   On Error GoTo CheckError
   
   Dim CheckThisLine As String
   Dim LineNumber As Long
   Dim FormIndex As Long, ConstraintIndex As Long, AprioriRankingsIndex As Long, i As Long

   Dim blnUserFrequency As Boolean  'KZ: keeps track of whether user entered a frequency.
   Dim TradInFile As Long
   
   Let TradInFile = FreeFile
    
    'Don't try to open a file if it's not there.
        If Dir(gInputFilePath + gFileName + gFileSuffix) <> "" Then
            Open gInputFilePath + gFileName + gFileSuffix For Input As #TradInFile
        Else
            MsgBox "Sorry, I can't find the file " + gInputFilePath + gFileName + gFileSuffix + ".  Click OK to continue.", vbExclamation
            Let gHaveIOpenedTheFile = False
            Exit Sub
        End If
    
    'Initialize the populations of these things, for second uses.
        Let mNumberOfForms = 0
        For FormIndex = 1 To mMaximumNumberOfForms
            Let mNumberOfRivals(FormIndex) = 0
        Next FormIndex
        Let mNumberOfConstraints = 0
        Let mNumberOfAPrioriRankings = 0
        
    Do
BlankLineReturnPoint:
       If EOF(TradInFile) Then Exit Do
       Input #TradInFile, CheckThisLine
       Let LineNumber = LineNumber + 1
InputReturnPoint:
       Let CheckThisLine = Trim(CheckThisLine)
       If LCase$(Left$(CheckThisLine, 10)) = "user name:" Then
          Let mUserName = Mid$(CheckThisLine, 13)
       ElseIf Trim(CheckThisLine) = "" Then
          GoTo BlankLineReturnPoint
       ElseIf LCase$(Left$(CheckThisLine, 13)) = "constraint:  " Then
             'You've hit a constraint entry.  Record the constraint, along with
           '  its abbreviation.
                Let mNumberOfConstraints = mNumberOfConstraints + 1
                Let mConstraintName(mNumberOfConstraints) = Mid$(CheckThisLine, 14)
           'Look for the abbreviation.
            Let LineNumber = LineNumber + 1
            If EOF(TradInFile) Then
               MsgBox "There's a problem at line " + Str(LineNumber) + "of your input" + _
                    "file.   I need a line that begins Abbreviation: followed by two spaces." + _
                    "but I reached the end of the file before finding it. Please fix your input file and try again.", vbCritical
               End
            Else
RestartPoint222:
               Input #TradInFile, CheckThisLine
               If CheckThisLine = "" Then GoTo RestartPoint222
               If LCase$(Left$(CheckThisLine, 15)) <> "abbreviation:  " Then
                  MsgBox "There's a problem at line " + Str(LineNumber) + "of your input file." + _
                  "I need a line that begins: Abbreviation: followed by two spaces." + _
                  "But instead I got:  " + CheckThisLine + "  Please fix your input file and try again.", vbCritical
                  End
               Else
                  Let mAbbrev(mNumberOfConstraints) = Mid$(CheckThisLine, 16)
               End If
        
            End If
       ElseIf LCase$(Left$(CheckThisLine, 17)) = "a priori ranking:" Then
          Let mNumberOfAPrioriRankings = mNumberOfAPrioriRankings + 1
          Input #TradInFile, mAPrioriRankingsList(mNumberOfAPrioriRankings, 0)
          Input #TradInFile, mAPrioriRankingsList(mNumberOfAPrioriRankings, 1)
            'The a priori rankings were done differently in the old .in files.  Convert, so they can be used.
              For i = 1 To mNumberOfAPrioriRankings
                  Let gAPrioriRankingsTable(mAPrioriRankingsList(AprioriRankingsIndex, 0), mAPrioriRankingsList(AprioriRankingsIndex, 1)) = True
              Next i
       ElseIf LCase$(Left$(CheckThisLine, 8)) = "input:  " Then
            'You've just reached a line labeled 'input', i.e. a form used as data.
            '  Extract the input, the winner, the rivals, and the constraints they violate.
                 Let mNumberOfForms = mNumberOfForms + 1
                 Let mInputForm(mNumberOfForms) = Trim(Mid$(CheckThisLine, 9))
                 Input #TradInFile, CheckThisLine
                      If EOF(TradInFile) Then
                            MsgBox "There is a problem with your input file:  the program reached the end of the file, without getting the full information for a particular form.  Check the file and fix.", vbCritical
                            End
                      End If
              'Extract the winner and its violations.
                  Let LineNumber = LineNumber + 1
                  Let CheckThisLine = Trim(CheckThisLine)
                  If LCase$(Left$(CheckThisLine, 9)) <> "winner:  " Then
                     MsgBox "At line " + Str(LineNumber) + "I need a line that begins Winner: followed by two spaces.  Please fix input file and try again.", vbCritical
                     Stop
                  Else
                     'KZ: added the following line because winner is being treated as a rival (see right below):
                        Let mNumberOfRivals(mNumberOfForms) = mNumberOfRivals(mNumberOfForms) + 1
                     'KZ: changed this from Winner() to mRival(); DeduceWinners sub later figures out the real winner:
                        Let mRival(mNumberOfForms, mNumberOfRivals(mNumberOfForms)) = Trim(Mid$(CheckThisLine, 10))
                     'KZ: we need to start labelling the rivals at 1 so that DeduceWinnersFromRivals will work.
                        Let mNumberOfRivals(mNumberOfForms) = 1
                     'Now extract the violations for all of the constraints, checking at all stages that you have the right data.
                        Let blnUserFrequency = False 'KZ: did user specify a frequency?
                         For ConstraintIndex = 1 To mNumberOfConstraints
                            Let LineNumber = LineNumber + 1
                            Input #TradInFile, CheckThisLine
                            'Ignore Frequency information, for just now:
                                If LCase$(Left$(CheckThisLine, 9)) = "frequency" Then
                                'KZ: read in frequency:
                                    Let mRivalFrequency(mNumberOfForms, mNumberOfRivals(mNumberOfForms)) = _
                                     Trim(Mid$(CheckThisLine, 13))
                                    blnUserFrequency = True
                                    Let ConstraintIndex = ConstraintIndex - 1
                                Else
                                    If CheckThisLine <> mAbbrev(ConstraintIndex) Then
                                       MsgBox "I'm stuck at or near line " + Str(LineNumber) + "of the file. I'm looking for the abbreviated constraint label " + mAbbrev(ConstraintIndex) + ", but I got " + CheckThisLine + " instead.  Please fix your file then restart.", vbCritical
                                       End
                                    Else
                                       'KZ: if no frequency entered for winner, assume it's 1
                                        If blnUserFrequency = False Then 'KZ: don't overwrite an existing frequency.
                                             Let mRivalFrequency(mNumberOfForms, mNumberOfRivals(mNumberOfForms)) = 1
                                        End If
                                       'KZ: so that DeduceWinnerFromFrequencies will work, intall this as a rival (#TradInFile)
                                            Input #TradInFile, mRivalViolations(mNumberOfForms, 1, ConstraintIndex)
                                       'KZ: but just to be on the safe side, also make it the winner:
                                            Let mWinnerViolations(mNumberOfForms, ConstraintIndex) = mRivalViolations(mNumberOfForms, 1, ConstraintIndex)
                                    End If
                                End If
                         Next ConstraintIndex
                  End If                'Are we processing a "winner"?
                 
              'Extract the rivals and their violations.
                    Do
                        Let LineNumber = LineNumber + 1
Line364RestartPoint:
                      If EOF(TradInFile) Then Exit Do
                      Input #TradInFile, CheckThisLine
                      Let CheckThisLine = Trim(CheckThisLine)
                      If LCase$(Left$(CheckThisLine, 8)) = "input:  " Then
                         'Go back to processing inputs:
                            GoTo InputReturnPoint
                      ElseIf LCase$(Left$(CheckThisLine, 15)) = "free variant:  " Then
                         'Process a free variant (no longer used, but what if there are old input files that have it?
                      ElseIf CheckThisLine = "" Then
                         'Ignore blank lines
                            GoTo Line364RestartPoint
                      ElseIf LCase$(Left$(CheckThisLine, 8)) <> "rival:  " Then
                         'Complain if there is something you can't identify.
                            MsgBox "At line " + Str(LineNumber) + "I need line that begins   Rival:    followed by two spaces.  Please fix input file and try again.", vbCritical
                            End
                      Else
                         'Process a rival candidate.
                            Let mNumberOfRivals(mNumberOfForms) = mNumberOfRivals(mNumberOfForms) + 1
                            Let mRival(mNumberOfForms, mNumberOfRivals(mNumberOfForms)) = Trim(Mid$(CheckThisLine, 9))
                         'Extract the violations for all of the constraints, checking at all stages that you have the right data.
                            Let blnUserFrequency = False 'KZ: did user specify frequency?
                            For ConstraintIndex = 1 To mNumberOfConstraints
                               Let LineNumber = LineNumber + 1
                               Input #TradInFile, CheckThisLine
                                   If LCase$(Left$(CheckThisLine, 9)) = "frequency" Then
                                        'KZ: read in frequency:
                                            Let mRivalFrequency(mNumberOfForms, mNumberOfRivals(mNumberOfForms)) = Trim(Mid$(CheckThisLine, 13))
                                            Let blnUserFrequency = True
                                            Let ConstraintIndex = ConstraintIndex - 1
                                   Else
                                       If CheckThisLine <> mAbbrev(ConstraintIndex) Then
                                          MsgBox "I'm stuck at or near line " + Str(LineNumber) + "of the file.  I'm looking for the short constraint label " + mAbbrev(ConstraintIndex) + ", but I got " + CheckThisLine + " instead.  Please fix your file and restart.", vbCritical
                                          End
                                       Else
                                          'KZ: if no frequency entered for rival, assume it's 0:
                                            If blnUserFrequency = False Then  'KZ: don't overwrite existing frequency.
                                                  Let mRivalFrequency(mNumberOfForms, mNumberOfRivals(mNumberOfForms)) = 0
                                            End If
                                          'KZ: modified for debug
                                            Input #TradInFile, CheckThisLine
                                            Let mRivalViolations(mNumberOfForms, mNumberOfRivals(mNumberOfForms), ConstraintIndex) = CheckThisLine
                                       End If
                               End If
                            Next ConstraintIndex
                      End If                        'Grand if:  possible headings in the input file.
                  Loop
       Else
          MsgBox "I couldn't understand line number " + Str(LineNumber) + " in your file: " + CheckThisLine + "Please check your input file for correct format.", vbCritical
          End
       End If
    Loop
    
    
    
    
    
    'Close the input file.
        Close #TradInFile
    'Record that you've opened this file.
        Let gHaveIOpenedTheFile = True
    
    Exit Sub

CheckError:
    MsgBox "I can't find this input file.  Please look for it using the Work With Different File button.", vbCritical
    Exit Sub

End Sub


Sub CreateAFolderForOutputFiles()

    'There are now so many output files that they should be in a separate folder.
        
        On Error GoTo CheckError
        
        'This must come first, so you know what the path is, even if it already exists.
            Let gOutputFilePath = gInputFilePath + "FilesFor" + gFileName + "\"
        'Make the new folder.  If it already exists, the error will be handled.
            MkDir gInputFilePath + "FilesFor" + gFileName + "\"
        
        Exit Sub
        
CheckError:
    
    Select Case Err.Number
        Case 75 'Path/file access error.  Almost certainly:  because folder already exists.  Just move on.
            Exit Sub
        Case 76 'Path not found.  Probably mother folder doesn't exist.
            Exit Sub
        Case Else
            MsgBox "Program error:  please contact Bruce Hayes at bhayes@humnet.ucla.edu, providing a copy of your input file and specifying error #50051.", vbCritical
            End
    End Select
        
End Sub

Function RedimensionTheArrays() As Boolean

    'Open the target input file, and see how big it is, so you can
    '  dimension the arrays at an appropriate size.
    
        'KZ: I changed this from a sub to a boolean function so that it can tell
        'the function that called it not to proceed if the input file couldn't be opened
        '(false means abort, true means proceed).

    'On Error GoTo CheckError
    
    Dim CheckThisLine As String
    Dim LocalNumberOfRivals As Long
    
    Dim TradInFile As Long
    Let TradInFile = FreeFile

    'Never open a file that isn't there.
        If Dir(gInputFilePath + gFileName + gFileSuffix) <> "" Then
            Open gInputFilePath + gFileName + gFileSuffix For Input As #TradInFile
        Else
            MsgBox "Sorry, I can't find the file " + _
                gInputFilePath + gFileName + gFileSuffix + ".  Click OK to continue.", vbExclamation
            Let RedimensionTheArrays = False
            Exit Function
        End If

    'Initialize the various maxima, in case you've opened a file before.
    
        Let mMaximumNumberOfForms = 0
        Let MaximumNumberOfConstraints = 0
        Let mMaximumNumberOfRivals = 0
        Let mMaximumNumberOfAPrioriRankings = 100000

    'Do different versions, depending on the kind of input file.
    
        Select Case gFileSuffix
            Case ".in"
                Do
                   If EOF(TradInFile) Then Exit Do
                   Input #TradInFile, CheckThisLine
                   Let CheckThisLine = Trim(CheckThisLine)
            
                   If LCase$(Left$(CheckThisLine, 13)) = "constraint:  " Then
                      Let MaximumNumberOfConstraints = MaximumNumberOfConstraints + 1
                   ElseIf LCase$(Left$(CheckThisLine, 8)) = "input:  " Then
                      Let mMaximumNumberOfForms = mMaximumNumberOfForms + 1
                      If LocalNumberOfRivals > mMaximumNumberOfRivals Then
                         Let mMaximumNumberOfRivals = LocalNumberOfRivals
                      End If
                      Let LocalNumberOfRivals = 0
                   ElseIf LCase$(Left$(CheckThisLine, 8)) = "rival:  " Then
                      Let LocalNumberOfRivals = LocalNumberOfRivals + 1
                   ElseIf LCase$(Left$(CheckThisLine, 17)) = "a priori ranking:" Then
                      Let mMaximumNumberOfAPrioriRankings = mMaximumNumberOfAPrioriRankings + 1
                      'Debug.Print "Max 3 of a priori rankings is now "; mMaximumNumberOfAPrioriRankings
                   End If
                Loop
            
            Case ".txt"
                'Same job, for tab-delimited text files.
                    Dim ConstraintsOver As Boolean
                    Let ConstraintsOver = False
                    Dim AbbreviationsOver As Boolean
                    Let AbbreviationsOver = False
                    Dim Buffer As String
                    Dim LastInput As String
                    Dim i As Long
                    
                    'To count constraints, start at -2 and add the number of tabs
                    '  on the first line.  This is because there are three blank
                    '  columns (-3), but the last guy is terminal and is not followed
                    '  by a tab.
                        Let MaximumNumberOfConstraints = -2
                        Let mMaximumNumberOfForms = 0
                        Let mMaximumNumberOfRivals = 0
                    
                    Do
RestartPoint:
                        If EOF(TradInFile) Then Exit Do
                        Line Input #TradInFile, CheckThisLine
                        Let CheckThisLine = Trim(CheckThisLine)
                        'Tolerate blank lines.
                            If CheckThisLine = "" Then GoTo RestartPoint
                        Let Buffer = ""
                        For i = 1 To Len(CheckThisLine)
                            'Until you hit a tab, merely augment the buffer:
                            If Mid(CheckThisLine, i, 1) <> Chr(9) Then
                                Let Buffer = Buffer + Mid(CheckThisLine, i, 1)
                            Else
                                If ConstraintsOver = False Then
                                    'You're on the first line of the file, so count constraints.
                                    Let MaximumNumberOfConstraints = MaximumNumberOfConstraints + 1
                                ElseIf AbbreviationsOver = True Then
                                    'Look to count inputs.
                                    Let Buffer = Trim(Buffer)
                                    If Buffer <> "" Then   'let user have repeated inputs, so remove:  And Buffer <> LastInput Then
                                        'You have a new input.
                                            Let mMaximumNumberOfForms = mMaximumNumberOfForms + 1
                                            Let LocalNumberOfRivals = 0
                                            Let LastInput = Buffer
                                            'Only the first column matters for inputs and rivals.
                                                Exit For
                                    Else
                                        'You're continuing on rivals for the same input
                                            Let LocalNumberOfRivals = LocalNumberOfRivals + 1
                                            If LocalNumberOfRivals > mMaximumNumberOfRivals Then
                                               Let mMaximumNumberOfRivals = LocalNumberOfRivals
                                            End If
                                            'Only the first column matters for inputs and rivals.
                                                Exit For
                                    End If
                                    Exit For
                                End If
                            End If
                        Next i
                        
                        If ConstraintsOver = True Then
                            Let AbbreviationsOver = True
                        End If
                        Let ConstraintsOver = True
                                             
                    Loop
                    
        End Select
      
      'It's possible that the biggest number of rival candidates is
      '  for the last input form in the file.  So check:

          If LocalNumberOfRivals > mMaximumNumberOfRivals Then
             Let mMaximumNumberOfRivals = LocalNumberOfRivals
          End If
      
      'Close the file, so that when you open it again, you'll be at the beginning.
          Close #TradInFile
    
    'Do the actual redimensioning.
    '   Returns false if failed.
        'this is causing trouble!
        Call RedimensioningFinalStep
        Let RedimensionTheArrays = True
        'If RedimensioningFinalStep = False Then
        '    'Let RedimensionTheArrays = False
        'Else
        '    Let RedimensionTheArrays = True
        'End If
       '
       Exit Function
    
CheckError:
    
    Select Case Err.Number
        Case 53 'File not found.
            MsgBox "I can't find your input file, which is supposed to be here:" + _
            Chr(10) + Chr(10) + _
            gInputFilePath + gFileName + gFileSuffix + _
            Chr(10) + Chr(10) + _
            "Please look for the file you want to rank using the Work With Different File button.", vbCritical
            Let RedimensionTheArrays = False    'KZ: let calling function know that the file
                                                'wasn't found.
            Exit Function
        Case Else
            MsgBox "Program error.  Please contact Bruce Hayes as bhayes@humnet.ucla.edu, including a copy of your input file and specifying error #77739.", vbCritical
    End Select
      
End Function


Function FindMaximumNumberOfRivals(MyArray() As Long) As Long

    'This gets called by other routines, so we don't have to pass to them the predictable parameter of
    '   MaximumNumberOfRivals.
    
    Dim i As Long
    For i = 1 To UBound(MyArray())
        If MyArray(i) > FindMaximumNumberOfRivals Then
            Let FindMaximumNumberOfRivals = MyArray(i)
        End If
    Next i

End Function

Function RedimensioningFinalStep() As Boolean
      
    'Do the actual redimensioning.  Returns false if it fails.

    'Since in a factorial typology run, the winner is going to be installed as
    '  a mere rival, increase the number of rivals by one.
      
    'Here is a point where it may be possible to detect non-OTSoft .txt files, and
    ' gracefully point them out before exiting.  This only works for the .txt and .in
    '   files, which (for reasons I can't remember, sigh) get predimensioned.
        If gFileSuffix = ".in" Or gFileSuffix = ".txt" Then
            If MaximumNumberOfConstraints <= 0 Then
                MsgBox "There's a problem with your input file:  I couldn't detect any constraints in it.  Please edit the file and try again." + _
                    Chr(10) + Chr(10) + _
                    "Click OK to exit OTSoft.", vbCritical
                    End
            End If
            If mMaximumNumberOfForms <= 0 Then
                MsgBox "There's a problem with your input file:  I couldn't detect any input forms in it.  Please edit the file and try again." + _
                    Chr(10) + Chr(10) + _
                    "Click OK to exit OTSoft.", vbCritical
                    End
            End If
        End If
    
    'Experimental:  (I think this must have something to do with structural descriptions.)
        Let mMaximumNumberOfForms = mMaximumNumberOfForms + MaximumNumberOfConstraints
    
    'A safety margin; more get added later.
       Let mMaximumNumberOfRivals = mMaximumNumberOfRivals + 3

        ReDim mConstraintName(MaximumNumberOfConstraints)
        ReDim mAbbrev(MaximumNumberOfConstraints)
        ReDim BackupConstraintName(MaximumNumberOfConstraints)
        ReDim BackupAbbrev(MaximumNumberOfConstraints) As String
        ReDim mAPrioriRankingsList(mMaximumNumberOfAPrioriRankings, 1)
        ReDim gAPrioriRankingsTable(MaximumNumberOfConstraints, MaximumNumberOfConstraints)
        ReDim mInputForm(mMaximumNumberOfForms)
        ReDim mWinner(mMaximumNumberOfForms)
        ReDim mWinnerFrequency(mMaximumNumberOfForms)
        ReDim mNumberOfRivals(mMaximumNumberOfForms)
        ReDim mRival(mMaximumNumberOfForms, mMaximumNumberOfRivals)
        ReDim mRivalFrequency(mMaximumNumberOfForms, mMaximumNumberOfRivals)
        ReDim mWinnerViolations(mMaximumNumberOfForms, MaximumNumberOfConstraints)
        ReDim mRivalViolations(mMaximumNumberOfForms, mMaximumNumberOfRivals, MaximumNumberOfConstraints)
        ReDim StillInformative(mMaximumNumberOfForms, mMaximumNumberOfRivals)
        ReDim Demotable(MaximumNumberOfConstraints)
        ReDim LocalRival(mMaximumNumberOfForms, mMaximumNumberOfRivals)
        ReDim LocalRivalViolations(mMaximumNumberOfForms, mMaximumNumberOfRivals, MaximumNumberOfConstraints) As Long
        ReDim LocalNumberOfRivals(mMaximumNumberOfForms)
        ReDim mFaithfulness(MaximumNumberOfConstraints)
        
        
    'Report success
        Let RedimensioningFinalStep = True

End Function

Sub SaveInputFile()
     
   'Save the file as it stands.  There are options both for traditional .in files, for
   '    tab-delimited text, and for Excel files.

    Select Case gFileSuffix
        Case ".in"
            Call mnuSaveAsIn_Click
        Case ".txt"
            Call SaveAsTxt(gFileName, False, mWinner())
        Case ".xls"
            MsgBox "Sorry, this program can't save your file as an Excel file.  Please use the Tab-delimited text option, open the resulting .txt file in Excel, then save the result as an Excel Workbook.", vbExclamation
        Case Else
            MsgBox "Illegal file suffix type.", vbExclamation
    End Select
    
End Sub

Sub RefreshRecentlyOpenedFiles(MyFileName As String)
    
    'Update the OTSoftRecentlyOpenedFiles.txt file.
        
        Dim ROFIndex As Long
        Dim InnerROFIndex As Long
        
        Call ReadTheRecentlyOpenedFilesFile
        Let MyFileName = gInputFilePath + gFileName + gFileSuffix
        
        For ROFIndex = 1 To 6
            Select Case mRecentlyOpenedFiles(ROFIndex)
                Case ""
                    'There's a slot at the end, which your should fill.
                        Let mRecentlyOpenedFiles(ROFIndex) = MyFileName
                        Call WriteRecentlyOpenedFilesFile
                        Exit Sub
                Case MyFileName
                    'It's already here, but not at the top.
                        'Demote the intervenors.
                        For InnerROFIndex = ROFIndex To 2 Step -1
                            Let mRecentlyOpenedFiles(InnerROFIndex) = mRecentlyOpenedFiles(InnerROFIndex - 1)
                        Next InnerROFIndex
                        'Install the new file name at the top.
                            Let mRecentlyOpenedFiles(1) = MyFileName
                            Call WriteRecentlyOpenedFilesFile
                            Exit Sub
            End Select
        Next ROFIndex

        'If you got this far, the new file was not on the old list.
        '   Demote everybody, and put it at the top.
            For ROFIndex = 6 To 2 Step -1
                Let mRecentlyOpenedFiles(ROFIndex) = mRecentlyOpenedFiles(ROFIndex - 1)
            Next ROFIndex
            'Install the new file name at the top.
                Let mRecentlyOpenedFiles(1) = MyFileName
                Call WriteRecentlyOpenedFilesFile

End Sub

Sub WriteRecentlyOpenedFilesFile()

    Dim ROFFile As Long
    Dim i As Long
    
    Let ROFFile = FreeFile
    Open gSafePlaceToWriteTo + "\OTSoftRecentlyOpenedFiles.txt" For Output As ROFFile
    For i = 1 To 6
        If Trim(mRecentlyOpenedFiles(i)) <> "" Then
            Print #ROFFile, mRecentlyOpenedFiles(i)
        Else
            Exit For
        End If
    Next i
    Close ROFFile
    
    'Put them on the menu labels.
        Call UpdateOpenRecentMenu
    
End Sub

Sub ReadTheRecentlyOpenedFilesFile()
    
    On Error GoTo CheckError
    
    Dim ROFFile As Long
    Dim i As Long
    Dim Buffer As String
    
    'If OTSoft has only recently been installed, then OTSoftRecentlyOpenedFiles.txt might
    '   not be in the correct location; i.e. the gSafePlaceToWriteTo.
    '   May 2005:  Ack, this is overwriting the correct file!  Let's try = instead of <>.
        
        'If Dir(gSafePlaceToWriteTo + "\OTSoftRecentlyOpenedFiles.txt") <> "" Then
        If Dir(gSafePlaceToWriteTo + "\OTSoftRecentlyOpenedFiles.txt") = "" Then
            If gSafePlaceToWriteTo <> App.Path Then
                FileCopy App.Path + "\OTSoftRecentlyOpenedFiles.txt", gSafePlaceToWriteTo + "\OTSoftRecentlyOpenedFiles.txt"
            End If
        End If
        
    'Initialize the list:
        For i = 1 To 6
            Let mRecentlyOpenedFiles(i) = ""
        Next i
    'Read as many as there are on the list:
        Let ROFFile = FreeFile
        
        'Don't open a file that doesn't exist; priority here is very low.
            If Dir(gSafePlaceToWriteTo + "\OTSoftRecentlyOpenedFiles.txt") <> "" Then
                Open gSafePlaceToWriteTo + "\OTSoftRecentlyOpenedFiles.txt" For Input As ROFFile
            Else
                Exit Sub
            End If
        
        Let i = 0
        Do While Not EOF(ROFFile)
            Line Input #ROFFile, Buffer
            If Trim(Buffer) <> "" Then
                Let i = i + 1
                Let mRecentlyOpenedFiles(i) = Buffer
            Else
                Exit Do
            End If
        Loop
        Close ROFFile
        
        Exit Sub
        
CheckError:
    
    'Try getting it from the program folder:
        If Dir(App.Path + "\OTSoftRecentlyOpenedFiles.txt") <> "" Then
            Open App.Path + "\OTSoftRecentlyOpenedFiles.txt" For Input As ROFFile
        Else
            Exit Sub
        End If
        
End Sub


Sub UpdateOpenRecentMenu()

    'Make visible, and put appropriate items on the Open Recent menu.
    
        Call ReadTheRecentlyOpenedFilesFile
        
        If mRecentlyOpenedFiles(1) <> "" Then
            Let mnuOpenRecent1.Caption = "&1 " + mRecentlyOpenedFiles(1)
        Else
            Let mnuOpenRecent1.Caption = ""
        End If
        If mRecentlyOpenedFiles(1) <> "" Then
            Let mnuOpenRecent2.Caption = "&2 " + mRecentlyOpenedFiles(2)
        Else
            Let mnuOpenRecent2.Caption = ""
        End If
        If mRecentlyOpenedFiles(3) <> "" Then
            Let mnuOpenRecent3.Caption = "&3 " + mRecentlyOpenedFiles(3)
        Else
            Let mnuOpenRecent3.Caption = ""
        End If
        If mRecentlyOpenedFiles(4) <> "" Then
            Let mnuOpenRecent4.Caption = "&4 " + mRecentlyOpenedFiles(4)
        Else
            Let mnuOpenRecent4.Caption = ""
        End If
        If mRecentlyOpenedFiles(5) <> "" Then
            Let mnuOpenRecent5.Caption = "&5 " + mRecentlyOpenedFiles(5)
        Else
            Let mnuOpenRecent5.Caption = ""
        End If
        If mRecentlyOpenedFiles(6) <> "" Then
            Let mnuOpenRecent6.Caption = "&6 " + mRecentlyOpenedFiles(6)
        Else
            Let mnuOpenRecent6.Caption = ""
        End If
        
End Sub

Private Sub mnuSaveAsIn_Click()

   'For what it's worth, save the input file in the traditional Ranker format of the mid 1990's.
   
    Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
    Dim PositionIndex As Long
    Dim APrioriRankingIndex As Long
     
    Dim TradInFile As Long
    Let TradInFile = FreeFile
    
    Call CreateAFolderForOutputFiles
    
    Open gInputFilePath + gFileName + ".in" For Output As #TradInFile
    Print #TradInFile, "User name:  "; mUserName

   For ConstraintIndex = 1 To mNumberOfConstraints
      Print #TradInFile, "Constraint:  ";
      Print #TradInFile, mConstraintName(ConstraintIndex)
      Print #TradInFile, "  Abbreviation:  ";
      Print #TradInFile, mAbbrev(ConstraintIndex)
   Next ConstraintIndex

   For APrioriRankingIndex = 1 To mNumberOfAPrioriRankings
      Print #TradInFile, "a priori ranking:"
      Print #TradInFile, "  "; mAPrioriRankingsList(APrioriRankingIndex, 0)
      Print #TradInFile, "  "; mAPrioriRankingsList(APrioriRankingIndex, 1)
   Next APrioriRankingIndex

   For FormIndex = 1 To mNumberOfForms
      Print #TradInFile, "Input:  ";
      Print #TradInFile, mInputForm(FormIndex)

      Print #TradInFile, "  Winner:  ";
      Print #TradInFile, mWinner(FormIndex)

      For ConstraintIndex = 1 To mNumberOfConstraints
         Print #TradInFile, "    ";
         Print #TradInFile, mAbbrev(ConstraintIndex); ",";
         For PositionIndex = Len(mAbbrev(ConstraintIndex)) To 7
            Print #TradInFile, " ";
         Next PositionIndex
         Print #TradInFile, mWinnerViolations(FormIndex, ConstraintIndex)
      Next ConstraintIndex

      For RivalIndex = 1 To mNumberOfRivals(FormIndex)
         Print #TradInFile, "  Rival:  ";
         Print #TradInFile, mRival(FormIndex, RivalIndex)
         For ConstraintIndex = 1 To mNumberOfConstraints
            Print #TradInFile, "    ";
            Print #TradInFile, mAbbrev(ConstraintIndex); ",";
            For PositionIndex = Len(mAbbrev(ConstraintIndex)) To 7
               Print #TradInFile, " ";
            Next PositionIndex
            Print #TradInFile, mRivalViolations(FormIndex, RivalIndex, ConstraintIndex)
         Next ConstraintIndex
      Next RivalIndex
   
   Next FormIndex

End Sub


Private Sub SaveAsTxt(MyFileName As String, SortByRank As Boolean, Winner() As String)

    'Save the input file.
    
    'On Error GoTo CheckError
    
    Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
    Dim PositionIndex As Long
    Dim ReportErrorFileName As String
    
    'Close #TextFile
    Dim TextFile As Long
    Let TextFile = FreeFile
    
    Call CreateAFolderForOutputFiles
    
    Let ReportErrorFileName = MyFileName + ".txt"
    
    'Backups go in the downstairs folder; straightforward saves go in the upstairs folder.
        If LCase(Right(MyFileName, 6)) = "backup" Then
            Open gOutputFilePath + MyFileName + ".txt" For Output As #TextFile
        Else
            Open gInputFilePath + MyFileName + ".txt" For Output As #TextFile
        End If
    
    'This needs to be a separate bit of code because other places in OTSoft use it.
        Call PrintContentOfAnInputFile(False, TextFile, mNumberOfConstraints, mConstraintName(), mAbbrev(), _
            mNumberOfForms, mInputForm(), mWinner(), mWinnerFrequency(), mWinnerViolations(), mNumberOfRivals(), _
            mRival(), mRivalFrequency(), mRivalViolations())
        
        Close #TextFile
        
    Exit Sub
    
CheckError:
    Select Case Err.Number  ' Evaluate error number.
        Case 70 ' "File already open" error.
            MsgBox "Error.  Probably what is happening is this:  I'm trying to open the file " + _
                gInputFilePath + ReportErrorFileName + " for purposes of storing my results, but a file of this name is already open.  I suggest you try to find this file, close it, then click OK.", vbExclamation
            Resume
        Case 75 ' "File access error
            MsgBox "Error.  I conjecture that " + ReportErrorFileName + " already exists in " + _
                gInputFilePath + " as a Read-Only file.  Try deleting this file (or right click, Properties, decheck Read-Only) and rerunning OTSoft.", vbExclamation
            End
        Case Else
            MsgBox "Program error.  You can ask for help at bhayes@humnet.ucla.edu.  Please send a copy of your input file with your message.", vbCritical
            End
    End Select

End Sub

    

Sub DeleteTmpFiles()

    On Error GoTo ErrorLine

    If Dir(gOutputFilePath + gFileName + "DraftOutput.txt") <> "" Then
        Kill gOutputFilePath + gFileName + "DraftOutput.txt"
    End If
    'If Dir(gOutputFilePath + "HowIRanked" + gFileName + ".txt") <> "" Then
    '    Kill gOutputFilePath + "HowIRanked" + gFileName + ".txt"
    'End If
    If Dir(gOutputFilePath + gFileName + "Hasse.txt") <> "" Then
        Kill gOutputFilePath + gFileName + "Hasse.txt"
    End If
    If Dir(gOutputFilePath + gFileName + ".sav") <> "" Then
        Kill gOutputFilePath + gFileName + ".sav"
    End If
    

ErrorLine:
    Exit Sub

End Sub


'===============================CONSTRAINT RANKING=================================
'==================================================================================

Sub Rank()
      
    'Call the routines needed to do discrete ranking. (The stochastic algorithms have their own interfaces.)
         
       'On Error GoTo CheckError
            
        Dim ReportErrorFileName As String
        Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
       
        Dim MyRankingResult As String
            'Holds the algorithm result in the form of a string;
            'For ease of communication between modules; I find it very hard to pass arrays across modules in VB.
            'Format is tab separated text.
            '   (True/False), Number of strata, stratum of each constraint) where True means "converged"
            'So each ranking algorithm is a function from the violation arrays to a single string.
    
    'Avoid confusion by deactivating the Rank and Factorial Typology buttons.
        Let cmdRank.Enabled = False
        Let cmdFacType.Enabled = False
    
    'If the user has gotten this far, (s)he probably wants the settings saved.
        Call SaveUserChoices
        
    'The user has chosen what algorithm to use; record this in the global variable hithere
        Call RecordAlgorithmChoice
    
    'Open the output files.  Know in advance what file this will be, in case it's already open.
    
        'First, make sure there is a folder for these files, a daughter of the
        '   folder in which the input file is located.
            Call CreateAFolderForOutputFiles
   '         'The quick, draft output:
                Let ReportErrorFileName = gFileName + "DraftOutput.txt"
                Let mTmpFile = FreeFile
                Open gOutputFilePath + gFileName + "DraftOutput.txt" For Output As #mTmpFile
   '             'Initialize the header numbers, in case this isn't the first run.
                    Let gLevel1HeadingNumber = 0
            'The output for pretty Word file conversion:
                Let ReportErrorFileName = gFileName + "QualityOutput.txt"
                Let mDocFile = FreeFile
                Open gOutputFilePath + gFileName + "QualityOutput.txt" For Output As #mDocFile
            'The HTML output:
                Let ReportErrorFileName = "ResultsFor" + gFileName + ".htm"
                Let mHTMFile = FreeFile
                Open gOutputFilePath + "ResultsFor" + gFileName + ".htm" For Output As #mHTMFile
                Call PrintTableaux.InitiateHTML(mHTMFile)
               
        If gHaveIOpenedTheFile = False Then
            If DigestTheInputFile(gInputFilePath, gFileName, gFileSuffix) = False Then
                Close
                Let cmdRank.Enabled = True
                Let cmdFacType.Enabled = True
                Exit Sub    'KZ: false means the file couldn't be opened
            End If
        End If
        Let gHaveIOpenedTheFile = True
    
    'Launch the ranking algorithm that the user selected.
        If optGLA.Value = True Then
        
                'Debug  1/1/26
                '    Dim DebugFile As Long
                '    Let DebugFile = FreeFile
                '    Open gOutputFilePath + "DebugOnWayToGLA.txt" For Output As #DebugFile
                '    Dim i As Long, j As Long
                '    For i = 1 To mNumberOfForms
                '        Print #DebugFile, "Input:"; vbTab; mInputForm(i)
                '        Print #DebugFile, vbTab; "Winner:  ["; vbTab; mWinner(i); "]"; vbTab; "Frequency"; vbTab; mWinnerFrequency(i)
                '        For j = 1 To mNumberOfRivals(mNumberOfForms)
                '            Print #DebugFile, vbTab; "RivalIndex:"; vbTab; Trim(Str(j)); vbTab; "Rival:"; vbTab; mRival(i, j); vbTab; "Frequency:"; vbTab; mRivalFrequency(i, j)
                '        Next j
                '        Print #DebugFile,
                '    Next i
                '    Close #DebugFile
        
            'Specfically, we will use the GLA to do Stochastic OT.
                Let GLA.optStochasticOT.Value = True
            'Call it.
                Call GLA.Main(mNumberOfForms, mInputForm(), mWinner(), mWinnerFrequency(), mWinnerViolations(), _
                     mNumberOfRivals(), mRival(), mRivalFrequency(), mRivalViolations(), _
                     mNumberOfConstraints, mConstraintName(), mAbbrev, _
                     mTmpFile, mDocFile, mHTMFile, "StochasticOT")
            'Crucial that the user not be stranded by inability to click on Rank or FacType.
                Let cmdRank.Enabled = True
                Let cmdFacType.Enabled = True
           Exit Sub
        ElseIf optMaximumEntropy.Value = True Then
            'This no longer calls the mediocre batch MaxEnt; which can be accessed from the GLA screen.
            'Instead, it calls the GLA screen, specifying MaxEnt as the relevant option.
                Let GLA.optMaxEnt.Value = True
                Call GLA.Main(mNumberOfForms, mInputForm(), mWinner(), mWinnerFrequency(), mWinnerViolations(), _
                     mNumberOfRivals(), mRival(), mRivalFrequency(), mRivalViolations(), _
                     mNumberOfConstraints, mConstraintName(), mAbbrev, _
                     mTmpFile, mDocFile, mHTMFile, "MaxEnt")
                 'Crucial that the user not be stranded by inability to click on Rank or FacType.
                     Let cmdRank.Enabled = True
                     Let cmdFacType.Enabled = True
                Exit Sub
                    'Debug
                    'Dim DebugFile As Long
                    'Let DebugFile = FreeFile
                    'Open gOutputFilePath + "DebugOnWayToMaxent.txt" For Output As #DebugFile
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

                    'Old code:
                    'Call MyMaxEnt.Main(mNumberOfForms, mInputForm(), mWinner(), mWinnerFrequency(), mWinnerViolations(), _
                    '    mNumberOfRivals(), mRival(), mRivalFrequency(), mRivalViolations(), _
                    '    mNumberOfConstraints, mConstraintName(), mAbbrev, _
                    '    mTmpFile, mDocFile, mHTMFile)
                    '    Let cmdRank.Enabled = True
                    '    Let cmdFacType.Enabled = True
                    'Exit Sub
        ElseIf optNoisyHarmonicGrammar.Value = True Then
                'Debug
                '    'Dim DebugFile As Long
                '    Let DebugFile = FreeFile
                '    Open gOutputFilePath + "DebugOnWayToNHG.txt" For Output As #DebugFile
                '    'Dim i As Long, j As Long
                '    For i = 1 To mNumberOfForms
                '        Print #DebugFile, "Input:"; vbTab; mInputForm(i)
                '        Print #DebugFile, vbTab; "Winner:  ["; vbTab; mWinner(i); "]"; vbTab; "Frequency"; vbTab; mWinnerFrequency(i)
                '        For j = 1 To mNumberOfRivals(mNumberOfForms)
                '            Print #DebugFile, vbTab; "RivalIndex:"; vbTab; Trim(Str(j)); vbTab; "Rival:"; vbTab; mRival(i, j); vbTab; "Frequency:"; vbTab; mRivalFrequency(i, j)
                '        Next j
                '        Print #DebugFile,
                '    Next i
                '    Close #DebugFile
            Call NoisyHarmonicGrammar.Main(mNumberOfForms, mInputForm(), mWinner(), mWinnerFrequency(), mWinnerViolations(), _
                mNumberOfRivals(), mRival(), mRivalFrequency(), mRivalViolations(), _
                mNumberOfConstraints, mConstraintName(), mAbbrev, _
                mTmpFile, mDocFile, mHTMFile)
                Let cmdRank.Enabled = True
                Let cmdFacType.Enabled = True
        Else
            'It will be some version of Recursive Constraint Demotion.
            'First, exercise caution:  if there are multiple winners in the input, give the user the opportunity to cancel.
                If CheckChoiceOfAlgorithm = False Then
                    Let cmdRank.Enabled = True
                    Let cmdFacType.Enabled = True
                    Close #mTmpFile
                    Close #mDocFile
                    Close #mHTMFile
                    Exit Sub
                End If
                
            'Now call the appropriate algorithm.  These are all functions that yield strings;
            '   see beginning of this procedure for the rationale.
                'This will be some version of Recursive Constraint Demotion, depending on menu choices made.
                    If mnuLowFaithfulness.Checked Then
                        'BH's Low Faithfulness Constraint Demotion.
                             Let MyRankingResult = LowFaithfulnessConstraintDemotion.Main(mNumberOfForms, mInputForm(), _
                                     mNumberOfRivals(), _
                                     mWinnerViolations(), mRival(), mRivalViolations(), _
                                     mNumberOfConstraints, mAbbrev(), mConstraintName())
                    ElseIf mnuBiasedConstraintDemotion.Checked Then
                        'The code from Bruce Tesar, implementing his and Prince's Biased Constraint Demotion.
                            Let MyRankingResult = BCD.Main(mNumberOfForms, mNumberOfRivals(), _
                                    mWinnerViolations(), mRivalViolations(), _
                                    mNumberOfConstraints, mAbbrev(), mConstraintName())
                    Else
                        'Classical Batch Recursive Constraint Demotion (Tesar and Smolensky 1993)
                            Let MyRankingResult = RecursiveConstraintDemotion.Main(mNumberOfForms, mInputForm(), mNumberOfRivals(), _
                                    mWinnerViolations(), mRival(), mRivalViolations(), mNumberOfConstraints, mAbbrev(), mConstraintName())
                    End If
             
            'Decode the ranking result from a single string to something more usable.
                'Convergence:
                    Select Case s.Chomp(MyRankingResult)
                        Case "True"
                            Let TSResult.Converged = True
                        Case Else
                            Let TSResult.Converged = False
                    End Select
                'Number of strata:
                    Let MyRankingResult = s.Residue(MyRankingResult)
                    Let TSResult.NumberOfStrata = Val(s.Chomp(MyRankingResult))
                'Stratal membership of each constraint:
                    For ConstraintIndex = 1 To mNumberOfConstraints
                        Let MyRankingResult = s.Residue(MyRankingResult)
                        Let TSResult.Stratum(ConstraintIndex) = Val(s.Chomp(MyRankingResult))
                    Next ConstraintIndex

            'Print a header on the output files
                Call PrintRankerHeader
            
            'Do the things that get irrespective of whether the algorithm converged or not.
                'This is a function, so you can skip hooey if it turns out to be true.
                    If CheckSubsetRelations = True Then
                        GoTo SkipHooeyPoint
                    End If
             
            'Report the result.
                Call SummarizeRankerResults(TSResult.Converged, TSResult.Stratum(), TSResult.NumberOfStrata, mConstraintName())
            
            'Take appropriate action depending on whether the algorithm converged or not.
                
                If TSResult.Converged = True Then
                    'Print the grammar (strata).
                    'Print tableaux.  The zero at the end means, "not doing factorial typology"
                    
                        Call PrintLevel1Header(mDocFile, mTmpFile, mHTMFile, "Tableaux")
                        Call PrintTableaux.Main(mNumberOfForms, mNumberOfConstraints, mConstraintName(), _
                            mAbbrev(), TSResult.Stratum(), mInputForm(), mWinner(), mWinnerFrequency(), mWinnerViolations(), _
                            mMaximumNumberOfRivals, mNumberOfRivals(), mRival(), mRivalFrequency(), mRivalViolations(), mTmpFile, mDocFile, mHTMFile, _
                            gAlgorithmName, RunningFactorialTypology, 0, True, True)
                     
                     'Call the unnecessary constraint finder.
                        Call FindUnnecessaryConstraints(mConstraintName(), mAbbrev(), _
                            mWinnerViolations(), mNumberOfRivals(), mRivalViolations())
    
                     'Save the a priori rankings if it was so requested.
                        If SaveStrataAsAPrioriRankings = True Then
                            Call APrioriRankings.PrintAPrioriRankingsConvertedFromStrataAsTable(TSResult.NumberOfStrata, TSResult.Stratum, mNumberOfConstraints, mAbbrev())
                        End If
                 
                    'Call the arguer if the button is checked.
                        If chkArguerOn.Value = Checked Then
                            'Frequencies are bogus here, but needed for printing tableaux.
                                ReDim mWinnerFrequency(mNumberOfForms)
                                Call Fred.Main(mNumberOfForms, mInputForm(), mWinner(), mWinnerFrequency(), mRival(), mRivalFrequency(), mMaximumNumberOfRivals, mNumberOfRivals(), _
                                    mNumberOfConstraints, mConstraintName(), mAbbrev(), TSResult.Stratum(), mWinnerViolations(), mRivalViolations(), _
                                    RunningFactorialTypology, 0, mTmpFile, mDocFile, mHTMFile)
                        End If
                                
                Else            'That is to say, if TSResult.Converged = False
                    
                    'If ranking fails, and the user wanted it, provide diagnostics.
                         If chkDiagnosticTableaux.Value = vbChecked Then
                            'The function LookForMinimalPairs will look for, and if found, print, the best kind of diagnostic:  minimal pairs.
                                If LookForMinimalPairs() = False Then
                                    Call PrepareDiagnosticTableaux(mNumberOfConstraints, mAbbrev(), mConstraintName(), TSResult.NumberOfStrata, TSResult.Stratum(), mNumberOfForms, _
                                        mInputForm(), mWinner(), mWinnerViolations(), mWinnerFrequency(), mNumberOfRivals(), mMaximumNumberOfRivals, mRival(), mRivalViolations(), mRivalFrequency(), _
                                        mTmpFile, mDocFile, mHTMFile)
                                End If
                         End If     'Did user want diagnostic tableaux?
                End If     'Did the algorithm converge?
             
SkipHooeyPoint:
            
            'Close output files.
                Close #mTmpFile
                Close #mDocFile
                Print #mHTMFile, "</BODY>"
                Close #mHTMFile
            
            'Announce you're done, and guide user to the View Results button.
                Let lblProgressWindow.Caption = "I'm done."
                Let cmdViewResults.Default = True
                Let cmdViewResults.Font.Size = 10
                Let cmdViewResults.FontBold = True
                Let cmdRank.Caption = "Rank " + gFileName + gFileSuffix

            'Reactivate the Rank and Factorial Typology buttons.
                Let cmdRank.Enabled = True
                Let cmdFacType.Enabled = True
                
            'Get ready to View Results
                cmdViewResults.SetFocus
                Let gHasTheProgramBeenRun = True
          
        End If  'choice of stochastic algorithms versus versions of Constraint Demotion
        
    'Since you've now loaded the file, the user might be interested in reloading an updated version.
        Let mnuReload.Caption = "Reload " + gFileName + gFileSuffix
        Let mnuReload.Visible = True
       
        Exit Sub
        
CheckError:
    Select Case Err.Number  ' Evaluate error number.
        Case 70 ' "File already open" error.
            MsgBox "Error.  Probably what is happening is this:  I'm trying to open the file " + _
                gOutputFilePath + ReportErrorFileName + " for purposes of storing my results, but a file of this name is already open.  I suggest you try to find this file, close it, then click OK.", vbExclamation
            Resume
        Case 71 'Can't find disk.
            MsgBox "Error.  Windows reports to me that " + Chr(34) + "a disk is not ready." + Chr(34) + "  Try inserting the relevant disk, or looking for your file by clicking the Work With A Different File button." + _
            Chr(10) + Chr(10) + _
            "Click OK to continue.", vbExclamation
            Exit Sub
        Case 75 ' "File access error
            MsgBox "Error.  I conjecture that " + ReportErrorFileName + " already exists in " + _
                gOutputFilePath + " as a Read-Only file.  Try deleting this file (or right click, Properties, decheck Read-Only) and rerunning OTSoft.", vbExclamation
            End
        Case 76 'Path not found error.
            MsgBox "Error.  I conjecture that you have moved your input file to a different directory.  Try clicking the Work With Different File button to relocate your input file.", vbExclamation
            Exit Sub
        Case 53
            MsgBox "I can't find this input file.  Please look for it using the Work With Different File button.", vbExclamation
        Case Else
            MsgBox "Program error, which the software calls " + Err.Description + ".  Please contact me at bhayes@humnet.ucla.edu, including a copy of your input file, and the error code 15092.", vbCritical
    End Select
            
End Sub

Sub RecordAlgorithmChoice()
            
    'This comes from the interface and is listed in a single string variable.
        If optGLA.Value = True Then
           Let gAlgorithmName = "GLA"
        ElseIf optMaximumEntropy.Value = True Then
           Let gAlgorithmName = "GLA"
        ElseIf optNoisyHarmonicGrammar.Value = True Then
            Let gAlgorithmName = "Noisy Harmonic Grammar"
        ElseIf optConstraintDemotion.Value = True Then
            If mnuLowFaithfulness.Checked Then
                Let gAlgorithmName = "Low Faithfulness Constraint Demotion"
            ElseIf mnuBiasedConstraintDemotion.Checked Then
                Let gAlgorithmName = "Biased Constraint Demotion"
            Else
                Let gAlgorithmName = "Recursive Constraint Demotion"
            End If
        Else
            MsgBox "Fatal error #44501.  Please contact Bruce Hayes at bhayes@humnet.ucla.edu for help."
        End If

End Sub


Sub PrintRankerHeader()

     'Print a header for the output file:

         '\ft means 'first time'.  It's to prevent all the characters from being converted
         '  to Times Roman in cases where (due to lack of memory) the converter macro
         '  has to be run twice.
                  
         Call PrintTopLevelHeader(mDocFile, mTmpFile, mHTMFile, "Results of Applying " + gAlgorithmName + " to " + gFileName + gFileSuffix)

         'If the specific-gets-priority provision of BCD was used, say so:
            If gAlgorithmName = "Biased Constraint Demotion" Then
                If mnuSpecificBCD.Checked Then
                    Call PrintPara(mDocFile, mTmpFile, mHTMFile, "Version of BCD:  specific Faithfulness constraints get priority.")
                End If
            End If
            
     'Print a header and diacritic to trigger a page number
         Print #mDocFile, "\hrRankings for "; gFileName; gFileSuffix; Chr$(9); NiceDate; Chr$(9); "\pn"

     'Print centered date and time.
         Print #mTmpFile,
         Print #mTmpFile, NiceDate; ", "; NiceTime
         Print #mTmpFile,
         Print #mTmpFile, "OTSoft " + gMyVersionNumber + ", release date " + gMyReleaseDate
         Print #mTmpFile,
         Print #mDocFile,
         Print #mDocFile, "\cn"; NiceDate; ", "; NiceTime
         Print #mDocFile,
         Print #mDocFile, "\cnOTSoft " + gMyVersionNumber
         Print #mDocFile, "\cnRelease date " + gMyReleaseDate
         Print #mDocFile,
         Print #mHTMFile, "<p><p>"; NiceDate; ", "; NiceTime
         Print #mHTMFile, "<p><p>"; "OTSoft " + gMyVersionNumber + ", release date " + gMyReleaseDate

End Sub

Sub SummarizeRankerResults(ConvergenceFlag As Boolean, Stratum() As Long, _
    NumberOfStrata As Long, Constraint() As String)

    Dim Flag As Boolean
    Dim ConstraintIndex As Long
    Dim StratumIndex As Long
    Dim SpaceIndex As Long
    Dim RowCount As Long
    
   'Warn about ties.
        If mMoreThanOneWinner = True Then
            'Give the basic warning:
                Call PrintLevel1Header(mDocFile, mTmpFile, mHTMFile, "Cautionary Note")
                Call PrintPara(mDocFile, mTmpFile, mHTMFile, "The " + gAlgorithmName + " algorithm assumes that for each input, PARAthere is only one winning candidate.  However, your input file included PARAmultiple winners. " + _
                    "PARAFor purposes of ranking, OTSoft adopted the highest frequency candidate as PARAthe single " + Chr(34) + "winner" + Chr(34) + ".")
            
            'The warning is worse if there are frequency ties for the winner:
                If mGlobalTie = True Then
                    Call PrintPara(mDocFile, mTmpFile, mHTMFile, "Moreover, in at least one case, two candidate were tied for the highest PARAfrequency. OTSoft selected the first of these as the winner for purposes PARAof computation.")
                End If
            'Continue the basic warning:
                Call PrintPara(mDocFile, mTmpFile, mHTMFile, "Unless this is actually what you want to have happen, you may find it " + _
                    "PARAbetter to use a stochastic algorithm (GLA, MaxEnt, NHG) instead.")
        End If
      
   'Print good or bad news, depending on the contradiction flag.
       Select Case ConvergenceFlag
   
        'Good news:
            Case True
                Call PrintLevel1Header(mDocFile, mTmpFile, mHTMFile, "Result")
                Call PrintPara(mDocFile, mTmpFile, mHTMFile, "A ranking was found that generates the correct outputs.")
                
                'Start a table with the constraints in their strata.
                    Dim MyHTMLTable() As String
                    Dim MyDocTable() As String
                
                'Header:
                    ReDim MyDocTable(3, 1)
                    Let MyDocTable(1, 1) = "Stratum"
                    Let MyDocTable(2, 1) = "Constraint Name"
                    Let MyDocTable(3, 1) = "Abbreviation"
                    ReDim MyHTMLTable(3, 1)
                    Let MyHTMLTable(1, 1) = "Stratum"
                    Let MyHTMLTable(2, 1) = "Constraint Name"
                    Let MyHTMLTable(3, 1) = "Abbreviation"
                    Let RowCount = 1
                
                'Constraints:
                    For StratumIndex = 1 To NumberOfStrata
                        Print #mTmpFile, "   Stratum #"; Trim(Str(StratumIndex))
                        Let Flag = False
                        For ConstraintIndex = 1 To mNumberOfConstraints
                           If Stratum(ConstraintIndex) = StratumIndex Then
                              Let RowCount = RowCount + 1
                              ReDim Preserve MyHTMLTable(3, RowCount)
                              ReDim Preserve MyDocTable(3, RowCount)
                              Print #mTmpFile, "      "; Constraint(ConstraintIndex);
                              For SpaceIndex = Len(Constraint(ConstraintIndex)) To 35
                                Print #mTmpFile, " ";
                              Next SpaceIndex
                              Print #mTmpFile, "      "; mAbbrev(ConstraintIndex)
                              'Use the Flag to print a blank cell underneath the stratum label:
                                 If Flag = True Then
                                    Let MyHTMLTable(1, RowCount) = "&nbsp"
                                 Else
                                    Let MyHTMLTable(1, RowCount) = "Stratum #" + Trim(Str(StratumIndex))
                                    Let MyDocTable(1, RowCount) = "Stratum #" + Trim(Str(StratumIndex))
                                 End If
                                 Let Flag = True
                              Let MyDocTable(2, RowCount) = SmallCapTag1 + Constraint(ConstraintIndex) + SmallCapTag2
                              Let MyDocTable(3, RowCount) = SmallCapTag1 + mAbbrev(ConstraintIndex) + SmallCapTag2
                              Let MyHTMLTable(2, RowCount) = Constraint(ConstraintIndex)
                              Let MyHTMLTable(3, RowCount) = mAbbrev(ConstraintIndex)
                              
                           End If
                        Next ConstraintIndex
                     Next StratumIndex
                     
                'Finish the table:
                    Call s.PrintHTMTable(MyHTMLTable(), mHTMFile, True, False, False)
                    Call s.PrintDocTable(MyDocTable(), mDocFile, True, False, False)
                 

        'Bad news:
        
            Case False
                Call PrintLevel1Header(mDocFile, mTmpFile, mHTMFile, "Result")
                Call PrintPara(mDocFile, mTmpFile, mHTMFile, "There is no ranking of the proposed constraints which yields the correct outputs.  " + _
                    "PARAYou can learn more about why the constraints fail by selecting Show How Ranking Was Done from the View menu.")
               
               If chkDiagnosticTableaux.Value <> vbChecked Then
                    Print #mTmpFile,
                    Call PrintPara(mDocFile, mTmpFile, mHTMFile, "To obtain diagnostic tableaux, try running the program again, first checking the " + _
                        "PARADiagnostics if Ranking Fails box.")
                    Call PrintPara(mDocFile, mTmpFile, mHTMFile, "You can also diagnose by selecting Show How Ranking Was Done from the View menu.")
               End If
   
   End Select

    'In either event, show the a priori rankings if appropriate.
        If mnuConstrainAlgorithmsByAPrioriRankings.Checked = True Then
            Call PrintOutTheAprioriRankings(mTmpFile, mDocFile, mHTMFile)
        End If

End Sub

Function CheckChoiceOfAlgorithm() As Boolean
    
    'If the user is using a stochastic-type input file, with multiple winners and/or ties, but
    '   isn't using a stochastic algorithm, give a warning and permit the user to cancel.
        
        Dim MessageString As String                   'For message box.
        
        If LCase(gAlgorithmName) <> "gla" Then
            If mMoreThanOneWinner = True Then
                Let MessageString = "Caution:  you've selected a ranking algorithm that can't deal with multiple winners, but your input file contains some.  " + _
                    Chr(10) + Chr(10)
                If mGlobalTie Then
                    Let MessageString = MessageString + _
                        "Moreover, for at least one of your input forms, there are multiple candidates that are tied for highest frequency." + _
                        Chr(10) + Chr(10) + _
                        "If you continue, I will choose the first-listed highest-frequency candidate for each of your inputs to serve as the sole " + Chr(34) + "winning candidate" + Chr(34) + _
                        " for purposes of learning, and ignore the positive input frequencies of all other candidates."
                Else
                    Let MessageString = MessageString + _
                        "If you continue, I will choose the highest-frequency candidate for each of your inputs to serve as the sole " + Chr(34) + "winning candidate" + Chr(34) + _
                        " for purposes of learning, and ignore the positive input frequencies of all other candidates."

                End If
                Let MessageString = MessageString + _
                    Chr(10) + Chr(10) + _
                    "You may find it more effective to chose the Gradual Learning Algorithm or Maxent (on the " + Chr(34) + "Choose Ranking Algorithm" + Chr(34) + " menu), " + _
                    "which can deal with multiple winners." + _
                    Chr(10) + Chr(10) + _
                    "Click Yes to continue with ranking, No to return to the main OTSoft screen."
                Select Case MsgBox(MessageString, vbYesNo + vbExclamation)
                    Case vbYes
                            Let CheckChoiceOfAlgorithm = True
                            Exit Function
                    Case vbNo
                            Let CheckChoiceOfAlgorithm = False
                            Exit Function
                End Select
            End If              'Was there more than one winner?
        End If                  'Was a discrete algorithm chosen?
        
        Let CheckChoiceOfAlgorithm = True
        
End Function

Function CheckSubsetRelations() As Boolean

   'If any rival candidate has the same violations as the winner, then there
   '  can be no successful grammar.

    Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
   
    Dim LoserBoundsWinner As Boolean
    Dim WinnerBoundsLoser As Boolean
    Dim SameViolationsFlag As Boolean
    Dim FailureFlag As Boolean
    Let FailureFlag = False
     
    For FormIndex = 1 To mNumberOfForms
        For RivalIndex = 1 To mNumberOfRivals(FormIndex)
            
            Let LoserBoundsWinner = True
            Let WinnerBoundsLoser = True
            For ConstraintIndex = 1 To mNumberOfConstraints
                If mRivalViolations(FormIndex, RivalIndex, ConstraintIndex) _
                    > mWinnerViolations(FormIndex, ConstraintIndex) Then
                    Let LoserBoundsWinner = False
                    Exit For
                End If
            Next ConstraintIndex
            For ConstraintIndex = 1 To mNumberOfConstraints
                If mRivalViolations(FormIndex, RivalIndex, ConstraintIndex) _
                    < mWinnerViolations(FormIndex, ConstraintIndex) Then
                    Let WinnerBoundsLoser = False
                    Exit For
                End If
            Next ConstraintIndex
            
            If LoserBoundsWinner = True And WinnerBoundsLoser = True Then
                Let SameViolationsFlag = True
            Else
                Let SameViolationsFlag = False
            End If

            'If the flag is still up, then this rival has the same number of
            '  constraint violations as the winner.  Tell the user this and
            '  quit.

            If SameViolationsFlag = True Or LoserBoundsWinner = True Then
                'Do you want to announce problem, or say "moreover"?
                If FailureFlag = False Then
                    Print #mTmpFile,
                    Print #mTmpFile, "There's a problem with the constraint set."
                    Print #mTmpFile,
                    Print #mTmpFile, "F";
                    Print #mHTMFile, "<p>"
                    Print #mHTMFile, "There's a problem with the constraint set."
                    Print #mHTMFile, "<p>"
                    Print #mHTMFile, "F";
                Else
                        Print #mTmpFile,
                    Print #mTmpFile, "Moreover, f";
                    Print #mHTMFile, "<p>"
                    Print #mHTMFile, "Moreover, f";
                End If
               
                Print #mTmpFile, "or input #"; Trim(Str(FormIndex)); ", /"; DumbSym(mInputForm(FormIndex)); "/, losing candidate ";
                Print #mTmpFile, "["; DumbSym(mRival(FormIndex, RivalIndex)); "]";
                Print #mTmpFile,
                Print #mHTMFile, "or input #"; Trim(Str(FormIndex)); ", /"; DumbSym(mInputForm(FormIndex)); "/, losing candidate ";
                Print #mHTMFile, "["; DumbSym(mRival(FormIndex, RivalIndex)); "]";
               
                Select Case SameViolationsFlag
                    Case True
                        Print #mTmpFile, "has exactly the same violations as winning candidate ";
                        Print #mHTMFile, " has exactly the same violations as winning candidate ";
                    Case False
                        Print #mTmpFile, "harmonically bounds winning candidate ";
                        Print #mHTMFile, " harmonically bounds winning candidate ";
                End Select
                
                Print #mTmpFile, "["; DumbSym(mWinner(FormIndex)); "]."
                Print #mHTMFile, "["; DumbSym(mWinner(FormIndex)); "]."
                Print #mHTMFile, "<p>"
                              
                Let FailureFlag = True
               
            End If
        Next RivalIndex
    Next FormIndex
   
    If FailureFlag = True Then
        Call PrintPara(mDocFile, mTmpFile, mHTMFile, "")
        Call PrintPara(mDocFile, mTmpFile, mHTMFile, "Therefore, a grammar with just these constraints won't work.")
        Let CheckSubsetRelations = True
    Else
        Let CheckSubsetRelations = False
    End If


    'Report stringency relations.
        Dim TotalViolations() As Long
        ReDim TotalViolations(mNumberOfConstraints)
        Dim VacuousConstraintsPresent As Boolean
        Dim Stringency() As String
        ReDim Stringency(mNumberOfConstraints ^ 2, 3)
        Dim NumberOfStringencyRelations As Long
        Dim OuterIsSupersetOfInner As Boolean, InnerIsSupersetOfOuter As Boolean
        Dim OuterConstraintIndex As Long, InnerConstraintIndex As Long, EntryIndex As Long
        Dim StrinFile As Long
        Let StrinFile = FreeFile
        
        'Open an output file.
            Open gOutputFilePath + "StringencyRelationsAmongContraintsFor" + gFileName + ".txt" For Output As #StrinFile
        'Print a header.
            Print #StrinFile, "The following constraints show a stringency relationship -- not necessarily a logical one, but in the data given."
            Print #StrinFile,
        'To avoid vacuous listings, find any constrants that have no violations (sometimes found in machine-generated files).
            For ConstraintIndex = 1 To mNumberOfConstraints
                For FormIndex = 1 To mNumberOfForms
                    Let TotalViolations(ConstraintIndex) = TotalViolations(ConstraintIndex) + mWinnerViolations(FormIndex, ConstraintIndex)
                    For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                        Let TotalViolations(ConstraintIndex) = TotalViolations(ConstraintIndex) + mRivalViolations(FormIndex, RivalIndex, ConstraintIndex)
                    Next RivalIndex
                Next FormIndex
                If TotalViolations(ConstraintIndex) = 0 Then Let VacuousConstraintsPresent = True
            Next ConstraintIndex
        
        'Search for stringency relations among all possible constraint pairs.
            For OuterConstraintIndex = 1 To mNumberOfConstraints
                If TotalViolations(OuterConstraintIndex) > 0 Then
                    For InnerConstraintIndex = OuterConstraintIndex + 1 To mNumberOfConstraints
                        If TotalViolations(InnerConstraintIndex) > 0 Then
                            'Set the flags at their default values, which will be falsified whenever the constraints have unmatched violations.
                                Let OuterIsSupersetOfInner = True
                                Let InnerIsSupersetOfOuter = True
                            'Go through all the data.
                                For FormIndex = 1 To mNumberOfForms
                                    'First, look for unmatched violations in the winning candidate.
                                        Select Case mWinnerViolations(FormIndex, OuterConstraintIndex) - mWinnerViolations(FormIndex, InnerConstraintIndex)
                                            Case Is > 0
                                                Let InnerIsSupersetOfOuter = False
                                            Case Is < 0
                                                Let OuterIsSupersetOfInner = False
                                        End Select
                                    'Now, in the rival candidates.
                                        For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                                            Select Case mRivalViolations(FormIndex, RivalIndex, OuterConstraintIndex) - mRivalViolations(FormIndex, RivalIndex, InnerConstraintIndex)
                                                Case Is > 0
                                                    Let InnerIsSupersetOfOuter = False
                                                Case Is < 0
                                                    Let OuterIsSupersetOfInner = False
                                            End Select
                                            'We can speed execution by stopping early when it has been established that there is no superset relation in either direction.
                                                If InnerIsSupersetOfOuter = False And OuterIsSupersetOfInner = False Then GoTo ExitPoint
                                        Next RivalIndex
                                Next FormIndex
                            'If you reach this bit of code, you have a  positive finding, which must be recorded.
                                Let NumberOfStringencyRelations = NumberOfStringencyRelations + 1
                                'The constraint involved:
                                    Let Stringency(NumberOfStringencyRelations, 1) = mAbbrev(OuterConstraintIndex)
                                    Let Stringency(NumberOfStringencyRelations, 2) = mAbbrev(InnerConstraintIndex)
                                'The nature of the stringency relationship (either direction, or identity)
                                    If InnerIsSupersetOfOuter = True And OuterIsSupersetOfInner = True Then
                                        Let Stringency(NumberOfStringencyRelations, 3) = "same violations"
                                    ElseIf InnerIsSupersetOfOuter = False And OuterIsSupersetOfInner = True Then
                                        Let Stringency(NumberOfStringencyRelations, 3) = "strict superset"
                                    Else
                                        Let Stringency(NumberOfStringencyRelations, 3) = "strict subset"
                                    End If
ExitPoint:                  'The early-exit point if you found nothing.
                        End If                      'Don't bother with no-violation constraints.
                    Next InnerConstraintIndex
                End If                              'Don't bother with no-violation constraints.
            Next OuterConstraintIndex               'Go through all the possible constraint pairs.
        'Print what you learned.
            'Any never-violated constraints?
                If VacuousConstraintsPresent Then
                    Print #StrinFile, "The following constraints are never violated (and are ignored in the listing below)."
                    For ConstraintIndex = 1 To mNumberOfConstraints
                        If TotalViolations(ConstraintIndex) = 0 Then
                            Print #StrinFile, mAbbrev(ConstraintIndex)
                        End If
                    Next ConstraintIndex
                    Print #StrinFile,
                End If
            If NumberOfStringencyRelations = 0 Then
                If VacuousConstraintsPresent Then
                    Print #StrinFile, "(No such constraint pairs were found, other than vacuous ones.)"
                Else
                    Print #StrinFile, "(No such constraint pairs were found.)"
                End If
            Else
                For EntryIndex = 1 To NumberOfStringencyRelations
                    If Stringency(EntryIndex, 3) = "strict subset" Then
                        Print #StrinFile, Stringency(EntryIndex, 1); Chr(9); " has a strict subset of the violations of "; Chr(9); Stringency(EntryIndex, 2)
                    ElseIf Stringency(EntryIndex, 3) = "strict superset" Then
                        Print #StrinFile, Stringency(EntryIndex, 2); Chr(9); " has a strict subset of the violations of "; Chr(9); Stringency(EntryIndex, 1)
                    End If
                Next EntryIndex
                Print #StrinFile,
                For EntryIndex = 1 To NumberOfStringencyRelations
                    If Stringency(EntryIndex, 3) = "same violations" Then
                        Print #StrinFile, Stringency(EntryIndex, 1); Chr(9); " has exactly the same violations as  "; Chr(9); Stringency(EntryIndex, 2)
                    End If
                Next EntryIndex
            End If
            Close #StrinFile


End Function

Function FaithfulnessConstraint(MyConstraint As String) As Boolean

    'Default is false.
        Let FaithfulnessConstraint = False
    'Initial strings of length 3:
        Select Case LCase(Left(MyConstraint, 3))
            Case "ide", "fai", "id(", "max", "dep", "map", "*ma"
                Let FaithfulnessConstraint = True
        End Select
    'Prince and Tesar's "F:", length 2.  Must match case.
        Select Case Left(MyConstraint, 2)
            Case "F:"
                Let FaithfulnessConstraint = True
        End Select
    
End Function



Sub FindUnnecessaryConstraints(ConstraintName() As String, Abbrev() As String, _
    WinnerViolations() As Long, NumberOfRivals() As Long, RivalViolations() As Long)
        
    'Go through the constraint set, and see if the correct outcomes are obtained even when a constraint is deleted.

    'After doing this, then try the whole set of "deletable" constraints at once, and see if they can go.

    'This version of this routine was completely written by BH in March 2008, in response to, gack, the bug
    '   reported on ROA by Alan Prince.
    '   The strategy is now totally straightforward:  do everything from scratch.
    '   The old version tried to save time, with fatal errors.  I doubt this routine takes much time, since
    '   Constraint Demotion is so fast.
    
        Dim TrulyNeeded() As Boolean
        ReDim TrulyNeeded(mNumberOfConstraints)
        Dim ViolatedInWinner() As Boolean
        ReDim ViolatedInWinner(mNumberOfConstraints)
        
        Dim NumberOfDeletableConstraints As Long
        Dim MassDeletionIsPossible As Boolean

        Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
        Dim InnerConstraintIndex As Long, OuterConstraintIndex As Long
        Dim StratumIndex As Long
        Dim SpaceIndex As Long
        
        Dim ConstraintUnderAssessment As Long
        Dim DifferentFlag As Boolean
        
        Dim LocalWinnerViolations() As Long
            ReDim LocalWinnerViolations(mNumberOfForms, mNumberOfConstraints)
        Dim LocalRivalViolations() As Long
            ReDim LocalRivalViolations(mNumberOfForms, mMaximumNumberOfRivals, mNumberOfConstraints)
            
        'For printing:
            Dim LengthOfLongestAbbreviation As Long
        
    For ConstraintUnderAssessment = 1 To mNumberOfConstraints

        'Construct a basic set of ranking data, altering the violations of the target constraint to zero
        '   and leaving the others unaltered.
            For ConstraintIndex = 1 To mNumberOfConstraints
                If ConstraintIndex = ConstraintUnderAssessment Then
                    For FormIndex = 1 To mNumberOfForms
                        Let LocalWinnerViolations(FormIndex, ConstraintIndex) = 0
                        For RivalIndex = 1 To NumberOfRivals(FormIndex)
                            Let LocalRivalViolations(FormIndex, RivalIndex, ConstraintIndex) = 0
                        Next RivalIndex
                    Next FormIndex
                Else
                    For FormIndex = 1 To mNumberOfForms
                        Let LocalWinnerViolations(FormIndex, ConstraintIndex) = WinnerViolations(FormIndex, ConstraintIndex)
                        For RivalIndex = 1 To NumberOfRivals(FormIndex)
                            Let LocalRivalViolations(FormIndex, RivalIndex, ConstraintIndex) = RivalViolations(FormIndex, RivalIndex, ConstraintIndex)
                        Next RivalIndex
                    Next FormIndex
                End If
            Next ConstraintIndex
            
        'Check if, under the changed violation set, any winner has the same violations as a rival,
        '   in which case, the constraint is plainly necessary.
            For FormIndex = 1 To mNumberOfForms
                For RivalIndex = 1 To NumberOfRivals(FormIndex)
                    For ConstraintIndex = 1 To mNumberOfConstraints
                        If LocalWinnerViolations(FormIndex, ConstraintIndex) <> LocalRivalViolations(FormIndex, RivalIndex, ConstraintIndex) Then
                            'This rival is ok, cease to consider it.
                                GoTo RivalEscapePoint
                        End If
                    Next ConstraintIndex
                    'Uh-oh - the winner has the same number of violations for every constraint that a rival has,
                    '   when the current constraint is zeroed out.  So the grammar won't work if the constraint is
                    '   omitted.
                        Let TrulyNeeded(ConstraintUnderAssessment) = True
                        GoTo ConstraintEscapePoint
RivalEscapePoint:
                Next RivalIndex
            Next FormIndex
        
        'Ok, so there are no winner-rival duplications for this constraint.
        '   Now we need to know if the other constraints could do the job on their own.
            Select Case mnuConstrainAlgorithmsByAPrioriRankings.Checked
                Case False
                    If FastRCD(mNumberOfForms, NumberOfRivals(), _
                        LocalWinnerViolations(), LocalRivalViolations()) = True Then
                            'Algorithm converges even when the violations of the constraint under
                            '   assessment are set to zero, so the constraint is not needed.
                            Let TrulyNeeded(ConstraintUnderAssessment) = False
                    Else
                            Let TrulyNeeded(ConstraintUnderAssessment) = True
                    End If
                Case True
                    If FastRCDWithAPrioriRankings(mNumberOfForms, NumberOfRivals(), _
                        LocalWinnerViolations(), LocalRivalViolations()) = True Then
                            Let TrulyNeeded(OuterConstraintIndex) = False
                    Else
                            Let TrulyNeeded(OuterConstraintIndex) = True
                    End If
            End Select
        
ConstraintEscapePoint:                      'Line label, for constraints shown early to be unnecessary.
    Next ConstraintUnderAssessment          'Assess each constraint for necessitude.
        
        'Now count how many constraints are deletable.
            For ConstraintIndex = 1 To mNumberOfConstraints
                If TrulyNeeded(ConstraintIndex) = False Then
                    Let NumberOfDeletableConstraints = NumberOfDeletableConstraints + 1
                End If
            Next ConstraintIndex
        
        'If you found more than one removable constraint, see if they can all be removed.
        
            'Don't bother if there was just one or none.
                If NumberOfDeletableConstraints < 2 Then GoTo EscapePointForAll
                
            'Default for mass-removability is yes, which then gets disproven if appropriate.
                Let MassDeletionIsPossible = True
            
            'To test this, zero out the violations of all the individually-unneeded constraints.
                For ConstraintIndex = 1 To mNumberOfConstraints
                    If TrulyNeeded(ConstraintIndex) = False Then
                        For FormIndex = 1 To mNumberOfForms
                            Let LocalWinnerViolations(FormIndex, ConstraintIndex) = 0
                            For RivalIndex = 1 To NumberOfRivals(FormIndex)
                                Let LocalRivalViolations(FormIndex, RivalIndex, ConstraintIndex) = 0
                            Next RivalIndex
                        Next FormIndex
                    Else
                        For FormIndex = 1 To mNumberOfForms
                            Let LocalWinnerViolations(FormIndex, ConstraintIndex) = WinnerViolations(FormIndex, ConstraintIndex)
                            For RivalIndex = 1 To NumberOfRivals(FormIndex)
                                Let LocalRivalViolations(FormIndex, RivalIndex, ConstraintIndex) = RivalViolations(FormIndex, RivalIndex, ConstraintIndex)
                            Next RivalIndex
                        Next FormIndex
                    End If
                Next ConstraintIndex
            
            'Check if, under the changed violation set, any winner has the same violations as a rival,
            '   in which case, we cannot eliminate the entire batch of constraints.
            '   Note that this was checked earlier for the full constraint set.
                For FormIndex = 1 To mNumberOfForms
                    For RivalIndex = 1 To NumberOfRivals(FormIndex)
                        For ConstraintIndex = 1 To mNumberOfConstraints
                            If LocalWinnerViolations(FormIndex, ConstraintIndex) <> LocalRivalViolations(FormIndex, RivalIndex, ConstraintIndex) Then
                                'This rival is ok, cease to consider it.
                                    GoTo RivalEscapePointForAll
                            End If
                        Next ConstraintIndex
                        'The winner has been shown to have the same number of violations for every constraint
                        '   that a rival has, when the current constraint list is zeroed out.
                        '   So the grammar won't work if the constraint is omitted.
                            Let MassDeletionIsPossible = False
                            GoTo EscapePointForAll
RivalEscapePointForAll:
                    Next RivalIndex
                Next FormIndex
            'Ok, so there are no winner-rival duplications when the set of unneeded constraints is zeroed out.
            '   Now we need to know if the unzeroed constraints could do the job on their own.
            '   Mass deletion is possible when the algorithm converged, so we just copy over the Boolean output.
                Select Case mnuConstrainAlgorithmsByAPrioriRankings.Checked
                    Case False
                        Let MassDeletionIsPossible = FastRCD(mNumberOfForms, NumberOfRivals(), _
                            LocalWinnerViolations(), LocalRivalViolations())
                    Case True
                        Let MassDeletionIsPossible = FastRCDWithAPrioriRankings(mNumberOfForms, NumberOfRivals(), _
                            LocalWinnerViolations(), LocalRivalViolations())
                End Select
        
EscapePointForAll:      'Go here if you found that mass deletion is impossible or irrelevant.
        
    'Check which faithfulness constraints are violated in winners; since we want to advise users to retain them.
        For ConstraintIndex = 1 To mNumberOfConstraints
            Let ViolatedInWinner(ConstraintIndex) = False
            Let mFaithfulness(ConstraintIndex) = Form1.FaithfulnessConstraint(mConstraintName(ConstraintIndex))
            If mFaithfulness(ConstraintIndex) = True Then
               For FormIndex = 1 To mNumberOfForms
                    If WinnerViolations(FormIndex, ConstraintIndex) > 0 Then
                        Let ViolatedInWinner(ConstraintIndex) = True
                        Exit For
                    End If
               Next FormIndex
            End If
        Next ConstraintIndex
    
    'PRINTING OUT THE RESULT
    
    
    'Header.
        Call PrintLevel1Header(mDocFile, mTmpFile, mHTMFile, "Status of Proposed Constraints:  Necessary or Unnecessary")
        
    'Find length of longest constraint abbreviation, for producing the plain-text output.
        For ConstraintIndex = 1 To mNumberOfConstraints
            If Len(mAbbrev(ConstraintIndex)) > LengthOfLongestAbbreviation Then
                Let LengthOfLongestAbbreviation = Len(mAbbrev(ConstraintIndex))
            End If
        Next ConstraintIndex
    
    'Start a table in pretty and htm versions.
        Print #mDocFile, "\ts2"
        Print #mDocFile, "\ntConstraint"; Chr$(9); "Status"
        Dim MyHTMLTable() As String
        ReDim MyHTMLTable(2, mNumberOfConstraints + 1)
        Dim HTMRowCount As String
        Let HTMRowCount = 1
        Let MyHTMLTable(1, 1) = "Constraint"
        Let MyHTMLTable(2, 1) = "Status"
    
    'The needed constraints:
        For ConstraintIndex = 1 To mNumberOfConstraints
           If TrulyNeeded(ConstraintIndex) = True Then
                 Print #mDocFile, SmallCapTag1; ConstraintName(ConstraintIndex); SmallCapTag2; Chr$(9);
                 Print #mTmpFile, "   "; Abbrev(ConstraintIndex);
                 For SpaceIndex = Len(Abbrev(ConstraintIndex)) To LengthOfLongestAbbreviation + 1
                    Print #mTmpFile, " ";
                 Next SpaceIndex
                 Print #mTmpFile, "Necessary"
                 Print #mDocFile, "Necessary"
                 Let HTMRowCount = HTMRowCount + 1
                 Let MyHTMLTable(1, HTMRowCount) = ConstraintName(ConstraintIndex)
                 Let MyHTMLTable(2, HTMRowCount) = "Necessary"
           End If
        Next ConstraintIndex
    
    'The unneeded constraints, justifiable as an illustration of Faithfulness constraints violated by winners.
        For ConstraintIndex = 1 To mNumberOfConstraints
            If TrulyNeeded(ConstraintIndex) = False Then
                If ViolatedInWinner(ConstraintIndex) = True Then
                    Print #mDocFile, SmallCapTag1; ConstraintName(ConstraintIndex); SmallCapTag2; Chr$(9);
                    Print #mTmpFile, "   "; Abbrev(ConstraintIndex);
                    For SpaceIndex = Len(Abbrev(ConstraintIndex)) To LengthOfLongestAbbreviation + 1
                       Print #mTmpFile, " ";
                    Next SpaceIndex
                    Print #mDocFile, "Not necessary (but included to show Faithfulness violations of a winning candidate)"
                    Print #mTmpFile, "Not necessary (but included to show Faithfulness violations"
                    For SpaceIndex = 1 To LengthOfLongestAbbreviation + 8
                       Print #mTmpFile, " ";
                    Next SpaceIndex
                    Print #mTmpFile, "of a winning candidate)"
                    Let HTMRowCount = HTMRowCount + 1
                    Let MyHTMLTable(1, HTMRowCount) = ConstraintName(ConstraintIndex)
                    Let MyHTMLTable(2, HTMRowCount) = "Not necessary (but included to show Faithfulness violations of a winning candidate)"
                End If
            End If
        Next ConstraintIndex
    
    'The utterly useless constraints.
        For ConstraintIndex = 1 To mNumberOfConstraints
           If TrulyNeeded(ConstraintIndex) = False Then
                If ViolatedInWinner(ConstraintIndex) = False Then
                    Print #mDocFile, SmallCapTag1; ConstraintName(ConstraintIndex); SmallCapTag2; Chr$(9);
                    Print #mTmpFile, "   "; Abbrev(ConstraintIndex);
                    For SpaceIndex = Len(Abbrev(ConstraintIndex)) To LengthOfLongestAbbreviation + 1
                       Print #mTmpFile, " ";
                    Next SpaceIndex
                    Print #mDocFile, "Not necessary"
                    Print #mTmpFile, "Not necessary"
                    Let HTMRowCount = HTMRowCount + 1
                    Let MyHTMLTable(1, HTMRowCount) = ConstraintName(ConstraintIndex)
                    Let MyHTMLTable(2, HTMRowCount) = "Not necessary"
                End If
           End If
        Next ConstraintIndex
        Print #mTmpFile,
        Print #mDocFile, "\te"
        
        Call s.PrintHTMTable(MyHTMLTable(), mHTMFile, True, False, False)
    
    'If relevant, report whether mass deletion of unneeded constraints is possible.
        If NumberOfDeletableConstraints >= 2 Then
            If MassDeletionIsPossible = True Then
                Call PrintPara(mDocFile, mTmpFile, mHTMFile, "A check has determined that the grammar will still work even if the " + _
                    "PARAconstraints marked above as unnecessary are removed en masse.")
            Else
                Print #mTmpFile,
                Call PrintPara(mDocFile, mTmpFile, mHTMFile, "A check has determined that, although the grammar will still work with the" + _
                    "PARAremoval of ANY ONE of the constraints marked above as unnecessary, the" + _
                    "PARAgrammar will NOT work if they are removed en masse.")
           End If
        End If                  'Only worth doing if there are at least two removable constraints.
       
End Sub


Public Function FastRCD(ByVal NumberOfForms As Long, NumberOfRivals() As Long, _
    WinnerViolations() As Long, RivalViolations() As Long) As Boolean

   'Dimension the local variables.
   
      Dim Stratum() As Long
      ReDim Stratum(mNumberOfConstraints)
      
      Dim CurrentStratum As Long
   
      Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
   
      Dim SomeAreNonDemotible As Boolean
      Dim SomeAreDemotible As Boolean
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
                         If WinnerViolations(FormIndex, ConstraintIndex) > mRivalViolations(FormIndex, RivalIndex, ConstraintIndex) Then
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
   
      Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
      Dim InnerConstraintIndex As Long, OuterConstraintIndex As Long
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

Function LookForMinimalPairs() As Boolean

   'Look at the basic violation patterns, seeking the classical tableau configuration:

   '               Winner Rival
   '                ...    ...
   '                 ok     *
   '                ...    ...
   '                 *     ok
   '     where ...    ...  are identical

   '  This kind of setup which tells you that a ranking is crucial.
   '  Note that such an arrangement will be more likely to arise if the user has been
   '     careful about including all the relevant candidates.

    Dim MinimalPairEvidence() As Long
    ReDim MinimalPairEvidence(mNumberOfConstraints, mNumberOfConstraints, 5, 2)
    Dim NumberOfRankingArguments() As Long
    ReDim NumberOfRankingArguments(mNumberOfConstraints, mNumberOfConstraints)
    
    Dim Flag As Boolean
    Dim FoundAContradictionFlag As Boolean
    Let FoundAContradictionFlag = False
    Dim MinimalPairsSucceededFlag As Long
    
    Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
    Dim HigherConstraintIndex As Long, LowerConstraintIndex As Long
    Dim OuterConstraintIndex As Long, InnerConstraintIndex As Long
    Dim LocalFormIndex As Long, LocalRivalIndex As Long
    Dim RankingArgumentIndex As Long
   
    'Fake strata for making tiny tableaux.
        Dim FakeStrata() As Long
        ReDim FakeStrata(mNumberOfConstraints)
   
   'A form is entered into a crucial data structure, described below, if the
   '  Rival violates the dominating constraint, the Winner violates the
   '  dominated constraint, and their violations are otherwise *identical*.

   'Outer loops:  go through all constraint pairs.

   For HigherConstraintIndex = 1 To mNumberOfConstraints
      For LowerConstraintIndex = 1 To mNumberOfConstraints
         'Don't bother where these are the very same constraint:
         If HigherConstraintIndex <> LowerConstraintIndex Then
            'Then, go through the forms, looking for Winner/Rival pairs that
            '  provide an unambiguous ranking.
            For FormIndex = 1 To mNumberOfForms
               For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                   
                   'Is there a ranking argument here?:
                   If mRivalViolations(FormIndex, RivalIndex, HigherConstraintIndex) > mWinnerViolations(FormIndex, HigherConstraintIndex) Then
                      If mRivalViolations(FormIndex, RivalIndex, LowerConstraintIndex) < mWinnerViolations(FormIndex, LowerConstraintIndex) Then
                         
                         'So far, there could be.  But the pair must be
                         '  truly minimal.  Check that the Winner and Rival
                         '  don't differ on any of the other constraints.
                            Let Flag = True
                            For ConstraintIndex = 1 To mNumberOfConstraints
                               'Check that this is one of the constraints
                               '  that could spoil the minimal pair.
                                  If ConstraintIndex <> HigherConstraintIndex And ConstraintIndex <> LowerConstraintIndex Then
                                     If mWinnerViolations(FormIndex, ConstraintIndex) <> mRivalViolations(FormIndex, RivalIndex, ConstraintIndex) Then
                                        Let Flag = False
                                        Exit For
                                     End If
                                  End If
                            Next ConstraintIndex
                            
                            'Now, if you have found a true minimal pair, record the evidence for it.
                            '  The evidence is recorded in an array whose
                            '     dimensions correspond to:
                            '         1. the dominating constraint
                            '         2. the dominated constraint
                            '         3. an index:  "this is the nth case
                            '               I've found to show this."
                            '         4. an index:  1 = input 2 = rival

                            If Flag = True Then
                                 
                               'To avoid an unbounded array, the program
                               '  is willing to store only five ranking
                               '  arguments per ranking, so keep count.
                                 Let NumberOfRankingArguments(HigherConstraintIndex, LowerConstraintIndex) = NumberOfRankingArguments(HigherConstraintIndex, LowerConstraintIndex) + 1
                               'If there is room to store this ranking argument, store it.
                                  If NumberOfRankingArguments(HigherConstraintIndex, LowerConstraintIndex) < 5 Then
                                     Let MinimalPairEvidence(HigherConstraintIndex, LowerConstraintIndex, NumberOfRankingArguments(HigherConstraintIndex, LowerConstraintIndex), 1) = FormIndex
                                     Let MinimalPairEvidence(HigherConstraintIndex, LowerConstraintIndex, NumberOfRankingArguments(HigherConstraintIndex, LowerConstraintIndex), 2) = RivalIndex
                                  End If
                            End If  'This IF detected the true minimal pair.
                      End If
                   End If       'These two IF's detected the crucial ranking configuration.
               Next RivalIndex
            Next FormIndex      'Go through all the data.
         End If                 'Don't test a constraint against itself.
      Next LowerConstraintIndex
   Next HigherConstraintIndex   'Go through all constraint pairs.
   
   'Now, you have the evidence in your hands.  Comb through it for *contradictions*, and print them out.

      For OuterConstraintIndex = 1 To mNumberOfConstraints - 1
         For InnerConstraintIndex = OuterConstraintIndex + 1 To mNumberOfConstraints
            'Check this assignment for contradictions.
              If MinimalPairEvidence(OuterConstraintIndex, InnerConstraintIndex, 1, 1) > 0 Then
                 If MinimalPairEvidence(InnerConstraintIndex, OuterConstraintIndex, 1, 1) > 0 Then
                    'There is a contradiction.  Print out the bad news, the first time around.
                       If FoundAContradictionFlag = False Then
                            Call PrintLevel1Header(mDocFile, mTmpFile, mHTMFile, "Contradiction Located")
                            Call PrintPara(mDocFile, mTmpFile, mHTMFile, "The problem can be localized in the form of one or more minimal pairs that cannot be consistently ranked.")
                       End If
                       Let FoundAContradictionFlag = True

                       Call PrintPara(mDocFile, mTmpFile, mHTMFile, "The following is a contradiction:")
                       
                       Print #mTmpFile,
                       Print #mTmpFile, "The ranking "; mConstraintName(OuterConstraintIndex); " >> "; mConstraintName(InnerConstraintIndex); " is needed, because ";
                       Print #mDocFile,
                       Print #mDocFile, "   The ranking "; mConstraintName(OuterConstraintIndex); " >> "; mConstraintName(InnerConstraintIndex); " is needed:"
                       Print #mHTMFile, "<p>The ranking "; mConstraintName(OuterConstraintIndex); " >> "; mConstraintName(InnerConstraintIndex); " is needed, because ";
                       For RankingArgumentIndex = 1 To NumberOfRankingArguments(OuterConstraintIndex, InnerConstraintIndex)
                            'I set a limit of 5; don't know why it isn't being respected ...
                                If RankingArgumentIndex <= 5 Then
                                    Let LocalFormIndex = MinimalPairEvidence(OuterConstraintIndex, InnerConstraintIndex, RankingArgumentIndex, 1)
                                    Let LocalRivalIndex = MinimalPairEvidence(OuterConstraintIndex, InnerConstraintIndex, RankingArgumentIndex, 2)
                                    Call PrintContradictoryDerivation(SymbolTag1, SymbolTag2, mInputForm(), LocalFormIndex, mWinner(), mRival(), LocalRivalIndex)
                                End If
                          'Provide a minitableau to show this:
                              Let FakeStrata(OuterConstraintIndex) = 1
                              Let FakeStrata(InnerConstraintIndex) = 2
                              Call PrintMiniTableaux.Main(mNumberOfConstraints, mAbbrev(), LocalFormIndex, LocalRivalIndex, 2, FakeStrata(), mInputForm(), mWinner(), mWinnerViolations(), mRival(), mRivalViolations(), _
                                mTmpFile, mDocFile, mHTMFile)
                       Next RankingArgumentIndex
                       
                       Print #mTmpFile,
                       Print #mTmpFile, "The ranking "; mConstraintName(InnerConstraintIndex); " >> "; mConstraintName(OuterConstraintIndex); " is needed, because ";
                       Print #mDocFile,
                       Print #mDocFile, "   The ranking "; mConstraintName(InnerConstraintIndex); " >> "; mConstraintName(OuterConstraintIndex); " is needed:"
                       Print #mHTMFile, "<p>The ranking "; mConstraintName(InnerConstraintIndex); " >> "; mConstraintName(OuterConstraintIndex); " is needed, because ";
                       For RankingArgumentIndex = 1 To NumberOfRankingArguments(InnerConstraintIndex, OuterConstraintIndex)
                          Let LocalFormIndex = MinimalPairEvidence(InnerConstraintIndex, OuterConstraintIndex, RankingArgumentIndex, 1)
                          Let LocalRivalIndex = MinimalPairEvidence(InnerConstraintIndex, OuterConstraintIndex, RankingArgumentIndex, 2)
                          Call PrintContradictoryDerivation(SymbolTag1, SymbolTag2, mInputForm(), LocalFormIndex, _
                            mWinner(), mRival(), LocalRivalIndex)
                          'Provide a minitableau to show this:
                              Let FakeStrata(InnerConstraintIndex) = 1
                              Let FakeStrata(OuterConstraintIndex) = 2
                              Call PrintMiniTableaux.Main(mNumberOfConstraints, mAbbrev(), LocalFormIndex, LocalRivalIndex, 2, FakeStrata(), mInputForm(), mWinner(), mWinnerViolations(), mRival(), mRivalViolations(), _
                                  mTmpFile, mDocFile, mHTMFile)
                       Next RankingArgumentIndex

                    'Also, note that this search succeeded, so you won't have to do a less effective search.
                       Let LookForMinimalPairs = True

                 End If
              End If
         Next InnerConstraintIndex
      Next OuterConstraintIndex

End Function

Sub PrintContradictoryDerivation(SymbolTag1 As String, SymbolTag2 As String, InputForm() As String, _
    LocalFormIndex As Long, Winner() As String, Rival() As String, LocalRivalIndex As Long)

    Print #mDocFile, "     /"; SymbolTag1;
    Print #mDocFile, InputForm(LocalFormIndex);
    Print #mDocFile, SymbolTag2; "/  "; Chr$(9); "\ys"; Chr(174); "\ye  ["; SymbolTag1;
    Print #mDocFile, Winner(LocalFormIndex);
    Print #mDocFile, SymbolTag2; "], not *["; SymbolTag1;
    Print #mDocFile, Rival(LocalFormIndex, LocalRivalIndex);
    Print #mDocFile, SymbolTag2; "]"
    Print #mTmpFile, "/";
    Print #mTmpFile, DumbSym(InputForm(LocalFormIndex));
    Print #mTmpFile, "/  -->  [";
    Print #mTmpFile, DumbSym(Winner(LocalFormIndex));
    Print #mTmpFile, "], not *[";
    Print #mTmpFile, DumbSym(Rival(LocalFormIndex, LocalRivalIndex));
    Print #mTmpFile, "]"

    Print #mHTMFile, "/";
    Print #mHTMFile, DumbSym(InputForm(LocalFormIndex));
    Print #mHTMFile, "/  -->  [";
    Print #mHTMFile, DumbSym(Winner(LocalFormIndex));
    Print #mHTMFile, "], not *[";
    Print #mHTMFile, DumbSym(Rival(LocalFormIndex, LocalRivalIndex));
    Print #mHTMFile, "]<p>"

End Sub

Sub RunATTDot()
        
    'Call the ATT Dot program, in order to make a Hasse diagram.
        
    'On Error GoTo CheckError
        
    'Make sure that a copy of Dot.exe exists on this computer.
    '   If not, there will be no gFileName.gif file, and an error will be given
    '   to the user when (s)he asks to see the Hasse diagram.
    '   If the user never asks, then it won't matter.
        If Dir(gDotExeLocation) = "" Then
            Exit Sub
        End If
        
    'Make sure that the necessary Hasse.txt file exists.  Else leave.  The problem gets
    '   reported elsewhere, so as not to perturb the dot.exe-less user.
        If Dir(gOutputFilePath + gFileName + "Hasse.txt") = "" Then
            Exit Sub
        End If
        
    'Report what you're doing, in whatever window the user can see.
        Select Case gAlgorithmName
            Case "GLA"
                'Do nothing; let's try putting it in GLA.
            Case Else
                Let lblProgressWindow.Caption = "Creating Hasse diagram..."
                DoEvents
        End Select
    
    'We need to know when dot.exe is done.  Do this by deleting the old hasse file,
    '   then looping until the new one is done.
    '   And, moreover, we need to know when deletion is complete.
        If Dir(gOutputFilePath + gFileName + "hasse.gif") <> "" Then
            Kill (gOutputFilePath + gFileName + "hasse.gif")
        End If
        'If Dir("c:\UnnecessaryOTSoftFileDeleteMe.gif") <> "" Then
        '    Kill ("c:\UnnecessaryOTSoftFileDeleteMe.gif")
        'End If
        Dim MyTimer As Long
        Let MyTimer = Timer
        Do
            If Dir(gOutputFilePath + gFileName + "hasse.gif") = "" Then Exit Do
            If Timer - MyTimer > 30 Then
                MsgBox "Sorry, I can't plot a Hasse diagram.", vbExclamation
                'Let it be known that you have not created a Hasse diagram.
                    Let HasseDiagramCreated = False
                Exit Sub
            End If
            DoEvents
        Loop
    
    'Use Dot.exe to make the Hasse file.
        Dim Dummy As Long    'Needed to assign a value of the Shell() function.
        'Note:  Shell *must* have chr(34), the quotation mark, around all filename
        '   strings--else if they have spaces, the command will fail.
            'Experiment also output a Postscript file, for better resolution:
            Let Dummy = Shell(gDotExeLocation + " dot -Tps " + _
                Chr(34) + gOutputFilePath + gFileName + "Hasse.txt" + Chr(34) + _
                " -o " + _
                Chr(34) + gOutputFilePath + gFileName + "Hasse.ps" + Chr(34))
        
            Let Dummy = Shell(gDotExeLocation + " dot -Tgif " + _
                Chr(34) + gOutputFilePath + gFileName + "Hasse.txt" + Chr(34) + _
                " -o " + _
                Chr(34) + gOutputFilePath + gFileName + "Hasse.gif" + Chr(34))
        
    'Copy the Hasse graphics file to the root for purposes of pretty display later on.
    
    'Check for completion--you can't copy the file until you have the original.
        
        Let MyTimer = Timer
        Do
            'Does the (copied version of the) Hasse .gif file exist?
            If Dir(gOutputFilePath + gFileName + "Hasse.gif") <> "" Then
                'Note:  dot.exe first creates the file, then fills it.
                If FileLen(gOutputFilePath + gFileName + "Hasse.gif") > 0 Then
                    'You've got a legitimate .gif file, so now you can copy it to the
                    '   root for purposes of a pretty output file.
                       ' FileCopy gOutputFilePath + gFileName + "Hasse.gif", "c:\UnnecessaryOTSoftFileDeleteMe.gif"
                        Exit Do
                Else
                    'Don't wait forever for the basic file to appear.
                    If Timer - MyTimer > 10 Then
                        MsgBox "Sorry, I can't plot a Hasse diagram.", vbExclamation
                        'Let it be known that you have not created a Hasse diagram.
                            Let HasseDiagramCreated = False
                            Exit Sub
                    End If
                End If
            Else
                'Don't wait forever for the basic file to be long enough.
                    If Timer - MyTimer > 10 Then
                        MsgBox "Sorry, I can't plot a Hasse diagram.", vbExclamation
                        'Let it be known that you have not created a Hasse diagram.
                            Let HasseDiagramCreated = False
                            Exit Sub
                    End If
            End If
            DoEvents
        Loop
    
    'Let it be known that you have created a Hasse diagram.
        Let HasseDiagramCreated = True
    'And that it is unnecessary to replot, in case some editing happened earlier.
        Let ReplotFirst = False
            
    Exit Sub
            
CheckError:
    MsgBox "Program error:  trouble in making a Hasse diagram.   I would appreciate your letting me know the about the problem.  Email me at bhayes@humnet.ucla.edu, specifying error #85182, and including a copy of your input file." + _
        "Click ok and the program will continue."

End Sub

Sub InsertHasseDiagramIntoOutputFile(DocFile As Long, HTMFile As Long)

    'If there is a graphic file with the Hasse diagram, insert a diacritic that will
    '   cause the Word macro to insert it.  This cannot be done with the temp file.
        
        If Dir("c:\UnnecessaryOTSoftFileDeleteMe.gif") <> "" Then
            
            Print #mDocFile, "\ks"
            Call PrintLevel1Header(DocFile, -1, HTMFile, "Hasse Diagram")
            Call PrintPara(DocFile, -1, HTMFile, "The following Hasse diagram (in output folder at " + gFileName + "Hasse.gif) summarizes the rankings obtained.")
            Print #DocFile,
            Print #DocFile, "\hf"
            Print #DocFile, "\ke"
            'Print #HTMFile, "<img src=" + Chr(34) + "c:\UnnecessaryOTSoftFileDeleteMe.gif" + Chr(34) + " alt=" + Chr(34) + "Hasse diagram for " + Chr(34) + " + gfilename" + Chr(34) + " width=" + Chr(34) + "160" + Chr(34) + " height=" + Chr(34) + "120" + Chr(34) + " hspace=" + Chr(34) + "10" + Chr(34) + " vspace=" + Chr(34) + "10" + Chr(34) + " align=" + Chr(34) + "left" + Chr(34) + " border=" + Chr(34) + "0" + Chr(34) + " />"
            Print #HTMFile, "<p><p><img src=" + Chr(34) + gFileName + "Hasse.gif" + Chr(34) + " border=" + Chr(34) + "1" + Chr(34) + ">"
        End If
        
        
End Sub
           

Sub PrintOutTheAprioriRankings(TmpFile As Long, DocFile As Long, HTMFile As Long)

    Dim Table() As String
    Dim DocTable() As String
    Dim ConstraintIndex As Long, InnerConstraintIndex As Long
         
    Call PrintLevel1Header(DocFile, TmpFile, HTMFile, "A Priori Rankings")
    
    'Roman numerals in header for factorial typology, else Arabic:
    '    Select Case RunningFactorialTypology
    '        Case True
    '            Print #TmpFile, RomanNumeral(gLevel1HeadingNumber); ". A Priori Rankings"
    '        Case False
    '            Print #TmpFile, Trim(gLevel1HeadingNumber); ". A Priori Rankings"
    '    End Select
    Call PrintPara(DocFile, TmpFile, HTMFile, "In the following table, " + Chr(34) + "yes" + Chr(34) + " means that the constraint of the indicated row PARAwas marked a priori to dominate the constraint in the given column.")
    
    'Print the a priori rankings.
    
        ReDim Table(mNumberOfConstraints + 1, mNumberOfConstraints + 1)
        ReDim DocTable(mNumberOfConstraints + 1, mNumberOfConstraints + 1)
        For ConstraintIndex = 1 To mNumberOfConstraints
            Let Table(1, ConstraintIndex + 1) = mAbbrev(ConstraintIndex)
            Let Table(ConstraintIndex + 1, 1) = mAbbrev(ConstraintIndex)
            Let DocTable(1, ConstraintIndex + 1) = SmallCapTag1 + mAbbrev(ConstraintIndex) + SmallCapTag2
            Let DocTable(ConstraintIndex + 1, 1) = SmallCapTag1 + mAbbrev(ConstraintIndex) + SmallCapTag2
            For InnerConstraintIndex = 1 To mNumberOfConstraints
                If gAPrioriRankingsTable(ConstraintIndex, InnerConstraintIndex) = True Then
                    'Switch axes -- we needed to do this kludgy business because only the last index is redimensionable.
                    Let Table(InnerConstraintIndex + 1, ConstraintIndex + 1) = "yes"
                    Let DocTable(InnerConstraintIndex + 1, ConstraintIndex + 1) = "yes"
                End If
            Next InnerConstraintIndex
        Next ConstraintIndex
        
        Call s.PrintDocTable(DocTable(), DocFile, False, False, True)
        Call s.PrintTable(-1, TmpFile, HTMFile, Table(), False, False, True)

    'Report just how a priori ranking was implemented numerically for the GLA.
        If gAlgorithmName = "GLA" Then
            Print #TmpFile,
            Call PrintPara(DocFile, TmpFile, HTMFile, "  An a priori ranking was implemented as a minimal difference PARA  in ranking values of " + Trim(GLA.txtValueThatImplementsAPrioriRankings.Text) + ".")
        End If


        'Old code.  Maybe people will like it better this way, so keep.
             
            'Dim SpacesNeeded As Long, Dim SpaceIndex As Long
            'Figure out how wide the first column has to be.
                 'Let SpacesNeeded = 0
                 'For ConstraintIndex = 1 To mNumberOfConstraints
                 '    If Len(mAbbrev(ConstraintIndex)) + Len(Str(ConstraintIndex)) > SpacesNeeded Then
                 '        Let SpacesNeeded = Len(mAbbrev(ConstraintIndex)) + Len(Str(ConstraintIndex))
                 '    End If
                 'Next ConstraintIndex
             
             'Print #DocFile, "\ts2\ks"
            '
            ' Dim Table() As String
            ' ReDim Table(2, mNumberOfConstraints + 1)
            '
            ' Print #DocFile, "Constraint"; Chr(9); "Dominees"
            ' Let Table(1, 1) = "Constraint"
            ' Let Table(2, 1) = "Dominees"
            '
            ' For ConstraintIndex = 1 To mNumberOfConstraints
            '     Print #TmpFile, "  "; Trim(Str(ConstraintIndex)); ". "; mAbbrev(ConstraintIndex);
            '     Print #DocFile, Trim(Str(ConstraintIndex)); ". "; SmallCapTag1; mAbbrev(ConstraintIndex); SmallCapTag2; Chr$(9);
            '     Let Table(1, ConstraintIndex + 1) = Trim(Str(ConstraintIndex)) + ". " + mAbbrev(ConstraintIndex)
            '     For SpaceIndex = Len(Str(ConstraintIndex)) + Len(mAbbrev(ConstraintIndex)) To SpacesNeeded + 2
            '         Print #TmpFile, " ";
            '     Next SpaceIndex
            '     For InnerConstraintIndex = 1 To mNumberOfConstraints
            '         If gAPrioriRankingsTable(ConstraintIndex, InnerConstraintIndex) = True Then
            '             Print #TmpFile, Trim(Str(InnerConstraintIndex)); " ";
            '             Print #DocFile, Trim(Str(InnerConstraintIndex)); " ";
            '             Let Table(2, ConstraintIndex + 1) = Table(2, ConstraintIndex + 1) + Trim(Str(InnerConstraintIndex)) + " "
            '         End If
            '     Next InnerConstraintIndex
            '     Print #TmpFile,
            '     Print #DocFile,
            ' Next ConstraintIndex
            ' Print #DocFile, "\te\ke"

End Sub


Sub PrepareDiagnosticTableaux(NumberOfConstraints As Long, Abbrev() As String, Constraint() As String, NumberOfStrata As Long, Stratum() As Long, _
    NumberOfForms As Long, InputForm() As String, Winner() As String, WinnerViolations() As Long, WinnerFrequency() As Single, NumberOfRivals() As Long, _
    MaximumNumberOfRivals As Long, Rival() As String, RivalViolations() As Long, _
    RivalFrequency() As Single, TmpFile As Long, DocFile As Long, HTMFile As Long)

    'Prepare tableaux that contain only the candidates that die on the last, problematic stratum, with constraints that prefer winners or losers on that stratum.

    'We'll use the standard tableau-printing code, and dimension temporary arrays that are sent to it.
        Dim SentNumberOfForms As Long
        Dim SentNumberOfConstraints As Long
        Dim SentConstraintName() As String
        ReDim SentConstraintName(NumberOfConstraints)
        Dim SentAbbrev() As String
        ReDim SentAbbrev(NumberOfConstraints)
        Dim SentStratum() As Long
        ReDim SentStratum(NumberOfConstraints)
        Dim SentInputForm() As String
        ReDim SentInputForm(NumberOfForms)
        Dim SentWinner() As String
        ReDim SentWinner(NumberOfForms)
        Dim SentWinnerFrequency() As Single
        ReDim SentWinnerFrequency(NumberOfForms)
        Dim SentWinnerViolations() As Long
        ReDim SentWinnerViolations(NumberOfForms, NumberOfConstraints)
        Dim SentMaximumNumberOfRivals As Long
        Dim SentNumberOfRivals() As Long
        ReDim SentNumberOfRivals(NumberOfForms)
        Dim SentRival() As String
        ReDim SentRival(NumberOfForms, MaximumNumberOfRivals)
        Dim SentRivalFrequency() As Single
        ReDim SentRivalFrequency(NumberOfForms, MaximumNumberOfRivals)
        Dim SentRivalViolations() As Long
        ReDim SentRivalViolations(NumberOfForms, MaximumNumberOfRivals, NumberOfConstraints)
        Dim SentTmpFile As Long, SentDocFile As Long, SentHTMFile As Long
        Dim SentAlgorithmName As String
        Dim SentRunningFactorialTypology As Boolean
        Dim SentFactorialTypologyIndex As Long
        Dim SentShadingChoice As Boolean
        Dim SentExclamationPointChoice As Boolean
        
       
        Dim RivalOKForMinitableauxFlag As Long
        Dim WinnerOKForMinitableauxFlag As Long
        Dim RivalOkForMinitableaux() As Boolean
        ReDim RivalOkForMinitableaux(mNumberOfForms, mMaximumNumberOfRivals)
        Dim FormOKForMinitableaux() As Boolean
        ReDim FormOKForMinitableaux(mNumberOfForms)
        Dim ConstraintOkForMinitableaux() As Boolean
        ReDim ConstraintOkForMinitableaux(mNumberOfConstraints)
        Dim LocalConstraintCounter As Long
        Dim LocalRivalCounter As Long

        Dim NumberOfColumns As Long
        Dim FirstColumnWidth As Long
        Dim SpaceIndex As Long
    
        Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long
    
        Dim FatalityFlag As Boolean
    
    'Initialize RivalOkForMinitableaux(), with default value True.
        For FormIndex = 1 To NumberOfForms
            For RivalIndex = 1 To NumberOfRivals(FormIndex)
                Let RivalOkForMinitableaux(FormIndex, RivalIndex) = True
            Next RivalIndex
        Next FormIndex
    
    'Take away any candidates that are ruled out on a stratum found by the algorithm.  The latter all have indices greater than zero.
        For FormIndex = 1 To NumberOfForms
            For RivalIndex = 1 To NumberOfRivals(FormIndex)
                For ConstraintIndex = 1 To NumberOfConstraints
                    If Stratum(ConstraintIndex) < NumberOfStrata Then
                        'Is this a winner-preferrer?  If so, eliminate the rival.
                            If RivalViolations(FormIndex, RivalIndex, ConstraintIndex) > WinnerViolations(FormIndex, ConstraintIndex) Then
                                Let RivalOkForMinitableaux(FormIndex, RivalIndex) = False
                            End If
                    End If
                Next ConstraintIndex
            Next RivalIndex
        Next FormIndex
    
    'In addition, a rival candidate is not worth printing if no constraint in the bottom "stratum" prefers it over the winner:
        For FormIndex = 1 To mNumberOfForms
            For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                'Consider only the rivals that are still in contention; i.e. those which need to be ruled out in the bottom stratum:
                If RivalOkForMinitableaux(FormIndex, RivalIndex) = True Then
                    For ConstraintIndex = 1 To NumberOfConstraints
                        If Stratum(ConstraintIndex) = NumberOfStrata Then
                            If WinnerViolations(FormIndex, ConstraintIndex) > RivalViolations(FormIndex, RivalIndex, ConstraintIndex) Then
                                'This rival is preferred by some constraint, so it is worth printing.
                                GoTo ExitPoint
                            End If
                        End If
                    Next ConstraintIndex
                    'If you've gotten this far, then no constraint prefers the loser, so don't include it.
                        Let RivalOkForMinitableaux(FormIndex, RivalIndex) = False
                End If
ExitPoint:          'Some constraint preferred this rival, so keep it.
            Next RivalIndex
        Next FormIndex
    
    'A form is worth printing if it includes a rival candidate worth printing.
        For FormIndex = 1 To mNumberOfForms
            Let FormOKForMinitableaux(FormIndex) = False
            For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                If RivalOkForMinitableaux(FormIndex, RivalIndex) = True Then
                    Let FormOKForMinitableaux(FormIndex) = True
                    Exit For
                End If
            Next RivalIndex
        Next FormIndex
        
    'A constraint is worth printing if it is in the last stratum and prefers a winner or loser among the relevant set.
        For FormIndex = 1 To mNumberOfForms
            If FormOKForMinitableaux(FormIndex) Then
                For RivalIndex = 1 To mNumberOfRivals(FormIndex)
                    If RivalOkForMinitableaux(FormIndex, RivalIndex) Then
                        For ConstraintIndex = 1 To NumberOfConstraints
                            If Stratum(ConstraintIndex) = NumberOfStrata Then
                                If WinnerViolations(FormIndex, ConstraintIndex) > RivalViolations(FormIndex, RivalIndex, ConstraintIndex) Then
                                    Let ConstraintOkForMinitableaux(ConstraintIndex) = True
                                ElseIf WinnerViolations(FormIndex, ConstraintIndex) < RivalViolations(FormIndex, RivalIndex, ConstraintIndex) Then
                                    Let ConstraintOkForMinitableaux(ConstraintIndex) = True
                                End If
                            End If      'Is this constraint in the last stratum?
                        Next ConstraintIndex
                    End If          'Is rival ok for minitableaux?
                Next RivalIndex
            End If                  'Is form ok for minitableaux?
        Next FormIndex
        
     'Print a header
        Call PrintLevel1Header(DocFile, TmpFile, HTMFile, "Diagnostic Tableaux")
        Call PrintPara(DocFile, TmpFile, HTMFile, "The following tables are provided for diagnosis.  They omit all constraints PARAthat were ranked before the algorithm crashed, and all data that are explained PARAby rankable constraints.  They also exclude all constraints that prefer neither winners nor losers in the remaining data.")
     
     'Now fill the variables for printing.
        'Count how many forms will be printed, and install these forms in the "sent" arrays, along with the winners, rivals, and violations.
            'Indices must be counted carefully in order to rescale all of these things right.
            For FormIndex = 1 To NumberOfForms
                If FormOKForMinitableaux(FormIndex) Then
                    Let SentNumberOfForms = SentNumberOfForms + 1
                    Let SentInputForm(SentNumberOfForms) = InputForm(FormIndex)
                    Let SentWinner(SentNumberOfForms) = Winner(FormIndex)
                    Let SentWinnerFrequency(SentNumberOfForms) = WinnerFrequency(FormIndex)
                    'Winner violations.
                        Let LocalConstraintCounter = 0
                        For ConstraintIndex = 1 To NumberOfConstraints
                            If ConstraintOkForMinitableaux(ConstraintIndex) Then
                                Let LocalConstraintCounter = LocalConstraintCounter + 1
                                Let SentWinnerViolations(SentNumberOfForms, LocalConstraintCounter) = WinnerViolations(FormIndex, ConstraintIndex)
                            End If
                        Next ConstraintIndex
                    'Rivals and rival violations.
                        Let LocalRivalCounter = 0
                        For RivalIndex = 1 To NumberOfRivals(FormIndex)
                            If RivalOkForMinitableaux(FormIndex, RivalIndex) Then
                                Let LocalRivalCounter = LocalRivalCounter + 1
                                Let SentRival(SentNumberOfForms, LocalRivalCounter) = Rival(FormIndex, RivalIndex)
                                Let SentRivalFrequency(SentNumberOfForms, LocalRivalCounter) = RivalFrequency(FormIndex, RivalIndex)
                                Let SentNumberOfRivals(SentNumberOfForms) = SentNumberOfRivals(SentNumberOfForms) + 1
                                'Now the violations:
                                    Let LocalConstraintCounter = 0
                                    For ConstraintIndex = 1 To NumberOfConstraints
                                        If ConstraintOkForMinitableaux(ConstraintIndex) Then
                                            Let LocalConstraintCounter = LocalConstraintCounter + 1
                                            Let SentRivalViolations(SentNumberOfForms, LocalRivalCounter, LocalConstraintCounter) = RivalViolations(FormIndex, RivalIndex, ConstraintIndex)
                                        End If
                                    Next ConstraintIndex
                            End If
                        Next RivalIndex
                        'Augment maximum number of rivals if appropriate.
                            If LocalRivalCounter > SentMaximumNumberOfRivals Then
                                Let SentMaximumNumberOfRivals = LocalRivalCounter
                            End If
                End If              'Is this form ok for minitableaux?
            Next FormIndex
            
        'Count how many constraints will be printed, and create arrays of abbreviations and constraint names that fit in this category.
        '   Definitionally, they all fit in the same stratum.
            For ConstraintIndex = 1 To NumberOfConstraints
                If ConstraintOkForMinitableaux(ConstraintIndex) Then
                    Let SentNumberOfConstraints = SentNumberOfConstraints + 1
                    Let SentAbbrev(SentNumberOfConstraints) = Abbrev(ConstraintIndex)
                    Let SentConstraintName(SentNumberOfConstraints) = Constraint(ConstraintIndex)
                    Let SentStratum(SentNumberOfConstraints) = 1
                End If
            Next ConstraintIndex
            
        'Trivial other sent material, done for consistency.
            Let SentTmpFile = TmpFile
            Let SentDocFile = DocFile
            Let SentHTMFile = HTMFile
            Let SentAlgorithmName = gAlgorithmName
            Let SentRunningFactorialTypology = False
            Let SentFactorialTypologyIndex = 0
            Let SentShadingChoice = False
            Let SentExclamationPointChoice = False
            
        'Launch tableaux printing.
            Call PrintLevel1Header(DocFile, TmpFile, HTMFile, "Tableaux")
            Call PrintTableaux.Main(SentNumberOfForms, SentNumberOfConstraints, SentConstraintName(), SentAbbrev(), SentStratum(), _
            SentInputForm(), SentWinner(), SentWinnerFrequency(), SentWinnerViolations(), _
            SentMaximumNumberOfRivals, SentNumberOfRivals(), SentRival(), SentRivalFrequency(), SentRivalViolations(), _
            SentTmpFile, SentDocFile, SentHTMFile, SentAlgorithmName, SentRunningFactorialTypology, SentFactorialTypologyIndex, _
            SentShadingChoice, SentExclamationPointChoice)
     
End Sub


Sub PrintDraftCopy()
        
    'Print a rough copy of the FileNameDraftOutput.txt file, for users who are having trouble with Word.
    
    On Error GoTo CheckError
        
    Dim MyLine As String
    Dim PageNumber As Long
    Dim PageBreakThreshold As Single
    Dim PageBreakFlag As Boolean
    Dim EarlyPageBreakThreshold As Single
    Dim RightMarginValue As Single
    Dim RoughCopyFile As Long
    Let RoughCopyFile = FreeFile
    
    'Report progress
        Let lblProgressWindow.Caption = "Printing..."
    
    'Open the file you will print.
        Close
        'Don't ever open a file that doesn't exist.
            If Dir(gOutputFilePath + gFileName + "DraftOutput.txt") <> "" Then
                Open gOutputFilePath + gFileName + "DraftOutput.txt" For Input As #RoughCopyFile
            Else
                MsgBox "Sorry, I can't find the file needed to draft-print your results, " + _
                    gOutputFilePath + gFileName + "DraftOutput.txt.  Click OK to continue.", vbExclamation
                Exit Sub
            End If
    
    'Set up values needed for printing.
        Dim Zooming As Single
        Let Zooming = Val(frmPrinting.txtReduction.Text)
        Let Printer.Zoom = Zooming
        Let PageNumber = 1
    
    'Set up orientation.
        If frmPrinting.optPortrait.Value = True Then
            Let Printer.Orientation = vbPRORPortrait
        Else
            Let Printer.Orientation = vbPRORLandscape
        End If
    
    'Set the thresholds for page breaks and page number location, based on orientation.
        Select Case Printer.Orientation
            Case vbPRORPortrait
                Let EarlyPageBreakThreshold = 9
                Let PageBreakThreshold = 10
                Let RightMarginValue = 7
            Case vbPRORLandscape
                Let EarlyPageBreakThreshold = 6
                Let PageBreakThreshold = 7.5
                Let RightMarginValue = 10
        End Select
    
    'Put in a top margin
        Let Printer.CurrentY = 900 * (100 / Zooming)
    
    'Go through the printable file, printing it line by line.
        Do While Not EOF(RoughCopyFile)
            Line Input #RoughCopyFile, MyLine
            'A crude method of break control:  put one in if there's a blank line and you're
            '  at least nine (or six, for landscape) inches down the page.
                If Printer.CurrentY > EarlyPageBreakThreshold * 1440 * (100 / Zooming) Then
                    If MyLine = "" Then
                        Let PageBreakFlag = True
                    End If
                End If
            'Put in a break nonetheless, even if you have to break a continuous sequence:
                If Printer.CurrentY > PageBreakThreshold * 1440 * (100 / Zooming) Then
                    Let PageBreakFlag = True
                End If
            'Print the page break if either of these conditions was met.
                If PageBreakFlag = True Then
                   Printer.NewPage
                   Let PageNumber = PageNumber + 1
                   Let Printer.CurrentY = 900 * (100 / Zooming)
                   Let Printer.CurrentX = 1440 * RightMarginValue * (100 / Zooming)
                   Printer.Print "p. " + Str(PageNumber)
                   Printer.Print
                End If
                Let PageBreakFlag = False
            'Otherwise, just print the next line, with standard indentation.
                Let Printer.CurrentX = 800
                Printer.Print MyLine
        Loop
    
    'Tell the printer you're done.
        Printer.EndDoc
            
    'Report progress
        Let lblProgressWindow.Caption = "File " + gFileName + "DraftOutput.txt has been sent to the printer " + _
            Printer.DeviceName + "."
    
    Exit Sub
    
CheckError:
    Select Case Err.Number
        Case 53
            MsgBox ("I can't print your results.  Possible cause:  you need to run the program first, to get an output file I can print.  Try clicking the Rank button."), vbExclamation
        Case 57
            MsgBox ("I can't print your results.  Possible cause:  I'm having trouble accessing the printer."), vbExclamation
        Case Else
            MsgBox ("I can't print your results, and I don't know why.  Contact product support (bhayes@humnet.ucla.edu)."), vbCritical
    End Select
    
End Sub 'BH:  PrintDraftCopy


Public Sub PrintResultsOfRankingSoFar(CurrentStratum As Long, Stratum() As Long, ShowMe As Long, Faithfulness() As Boolean)

    'For algorithms reporting what they have, print out the results of learning for a single stratum.
    
    Dim LocalStratumIndex As Long, ConstraintIndex As Long
    Dim FoundOne As Boolean
    Dim LastStratum As Long
    
    Print #ShowMe,
    Print #ShowMe, "Results so far:"
    Print #ShowMe,
    'Already ranked, sorted by stratum:
        Let LastStratum = 0
        For LocalStratumIndex = 1 To CurrentStratum - 1
            For ConstraintIndex = 1 To mNumberOfConstraints
                If Stratum(ConstraintIndex) = LocalStratumIndex Then
                    If LocalStratumIndex = LastStratum Then
                        'Graceful blank space for multiple constraints in same stratum.
                        Print #ShowMe, "                               "; mAbbrev(ConstraintIndex)
                    Else
                        Print #ShowMe, "  Stratum "; Trim(Str(LocalStratumIndex)); " (already ranked):  "; mAbbrev(ConstraintIndex)
                        Let LastStratum = LocalStratumIndex
                    End If
                End If
            Next ConstraintIndex
        Next LocalStratumIndex
    'Ranked in the current stratum:
            Let FoundOne = False
            For ConstraintIndex = 1 To mNumberOfConstraints
                If Stratum(ConstraintIndex) = CurrentStratum Then
                    If FoundOne = False Then
                        Print #ShowMe, "  Stratum "; Trim(Str(CurrentStratum)); " (newly ranked):    "; mAbbrev(ConstraintIndex)
                        Let FoundOne = True
                    Else
                        Print #ShowMe, "                               "; mAbbrev(ConstraintIndex)
                    End If
                End If
            Next ConstraintIndex
        Print #ShowMe,
        
    'Remaining Markedness constraints:
        Let FoundOne = False
        Print #ShowMe, "  Markedness constraints still unranked:"
        For ConstraintIndex = 1 To mNumberOfConstraints
            If Stratum(ConstraintIndex) = 0 And Faithfulness(ConstraintIndex) = False Then
                Print #ShowMe, "    "; mAbbrev(ConstraintIndex)
                Let FoundOne = True
            End If
        Next ConstraintIndex
        If FoundOne = False Then
            Print #ShowMe, "    (none)"
        End If
    'Remaining Faithfulness constraints:
        Let FoundOne = False
        Print #ShowMe, "  Faithfulness constraints still unranked:"
        For ConstraintIndex = 1 To mNumberOfConstraints
            If Stratum(ConstraintIndex) = 0 And Faithfulness(ConstraintIndex) = True Then
                Print #ShowMe, "    "; mAbbrev(ConstraintIndex)
                Let FoundOne = True
            End If
        Next ConstraintIndex
        If FoundOne = False Then
            Print #ShowMe, "    (none)"
        End If

End Sub


