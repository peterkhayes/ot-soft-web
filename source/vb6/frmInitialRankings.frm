VERSION 5.00
Begin VB.Form frmInitialRankings 
   Caption         =   "Customize Initial Constraint Rankings"
   ClientHeight    =   3960
   ClientLeft      =   7365
   ClientTop       =   3255
   ClientWidth     =   4440
   ControlBox      =   0   'False
   Icon            =   "frmInitialRankings.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3960
   ScaleWidth      =   4440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2640
      TabIndex        =   9
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   2640
      TabIndex        =   8
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdRevertFaith 
      Caption         =   "Revert to default for &faithfulness"
      Height          =   615
      Left            =   2640
      TabIndex        =   7
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdRevertMark 
      Caption         =   "Revert to default for &markedness"
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtFaith 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3000
      TabIndex        =   3
      Text            =   "100"
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtMark 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Text            =   "100"
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   $"frmInitialRankings.frx":030A
      Height          =   1335
      Left            =   360
      TabIndex        =   5
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Initial constraint ranking ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Faithfulness constraints"
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Markedness constraints"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmInitialRankings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===================================frmInitialRankings=======================================

    'Let the user separate initial rankings or weights for Faithfulness and Markedness.

    Option Explicit

    Dim DefaultMark As Single, DefaultFaith As Single


Private Sub Form_Load()

    'Install caption
        Let Me.Caption = "OTSoft " + gMyVersionNumber + " - Customize Initial Rankings or Weights"
    
    'Center the form on the screen.
        Let Me.Left = (Screen.Width - Me.Width) / 2
        Let Me.Top = (Screen.Height - Me.Height) / 2
    
    'KZ: record the default ranking values
        If blnCustomRankCreated = False Then  'KZ: don't do this if the user has
                                              'already created a custom noise schedule.
            DefaultMark = txtMark.Text
            DefaultFaith = txtFaith.Text
        End If

End Sub

Private Sub cmdOK_Click()
    
    'KZ: record the ranking values entered by the user and exit to GLA.
    
    'KZ: first check if a valid number was entered in every box
        If Not GoodNumberPerhapsNegative(txtFaith.Text) Or _
            Not GoodNumberPerhapsNegative(txtMark.Text) Then
                MsgBox "You must enter a number in each box.", vbExclamation
                Exit Sub    'KZ: don't unload the form.
        End If
    
    'KZ: now record the values
        Let blnCustomRankCreated = True
        Let gCustomRankMark = Val(txtMark.Text)
        Let gCustomRankFaith = Val(txtFaith.Text)
    
    'KZ: note in GLA that custom values are being used
        Let GLA.lblCustomRank.Visible = True
        Let GLA.lblCustomRank.Caption = "Separate initial rankings or weights for Markedness and Faithfulness"
    
    'KZ: also in GLA, assume user will want to use this schedule.
        Let GLA.mnuUseSeparateMarkFaithInitialRankings.Visible = True
        Let GLA.mnuUseSeparateMarkFaithInitialRankings.Checked = True
        
    'Decheck the rival options.
        Let GLA.mnuUseDefaultInitialRankingValues.Checked = False
        Let GLA.mnuUseFullyCustomizedInitialRankingValues.Checked = False
        Let GLA.mnuUsePreviousResultsAsInitialRankingValues.Checked = False

    'Remember the choice as a parameter.
        Let InitialRankingChoice = MarkednessFaithfulness
    
        Unload Me

End Sub

Private Sub cmdRevertMark_Click()
    'KZ: Revert to default ranking for markedness constraints
        Let txtMark.Text = DefaultMark
End Sub

Private Sub cmdRevertFaith_Click()
    'KZ: Revert to default noise values for faithfulness constraints
        Let txtFaith.Text = DefaultFaith
End Sub

Private Sub cmdCancel_Click()
    'KZ: go back to GLA.frm without recording any custom ranking values
        Unload Me
End Sub





