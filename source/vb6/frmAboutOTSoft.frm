VERSION 5.00
Begin VB.Form frmAboutOTSoft 
   BackColor       =   &H00FFFFFF&
   Caption         =   "AboutOTSoft"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6840
   Icon            =   "frmAboutOTSoft.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3225
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAcknowledgements 
      Caption         =   "&Acknowledgments"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   2640
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3600
      Picture         =   "frmAboutOTSoft.frx":030A
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      Height          =   375
      Left            =   4320
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   2640
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "frmAboutOTSoft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'About OTSoft


Private Sub Form_Load()
    Let Me.Top = (Screen.Height - Me.Height) / 2
    Let Me.Left = (Screen.Width - Me.Width) / 2
    Let Me.Caption = "OTSoft " + gMyVersionNumber
    Let Label1.Caption = "OTSoft is a freeware program for Windows, intended to facilitate research and training " + _
        "in Optimality Theory and similar frameworks descended from OT. OTSoft was programmed by Bruce Hayes of UCLA (bhayes@humnet.ucla.edu), " + _
        "with contributions from Bruce Tesar of Rutgers University, and Kie Zuraw of UCLA." + _
        Chr(10) + Chr(10) + _
        "If you use OTSoft " + gMyVersionNumber + " in verifying a published analysis, you may cite it as follows:" + _
        Chr(10) + Chr(10) + _
        "Hayes, Bruce, Bruce Tesar, and Kie Zuraw (2013) " + Chr(34) + "OTSoft " + gMyVersionNumber + "," + Chr(34) + " software package, https://brucehayes.org/otsoft/."
    Show
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdAcknowledgements_Click()
    MsgBox "Thanks to " + _
            Chr(10) + Chr(10) + _
        "   Bruce Tesar for programming Biased Constraint Demotion" + _
            Chr(10) + _
        "   Kie Zuraw for many improvements to the GLA code and elsewhere" + _
            Chr(10) + _
        "   Paul Boersma for Stochastic OT advice" + _
            Chr(10) + _
        "   Gerhard Jäger for serial-MaxEnt advice" + _
            Chr(10) + _
        "   Lukas Pietsch for useful ideas and improvements" + _
            Chr(10) + _
        "   Taesun Moon for some clever debugging" + _
            Chr(10) + Chr(10) + _
        "and to all the users who have reported bugs or made suggestions concerning previous versions of OTSoft.", _
        vbInformation, "OTSoft " + gMyVersionNumber
End Sub

Private Sub Picture1_Click()
    Unload Me
End Sub
