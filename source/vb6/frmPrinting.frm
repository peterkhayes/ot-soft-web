VERSION 5.00
Begin VB.Form frmPrinting 
   Caption         =   "OTSoft - Options for Draft Printing"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4725
   Icon            =   "frmPrinting.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4875
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print draft output for the current file"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   8
      Top             =   3720
      Width           =   4095
   End
   Begin VB.CommandButton cmdSwitchPrinters 
      Caption         =   "&Switch printers"
      Height          =   735
      Left            =   360
      TabIndex        =   7
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit this menu"
      Height          =   735
      Left            =   2520
      TabIndex        =   6
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Frame frameOrientation 
      Caption         =   "Page Orientation"
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   4455
      Begin VB.OptionButton optLandscape 
         Caption         =   "Landscape"
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton optPortrait 
         Caption         =   "Portrait"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame frmReduction 
      Caption         =   "Reduction"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4455
      Begin VB.TextBox txtReduction 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   3720
         TabIndex        =   2
         Text            =   "80"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblReduction 
         Caption         =   "Specify a number for how much reduction you want in the printed output.  100 = full size; less than 100 is reduced-size print."
         Height          =   615
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmPrinting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================================frmPrinting=======================================

'Let the user control the draft-print function.

Private Sub Form_Load()
    'Set up the form
        Let Me.Left = (Screen.Width - Me.Width) / 2
        Let Me.Top = (Screen.Height - Me.Height) / 2
        Let Me.Caption = "OTSoft " + gMyVersionNumber + " - Print in Draft Mode"
End Sub

Private Sub cmdSwitchPrinters_Click()

    'Let the user switch printers for printing drafts.
    
    MsgBox "Switching printers is now a disabled function -- sorry!  You can use Windows to pick a default printer, or else use your word processor to open one of your output files (see folder FilesForXXX) and print it from there. Click OK to continue."
    
    'Diabled, as a function of the CommonDialog toolkit.
    
    '    Form1.CommonDialog1.ShowPrinter
End Sub

Private Sub cmdPrint_Click()
    Call Form1.PrintDraftCopy
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

