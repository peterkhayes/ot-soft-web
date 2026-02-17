VERSION 5.00
Begin VB.Form frmHTMLOptions 
   Caption         =   "HTML Options"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7500
   LinkTopic       =   "Form2"
   ScaleHeight     =   3705
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExitOnly 
      Caption         =   "E&xit without saving"
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   2640
      Width           =   3135
   End
   Begin VB.CommandButton cmdSaveAndExit 
      Caption         =   "&Save and exit"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   2640
      Width           =   3135
   End
   Begin VB.TextBox txtHTMLColorCode 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox txtGrayness 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Text            =   "50"
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   $"frmHTMLOptions.frx":0000
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Width           =   6495
   End
   Begin VB.Label Label1 
      Caption         =   "Pick how dark you want shading in HTML (Web page) tableaux to be.  1 = white, 100 = black."
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   6495
   End
End
Attribute VB_Name = "frmHTMLOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Enter the user's preferred information about HTML tableaux.

Private Sub Form_Load()

    'Center the form on the screen.
        Let Me.Left = 0.5 * (Screen.Width - Me.Width)
        Let Me.Top = 0.5 * (Screen.Height - Me.Height)
        
    'Put the current program version number on the form.
        Let Me.Caption = "OTSoft " + gMyVersionNumber + " - " + gInputFilePath + gFileName + gFileSuffix
        
        Let txtHTMLColorCode.Text = gShadingColor
   
End Sub

Private Sub txtHTMLColorCode_Change()
    'White this window, gray the other window when information is entered in this one.
        Let txtHTMLColorCode.BackColor = &HFFFFFF   'white
        Let txtGrayness.BackColor = &HE0E0E0        'gray
        
End Sub
Private Sub txtGrayness_Change()
    'White this window, gray the other window when information is entered in this one.
        Let txtHTMLColorCode.BackColor = &HE0E0E0   'gray
        Let txtGrayness.BackColor = &HFFFFFF        'white
End Sub

Sub Main()
    'Show this form.
    Me.Show
End Sub

Private Sub cmdSaveAndExit_Click()

    'Check the entry for validity, then convert to HTML code where appropriate, record the answer, and exit.
    
    If txtGrayness.BackColor = &HFFFFFF Then
        'percentage gray
            If BadPercentage(txtGrayness.Text) Then
                MsgBox "Please enter a number between 0 and 100.", vbExclamation
                Exit Sub
            Else
                Let gShadingColor = ColorCodeForGrayness(txtGrayness.Text)
                MsgBox "Ok.  This change will take effect the next time you run a learning simulation.  (It will not change an output file already created.)", vbInformation
                Unload Me
            End If
    Else
        'HTM color code
            If BadColorCode(txtHTMLColorCode.Text) Then
                MsgBox "Please enter a valid HTML color code.  These generally are six characters long, and include any digit plus letters from A-F.", vbExclamation
                Exit Sub
            Else
                Let gShadingColor = Trim(txtHTMLColorCode)
                MsgBox "Ok.  This change will take effect the next time you run a learning simulation.  (It will not change an output file already created.)", vbInformation
                Unload Me
            End If
    End If

End Sub


Private Sub cmdExitOnly_Click()
    Unload Me
End Sub

Function BadPercentage(MyText As String) As Boolean

    'Vet that what was entered was a whole number between 0 and 100, giving advice to the user.
        If s.IsAnInteger(MyText) = False Then
            MsgBox "Please enter a whole number.", vbExclamation
            Let BadPercentage = True
        ElseIf Val(MyText) < 1 Then
            MsgBox "Please enter a whole number from the range 1-100.", vbExclamation
            Let BadPercentage = True
        ElseIf Val(MyText) > 100 Then
            MsgBox "Please enter a whole number from the range 1-100.", vbExclamation
            Let BadPercentage = True
        Else
            Let BadPercentage = False
        End If
        
End Function

Function BadColorCode(MyText As String) As Boolean
    
    'Vet that what was entered was a whole number between 0 and 100, giving advice to the user.
    
        Dim i As Long
        
        If Len(Trim(MyText)) <> 6 Then
            Let BadColorCode = True
        Else
            For i = 1 To 6
                Select Case Mid(MyText, i, 1)
                    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F"
                        'do nothing
                    Case Else
                        Let BadColorCode = True
                        Exit Function
                End Select
            Next i
        End If
        
End Function

Function ColorCodeForGrayness(MyGrayness As String) As String
    
    Dim Buffer As String
    
    'What this does:  convert to a number, divide it by 100 to get a proportion, multiply by 256 (the range of possible
    '   RBG component values), subtract one (since the real range is 0-255), convert to hex, then triplicate.
        Let Dummy = Hex((((101 - Int(Val(MyGrayness))) / 100) * 256) - 1)
        Let ColorCodeForGrayness = Dummy + Dummy + Dummy
        
End Function
