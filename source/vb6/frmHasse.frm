VERSION 5.00
Begin VB.Form frmHasse 
   Caption         =   "OTSoft - Hasse Diagram"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   Icon            =   "frmHasse.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgHasse 
      Height          =   1095
      Left            =   1320
      Top             =   840
      Width           =   2295
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuShrink 
         Caption         =   "&Fit image to screen"
      End
      Begin VB.Menu mnuOriginalSize 
         Caption         =   "View image at &original size"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPaint 
         Caption         =   "View image with MS&Paint"
      End
   End
   Begin VB.Menu mnuClose 
      Caption         =   "&Close this window"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit OTSoft"
   End
End
Attribute VB_Name = "frmHasse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================frmHasse============================================

'Display a Hasse diagram in its own window.

Option Explicit

Private Sub Form_Load()
    
    On Error GoTo CheckError
    
    'Label this window with the version number of OTSoft.
        Let Me.Caption = "OTSoft " + gMyVersionNumber + " - Hasse Diagram"
    
    'Make sure Dot.exe is installed.
        If Dir(gDotExeLocation) = "" Then
            MsgBox "OTSoft did not generate a graphics file for the Hasse diagram." + _
                Chr(10) + Chr(10) + _
                "I conjecture that this because OTSoft was unable to access the necessary ATT GraphViz software." + _
                Chr(10) + Chr(10) + _
                "If you have already installed GraphViz, open the file " + App.Path + "OTSoftAuxiliarySoftwareLocations.txt, in the OTSoft folder, and alter it to tell OTSoft where the crucial program of GraphViz, dot.exe, resides on your computer." + _
                Chr(10) + Chr(10) + _
                "If you have not installed GraphViz , you can get it for free by downloading it from " + gATTWebSite + "." + _
                Chr(10) + Chr(10) + _
                "Note that you can still view your results without the Hasse diagram." + _
                Chr(10) + Chr(10) + _
                "Click OK to proceed further.", vbExclamation
                Exit Sub
        End If
    
    'Also, make sure the file exists.
        If Dir(gOutputFilePath + gFileName + "Hasse.gif") = "" Then
            'The most obvious explanation is that the user didn't check the Ranking
            '   Arguments box.
                If Form1.chkArguerOn.Value <> vbChecked And gAlgorithmName <> "GLA" Then
                    MsgBox "OTSoft did not generate a graphics file for the Hasse diagram." + _
                        Chr(10) + Chr(10) + _
                        "To generate such a file, you must check the Include Ranking Arguments box before you rank." + _
                        Chr(10) + Chr(10) + _
                        "Note that you can still view your results without the Hasse diagram.", vbExclamation
                        'This crashes the program.  Need to let user close window (grr...).
                            'Unload Me
                        Exit Sub
                Else
                    'The user *did* check the box, and it's *not* the GLA.  Who knows
                    '   what's happening.
                    MsgBox "For undiagnosed reasons, OTSoft did not generate a graphics file for the Hasse diagram." + _
                        Chr(10) + Chr(10) + _
                        "Please report this error to Bruce Hayes at bhayes@humnet.ucla.edu, specifying error #88555." + _
                        Chr(10) + Chr(10) + _
                        "Note that you can still view your results without the Hasse diagram.", vbCritical
                        'This crashes the program.  Need to let user close window (grr...).
                            'Unload Me
                        Exit Sub
                End If
        End If
        
        'Assuming all is ok, show the Hasse diagram:
        
            'Load the image.
                Let imgHasse.Picture = LoadPicture(gOutputFilePath + gFileName + "Hasse.gif")
            'Size the picture window and its parent form.
                Let imgHasse.Left = 200
                Let imgHasse.Top = 0
                Let Me.Height = imgHasse.Height + 800
                Let Me.Width = imgHasse.Width + 500
            'Let's try putting it off to the right, so user can look at text on the left.
            '   Formerly, it was centered.
                Let Me.Left = Screen.Width - Me.Width   '(Screen.Width - Me.Width) / 2
                Let Me.Top = 0                          '(Screen.Height - Me.Height) / 2
                
            Exit Sub

CheckError:

    MsgBox "An error occurred in trying to display your Hasse diagram.  For diagnosis, try examining these two files:" + _
        Chr(10) + Chr(10) + _
        gOutputFilePath + gFileName + "Hasse.txt" + _
        Chr(10) + Chr(10) + _
        gOutputFilePath + gFileName + "Hasse.gif", vbCritical
    Exit Sub


End Sub


Private Sub mnuClose_Click()
    Unload Me
End Sub

Private Sub mnuExit_Click()
    Call Form1.SaveUserChoices
    If Form1.mnuDeleteTmpFiles.Checked = True Then Call Form1.DeleteTmpFiles
    End
End Sub

Private Sub mnuPaint_Click()

    'Call Paint to show the Hasse diagram more nicely.
        
        On Error GoTo CheckError
        
    'First make sure that there is a program available to look at the file.
        If Dir(gPaintLocation) = "" Then
            'Can't find Paint:
            Select Case MsgBox("I'm having trouble displaying your Hasse diagram.  The problem seems to be that I can't find the copy of Paint (or other graphics program) you specified." + _
                Chr(10) + Chr(10) + _
                "I'm currently looking for this program (and not finding it) in this location:" + _
                Chr(10) + Chr(10) + _
                gPaintLocation + _
                Chr(10) + Chr(10) + _
                "Please open the file" + _
                Chr(10) + Chr(10) + _
                App.Path + "\OTSoftAuxiliarySoftwareLocations.txt" + _
                Chr(10) + Chr(10) + _
                "and type in the location of your copy of Paint (or other graphics program) on the relevant line.  Then try again." + _
                Chr(10) + Chr(10) + _
                "If you like, I can try to find some other program to show your Hasse diagram.  Click Yes for this option, No to return to the main OTSoft window.", _
                vbYesNo + vbExclamation)
                'Old:  gSafePlaceToWriteTo + "\OTSoftAuxiliarySoftwareLocations.txt" + _

                
                Case vbYes
                    'Try opening with ShellExecute
                        Select Case UseWindowsPrograms.TryShellExecute(gOutputFilePath + gFileName + "Hasse.gif")
                            Case gShellExecuteWasSuccessful
                                'Do nothing
                            Case Else
                                MsgBox "Sorry, I can't still am having trouble displaying your Hasse file." + _
                                    Chr(10) + Chr(10) + _
                                    "Please report this program error to Bruce Hayes at bhayes@humnet.ucla, specifying error #70046." + _
                                    Chr(10) + Chr(10) + _
                                    "Click OK to return to the main OTSoft screen.", vbCritical
                                Exit Sub
                        End Select          'How did ShellExecute go?
                Case vbNo
                    'do nothing
            End Select
        Else
            'Since the Paint or other program is there, go ahead and use it:
            '   Note:  chr(34), the quote mark, is needed around all file names in a Shell command,
            '       else spaces will foil it.
                Dim Dummy As Long
                Let Dummy = Shell(gPaintLocation + " " + _
                    Chr(34) + gOutputFilePath + gFileName + "Hasse.gif" + Chr(34), _
                    vbNormalFocus)
        End If
        
        Exit Sub
    
CheckError:
        MsgBox "Program error.  I can't call up Paint (or other software) to view your Hasse diagram, for a reason not yet diagnosed.  For help please contact bhayes@humnet.ucla, specifying error #46172, and including a copy of your input file." + _
            Chr(10) + Chr(10) + _
            "Click OK to return to the main menu.", vbCritical
        Exit Sub
    
End Sub

Private Sub mnuShrink_Click()
    
    'Shrink (or stretch) the Hasse diagram on the screen.
    
        On Error GoTo CheckError
    
    'The strategy of actually resizing the .gif file with dot.exe doesn't seem
    '   to be working.
    
     '   Dim HasseFile As Long
     '   Let HasseFile = FreeFile
     '   Open gOutputFilePath + gFileName + "hasse.txt" For Input As HasseFile
     '   '   Dim HasseFileContent() As String
     '   Dim HasseFileLength As Long
     '   Dim MyLine As String
     '
     '   Do While Not EOF(HasseFile)
     '       Line Input #HasseFile, MyLine
     '       Let HasseFileLength = HasseFileLength + 1
     '       ReDim Preserve HasseFileContent(HasseFileLength)
     '       Let HasseFileContent(HasseFileLength) = MyLine
     '   Loop
     '
     '   Close #HasseFile
     '
     '   Open gOutputFilePath + gFileName + "hasse.txt" For Output As HasseFile
     '   'Repeat first line:
     '       Print #HasseFile, HasseFileContent(1)
     '   'Now, the crucial line with size:
     '       Print #HasseFile, "   size =" + Chr(34) + _
     '           Trim(Str(Str((Screen.Height - 800) / 1440))) + "," + _
     '           Trim(Str(Str((Screen.Width - 600) / 1440))) + _
     '           Chr(34) + ";"
     '   'Now, the rest of the original file.
     '       Dim HasseIndex As Long
     '       For HasseIndex = 2 To HasseFileLength
     '           Print #HasseFile, HasseFileContent(HasseIndex)
     '       Next HasseIndex
     '       Close #HasseFile
     '
     '       Call Form1.mnuReplotHasse_Click
     '
     '   Let imgHasse.Picture = LoadPicture(gOutputFilePath + gFileName + "Hasse.gif")
    
    'So we'll just use the Visual Basic resizing, which is crude...
    
    'Establish the maximum window size
        If Me.Height > Screen.Height - 200 Then
            Let Me.Height = Screen.Height - 200
        End If
        If Me.Width > Screen.Width Then
            Let Me.Width = Screen.Width
        End If

    'Let's try putting it off to the right, so user can look at text on the left.
    '   Formerly, it was centered.
        Let Me.Left = Screen.Width - Me.Width   '(Screen.Width - Me.Width) / 2
        Let Me.Top = 0                          '(Screen.Height - Me.Height) / 2

    'Shrink.
        Let imgHasse.Stretch = True

    'Size the picture window.
        Let imgHasse.Left = 200
        Let imgHasse.Top = 0
        Let imgHasse.Height = Me.Height - 600
        Let imgHasse.Width = Me.Width - 500
        
    'Toggle the menu options.
        Let mnuOriginalSize.Visible = True
        Let mnuShrink.Visible = False
        
CheckError:
    Select Case Err.Number
        Case 53
            Resume
        Case Else
            Exit Sub
    End Select
    
End Sub

Private Sub mnuOriginalSize_Click()

    'Shrink.
        Let imgHasse.Stretch = False

    'Toggle the menu options.
        Let mnuOriginalSize.Visible = True
        Let mnuShrink.Visible = False

    'Size the picture window and its parent form.
        Let imgHasse.Left = 200
        Let imgHasse.Top = 0
        Let Me.Height = imgHasse.Height + 800
        Let Me.Width = imgHasse.Width + 500
        
    'Let's try putting it off to the right, so user can look at text on the left.
    '   Formerly, it was centered.
        Let Me.Left = Screen.Width - Me.Width   '(Screen.Width - Me.Width) / 2
        Let Me.Top = 0                          '(Screen.Height - Me.Height) / 2

    'Toggle the menu options.
        Let mnuOriginalSize.Visible = False
        Let mnuShrink.Visible = True
        
End Sub

