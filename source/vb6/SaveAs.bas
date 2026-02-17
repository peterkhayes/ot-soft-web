Attribute VB_Name = "SaveAs"
Option Explicit

    Dim mPraatFileFlag As Boolean

Sub SaveAsPraat(MyFileName As String, SortByRank As Boolean, Winner() As String, NumberOfConstraints As Long, ConstraintName() As String, _
    NumberOfForms As Long, InputForm() As String, _
    NumberOfRivals() As Long, WinnerViolations() As Long, WinnerFrequency() As Single, Rival() As String, RivalViolations() As Long, _
    RivalFrequency() As Single)

    'Save the input file in Praat format.
    
    On Error GoTo CheckError
    
    Dim TotalNumberOfCandidates As Long
    Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long, PairIndex As Long
    Dim PositionIndex As Long
    Dim ReportErrorPraatFileName As String
    
    Dim PraatFile As Long
    Let PraatFile = FreeFile
    
    Call Form1.CreateAFolderForOutputFiles
    
    Let mPraatFileFlag = True
    Call Form1.DigestTheInputFile(gInputFilePath, gFileName, gFileSuffix)  'KZ: DigestTheInputFile is now a boolean function
    Let mPraatFileFlag = False

    
    Let ReportErrorPraatFileName = MyFileName + ".txt"
    
    'Two files must be created:  xxx.OTGrammar, and xxx.PairDistribution
    
    'Create xxx.OTGrammar:
    
        'Backups go in the downstairs folder; straightforward saves go in the upstairs folder.
            If LCase(Right(MyFileName, 6)) = "backup" Then
                Open gOutputFilePath + MyFileName + "ForPraat.OTGrammar" For Output As #PraatFile
            Else
                Open gOutputFilePath + MyFileName + "ForPraat.OTGrammar.txt" For Output As #PraatFile
            End If
        
        'Print stuff at top of file
            Print #PraatFile, "File type = " + Chr(34) + "ooTextFile" + Chr(34)
            Print #PraatFile, "Object class = " + Chr(34) + "OTGrammar 2" + Chr(34)
            Print #PraatFile,
            Print #PraatFile, "<OptimalityTheory>"
            Print #PraatFile, "0 ! leak"
            
        'Number of constraints:
            Print #PraatFile, Trim(Str(NumberOfConstraints)) + " constraints"
            
        'Print constraints, then a blank:
            For ConstraintIndex = 1 To NumberOfConstraints
                Print #PraatFile, "constraint [" + Trim(Str(ConstraintIndex)) + "]: " + Chr(34) + _
                    ConstraintName(ConstraintIndex) + Chr(34) + " 0.0 0.0 1 ! " + ConstraintName(ConstraintIndex)
            Next ConstraintIndex
            Print #PraatFile,
            
        'Fixed text (for now) and a blank:
            Print #PraatFile, "0 fixed rankings"
            Print #PraatFile,
            
        'Input count:
            Print #PraatFile, Trim(Str(NumberOfForms)) + " tableaus"
            
        'Inputs and candidates:
            For FormIndex = 1 To NumberOfForms
                'Print the input:
                    Print #PraatFile, "input [" + Trim(Str(FormIndex)) + "]: " + Chr(34) + InputForm(FormIndex) + Chr(34) + " " + Trim(Str(NumberOfRivals(FormIndex) + 1))
                'Print the winner as the first candidate:
                    Print #PraatFile, Chr(9); "candidate[1]: " + Chr(34) + Winner(FormIndex) + Chr(34);
                    'Its violations:
                        For ConstraintIndex = 1 To NumberOfConstraints
                            Print #PraatFile, " " + Trim(Str(WinnerViolations(FormIndex, ConstraintIndex)));
                        Next ConstraintIndex
                        Print #PraatFile,
                'Now the "rivals" as the remaining candidates:
                    For RivalIndex = 1 To NumberOfRivals(FormIndex)
                        'Print the rival candidate:
                            Print #PraatFile, Chr(9); "candidate[" + Trim(Str(RivalIndex + 1)) + "]: " + Chr(34) + Rival(FormIndex, RivalIndex) + Chr(34);
                    'Its violations:
                        For ConstraintIndex = 1 To NumberOfConstraints
                            Print #PraatFile, " " + Trim(Str(RivalViolations(FormIndex, RivalIndex, ConstraintIndex)));
                        Next ConstraintIndex
                        Print #PraatFile,
                    Next RivalIndex
            Next FormIndex
            
        'Close the file.
            Close #PraatFile
        
    'Create xxx.PairDistribution:
            
        'Backups go in the downstairs folder; straightforward saves go in the upstairs folder.
            If LCase(Right(MyFileName, 6)) = "backup" Then
                Open gOutputFilePath + "/" + MyFileName + "ForPraat.PairDistribution" For Output As #PraatFile
            Else
                Open gOutputFilePath + "/" + MyFileName + "ForPraat.PairDistribution.txt" For Output As #PraatFile
            End If
        
        'Print stuff at top of file
            Print #PraatFile, "File type = " + Chr(34) + "ooTextFile" + Chr(34)
            Print #PraatFile, "Object class = " + Chr(34) + "PairDistribution" + Chr(34)
            Print #PraatFile,
            
        'Compute the total number of candidates.
            For FormIndex = 1 To NumberOfForms
                'One for the winner.
                    Let TotalNumberOfCandidates = TotalNumberOfCandidates + 1
                'One for each rival.
                    For RivalIndex = 1 To NumberOfRivals(FormIndex)
                        Let TotalNumberOfCandidates = TotalNumberOfCandidates + 1
                    Next RivalIndex
            Next FormIndex
            
        'Print this information.
            Print #PraatFile, "pairs: size = "; Trim(Str(TotalNumberOfCandidates))
            
        'Now go through all inputs and rival sets, augmenting the count with PairIndex and printing the input-output "pairs".
            For FormIndex = 1 To NumberOfForms
                'The winner.
                    Let PairIndex = PairIndex + 1
                    Print #PraatFile, "pairs ["; Trim(Str(PairIndex)); "];"
                    Print #PraatFile, "    string1 = " + Chr(34) + InputForm(FormIndex) + Chr(34)
                    Print #PraatFile, "    string2 = " + Chr(34) + Winner(FormIndex) + Chr(34)
                    Print #PraatFile, "    weight = ";
                    'Praat can't handle decimals properly, so append a zero to values between 0 and 1.
                        If WinnerFrequency(FormIndex) > 0 And WinnerFrequency(FormIndex) < 1 Then
                            Print #PraatFile, "0";
                        End If
                    Print #PraatFile, Trim(Str(WinnerFrequency(FormIndex)))
                'The rivals.
                    For RivalIndex = 1 To NumberOfRivals(FormIndex)
                        Let PairIndex = PairIndex + 1
                        Print #PraatFile, "pairs ["; Trim(Str(PairIndex)); "];"
                        Print #PraatFile, "    string1 = " + Chr(34) + InputForm(FormIndex) + Chr(34)
                        Print #PraatFile, "    string2 = " + Chr(34) + Rival(FormIndex, RivalIndex) + Chr(34)
                        Print #PraatFile, "    weight = ";
                        'Praat can't handle decimals properly, so append a zero to values between 0 and 1.
                            If RivalFrequency(FormIndex, RivalIndex) > 0 And RivalFrequency(FormIndex, RivalIndex) < 1 Then
                                Print #PraatFile, "0";
                            End If
                        Print #PraatFile, Trim(Str(RivalFrequency(FormIndex, RivalIndex)))
                    Next RivalIndex
            Next FormIndex
            
        Close #PraatFile
        
        'Announce completion.
            MsgBox "Done!  I have created the two necessary Praat-format files, " + MyFileName + "ForPraat.OTGrammar and " + MyFileName + "ForPraat.PairDistribution.  They are in the folder " + gOutputFilePath + "."
        
    Exit Sub
    
CheckError:
    Select Case Err.Number  ' Evaluate error number.
        Case 70 ' "File already open" error.
            MsgBox "Error.  Probably what is happening is this:  I'm trying to open the file " + _
                gInputFilePath + ReportErrorPraatFileName + " for purposes of storing my results, but a file of this name is already open.  I suggest you try to find this file, close it, then click OK.", vbExclamation
            Resume
        Case 75 ' "File access error
            MsgBox "Error.  I conjecture that " + ReportErrorPraatFileName + " already exists in " + _
                gInputFilePath + " as a Read-Only file.  Try deleting this file (or right click, Properties, decheck Read-Only) and rerunning OTSoft.", vbExclamation
            End
        Case Else
            MsgBox "Program error.  You can ask for help at bhayes@humnet.ucla.edu.  Please send a copy of your input file with your message.", vbCritical
            End
    End Select

End Sub


Sub SaveAsR(MyFileName As String, SortByRank As Boolean, Winner() As String, NumberOfConstraints As Long, ConstraintName() As String, _
    NumberOfForms As Long, InputForm() As String, _
    NumberOfRivals() As Long, WinnerViolations() As Long, WinnerFrequency() As Single, Rival() As String, RivalViolations() As Long, _
    RivalFrequency() As Single)

    'Save the input file in a format that could be used for logistic regression in R.
    
    'On Error GoTo CheckError
    
    Dim TotalNumberOfCandidates As Long
    Dim FormIndex As Long, RivalIndex As Long, ConstraintIndex As Long, PairIndex As Long, CopyIndex As Long
    Dim PositionIndex As Long
    Dim ReportErrorFileName As String
    
    Dim CautionFlagAlreadyRaised As Boolean
    
    Dim RFile As Long
    Let RFile = FreeFile
    
    Call Form1.CreateAFolderForOutputFiles
    
    'Let mRFileFlag = True
    Call Form1.DigestTheInputFile(gInputFilePath, gFileName, gFileSuffix)  'KZ: DigestTheInputFile is now a boolean function
    'Let mRFileFlag = False

    Let ReportErrorFileName = MyFileName + ".txt"
    
        'Backups go in the downstairs folder; straightforward saves go in the upstairs folder.
            If LCase(Right(MyFileName, 6)) = "backup" Then
                Open gOutputFilePath + MyFileName + "ForR" For Output As #RFile
            Else
                Open gOutputFilePath + MyFileName + "ForR.txt" For Output As #RFile
            End If
        
        'Print stuff at top of file
            Print #RFile, "Input" + Chr(9) + "Winner" + Chr(9) + "WinnerIndex" + Chr(9) + "WhichToke" + Chr(9) + "HowManyTokens";
            For ConstraintIndex = 1 To NumberOfConstraints
                'Constraint names must be very conservative about what they include.
                Print #RFile, Chr(9); Cleaned(ConstraintName(ConstraintIndex));
            Next ConstraintIndex
            Print #RFile,
            
        'Inputs and candidates:
            For FormIndex = 1 To NumberOfForms
                'Make sure we're ok for logistic regression.
                    If NumberOfRivals(FormIndex) <> 1 Then
                        MsgBox "The form " + InputForm(FormIndex) + " has a nonbinary choice among candidates, so I cannot make a logistic regression file for you.  You can fix this by arranging an input file in which every input has two candidates, arranged in the same order for each input."
                        Close #RFile
                        Exit Sub
                    End If
                'Make sure the winner frequency is an integer.
                    If Int(WinnerFrequency(FormIndex)) <> WinnerFrequency(FormIndex) Then
                        MsgBox "The form " + InputForm(FormIndex) + " has a noninteger frequency for the first candidate.  Please readjust your input file so that all frequencies are integers."
                        Close #RFile
                        Exit Sub
                    End If
                'Make sure the rival frequency is an integer.
                    If Int(RivalFrequency(FormIndex, 1)) <> RivalFrequency(FormIndex, 1) Then
                        MsgBox "The form " + InputForm(FormIndex) + " has a noninteger frequency for the second candidate.  Please readjust your input file so that all frequencies are integers."
                        Close #RFile
                        Exit Sub
                    End If
                'Print as many copies as needed of the "winner" candidate, with violations.
                    For CopyIndex = 1 To WinnerFrequency(FormIndex)
                        Print #RFile, InputForm(FormIndex);
                        Print #RFile, Chr(9); Winner(FormIndex);
                        Print #RFile, Chr(9); "1";
                        Print #RFile, Chr(9); Trim(Str(CopyIndex));
                        Print #RFile, Chr(9); Trim(Str(WinnerFrequency(FormIndex)));
                        For ConstraintIndex = 1 To NumberOfConstraints
                            'Watch out for cases not appropriate for logistic regression.
                                If WinnerViolations(FormIndex, ConstraintIndex) > 0 And RivalViolations(FormIndex, 1, ConstraintIndex) > 0 Then
                                    If CautionFlagAlreadyRaised = False Then
                                        MsgBox "Caution:  some constraints are violated by both candidates; results will be unpredictable."
                                        Let CautionFlagAlreadyRaised = True
                                    End If
                                End If
                            Print #RFile, Chr(9); RivalViolations(FormIndex, 1, ConstraintIndex) - WinnerViolations(FormIndex, ConstraintIndex);
                        Next ConstraintIndex
                        Print #RFile,
                    Next CopyIndex
                'Print as many copies as needed of the "rival" candidate, with violations.
                    For CopyIndex = 1 To RivalFrequency(FormIndex, 1)
                        Print #RFile, InputForm(FormIndex);
                        Print #RFile, Chr(9); Rival(FormIndex, 1);
                        Print #RFile, Chr(9); "0";
                        'Number the copies.
                        Print #RFile, Chr(9); Trim(Str(CopyIndex));
                        Print #RFile, Chr(9); Trim(Str(RivalFrequency(FormIndex, 1)));
                        For ConstraintIndex = 1 To NumberOfConstraints
                            Print #RFile, Chr(9); RivalViolations(FormIndex, 1, ConstraintIndex) - WinnerViolations(FormIndex, ConstraintIndex);
                        Next ConstraintIndex
                        Print #RFile,
                    Next CopyIndex
            Next FormIndex
            
        'Close the file.
            Close #RFile
            
        'Announce completion.
            MsgBox "Done!  I have created a file in R-compatible format, for use in logistic regression.  The file name is " + MyFileName + "ForR.txt.  It is in the folder " + gOutputFilePath + "."
        
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


Function Cleaned(MyString As String) As String
    'Keep only letters and numbers in R headers
        Dim InBuffer As String, OutBuffer As String
        Dim i As Long
        Let InBuffer = MyString
        Select Case Left(InBuffer, 1)
            Case "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", " ", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"
                'do nothing
            Case Else
                Let InBuffer = "X"
        End Select
        For i = 1 To Len(InBuffer)
            Select Case Mid(InBuffer, i, 1)
                Case "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", " ", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", ".", "_", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
                    Let OutBuffer = OutBuffer + Mid(InBuffer, i, 1)
            End Select
        Next i
    If OutBuffer = "" Then
        MsgBox "I cannot make a legal variable name out of " + MyString + "; it needs to have some letters in it.  Please fix the file and rerun OTSoft."
        Close
        End
    Else
        Let Cleaned = OutBuffer
    End If
        
End Function

