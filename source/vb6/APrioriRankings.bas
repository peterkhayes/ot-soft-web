Attribute VB_Name = "APrioriRankings"
'=============================DEAL WITH A PRIORI RANKINGS==========================
'==================================================================================

Option Explicit
    
    
'Three routines:

    'PrintOutTemplateForAprioriRankings()                   for blank template
    'PrintAPrioriRankingsConvertedFromStrataAsTable()       for testing on broader data set
    'ReadAPrioriRankingsAsTable()                           read the input file
    '   calls:  Function FindAPrioriContradiction()         vets input file


Sub PrintOutTemplateForAPrioriRankings(NumberOfConstraints As Long, Abbrev() As String)

    'Print out a simple template file that the user can open to enter a priori rankings.
    
        Dim APrioriFile As Long     'Handle for output file.
        Dim i As Long, j As Long
    
    'First check if such a template already exists.
        If Dir(gOutputFilePath + gFileName + "APriori" + ".txt") <> "" Then
            'It already exists.  Try to open it.
                Call EditTemplateForAPrioriRankings(gOutputFilePath + gFileName + "APriori.txt")
        Else
            'You have to create a new template.
            
            'First, check if input file exists.
                If Dir(gInputFilePath + "\" + gFileName + gFileSuffix) = "" Then
                    MsgBox "I can't make an a priori rankings file, because I can't find the input file.  It is supposed to be located at:" + Chr(10) + Chr(10) + _
                        gInputFilePath + "\" + gFileName + gFileSuffix + Chr(10) + Chr(10) + _
                        "Try using the Work With Different File button to locate your input file.", vbExclamation
                    Exit Sub
                End If
            
            'Format:
                'First row:  gap, constraint names, in input file order
                'Subsequent rows:  constraint names, in input file order
                
            'Note that you're working.
                Let Form1.lblProgressWindow = "Working..."
                    
            'If you're going to do this, a file had better be open:
                If gHaveIOpenedTheFile = False Then
                    If Form1.DigestTheInputFile(gInputFilePath, gFileName, gFileSuffix) = False Then  'This calls the file opener.
                        Exit Sub    'KZ: false = input file couldn't be opened
                    End If
                End If
                Let gHaveIOpenedTheFile = True
        
            'Now that you have the constraints, open the template file:
                Let APrioriFile = FreeFile
                Open gOutputFilePath + gFileName + "APriori" + ".txt" For Output As #APrioriFile
                
                For i = 1 To NumberOfConstraints
                    Print #APrioriFile, Chr(9); Abbrev(i);
                Next i
                Print #APrioriFile,
                For i = 1 To NumberOfConstraints
                    Print #APrioriFile, Abbrev(i);
                    'A tab for each constraint:
                        For j = 1 To NumberOfConstraints
                            Print #APrioriFile, Chr(9);
                        Next j
                        Print #APrioriFile,
                Next i
                
                Close #APrioriFile
                
                Let Form1.lblProgressWindow = ""
                
            'Let the user edit the sheet if they wish.
                    Select Case MsgBox("Template for a priori rankings prepared. " + Chr(10) + Chr(10) + _
                        "You can find it at " + gOutputFilePath + gFileName + "APriori.txt." + Chr(10) + Chr(10) + "Add a " + _
                        "priori rankings with a spreadsheet program, inserting any nonnull " + _
                        "symbol to indicate that the column-label constraint must dominate the " + _
                        "row-label constraint." + Chr(10) + Chr(10) + "Click Yes to edit this file with your spreadsheet, no to return to OTSoft.", vbYesNo + vbInformation)
                            Case vbYes
                                Call EditTemplateForAPrioriRankings(gOutputFilePath + gFileName + "APriori.txt")
                            Case vbNo
                                'Do nothing.
                    End Select
                
        End If                  'Does a template already exist?
        
    Let Form1.mnuConstrainAlgorithmsByAPrioriRankings.Checked = True
        
CheckError:

    Exit Sub
        
End Sub


Sub EditTemplateForAPrioriRankings(MyFileName As String)

    'Edit the a priori rankings template--either just made, or modifying a previous version.
    
    On Error GoTo CheckError
    
        Dim Dummy As Long                   'To assign value of Shell() function.
    
    'First check if Excel is there.
        If Dir(gExcelLocation) <> "" Then
            'With Excel, if possible:
                Let Dummy = Shell(gExcelLocation + " " + Chr(34) + MyFileName + Chr(34), vbNormalFocus)
        Else
            'Else with whatever Windows says.
            '   Note:  chr(34), the quotation mark, needed around file names in Shell.
                UseWindowsPrograms.TryShellExecute (MyFileName)
        End If
    
        Exit Sub
        
CheckError:

    MsgBox "I'm having trouble editing the file for a priori rankings.  You may wish to try to open it outside OTSoft.  It can be found at:" + _
    Chr(10) + Chr(10) + _
    MyFileName + _
    Chr(10) + Chr(10) + _
    "Click OK to return to the main OTSoft screen.", vbExclamation

End Sub


Sub PrintAPrioriRankingsConvertedFromStrataAsTable(NumberOfStrata As Long, Stratum() As Long, NumberOfConstraints As Long, Abbrev() As String)

    'Save the a priori rankings, as derived by the algorithm, as a simple tabular text file,
    '  which the user can manipulate with Excel.
    
        Dim StratumIndex As Long, OuterStratumIndex As Long, InnerStratumIndex As Long
        Dim ConstraintIndex As Long, OuterConstraintIndex As Long, InnerConstraintIndex As Long
        
        Dim APrioriFile As Long
        Let APrioriFile = FreeFile
        
        Open gOutputFilePath + gFileName + "APriori" + ".txt" For Output As #APrioriFile
        
    'Print headings for dominated constraint, in descending stratum order.
        For StratumIndex = 1 To NumberOfStrata
            For ConstraintIndex = 1 To NumberOfConstraints
                If Stratum(ConstraintIndex) = StratumIndex Then
                    Print #APrioriFile, Chr(9); Abbrev(ConstraintIndex);  'Caution:  mabbrev not checked
                End If
            Next ConstraintIndex
        Next StratumIndex
        Print #APrioriFile,
        
    'Print row headings for dominating constraint, then a "x" for each domination
    '  relation.  Follow descending rank order of strata.
        
        For OuterStratumIndex = 1 To NumberOfStrata
            For OuterConstraintIndex = 1 To NumberOfConstraints
                If Stratum(OuterConstraintIndex) = OuterStratumIndex Then
                    Print #APrioriFile, Abbrev(OuterConstraintIndex);
                    For InnerStratumIndex = 1 To NumberOfStrata
                        For InnerConstraintIndex = 1 To NumberOfConstraints
                            If Stratum(InnerConstraintIndex) = InnerStratumIndex Then
                                If Stratum(OuterConstraintIndex) < Stratum(InnerConstraintIndex) Then
                                    Print #APrioriFile, Chr(9); "1";
                                Else
                                    Print #APrioriFile, Chr(9);
                                End If
                            End If
                        Next InnerConstraintIndex
                    Next InnerStratumIndex
                    Print #APrioriFile,
                End If
            Next OuterConstraintIndex
        Next OuterStratumIndex
        
        Close #APrioriFile

End Sub


Public Function ReadAPrioriRankingsAsTable(NumberOfConstraints As Long, Abbrev() As String) As Boolean

    'Read an priori rankings table.  See routine immediately above for format.
    
    On Error GoTo CheckError
        
        Dim LocalNumberOfConstraints As Long
        Dim LocalAbbrev() As String
        Dim MyLine As String
        Dim MyCell As String
        Dim Buffer As String
        Dim ConstraintIndex As Long, OuterConstraintIndex As Long, InnerConstraintIndex As Long
        Dim i As Long, j As Long
        
        Dim APFile As Long
        Let APFile = FreeFile
        
        'Never open a file that might not exist.
            If Dir(gOutputFilePath + gFileName + "APriori.txt") <> "" Then
                Open gOutputFilePath + gFileName + "APriori.txt" For Input As #APFile
            Else
                MsgBox "Error:  you're trying to do ranking or factorial typology without having first created a file of a priori rankings." + Chr(10) + Chr(10) + _
                    "To create such a file, select from the A Priori Rankings menu the item Make a New File Formatted for Entering A Priori Rankings" + _
                    "and edit this file with a spreadsheet.  For now, I will proceed without a priori rankings.", vbExclamation
                Let ReadAPrioriRankingsAsTable = False
                Exit Function
            End If
    
    'Read through the top line of this file, extracting the constraint names
    '  and storing them in LocalAbbrev().  These will then be checked to make
    '  sure they match the constraint names of this simulation.
    
        Line Input #APFile, MyLine
        'There is nothing in the upper left corner.
            Let MyLine = s.Residue(MyLine)
            'Check that you've got the right constraints in the column labels.
                For i = 1 To NumberOfConstraints
                        Let MyCell = s.Chomp(MyLine)
                        'MsgBox "Excel:  xxx" + Abbrev(i) + "xxx a priori:  xxx" + MyCell + "xxx"
                        'Dim f As Long
                        'Let f = FreeFile
                        'Open App.Path + "/DebugHorribleThing.txt" For Output As #f
                        'Print #f, Abbrev(i); Chr(9); "Excel"
                        'Print #f, MyCell; Chr(9); "a priori"
                        'Close #f
                        'Stop
                        If MyCell <> Abbrev(i) Then
                            MsgBox "Sorry, I can't do a priori rankings.  In your a priori rankings " + _
                            "file, " + gFileName + "APriori.txt, column " + Str(i + 1) + ", you have " + _
                            "the constraint name " + MyCell + ", but in the input file " + _
                            gFileName + gFileSuffix + ", the constraint that appears in this order is " + _
                            Abbrev(i) + ". Please fix the file " + gFileName + "APriori.txt before " + _
                            "proceeding further.  [Additional note 1/25/13:  trying running your file from a tabbed text format input file; this seems to be a hard (to me) problem caused by changes in Windows.]", vbExclamation
                            Exit Function
                        End If
                        Let MyLine = s.Residue(MyLine)
                Next i
            'Now read the rows and the a priori rankings.
                For i = 1 To NumberOfConstraints
                    Line Input #APFile, MyLine
                    Let MyCell = s.Chomp(MyLine)
                    'Make sure the row headers are correct:
                        If MyCell <> Abbrev(i) Then
                            MsgBox "Sorry, I can't do a priori rankings.  In your a priori rankings " + _
                            "file, " + gFileName + "APriori.txt, row " + Str(i + 1) + ", you have " + _
                            "the constraint name " + MyCell + ", but in the input file " + _
                            gFileName + gFileSuffix + ", the constraint that appears in this order is " + _
                            Abbrev(i) + ". Please fix the file " + gFileName + "APriori.txt if you want " + _
                            "a priori rankings; for now, I will proceed without it.", vbExclamation
                            Let ReadAPrioriRankingsAsTable = False
                            Exit Function
                        End If
                    'Now, read in the apriori rankings.
                        Let MyLine = s.Residue(MyLine)
                        For j = 1 To NumberOfConstraints
                            If s.Chomp(MyLine) <> "" Then Let gAPrioriRankingsTable(i, j) = True
                            Let MyLine = s.Residue(MyLine)
                        Next j
                Next i
                
        Close #APFile
        
    'Conduct some checks before you use this file.
    
        'No constraint can a priori dominate itself.
            For ConstraintIndex = 1 To NumberOfConstraints
                If gAPrioriRankingsTable(ConstraintIndex, ConstraintIndex) = True Then
                    MsgBox "Error:  your a priori rankings table, " + _
                        gFileName + "APriori.txt, posits a constraint, " + _
                        Abbrev(ConstraintIndex) + ", that a priori dominates itself. " + _
                        "Please fix this problem if you want a priori rankings; for now, I will proceed without them.", vbExclamation
                        Let ReadAPrioriRankingsAsTable = False
                        Exit Function
                End If
            Next ConstraintIndex
            
        'No two constraints can a priori dominate each other.
            For OuterConstraintIndex = 1 To NumberOfConstraints
                For InnerConstraintIndex = 1 To NumberOfConstraints
                    If gAPrioriRankingsTable(OuterConstraintIndex, InnerConstraintIndex) = True Then
                        If gAPrioriRankingsTable(InnerConstraintIndex, OuterConstraintIndex) = True Then
                            MsgBox "Error:  in your a priori rankings table, " + _
                                gFileName + "APriori.txt, the constraint " + _
                                Abbrev(OuterConstraintIndex) + " is posited to dominate the constraint " + _
                                Abbrev(InnerConstraintIndex) + " and vice versa, which is impossible. " + _
                                "Please fix this problem if you want a priori rankings; for now, I will proceed without them.", vbExclamation
                            Let ReadAPrioriRankingsAsTable = False
                            Exit Function
                        End If
                    End If
                Next InnerConstraintIndex
            Next OuterConstraintIndex
            
        'Nor can there be more extensive contradictions.
            If FindAPrioriContradiction(gAPrioriRankingsTable(), NumberOfConstraints, Abbrev()) = True Then
                MsgBox "Error:  in your a priori rankings table, " + _
                    gFileName + "APriori.txt, there is a contradiction, i.e. a chain of " + _
                    "domination relations that forms a loop.  To help find this loop, " + _
                    "go to the directory " + gOutputFilePath + ", and open the file called " + _
                    "BadLoopIn" + gFileName + "Priori.txt.  Fix the loop, and restart OTSoft.", vbExclamation
                End

            End If
        
        'All is well, so return True:
            Let ReadAPrioriRankingsAsTable = True

    
    'Debug this routine:
    '    Dim D As Long
    '    Let D = FreeFile
    '    Open gOutputFilePath + "DebugAPriori.txt" For Output As #D
    '    For i = 1 To NumberOfConstraints
    '        Print #D, Chr(9); Abbrev(i);
    '    Next i
    '    Print #D,
    '    For i = 1 To NumberOfConstraints
    '        Print #D, Abbrev(i);
    '        For j = 1 To NumberOfConstraints
    '            If gAPrioriRankingsTable(i, j) = True Then
    '                Print #D, Chr(9); "x";
    '            Else
    '                Print #D, Chr(9);
    '            End If
    '        Next j
    '        Print #D,
    '    Next i
    '
    '    Exit Sub
        
CheckError:

    Select Case Err.Number  ' Evaluate error number.
        Case 53             'A priori rankings file doesn't exist.
            MsgBox "Error:  I conjecture that you're trying to do ranking or factorial typology without having first created a file of a priori rankings." + Chr(10) + Chr(10) + _
                "The file that you need is called " + gFileName + "APriori.txt, and should be located in the same folder as your input file." + Chr(10) + Chr(10) + _
                "If the file does not exist, and you want to create it, select from the A Priori Rankings menu the item Make a New File Formatted for Entering A Priori Rankings" + _
                "and edit this file with a spreadsheet.", vbExclamation
                Exit Function
    End Select

    
End Function

Function FindAPrioriContradiction(MyAprioriRankingTable() As Boolean, NumberOfConstraints As Long, _
    Abbrev() As String) As Boolean

    'Inspect an apriori ranking table and determine whether it is free of
    '   contradiction.  This can be done by using Constraint Demotion,
    '   making use only of the a priori rankings to form the strata.
    
    Dim LocalStratum() As Long
    ReDim LocalStratum(NumberOfConstraints)
    Dim CurrentStratum As Long
    Dim StratumIndex As Long
    Dim Demotable() As Boolean
    ReDim Demotable(NumberOfConstraints)
    Dim i As Long, j As Long
    Dim ICanHaltNow As Boolean, IHaveCrashed As Boolean
    
    'Initialize the strata.
        For i = 1 To NumberOfConstraints
            Let LocalStratum(i) = 0
        Next i
        Let CurrentStratum = 0
    
    'Loop, creating strata.
        Do
        
            Let CurrentStratum = CurrentStratum + 1
            
            'Initialize Demotable.
                For i = 1 To NumberOfConstraints
                    Let Demotable(i) = False
                Next i
                
            'Find who should not be in this stratum.
                For i = 1 To NumberOfConstraints
                    'Only unranked constraints forced their dominees out of the next stratum.
                    If LocalStratum(i) = 0 Then
                        'Who should be forced out of stratum?
                            For j = 1 To NumberOfConstraints
                                'Only unranked constraints qualified for demotion.
                                    If LocalStratum(j) = 0 Then
                                        If MyAprioriRankingTable(i, j) = True Then
                                            Let Demotable(j) = True
                                        End If
                                    End If
                            Next j
                    End If
                Next i
                
            'Inspect what you've got.
                Let ICanHaltNow = True
                Let IHaveCrashed = True
                For i = 1 To NumberOfConstraints
                    If LocalStratum(i) = 0 Then
                        If Demotable(i) = True Then
                            Let ICanHaltNow = False
                        Else
                            Let IHaveCrashed = False
                            Let LocalStratum(i) = CurrentStratum
                        End If
                    End If
                Next i
        
            'If nobody is demotable, then the system is consistent and you can quit.
                If ICanHaltNow = True Then
                    Let FindAPrioriContradiction = False
                    Exit Function
                End If
                
            'If everybody is demotable, then report the contradiction.
                If IHaveCrashed = True Then
                    Let FindAPrioriContradiction = True
                    
                    'First, make sure there is a folder for these files, a daughter of the
                    '   folder in which the input file is located.
                        Call Form1.CreateAFolderForOutputFiles

                    Dim CF As Long
                    Let CF = FreeFile
                    Open gOutputFilePath + "BadLoopIn" + gFileName + "APriori.txt" For Output As #CF
                    
                    Print #CF, "Error:  in your a priori rankings table, " + _
                        gOutputFilePath + gFileName + "APriori.txt, there is a contradiction, specifically a chain of " + _
                        "domination relations that forms a loop.  To help find this loop, " + _
                        "I have carried out Constraint Demotion, relying solely on the " + _
                        "a priori rankings provided in your file.  The contradiction is located " + _
                        "somewhere in the bottom stratum below."
                    Print #CF,
                    
                    'Print all the strata that actually succeeded in being created.
                        For StratumIndex = 1 To CurrentStratum - 1   'Last batch is still in Stratum 0
                            Print #CF, "Stratum #"; Trim(StratumIndex); ":"
                            For i = 1 To NumberOfConstraints
                                If LocalStratum(i) = StratumIndex Then
                                    Print #CF, "   "; Abbrev(i)
                                End If
                            Next i
                            Print #CF,
                        Next StratumIndex
                        
                    'Print the bad last stratum.
                            Print #CF, "Stratum #"; Trim(CurrentStratum); " (contains contradiction):"
                            For i = 1 To NumberOfConstraints
                                If LocalStratum(i) = 0 Then
                                    Print #CF, "   "; Abbrev(i)
                                End If
                            Next i
                            Print #CF,
                            
                    'Print a list of the possible culprits:
                        Print #CF,
                        Print #CF, "The following list of a priori rankings from your file can be inspected to locate the contradiction:"
                        Print #CF,
                        For i = 1 To NumberOfConstraints
                            If LocalStratum(i) = 0 Then
                            For j = 1 To NumberOfConstraints
                                If LocalStratum(j) = 0 Then
                                    If MyAprioriRankingTable(i, j) = True Then
                                        Print #CF, "   "; Abbrev(i); " >> "; Abbrev(j)
                                    End If
                                End If
                            Next j
                            End If
                        Next i
                        
                        Close #CF
                        
                    Exit Function
                
                End If              'End of code for what to do in a crash.
                
            'If you've gotten this far, you're still looking for a contradiction.
        
        Loop
        
    'You should never get this far.
        MsgBox "Program error.  Please contact bhayes@humnet.ucla.edu, including a copy of your input file and specifying error #33488.", vbCritical

End Function
