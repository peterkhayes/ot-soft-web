Attribute VB_Name = "s"
Public Function TabTrim(MyLine As String) As String
    Dim Buffer As String
    Let Buffer = MyLine
Restart1:
    If Right(Buffer, 1) = 1 Then
        Let Buffer = Left(Buffer, Len(Buffer) - 1)
        GoTo Restart1
    End If
Restart2:
    If Left(Buffer, 1) = 1 Then
        Let Buffer = Mid(Buffer, 2)
        GoTo Restart2
    End If
    Let TabTrim = Buffer
End Function

Public Function Chomp(MyString) As String
    Dim i As Long
    For i = 1 To Len(MyString)
        If Mid(MyString, i, 1) = Chr(9) Then
            Let Chomp = Left(MyString, i - 1)
            Exit Function
        End If
    Next i
    Let Chomp = MyString
End Function
Public Function Residue(MyString) As String
    Dim i As Long
    For i = 1 To Len(MyString)
        If Mid(MyString, i, 1) = Chr(9) Then
            Let Residue = Mid(MyString, i + 1)
            Exit Function
        End If
    Next i
    Let Residue = ""
End Function

Public Function SlashFinalPath(MyString) As String
    'Make sure a path ends in \.
        Let MyString = Trim(MyString)
        Select Case Right(MyString, 1)
            Case "\", "/"
                Let SlashFinalPath = MyString
            Case Else
                Let SlashFinalPath = MyString + "\"
        End Select
End Function

Public Function NoSlashFinalPath(MyString) As String
    'Make sure a path does *not* end in \ or /.
        Let MyString = Trim(MyString)
        Select Case Right(MyString, 1)
            Case "\", "/"
                Let NoSlashFinalPath = Left(MyString, Len(MyString) - 1)
            Case Else
                Let NoSlashFinalPath = MyString
        End Select
End Function

Public Sub CheckThatFileIsPresent(MyFile As String, Description As String)

    'Warn, and abort, if a crucial file is absent.
        If Dir(Trim(MyFile)) = "" Then
            MsgBox "Problem:  you've specified " + Description + " in the location " + _
            Chr(10) + Chr(10) + _
            mLocationOfInputFiles + MyFile + _
            Chr(10) + Chr(10) + _
            "but I can't find this file.  Kindly correct this problem and then start again.  When you click OK, the program will exit.", vbCritical
            End
        End If

End Sub

Public Function IsAnInteger(MyString As String) As Boolean
    Dim Buffer As String, i As Long
    Let Buffer = Trim(MyString)
    For i = 1 To Len(Buffer)
        Select Case Mid(Buffer, i, 1)
            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                'do nothing
            Case Else
                Let IsAnInteger = False
                Exit Function
        End Select
    Next i
    Let IsAnInteger = True
    
End Function

Public Function CustomChomp(MyString As String, Delimitor As String) As String
    'Return the portion of MyString that precedes Delimitor.
    Dim i As Long
    For i = 1 To Len(MyString) - Len(Delimitor) + 1
        If Mid(MyString, i, Len(Delimitor)) = Delimitor Then
            Let CustomChomp = Left(MyString, i - 1)
            Exit Function
        End If
    Next i
    Let CustomChomp = MyString
End Function
Public Function CustomResidue(MyString As String, Delimitor As String) As String
    'Return the portion of MyString that follows Delimitor.
    Dim i As Long
    For i = 1 To Len(MyString) - Len(Delimitor) + 1
        If Mid(MyString, i, Len(Delimitor)) = Delimitor Then
            Let CustomResidue = Mid(MyString, i + Len(Delimitor))
            Exit Function
        End If
    Next i
    Let CustomResidue = ""
End Function

Public Function TrimToPath(MyPath)
    'Chop off an actual file name to get path.
    Dim i As Long
    For i = Len(MyPath) To 1 Step -1
        Select Case Mid(MyPath, i, 1)
            Case "/", "\"
                Let TrimToPath = Left(MyPath, i - 1)
                Exit Function
        End Select
    Next i
    Let TrimToPath = MyPath
End Function
Public Function TrimToFileName(MyPath)
    'Chop off a path to get the plain file name.
    Dim i As Long
    For i = Len(MyPath) To 1 Step -1
        Select Case Mid(MyPath, i, 1)
            Case "/", "\"
                Let TrimToFileName = Mid(MyPath, i + 1)
                Exit Function
        End Select
    Next i
    Let TrimToFileName = MyPath
End Function

Public Function IsPositiveReal(MyString As String) As Boolean

    'Check that a string is a positive real number.
    
    Dim i As Long, NumberOfDecimalPoints As Long
    
    For i = 1 To Len(MyString)
        Select Case Mid(MyString, i, 1)
            Case "."
                Let NumberOfDecimalPoints = NumberOfDecimalPoints + 1
            Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
                'Do nothing.
            Case Else
                Let IsPositiveReal = False
                Exit Function
        End Select
    Next i
    If NumberOfDecimalPoints <= 1 Then Let IsPositiveReal = True

End Function

Public Sub UnloadAllForms(Optional FormToIgnore As String = "")
  
  'From http://techrepublic.com.com/5100-3513_11-5533338.html#.
    'Try to unload all the forms so that you don't get crap in memory after
    '   execution--said crap is making it hard to copy and delete versions
    '   of the code.

  Dim f As Form
  For Each f In Forms
    If f.Name <> FormToIgnore Then
      Unload f
      Set f = Nothing
    End If
  Next f
    
End Sub
 

Public Function FileDate(FileSpec As String)
    
    On Error GoTo CheckError
    
    'From the VB6 help file.  Other stuff is there for perhaps future use.
        Dim fs, f, s
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set f = fs.GetFile(FileSpec)
        s = UCase(FileSpec) & vbCrLf
        s = s & "Created: " & f.DateCreated & vbCrLf
        s = s & "Last Accessed: " & f.DateLastAccessed & vbCrLf
        s = s & "Last Modified: " & f.DateLastModified
        Let FileDate = f.DateLastModified
        Exit Function
        
CheckError:
    Let FileDate = "unknown"
    Exit Function
    
End Function

Function NumberContainedIn(MyString As String, MyCharacter As String) As Long

    'How many spaces, hyphens, etc. does a string contain?
    
    Dim i As Long
    
    For i = 1 To Len(MyString)
        If Mid(MyString, i, 1) = MyCharacter Then
            Let NumberContainedIn = NumberContainedIn + 1
        End If
    Next i

End Function

Function ZeroToNull(MyNumber) As String

    If MyNumber = 0 Then
        Let ZeroToNull = ""
    Else
        Let ZeroToNull = Str(MyNumber)
    End If
    
End Function

Public Function RightJustifiedFill(MyString As Variant, MyLength As Long) As String

    'Return a string that is as many spaces as is necessary to fill out the input string
    '   to the specified length.  This can then be concatenated to produce pretty
    '   justification.
    
    'This version adds the spaces on the left, to produce right justification.
    
    Dim Base As String
    Dim Buffer As String
    Dim i As Long
    
    Let Base = Trim(MyString)
    Let Buffer = MyString
    For i = Len(MyString) + 1 To MyLength
        Let Buffer = " " + Buffer
    Next i
    Let RightJustifiedFill = Buffer

End Function

Public Function PrintTable(DocFile As Long, TmpFile As Long, HTMFile As Long, MyTable() As String, BoldTopRow As Boolean, BoldLeftColumn As Boolean, CenterCells As Boolean)

    If DocFile > 0 Then
        Call PrintDocTable(MyTable(), DocFile, BoldTopRow, BoldLeftColumn, CenterCells)
    End If
    
    If HTMFile > 0 Then
        Call PrintHTMTable(MyTable(), HTMFile, BoldTopRow, BoldLeftColumn, CenterCells)
    End If

    'This one must come last because it erases the shading diacritics.
        If TmpFile > 0 Then
            Call PrintTextTable(MyTable(), TmpFile, CenterCells)
        End If
    
End Function

Public Function PrintHTMTable(MyTable() As String, FileNum As Long, BoldTopRow As Boolean, BoldLeftColumn As Boolean, CenterCells As Boolean)

    'Presupposes that the file is already open, and that gHTMLTableSpecs (Module1) has the intended border width, etc.
    'The array MyTable() has two indices in intuitively-reversed order:  column followed by row.  This
    '   is to permit redimensioning in routines that call this one.
    
    Dim RowIndex As Long, ColumnIndex As Long
    Dim BoldfaceFlag As Boolean
    
    
    'Table start, with border width and spacing.
        Print #FileNum, "<p>"
        Print #FileNum, gHTMLTableSpecs
        
    'Go through the rows.
        For RowIndex = 1 To UBound(MyTable, 2)
            'Print an html row start.
                Print #FileNum, "<tr>"
            'Go through the columns
                For ColumnIndex = 1 To UBound(MyTable, 1)
                    'Print an entry start, according to boldfacing diacritics
                        If (BoldTopRow And RowIndex = 1) Or (BoldLeftColumn And ColumnIndex = 1) Then
                            Let BoldfaceFlag = True
                        Else
                            Let BoldfaceFlag = False
                        End If
                        Print #FileNum, "<td";
                    'Centering--only noninitial columns, and if this was requested.
                        If ColumnIndex > 1 Then
                            If CenterCells Then
                                Print #FileNum, " align=" + Chr(34) + "center" + Chr(34)
                            End If
                        End If
                    'Shading:  add the code and delete the diacritic /sh.
                        If Right(MyTable(ColumnIndex, RowIndex), 3) = "\sh" Then
                            Print #FileNum, " bgcolor=#" + gShadingColor
                            Let MyTable(ColumnIndex, RowIndex) = Left(MyTable(ColumnIndex, RowIndex), Len(MyTable(ColumnIndex, RowIndex)) - 3)
                        End If
                    'Finish the cell-start code.
                        Print #FileNum, ">"
                    'Handle boldfacing.
                        If BoldfaceFlag Then
                            Print #FileNum, "<b>"
                        End If
                        
                    'Print the entry.  Do a non-breaking space if null, to produce the borders.
                        If MyTable(ColumnIndex, RowIndex) = "" Then Let MyTable(ColumnIndex, RowIndex) = "&nbsp;"
                        Print #FileNum, MyTable(ColumnIndex, RowIndex)
                    'Print an entry end
                        If BoldfaceFlag Then
                            Print #FileNum, "</b></td>"
                        Else
                            Print #FileNum, "</td>"
                        End If
                Next ColumnIndex
            'Print an html row end.
                Print #FileNum, "</tr>"
        Next RowIndex
        
    'Print a table conclusion and line break.
        Print #FileNum, "</table>"
        Print #FileNum, "<p>"

End Function

Public Sub PrintDocTable(MyTable() As String, FileNum As Long, BoldTopRow As Boolean, BoldLeftColumn As Boolean, CenterCells As Boolean)

    Dim NumberOfRows As Long, NumberOfColumns As Long
    Dim RowIndex As Long, ColumnIndex As Long
    
    Let NumberOfColumns = UBound(MyTable(), 1)
    Let NumberOfRows = UBound(MyTable(), 2)
    
    
    'Keep with next.
        Print #FileNum, "\ks"
    'Marker for start table
        Print #FileNum, "\ts";
    'Number of columns
        Print #FileNum, Trim(Str(NumberOfColumns))
    'First Row
        'First cell
            If BoldTopRow Or BoldLeftColumn Then
                Print #FileNum, "\nt";
            End If
            Print #FileNum, MyTable(1, 1);
        'Other cells
            For ColumnIndex = 2 To NumberOfColumns
                Print #FileNum, Chr(9); MyTable(ColumnIndex, 1);
            Next ColumnIndex
            Print #FileNum,
    'remaining rows
        For RowIndex = 2 To NumberOfRows
            'First cell
                If BoldLeftColumn Then
                    Print #FileNum, "\nt"
                End If
                Print #FileNum, MyTable(1, RowIndex);
            'Other cells
                For ColumnIndex = 2 To NumberOfColumns
                    Print #FileNum, Chr(9); MyTable(ColumnIndex, RowIndex);
                Next ColumnIndex
                Print #FileNum,
        Next RowIndex
    'Marker for table end
        Print #FileNum, "\te";
    'Marked for keep next
        Print #FileNum, "\ke"


End Sub

Public Sub PrintTextTable(MyTable() As String, FileNum As Long, CenterCells As Boolean)

    Dim NumberOfRows As Long, NumberOfColumns As Long
    Dim RowIndex As Long, ColumnIndex As Long
    Dim ColumnWidths() As Long
    
    Let NumberOfColumns = UBound(MyTable(), 1)
    Let NumberOfRows = UBound(MyTable(), 2)
    ReDim ColumnWidths(NumberOfColumns)
    
    'Remove the shading diacritic; you can't shade plain text.
        For ColumnIndex = 1 To NumberOfColumns
            For RowIndex = 1 To NumberOfRows
                Let MyTable(ColumnIndex, RowIndex) = Replace(MyTable(ColumnIndex, RowIndex), "\sh", "")
                Let MyTable(ColumnIndex, RowIndex) = Replace(MyTable(ColumnIndex, RowIndex), "&nbsp;", " ")
            Next RowIndex
        Next ColumnIndex
    
    'Establish the column widths.
        For ColumnIndex = 1 To NumberOfColumns
            For RowIndex = 1 To NumberOfRows
                If Len(MyTable(ColumnIndex, RowIndex)) > ColumnWidths(ColumnIndex) Then
                    Let ColumnWidths(ColumnIndex) = Len(MyTable(ColumnIndex, RowIndex))
                End If
            Next RowIndex
        Next ColumnIndex
    
    'Content
        For RowIndex = 1 To NumberOfRows
            'First column, never centered.
                Print #FileNum, Fillout(MyTable(1, RowIndex), ColumnWidths(1));
            'Remaining columns
                For ColumnIndex = 2 To NumberOfColumns
                    If CenterCells Then
                        Print #FileNum, "  " + CenteredFillout(MyTable(ColumnIndex, RowIndex), ColumnWidths(ColumnIndex));
                    Else
                        Print #FileNum, "  " + Fillout(MyTable(ColumnIndex, RowIndex), ColumnWidths(ColumnIndex));
                    End If
                Next ColumnIndex
            Print #FileNum,
        Next RowIndex
        'A blank line.
            Print #FileNum,

End Sub

Function Fillout(MyString As String, MyLength As Long)
    
    'Add trailing spaces to create a string of specified length.
        Dim i As Long, Buffer As String
        Let Buffer = MyString
        For i = Len(Buffer) + 1 To MyLength
            Let Buffer = Buffer + " "
        Next i
        Let Fillout = Buffer

End Function
Function CenteredFillout(MyString As String, MyLength As Long)
    
    'Add spaces to center as well as possible, favoring leftward error in uneven cases.
        Dim i As Long, Buffer As String, StringLength As Long, TotalSpaces As Long
        
        Let StringLength = Len(MyString)
        Let TotalSpaces = MyLength - StringLength
        If Int(TotalSpaces / 2) = TotalSpaces / 2 Then
            'even number of spaces
                For i = 1 To TotalSpaces / 2
                    Let Buffer = Buffer + " "
                Next i
                Let Buffer = Buffer + MyString
                For i = 1 To TotalSpaces / 2
                    Let Buffer = Buffer + " "
                Next i
        Else
            'Odd number of spaces
                For i = 1 To Int(TotalSpaces / 2)
                    Let Buffer = Buffer + " "
                Next i
                Let Buffer = Buffer + MyString
                For i = 1 To Int(TotalSpaces / 2) + 1
                    Let Buffer = Buffer + " "
                Next i
        End If
        Let CenteredFillout = Buffer

End Function

Public Function Longest(MyStringArray() As String) As Long
    'Length of longest string in a array.
        Dim Buffer As Long, i As Long
        For i = 1 To UBound(MyStringArray())
            If Len(MyStringArray(i)) > Buffer Then
                Let Buffer = Len(MyStringArray(i))
            End If
        Next i
        Let Longest = Buffer
End Function

Public Function DumbSym(Candidate As String) As String

    'Substitute ad hoc dumb symbols for real IPA, for appearance on the Windows screen.
    
        Dim b As String     '"buffer"
        Let b = ""
        Dim i As Long
        Dim UpperBound As Long
        
        'You only do this when the user has selected the IPA font.
        
        If SymbolTag1 = "\ss" Then
       
            Let UpperBound = Len(Candidate)
            For i = 1 To UpperBound
                'If Mid(Candidate, i, 1) = ")" Then Stop
                
                Select Case Mid(Candidate, i, 1)
                    Case Chr(34)
                        Let b = b + "1"
                    Case Chr(34)
                        Let b = b + "-"
                    Case Chr(41)
                        Let b = b + "~"
                    Case Chr(45)
                        Let b = b + "#"
                    Case Chr(47)
                        Let b = b + "?"
                    Case Chr(48)
                        Let b = b + "~"
                    Case Chr(58)
                        Let b = b + "l?"
                    Case Chr(59)
                        Let b = b + "L"
                    ' <, the null code, is Ø
                    Case Chr(60)
                        Let b = b + Chr(216)
                    Case Chr(62)
                        Let b = b + "."
                    Case Chr(63)
                        Let b = b + "?"
                    Case Chr(65)
                        Let b = b + "a"
                    Case Chr(66)
                        Let b = b + "B"
                    Case Chr(67)
                        Let b = b + "C"
                    Case Chr(68)
                        Let b = b + "D"
                    Case Chr(69)
                        Let b = b + "E"
                    Case Chr(70)
                        Let b = b + "G"
                    Case Chr(71)
                        Let b = b + "G"
                    Case Chr(72)
                        Let b = b + "H"
                    Case Chr(73)
                        Let b = b + "I"
                    Case Chr(74)
                        Let b = b + "J"
                    Case Chr(75)
                        Let b = b + "H"
                    Case Chr(76)
                        Let b = b + "L"
                    Case Chr(77)
                        Let b = b + "M"
                    Case Chr(78)
                        Let b = b + "N"
                    Case Chr(79)
                        Let b = b + "O"
                    Case Chr(80)
                        Let b = b + "O"
                    Case Chr(81)
                        Let b = b + "A"
                    Case Chr(82)
                        Let b = b + "R"
                    Case Chr(83)
                        Let b = b + "S"
                    Case Chr(84)
                        Let b = b + "T"
                    Case Chr(85)
                        Let b = b + "U"
                    Case Chr(86)
                        Let b = b + "V"
                    Case Chr(87)
                        Let b = b + "W"
                    Case Chr(88)
                        Let b = b + "X"
                    Case Chr(89)
                        Let b = b + "Y"
                    Case Chr(90)
                        Let b = b + "Z"
                    Case Chr(91)
                        Let b = b + "["
                    Case Chr(92)
                        Let b = b + "\"
                    Case Chr(93)
                        Let b = b + "]"
                    Case Chr(123)
                        Let b = b + "R"
                    Case Chr(124)
                        Let b = b + "'"
                    Case Chr(125)
                        Let b = b + "R"
                    Case Chr(128)
                        Let b = b + "LM"
                    Case Chr(129)
                        Let b = b + "O"
                    Case Chr(130)
                        Let b = b + "C"
                    Case Chr(132)
                        Let b = b + "||"
                    Case Chr(133)
                        Let b = b + "|"
                    Case Chr(134)
                        Let b = b + "|"
                    Case Chr(135)
                        Let b = b + "0"
                    Case Chr(138)
                        Let b = b + "]"
                    Case Chr(139)
                        Let b = b + "^"
                    Case Chr(140)
                        Let b = b + "A"
                    Case Chr(141)
                        Let b = b + "O"
                    Case Chr(142)
                        Let b = b + "|"
                    Case Chr(145)
                        Let b = b + "H"
                    Case Chr(146)
                        Let b = b + "|"
                    Case Chr(149)
                        Let b = b + "M"
                    Case Chr(150)
                        Let b = b + "|"
                    Case Chr(151)
                        Let b = b + "!"
                    Case Chr(154)
                        Let b = b + "L"
                    Case Chr(155)
                        Let b = b + "|"
                    Case Chr(156)
                        Let b = b + "+"
                    Case Chr(158)
                        Let b = b + "#"
                    Case Chr(159)
                        Let b = b + "]"
                    Case Chr(160)
                        Let b = b + "C"
                    Case Chr(167)
                        Let b = b + "S"
                    Case Chr(168)
                        Let b = b + "R"
                    Case Chr(169)
                        Let b = b + "G"
                    Case Chr(171)
                        Let b = b + "@"
                    Case Chr(172)
                        Let b = b + "U"
                    Case Chr(174)
                        Let b = b + "I"
                    Case Chr(175)
                        Let b = b + "O"
                    Case Chr(178)
                        Let b = b + "N"
                    Case Chr(179)
                        Let b = b + "?"
                    Case Chr(180)
                        Let b = b + "L"
                    Case Chr(181)
                        Let b = b + "U"
                    Case Chr(184)
                        Let b = b + "F"
                    Case Chr(185)
                        Let b = b + "P"
                    Case Chr(186)
                        Let b = b + "B"
                    Case Chr(189)
                        Let b = b + "Z"
                    Case Chr(190)
                        Let b = b + "J"
                    Case Chr(191)
                        Let b = b + "O"
                    Case Chr(192)
                        Let b = b + "?"
                    Case Chr(194)
                        Let b = b + "L"
                    Case Chr(195)
                        Let b = b + "^"
                    Case Chr(196)
                        Let b = b + "G"
                    Case Chr(198)
                        Let b = b + "J"
                    Case Chr(199)
                        Let b = b + ","
                    Case Chr(200)
                        Let b = b + "'"
                    Case Chr(201)
                        Let b = b + "L"
                    Case Chr(204)
                        Let b = b + "/"
                    Case Chr(205)
                        Let b = b + ">"
                    Case Chr(206)
                        Let b = b + "3"
                    Case Chr(207)
                        Let b = b + "Q"
                    Case Chr(210)
                        Let b = b + "R"
                    Case Chr(211)
                        Let b = b + "R"
                    Case Chr(212)
                        Let b = b + "R"
                    Case Chr(213)
                        Let b = b + "="
                    Case Chr(214)
                        Let b = b + "-"
                    Case Chr(215)
                        Let b = b + "J"
                    Case Chr(216)
                        Let b = b + "^"
                    Case Chr(217)
                        Let b = b + "V"
                    Case Chr(223)
                        Let b = b + "^"
                    Case Chr(226)
                        Let b = b + "~"
                    Case Chr(227)
                        Let b = b + "W"
                    Case Chr(228)
                        Let b = b + "L"
                    Case Chr(229)
                        Let b = b + "Y"
                    Case Chr(231)
                        Let b = b + "Y"
                    Case Chr(232)
                        Let b = b + "^"
                    Case Chr(234)
                        Let b = b + "D"
                    Case Chr(235)
                        Let b = b + "D"
                    Case Chr(236)
                        Let b = b + "U"
                    Case Chr(237)
                        Let b = b + "v"
                    Case Chr(238)
                        Let b = b + "H"
                    Case Chr(239)
                        Let b = b + "J"
                    Case Chr(240)
                        Let b = b + "H"
                    Case Chr(241)
                        Let b = b + "L"
                    Case Chr(245)
                        Let b = b + "B"
                    Case Chr(246)
                        Let b = b + "I"
                    Case Chr(247)
                        Let b = b + "N"
                    Case Chr(248)
                        Let b = b + "N"
                    Case Chr(249)
                        Let b = b + ":"
                    Case Chr(250)
                        Let b = b + "H"
                    Case Chr(251)
                        Let b = b + "K"
                    Case Chr(252)
                        Let b = b + "Z"
                    Case Chr(253)
                        Let b = b + "G"
                    Case Chr(254)
                        Let b = b + "C"
                    Case Chr(255)
                        Let b = b + "T"
                    Case Else
                        Let b = b + Mid(Candidate, i, 1)
                End Select
                Let DumbSym = b
            Next i
        Else
            Let DumbSym = Candidate
        End If
        
End Function

Public Sub PrintContentOfAnInputFile(AllRivals As Boolean, MyFileNumber As Long, NumberOfConstraints As Long, ConstraintName() As String, Abbrev() As String, _
    NumberOfForms As Long, InputForm() As String, Winner() As String, WinnerFrequency() As Single, WinnerViolations() As Long, NumberOfRivals() As Long, _
    Rival() As String, RivalFrequency() As Single, RivalViolations() As Long)

    'Print the content of an input file to whatever file is invoked calling this subroutine.
    
        Dim LowerLimit As Long
        Dim ConstraintIndex As Long, FormIndex As Long, RivalIndex As Long
        
    'Print a row of constraints:
        Print #MyFileNumber, Chr(9); Chr(9);
        For ConstraintIndex = 1 To NumberOfConstraints
            Print #MyFileNumber, Chr(9); ConstraintName(ConstraintIndex);
        Next ConstraintIndex
        Print #MyFileNumber,
        
    'Print a row of Abbreviations:
        Print #MyFileNumber, Chr(9); Chr(9);
        For ConstraintIndex = 1 To NumberOfConstraints
            Print #MyFileNumber, Chr(9); Abbrev(ConstraintIndex);
        Next ConstraintIndex
        Print #MyFileNumber,
        
    'Inputs, winners, rivals, violations:
        For FormIndex = 1 To NumberOfForms
            
            'The Input is always printed:
                Print #MyFileNumber, InputForm(FormIndex);
            
            'There is a separate Winner category only for certain algorithms.  The others
            '   use "nothing but rivals", with the winner folded into slot 0.
                If AllRivals = False Then
                    Print #MyFileNumber, Chr(9); Winner(FormIndex);
                    If WinnerFrequency(FormIndex) = 0 Then
                        Print #MyFileNumber, Chr(9);
                    Else
                        Print #MyFileNumber, Chr(9); WinnerFrequency(FormIndex);
                    End If
                
                    For ConstraintIndex = 1 To NumberOfConstraints
                        If WinnerViolations(FormIndex, ConstraintIndex) = 0 Then
                            Print #MyFileNumber, Chr(9);
                        Else
                            Print #MyFileNumber, Chr(9); WinnerViolations(FormIndex, ConstraintIndex);
                        End If
                    Next ConstraintIndex
                    Print #MyFileNumber,
                End If
            
            'If this is a system in which the "winners" have been folded into the Rival(0) slot, start at 0.
                If AllRivals = True Then
                    Let LowerLimit = 0
                Else
                    Let LowerLimit = 1
                End If
            
            'With this established, rival information is the same for all algorithms.
                For RivalIndex = LowerLimit To NumberOfRivals(FormIndex)
                    
                    Print #MyFileNumber, Chr(9); Rival(FormIndex, RivalIndex);
                    
                    If RivalFrequency(FormIndex, RivalIndex) = 0 Then
                        Print #MyFileNumber, Chr(9);
                    Else
                        Print #MyFileNumber, Chr(9); RivalFrequency(FormIndex, RivalIndex);
                    End If
                        
                        For ConstraintIndex = 1 To NumberOfConstraints
                            If RivalViolations(FormIndex, RivalIndex, ConstraintIndex) = 0 Then
                                Print #MyFileNumber, Chr(9);
                            Else
                                Print #MyFileNumber, Chr(9); RivalViolations(FormIndex, RivalIndex, ConstraintIndex);
                            End If
                        Next ConstraintIndex
                        Print #MyFileNumber,
                        
                Next RivalIndex
            Next FormIndex
    
End Sub

Public Sub p(MyDocFile As Long, MyTmpFile As Long, Optional LeadString As String, Optional MyLine As String)
                    
    'This saves many lines of code by printing the same message to two files.
        Print #MyDocFile, MyLine
        Print #MyTmpFile, LeadString + MyLine

End Sub
