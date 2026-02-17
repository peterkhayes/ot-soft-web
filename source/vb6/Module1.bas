Attribute VB_Name = "Module1"
    
'==============Material for OTSoft that has to be shared across forms/modules================

Option Explicit



Type gLikelihoodCalculation
    
    'User-defined variable type.
    'Purpose is to return both a value and a warning if zero probability for a winner had to be converted
    '   artificially to .001.
    
    LogLikelihood As Single
    IncludesAZeroProbability As Boolean

End Type


'----------------------------Information about the Real World---------------------------------
    
    'This information must be updated, so it is put all in one place.
    
    'List the version number of this edition on all windows.
        Public Const gMyVersionNumber As String = "2.7"
        
    'Debugging will be helped if we know the release date of this version
        Public Const gMyReleaseDate As String = "2/1/2026"
  
    'Where you can get needed software.  For error message boxes.
        Public Const gATTWebSite As String = "http://www.graphviz.org/"
        Public Const gJavaWebSite As String = "http://www.java.com/en/download/manual.jsp"
        
'-------------------------------------Shell Execute---------------------------------------------
    
    'This is used to run other Windows programs.
    'It was created by Omar Mallat, and was obtained from
    '   http://www.experts-exchange.com/Programming/Programming_Languages/Visual_Basic/Q_20272579.html
    
        Public Type ShellExecuteInfo
            cbSize As Long
            fMask As Long
            hwnd As Long
            lpVerb As String
            lpFile As String
            lpParameters As String
            lpDirectory As String
            nShow As Long
            hInstApp As Long
            lpIDList As Long
            lpClass As String
            hkeyClass As Long
            dwHotKey As Long
            hIcon As Long
            hProcess As Long
        End Type
        
        Public Declare Function ShellExecuteEx Lib "shell32.dll" Alias "ShellExecuteExA" (lpExecInfo As _
           ShellExecuteInfo) As Long
        Public Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal _
           dwMilliseconds As Long) As Long
        
        Public Const SEE_MASK_NOCLOSEPROCESS = &H40
        Public Const SW_SHOWNORMAL = 1
        Public Const SE_ERR_FNF = 2
        'I think Omar committed a typo:
            'Public Const SE_ERR_NOASSOC = 31
            Public Const SE_NOASSOC = 31
        Public Const INFINITE = &HFFFF
        Public Const WAIT_TIMEOUT = &H102
        Public Const gShellExecuteWasSuccessful As Long = 55946
        
'------------------------------------------------Global variables-----------------------------------------------------
    
    'Variables for files.
        Public gSafePlaceToWriteTo As String
        Public gFileName As String
        Public gFileSuffix As String
        Public gInputFilePath As String
        Public gOutputFilePath As String
        Public gLevel1HeadingNumber As Long    'For numbering the sections of the .tmp file.
        
    'Variables for programs called up by OTSoft.
        Public gUsersWordProcessor As String
        Public gExcelLocation As String
        Public gPaintLocation As String
        Public gDotExeLocation As String
        
    'Name of algorithm run
        Public gAlgorithmName As String
            
    'Variable to catch cases where people look for results when
    '   program hasn't yet been run.
        Public gHasTheProgramBeenRun As Boolean
        Public gHaveIOpenedTheFile As Boolean
    
    'Contents of the Excel spreadsheet used as input:
        Public gRawColumns() As String
        Public gTotalNumberOfRows As Long
    
    'Variables for a priori rankings:
        Public gAPrioriRankingsTable() As Boolean    'First index:  dominator.  Second index:  dominatee.
    
    'Variables for GLA.  These must be global so they can be saved from session to session.
        Public gNumberOfDataPresentations As Long
        Public gCyclesToTest As Long
        Public gNegativeWeightsOK As Boolean
        'Plasticity can be made different for Markedness and Faithfulness constraints.
            Public gCoarsestPlastMark As Single
            Public gFinestPlastMark As Single

    'Maxima (for array sizes)
        Public MaximumNumberOfConstraints As Long
        Public MaximumSizeOfFactorialTypology As Long
    
    'Variables for printing:
        Public SymbolTag1 As String
        Public SymbolTag2 As String
        Public SmallCapTag1 As String
        Public SmallCapTag2 As String
            
    'KZ: variables for default plasticities
        Public Const DefaultUpperFaith = 2
        Public Const DefaultLowerFaith = 0.002
        Public Const DefaultUpperMark = 0.2
        Public Const DefaultLowerMark = 0.002
    
    'KZ: default and custom noise
        Public CustomNoiseFaith() As Single
        Public CustomNoiseMark() As Single
        
    'KZ: default and custom initial rankings or weights
        Public InitialRankingChoice As Long
            Public Const AllSame As Long = 1
            Public Const MarkednessFaithfulness = 2
            Public Const FullyCustomized = 3
            Public Const ValuesFromPreviousRun = 4
            Public Const UseReverseReliability = 5
        Public blnCustomRankCreated As Boolean
        Public gCustomRankFaith As Single
        Public gCustomRankMark As Single
        Public IncludeTableauxInGLAOutput As Boolean
        Public gExactProportionsForGLAEtc As Boolean
        Public gMagriUpdateRuleInEffect As Boolean
        Public gReportingFrequency As Long
        
    'Variables for Noisy Harmonic Grammar
        Public gNHGLateNoise As Boolean
        Public gNHGNegativeWeightsOK As Boolean
        Public gNHGNoiseAppliesToTableauCells As Boolean
        Public gNHGNoiseForZeroCells As Boolean
        Public gNHGNoiseIsAddedAfterMultiplication As Boolean
        Public gExponentialNHG As Boolean
        Public gDemiGaussianNHG As Boolean
        Public gResolveTiesBySkipping As Boolean
    
    'Variables for preparing Hasse diagrams with ATT dot.exe
        Public HasseFile As Long
        Public HasseDiagramCreated As Boolean       'remembers if you've made one
        Public Const HasseFontSize As Long = 14
        
    'Constant for table creation
        Public gHTMLTableSpecs As String
        Public gShadingColor As String '  = "929292" '"737373"  '"F5F5F5"
        
    'Other variables on the interface
        Public TestWugOnly As Boolean

'-----------------------------------Printing Utilities-----------------------------------------

Public Sub PrintLevel1Header(DocFile As Long, TmpFile As Long, HTMFile As Long, MyText As String)

    'Augment the index for headers.
        Let gLevel1HeadingNumber = gLevel1HeadingNumber + 1
    'Temp file:
        If TmpFile > 0 Then
            Print #TmpFile,
            Print #TmpFile, Trim(gLevel1HeadingNumber); ". "; MyText
            Print #TmpFile,
        End If
    'Doc file:
        If DocFile > 0 Then
            Print #DocFile,
            Print #DocFile, "\h1"; MyText
        End If
    'HTML file:
        If HTMFile > 0 Then
            Print #HTMFile, "<p><b>"; Trim(gLevel1HeadingNumber); ". "; MyText; "</b><p><p>"
        End If
    
End Sub

Public Sub PrintTopLevelHeader(DocFile As Long, TmpFile As Long, HTMFile As Long, MyText As String)

    If TmpFile > 0 Then
        Print #TmpFile, MyText
        Print #TmpFile,
    End If
        
    If DocFile > 0 Then
        Print #DocFile, "\ti\ft" + MyText
        Print #DocFile,
        End If
    
    If HTMFile > 0 Then
        Print #HTMFile, "<b><big>" + MyText + "</big></b><p>"
    End If

End Sub


Public Sub PrintPara(DocFile As Long, TmpFile As Long, HTMFile As Long, MyText As String, Optional LeadSpaces As String)

    'The string PARA is used to split up lines in the primitive tmp file.
    '   It must be removed in the more sophisticated files.
    
    Dim NoPARA As String
    Dim i As Long
    Dim InBuffer As String, OutBuffer As String
    
    'Passing negative file numbers causes them to be skipped.
    
    'Temp:
        If TmpFile > 0 Then
            Let InBuffer = MyText
RestartPoint:
            For i = 1 To Len(InBuffer)
                If Mid(InBuffer, i, 4) = "PARA" Then
                    If Trim(OutBuffer) <> "" Then
                        Print #TmpFile, LeadSpaces + OutBuffer
                    End If
                    Let InBuffer = Mid(InBuffer, i + 4)
                    Let OutBuffer = ""
                    GoTo RestartPoint
                Else
                    Let OutBuffer = OutBuffer + Mid(InBuffer, i, 1)
                End If
            Next i
            If Trim(OutBuffer) <> "" Then
                Print #TmpFile, LeadSpaces + OutBuffer;
            End If
            Print #TmpFile,
            Print #TmpFile,
        End If
        
    'Define NoPARA, eliminating PARA recursively (I think it has to be so.)
        Let NoPARA = MyText
        For i = 1 To 10
            Let NoPARA = Replace(NoPARA, "PARA", "")
        Next i
    
    'Doc:
        If DocFile > 0 Then
            Print #DocFile, NoPARA
        End If
    'HTM:
        If HTMFile > 0 Then
            Print #HTMFile, "<p><p>" + NoPARA
        End If
        
End Sub

'-----------------------------------String Utilities-----------------------------------------



Public Function GoodNumber(ByVal Text As String) As Boolean

    'KZ: checks if textbox contains valid number
    Dim CharCounter As Integer
    Dim TempString As String
    
    If Text = "" Then   'KZ: blank is no good
        Let GoodNumber = False
        Exit Function
    End If
    
    Let CharCounter = 1
    Do While CharCounter <= Len(Text) And Numeral(Text, CharCounter)
        'KZ: first run through all the initial numerals
            CharCounter = CharCounter + 1
    Loop
    
    If CharCounter <= Len(Text) Then
    'KZ:the next thing has to be a decimal
        Do While CharCounter <= Len(Text) And Mid(Text, CharCounter, 1) = "."
            CharCounter = CharCounter + 1
        Loop
        If CharCounter > Len(Text) Then
            GoodNumber = False  'KZ: number ends in a decimal. no good.
            Exit Function
        End If
    End If
    If CharCounter <= Len(Text) Then 'KZ: and then more numerals
        Do While CharCounter <= Len(Text) And Numeral(Text, CharCounter)
            CharCounter = CharCounter + 1
        Loop
    End If
    
    If CharCounter <= Len(Text) Then 'if we're not at the end now, there's
        'something wrong with the number
        GoodNumber = False
    Else
        GoodNumber = True
    End If

End Function


Public Function GoodNumberPerhapsNegative(ByVal Text As String) As Boolean

    'KZ: checks if textbox contains valid number
    'BH:  which, for this function, can be a negative number, unlike for GoodNumber().
    
    Dim CharCounter As Integer
    Dim TempString As String, Buffer As String
    
    Let Buffer = Text
    
    'KZ: blank is no good
        If Buffer = "" Then
            Let GoodNumberPerhapsNegative = False
            Exit Function
        End If
    'BH:  Detect negative numbers, which are ok.
        If Left(Buffer, 1) = "-" Then
            Let Buffer = Mid(Buffer, 2)
        End If
    'KZ:  first run through all the initial numerals
        Let CharCounter = 1
        Do While CharCounter <= Len(Buffer) And Numeral(Buffer, CharCounter)
                Let CharCounter = CharCounter + 1
        Loop
    'KZ:  The next thing has to be a decimal.
        If CharCounter <= Len(Buffer) Then
            'BH:  keep detecting consecutive .
                Do While CharCounter <= Len(Buffer) And Mid(Buffer, CharCounter, 1) = "."
                    Let CharCounter = CharCounter + 1
                Loop
            'KZ: number ends in a decimal. no good.
                If CharCounter > Len(Buffer) Then
                    Let GoodNumberPerhapsNegative = False
                    Exit Function
                End If
        End If
    
    'KZ: and then more numerals
        If CharCounter <= Len(Buffer) Then
            Do While CharCounter <= Len(Buffer) And Numeral(Buffer, CharCounter)
                Let CharCounter = CharCounter + 1
            Loop
        End If
    
    'KZ:  if we're not at the end now, there's something wrong with the number
        If CharCounter <= Len(Buffer) Then
            Let GoodNumberPerhapsNegative = False
        Else
            Let GoodNumberPerhapsNegative = True
        End If

End Function

Public Function GoodDecimal(Text As String) As Boolean

    'This checks if textbox contains a valid positive number, which can be a decimal.
    
    Dim i As Long, j As Long
    Dim Buffer As String
    
    Let Buffer = Text
    
    'KZ: blank is no good
        If Buffer = "" Then
            Let GoodDecimal = False
            Exit Function
        End If
    'Handle negatives.
        If Left(Buffer, 1) = "-" Then
            Let Buffer = Mid(Buffer, 2)
        End If
    'Check every value.
        For i = 1 To Len(Buffer)
            Select Case Mid(Buffer, i, 1)
                Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                    'Do nothing
                Case "."
                    'Scan the decimal part.
                        For j = i + 1 To Len(Buffer)
                            Select Case Mid(Buffer, j, 1)
                                Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                                    'Do nothing.
                                Case Else
                                    'Bad number, return False.
                                        Let GoodDecimal = False
                                        Exit Function
                            End Select
                        Next j
                        Exit For
                Case Else
                    'Bad number, return False.
                        Let GoodDecimal = False
                        Exit Function
            End Select
        Next i
    
    Let GoodDecimal = True

End Function



Public Function Numeral(ByVal Text As String, ByVal CharCounter As Integer) As Boolean

'KZ: Determine whether a character in a string is a numeral

    Dim ASCIICode As Integer    'KZ: I think it's marginally faster
                                '   to do it this way rather than calculating
                                '   Asc(Mid...) twice, but I could be wrong.
    
    'Characters not in the string are not numerals.
        If CharCounter > Len(Text) Then
            Let Numeral = False
            Exit Function
        End If
    
    'KZ: the above is necessary because VB seems to evaluate both sides
    '   of an AND even when the left side has already been found to be false (!!).
    '   this means this function can get called from GoodNumber even when we've
    '   passed the end of the string.
    
    Let ASCIICode = Asc(Mid(Text, CharCounter))

    'The numerals reside between 48 and 57 in the ASCII series.
        If ASCIICode > 57 Or ASCIICode < 48 Then
            Let Numeral = False
        Else
            Let Numeral = True
        End If

End Function

Public Function FillStringTo(MyString As Variant, MyLength As Long) As String

    'Return a string that is as many spaces as is necessary to fill out the input string
    '   to the specified length.  This can then be concatenated to produce pretty
    '   justification.
    
    Dim Base As String
    Dim Buffer As String
    Dim i As Long
    
    Let Base = Trim(MyString)
    Let Buffer = MyString
    For i = Len(MyString) + 1 To MyLength
        Let Buffer = Buffer + " "
    Next i
    Let FillStringTo = Buffer

End Function

Public Function TrueFalseToYesNo(MyBoolean As Boolean) As String
    'Return the strings "yes" or "no" according to the value of a Boolean variable.
        If MyBoolean = True Then
            Let TrueFalseToYesNo = "yes"
        Else
            Let TrueFalseToYesNo = "no"
        End If
End Function

Public Function NiceDate() As String
    'Take off initial zeroes from Date$.
        If Left(Date$, 1) = "0" Then
            Let NiceDate = Mid(Date$, 2)
        Else
            Let NiceDate = Date$
        End If
End Function

Public Function NiceTime() As String
    'Make a legible time out of what VB provides.
        Dim Buffer As String
        
        Let Buffer = Format(Time$, "hh:mm AMPM")
        If Left(Buffer, 1) = "0" Then
            Let Buffer = Mid(Buffer, 2)
        End If
        Select Case Right(Buffer, 2)
            Case "AM"
                Let NiceTime = Left(Buffer, Len(Buffer) - 2) + "a.m."
            Case "PM"
                Let NiceTime = Left(Buffer, Len(Buffer) - 2) + "p.m."
        End Select
End Function
Public Function RomanNumeral(MyLong As Long) As String
    Dim Thousands As Long, Hundreds As Long, Tens As Long, Ones As Long
    Dim Buffer As String
    If MyLong > 3999 Then
        Let RomanNumeral = ""
        Exit Function
    End If
    Let Thousands = Int(MyLong / 1000)
    Let MyLong = MyLong - 1000 * Thousands
    Let Hundreds = Int(MyLong / 100)
    Let MyLong = MyLong - 100 * Hundreds
    Let Tens = Int(MyLong / 10)
    Let Ones = MyLong - 10 * Tens
    Let Buffer = ""
    Select Case Thousands
        Case 1
            Let Buffer = Buffer + "M"
        Case 2
            Let Buffer = Buffer + "MM"
        Case 3
            Let Buffer = Buffer + "MMM"
    End Select
    Select Case Hundreds
        Case 1
            Let Buffer = Buffer + "C"
        Case 2
            Let Buffer = Buffer + "CC"
        Case 3
            Let Buffer = Buffer + "CCC"
        Case 4
            Let Buffer = Buffer + "CD"
        Case 5
            Let Buffer = Buffer + "D"
        Case 6
            Let Buffer = Buffer + "DC"
        Case 7
            Let Buffer = Buffer + "DCC"
        Case 8
            Let Buffer = Buffer + "DCCC"
        Case 9
            Let Buffer = Buffer + "CM"
    End Select
    Select Case Tens
        Case 1
            Let Buffer = Buffer + "X"
        Case 2
            Let Buffer = Buffer + "XX"
        Case 3
            Let Buffer = Buffer + "XXX"
        Case 4
            Let Buffer = Buffer + "XL"
        Case 5
            Let Buffer = Buffer + "L"
        Case 6
            Let Buffer = Buffer + "LX"
        Case 7
            Let Buffer = Buffer + "LXX"
        Case 8
            Let Buffer = Buffer + "LXXX"
        Case 9
            Let Buffer = Buffer + "XC"
    End Select
    Select Case Ones
        Case 1
            Let Buffer = Buffer + "I"
        Case 2
            Let Buffer = Buffer + "II"
        Case 3
            Let Buffer = Buffer + "III"
        Case 4
            Let Buffer = Buffer + "IV"
        Case 5
            Let Buffer = Buffer + "V"
        Case 6
            Let Buffer = Buffer + "VI"
        Case 7
            Let Buffer = Buffer + "VII"
        Case 8
            Let Buffer = Buffer + "VIII"
        Case 9
            Let Buffer = Buffer + "IX"
    End Select
    Let RomanNumeral = Buffer
End Function
    

'-----------------------------------Numeric Utilities-----------------------------------------

Function Factorial(MyLong As Long) As Single

    'The factorial of an integer.
    
        On Error GoTo CheckError
        
        Dim i As Single
        Dim Buffer As Single
            
        Let Buffer = 1
        For i = 2 To MyLong
            Let Buffer = Buffer * i
        Next i
        Let Factorial = Buffer
        Exit Function
    
CheckError:
        'Return -1 if you overflow.
            Let Factorial = -1

End Function


'----------------------Apparatus for calculating log likelihood------------------------------

Function CalculateLogLikelihood(ByRef NumberOfForms As Long, NumberOfRivals() As Long, Probability() As Single, Frequency() As Single) As gLikelihoodCalculation

    'Calculate the log likelihood of the data.
    
        Dim Buffer As Single, CandidateLikelihood As Single
        Dim FormIndex As Long, RivalIndex As Long
        
    'Go through all forms; for those with nonzero frequency, multiply frequency by log probability.
        
        For FormIndex = 1 To NumberOfForms
           For RivalIndex = 1 To NumberOfRivals(FormIndex)
           'For RivalIndex = 0 To NumberOfRivals(FormIndex)
               If Frequency(FormIndex, RivalIndex) > 0 Then
                    If Probability(FormIndex, RivalIndex) = 0 Then
                       'Log likelihood is negative infinity in this case, so adjust the probability instead to the
                       '   arbitrary value of .001 (per Zuraw and Hayes 2017, Lg.), and raise a flag of warning.
                            Let CandidateLikelihood = Log(0.001) * Frequency(FormIndex, RivalIndex)
                            Let CalculateLogLikelihood.IncludesAZeroProbability = True
                    Else
                        Let CandidateLikelihood = Log(Probability(FormIndex, RivalIndex)) * Frequency(FormIndex, RivalIndex)
                    End If
                    Let Buffer = Buffer + CandidateLikelihood
                End If
           Next RivalIndex
        Next FormIndex
    
        Let CalculateLogLikelihood.LogLikelihood = Buffer
        
End Function


