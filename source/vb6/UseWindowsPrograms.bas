Attribute VB_Name = "UseWindowsPrograms"
'--------------------------------USE WINDOWS PROGRAMS---------------------------------------

Option Explicit

Sub OpenWebBrowser(MyWebSite As String)

    On Error GoTo CheckError
    Call UseWindowsPrograms.TryShellExecute(MyWebSite)
    Exit Sub
    
CheckError:
    MsgBox "Sorry, attempt to download from the Web failed.", vbExclamation
    Exit Sub
    
End Sub

Function TryShellExecute(MyFileName As String) As Long
   
    'Open the program specified by Windows for a particular file type.
    
    'This code is largely from Omar Mallat, on the Web at
    '   http://www.experts-exchange.com/Programming/Programming_Languages/Visual_Basic/Q_20272579.html
    '   Mallat's comments are marked OM.
   
        Dim MyShellExecuteInfo As ShellExecuteInfo   'OM:  Structure used by the function
        Dim RetVal As Long                           'OM:  return value
   
   'OM: Load the information needed to open MyFileName into the structure.
        With MyShellExecuteInfo
            'OM: Size of the structure
                .cbSize = Len(MyShellExecuteInfo)
            'OM: Use the optional hProcess element of the structure.
                .fMask = SEE_MASK_NOCLOSEPROCESS
            'OM: Handle to the window calling this function.
                .hwnd = Form1.hwnd
            'OM: The action to perform: open the file.
                .lpVerb = "open"
            'OM: The file to open.
                .lpFile = MyFileName
            'OM: No additional parameters are needed here.
                .lpParameters = ""
            'OM: The default directory -- not really necessary in this case.
                .lpDirectory = "C:\"
            'OM: Simply display the window.
                .nShow = SW_SHOWNORMAL
            'OM: The other elements of the structure are either not used
            'OM: or will be set when the function returns.
        End With
   
   'OM: Open the file using its associated program.
        Let RetVal = ShellExecuteEx(MyShellExecuteInfo)
        Let TryShellExecute = MyShellExecuteInfo.hInstApp
        
    'BH:  Deal with failure or success, depending.
        If RetVal = 0 Then
            'BH:  This is now dealt with in the calling routine, which knows the
            '   value of MyShellExecuteInfo.hInstApp.
                'OM: The function failed, so report the error.  Err.LastDllError
                '    could also be used instead, if you wish.
                '    Select Case MyShellExecuteInfo.hInstApp
                '        Case SE_ERR_FNF
                '            MsgBox "File not found"
                '        Case SE_NOASSOC
                '             MsgBox "no program is associated to this file"
                '        Case Else
                '            MsgBox "An unexpected error occured."
                '    End Select
                
        Else
            
            'OM:  Wait for the opened process to close before continuing.  Instead
            '     of waiting once for a time of INFINITE, this example repeatedly checks to see if the
            '     is still open.  This allows the DoEvents VB function to be called, preventing
            '     our program from appearing to lock up while it waits.
            'BH:  I'm not sure this is helpful.  It seems to be doing no harm.
            'BH:  1/11/04:  This seems kind of messy--I'm getting multiple calls to
            '   the program.  In the absence of a known way to fix, I can at least
            '   keep these calls 20 seconds apart.
            
            'BH:  No more than 20 seconds, please.
                Dim MyTimer As Long
                Let MyTimer = Timer
            
            Do
                If Timer - MyTimer > 20 Then Exit Do
                DoEvents
                Let RetVal = WaitForSingleObject(MyShellExecuteInfo.hProcess, 0)
            Loop While RetVal = WAIT_TIMEOUT
            
            'MsgBox "the program is ended"
            Let TryShellExecute = gShellExecuteWasSuccessful
        
        End If
    
End Function

