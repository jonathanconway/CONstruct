Attribute VB_Name = "ModShellAndWait"
' modShellAndWait
' 1998/09/14 From MSDN HOWTO: 32-Bit App Can Determine When a Shelled Process Ends
'            Last reviewed: October 13, 1997
'            Article ID: Q129796
' 1998/10/07 Modified by Larry Rebich, larry@buygold.net
' 1998/11/18 Add Terminate and Log File

    Option Explicit
    DefLng A-Z
    
    Private Type STARTUPINFO
        cb              As Long
        lpReserved      As String
        lpDesktop       As String
        lpTitle         As String
        dwX             As Long
        dwY             As Long
        dwXSize         As Long
        dwYSize         As Long
        dwXCountChars   As Long
        dwYCountChars   As Long
        dwFillAttribute As Long
        dwFlags         As Long
        wShowWindow     As Integer
        cbReserved2     As Integer
        lpReserved2     As Long
        hStdInput       As Long
        hStdOutput      As Long
        hStdError       As Long
    End Type

    Private Type PROCESS_INFORMATION
        hProcess        As Long
        hThread         As Long
        dwProcessID     As Long
        dwThreadID      As Long
    End Type

    Public Type udtShellAndWait         'pass information here
        sCommand As String              'command line for Shell
        bShellAndWaitRunning As Boolean 'shell and wait is running
        bNoTerminate  As Boolean        'no forced termination [no DoEvents]
        lMilliseconds As Long           'interrupt this often in milliseconds [1000 is 1 second]
        bLogFile As Boolean             'do a log file
        sLogFile As String              'log file name
        lMaxSize As Long                'maximum log file size before deleting and starting over, default is 20K
        dStart   As Date                'date/time started the process
        dFinish  As Date                'date/time finished the process
        dDiff    As Date                'date/time difference
        bTerminated As Boolean          'true if terminated by ShellAndWaitTerminate
        dTerminated As Date             'date/time if terminated
        tProcess As PROCESS_INFORMATION 'above structures
        tStart   As STARTUPINFO
    End Type
    Const mclMaxDefault             As Long = 20& * 1024&   'maximum log file size default, 20K
    Const mclMillisecondsDefault    As Long = 1& * 1000&    'default lMilliseconds if bNoTerminate

    
    Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, _
        ByVal dwMilliseconds As Long) As Long

    Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
        lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
        lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
        ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
        ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
        lpStartupInfo As STARTUPINFO, lpProcessInformation As _
        PROCESS_INFORMATION) As Long

    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

    Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, _
        ByVal uExitCode As Long) As Long


    Private Const NORMAL_PRIORITY_CLASS As Long = &H20
    Private Const INFINITE              As Long = -1
    Private Const WAIT_TIMEOUT          As Long = &H102
    '

Public Function ShellAndWait(tShellAndWait As udtShellAndWait) As Boolean
' 1998/10/07 Add optional log and times
' 1998/10/27 Add DoEvents to allow the process to be terminated. Move tProcess to caller.
    
    Dim lRtn    As Long
    Dim iFN     As Integer
    Dim lMilliseconds As Long
    
    With tShellAndWait
        .dStart = Now                       'started now
        .bTerminated = False                'set not terminated yet into structure
        .dTerminated = 0                    'not terminated at this point
        lMilliseconds = .lMilliseconds      'local variable
        If lMilliseconds <= 0 Then          'zero or negative then use 1 second
            lMilliseconds = mclMillisecondsDefault  'default
        End If
        If .bNoTerminate Then               'don't allow terminate
            .lMilliseconds = INFINITE       'set to never return
        End If
        .bShellAndWaitRunning = True        'started
        
        ' Initialize the STARTUPINFO structure:
        .tStart.cb = Len(.tStart)
        
        ' Start the shelled application:
        lRtn = CreateProcessA(0&, .sCommand, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, .tStart, .tProcess)
        
        ' Wait for the shelled application to finish:
        Do
            lRtn = WaitForSingleObject(.tProcess.hProcess, lMilliseconds)   'wait milliseconds
            If lRtn <> WAIT_TIMEOUT Then
                Exit Do
            End If
            DoEvents                                'allow other processes
        Loop While True
        
        lRtn = CloseHandle(.tProcess.hProcess)
        If lRtn <> 0 Then
            ShellAndWait = True                     'report success
        End If
        .bShellAndWaitRunning = False               'ended
        
        On Error GoTo ShellAndWaitExit              'skip log if any error
        .dFinish = Now
        .dDiff = .dFinish - .dStart
        If .bLogFile Then                           'write the log
            KillLogFileIfTooBig tShellAndWait       'can't let it get too big
            iFN = FreeFile                          'get a file handle
            Open .sLogFile For Append As iFN        'Add to an existing file
            Print #iFN, Format$(.dStart, "general date"); ", ";     'started
            Print #iFN, Format$(.dFinish, "general date"); ", ";    'finished
            Print #iFN, Format$(.dDiff, "Hh.Nn.Ss"); " ";           'duration
            Print #iFN, .sCommand                   'command executed
            Close #iFN
        End If
    End With
ShellAndWaitExit:
End Function

Public Function ShellAndWaitTerminate(tShellAndWait As udtShellAndWait) As Boolean
' 1998/10/27 Allow tShellAndWait.tProcess.hProcess to be terminated.
    Dim lRtn As Long
    Dim iFN  As Integer
    Dim iLn  As Integer
    
    iLn = Len(Format$(Now, "general date")) + 2     'offset this amount for log
    
    With tShellAndWait
        lRtn = TerminateProcess(.tProcess.hProcess, "0")
        If lRtn <> 0 Then                           'success
            lRtn = CloseHandle(.tProcess.hProcess)  'close handle, don't know if this is really needed!
            .bTerminated = True                     'set terminated
            .dTerminated = Now                      'and date/time
            ShellAndWaitTerminate = True            'report success
        End If
    
        On Error GoTo ShellAndWaitTerminateExit     'skip log if any error
        If .bLogFile Then                           'write the log
            KillLogFileIfTooBig tShellAndWait       'can't let it get too big
            iFN = FreeFile                          'get a file handle
            Open .sLogFile For Append As iFN        'open for append
            Print #iFN, String$(iLn, " ");          'offset
            Print #iFN, Format$(.dTerminated, "general date");  'time terminated
            Print #iFN, ", Forced Termination"      'and the reason
            Close #iFN                              'close the file
        End If
    End With
ShellAndWaitTerminateExit:
End Function

Private Function KillLogFileIfTooBig(tShellAndWait As udtShellAndWait) As Boolean
' Don't let the log file get too big
    Dim lLen            As Long
    Dim lMax            As Long
    Dim iFN             As Integer
    
    With tShellAndWait
        If .bLogFile Then                           'any log file
            If .lMaxSize = 0 Then                   'any size in structure
                lMax = mclMaxDefault                'no, use default
            Else
                lMax = .lMaxSize                    'yes, use what is in the structure
            End If
            On Error GoTo KillLogFileIfTooBigExit   'don't let this fail
            iFN = FreeFile                          'get a handle
            Open .sLogFile For Input As #iFN        'open it to get file length
            lLen = LOF(iFN)                         'get the length
            Close #iFN                              'close it
            If lLen > lMax Then                     'too big?
                Kill .sLogFile                      'yes, try to delete it
                KillLogFileIfTooBig = True          'say success
            End If
        End If
    End With
KillLogFileIfTooBigExit:
End Function




