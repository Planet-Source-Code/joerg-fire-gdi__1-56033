Attribute VB_Name = "modPerfCount"
Option Explicit

Public Declare Function QueryPerformanceCounter Lib "kernel32" _
    (lpPerformanceCount As Currency) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" _
(lpFrequency As Currency) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private curStart As Currency
Private curFinish As Currency
Private curFreq As Currency
Private fTimeLag As Double
Private fStdError As Double

Private bRunning As Boolean
Private bWarned As Boolean

Public Sub PerfInit()
Dim I As Long
Dim fMillis As Double
Const COUNTS As Long = 1000
Dim fTest As Double


    'get the counter frequency
    
    If QueryPerformanceFrequency(curFreq) = 0 Then
        MsgBox "You are not able to access a Performance Counter, sorry."
        Exit Sub
    End If
    
    'now we need to find out how long it actually takes
    'to make the function calls, so we can subtract by that amount.
    
    fMillis = 0
    fTimeLag = 0
    For I = 1 To COUNTS
        PerfStart
        PerfFinish
        fMillis = fMillis + PerfElapsedInternal
        
    Next
    
    fTest = fMillis / COUNTS
    
    'Here fTimeLag is the mean time that it takes to call PerfStart then
    'PerfFinish, that is, how expensive the function calls are themselves.
    'We need to subtract this amount from our result to get a more accurate
    'number.
    
    fTimeLag = fTest
    
    'Now we are going to see how accurate our calls are by looking at the
    'standard deviation from the mean.  Here the mean function call time
    'will be fTimeLag.
    
    fMillis = 0
    fStdError = 0
    For I = 1 To COUNTS
        PerfStart
        PerfFinish
        
        'compute the square of the distance from the mean
        fMillis = fMillis + (PerfElapsedInternal - fTimeLag) * _
            (PerfElapsedInternal - fTimeLag)
    Next
    
    'now divide by number of iterations and take square root to get std deviation.
    'this is a measure of how accurate this Perf counter really is.
    
    fStdError = Sqr(fMillis / COUNTS)
    
    
End Sub

Public Function PerfTimeLag() As Single
    PerfTimeLag = MakeSignificant(fTimeLag, fStdError)
End Function

Public Function PerfStdError() As Single
    PerfStdError = MakeSignificant(fStdError, fStdError)
End Function


Public Sub PerfStart()

    'We only allow one start/finish session at a time here
    If bRunning Then Exit Sub
    
    'If curFreq is zero then either PerfInit has not yet been called,
    'or there is no performance counter on the equipment.
    
    If curFreq = 0 Then
        If Not bWarned Then
            MsgBox "Please Initialize with PerfInit before calling this Sub"
            bWarned = True
        End If
        Exit Sub
    End If
    
    'Flag the session as being in progress
    bRunning = True
    
    'Save the current count
    QueryPerformanceCounter curStart
End Sub

Public Sub PerfFinish()

    'save the current final count
    QueryPerformanceCounter curFinish
    
    'Flag the session as complete
    bRunning = False
End Sub

Public Function PerfElapsed() As Single
    'Note: for more accurate results, you should call
    'PerfFinish prior to calling PerfElapsed.  If you want
    'to use this to update a progress bar or something, then
    'calling this before PerfFinish might be OK.

    'Check for initialization and/or presence of a performance counter

    If curFreq = 0 Then
        If Not bWarned Then
            MsgBox "Please Initialize with PerfInit before calling this Sub"
            bWarned = True
        End If
        PerfElapsed = 0
        Exit Function
    End If
    
    PerfElapsed = MakeSignificant(PerfElapsedInternal, fStdError)
End Function

Private Function PerfElapsedInternal() As Double
Dim curTest As Currency
Dim fResult As Double

    'Make a quick check if the session is in progress, otherwise
    'use the value we got by calling PerfFinish (the better way)
    
    If bRunning Then
        QueryPerformanceCounter curTest
    Else
        curTest = curFinish
    End If
    
    'Note that we are dividing a Currency by another Currency, so the
    'factor of 10000 is going to cancel out in the division.  Multiply
    '1000 to get milliseconds, and subtract the time lag we found
    'in PerfInit.
    
    fResult = 1000 * (CDbl(curTest) - CDbl(curStart)) / CDbl(curFreq) - fTimeLag
    PerfElapsedInternal = fResult
End Function

Public Function MakeSignificant(fValue As Double, fError As Double) As Single
Dim fLog As Double
Dim fInt As Double

    'This function is used to strip all the bogus digits off a
    'result.  It uses the standard error to do this.
    If fError = 0 Then
        fLog = -4   'arbitrary, so we don't take log of 0
    Else
        fLog = Log10(fError)
    End If
    fInt = Int(fLog)
    If fInt < 0 Then
        MakeSignificant = CSng(Format(fValue, "0." & String$(-fInt, "0")))
    ElseIf fInt = 0 Then
        MakeSignificant = CSng(Format(fValue, "0"))
    Else
        MakeSignificant = (10 ^ fInt) * CSng(Format(fValue / (10 ^ fInt), "0"))
    End If
End Function

Private Function Log10(fValue As Double) As Double
    'log base 10
    Log10 = Log(fValue) / Log(10#)
End Function
