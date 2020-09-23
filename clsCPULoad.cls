VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPULoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Code by Benjamin Kunz. For more information visit:
' http://www.vbarchiv.net/archiv/tipp_1080.html

Option Explicit

Private Declare Function RegQueryValueEx& Lib "advapi32.dll" _
  Alias "RegQueryValueExA" ( _
  ByVal hKey&, _
  ByVal lpValueName$, _
  ByVal lpReserved&, _
  lpType&, _
  lpData As Any, _
  lpcbData&)
        
Private Declare Function RegOpenKey& Lib "advapi32.dll" _
  Alias "RegOpenKeyA" ( _
  ByVal hKey&, _
  ByVal lpSubKey$, _
  phkResult&)
        
Private Declare Function RegCloseKey& Lib "advapi32.dll" ( _
  ByVal hKey&)

Private Declare Sub CopyMemory Lib "kernel32.dll" _
  Alias "RtlMoveMemory" ( _
  Dest As Any, _
  Src As Any, _
  ByVal Length As Long)
        
Private Declare Function PdhOpenQuery Lib "PDH.DLL" ( _
  ByVal Reserved As Long, _
  ByVal dwUserData As Long, _
  ByRef hQuery As Long) As Long
        
Private Declare Function PdhCloseQuery Lib "PDH.DLL" ( _
  ByVal hQuery As Long) As Long
        
Private Declare Function PdhVbAddCounter Lib "PDH.DLL" ( _
  ByVal QueryHandle As Long, _
  ByVal CounterPath As String, _
  ByRef CounterHandle As Long) As Long
        
Private Declare Function PdhCollectQueryData Lib "PDH.DLL" ( _
  ByVal QueryHandle As Long) As Long
        
Private Declare Function PdhVbGetDoubleCounterValue Lib "PDH.DLL" ( _
  ByVal CounterHandle As Long, _
  ByRef CounterStatus As Long) As Double

Private Declare Function GetVersionEx Lib "kernel32" _
  Alias "GetVersionExA" ( _
  VersionInfo As OSVERSIONINFOEX) As Long

Private Const OSVERSIONINFOSIZE = 148

Private Const PDH_CSTATUS_VALID_DATA = &H0
Private Const PDH_CSTATUS_NEW_DATA = &H1
Private Const ERROR_SUCCESS = 0
Private Const VER_PLATFORM_WIN32_NT = 2

Private Type OSVERSIONINFOEX
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
  wServicePackMajor As Integer
  wServicePackMinor As Integer
  wSuiteMask As Integer
  bProductType As Byte
  bReserved As Byte
End Type

Private hQuery As Long
Private hCounter As Long
Private RetVal As Long

Private Stack() As Long
Private StackPointer As Long
Private m_StackSize As Long
Private m_Sum As Long

Private m_BandWidth As Long

Private Sub Class_Initialize()

  m_StackSize = 5
  ReDim Stack(0 To m_StackSize - 1)
  StackPointer = 0
       
  If IsNT Then
    RetVal = PdhOpenQuery(0, 1, hQuery)
    If RetVal = 0 Then
      ' Define performance-counter
      RetVal = PdhVbAddCounter(hQuery, _
        "\Prozessor(0)\Prozessorzeit (%)", hCounter)
      
      ' Close Query on Error
      If RetVal <> 0 Then PdhCloseQuery hQuery
    End If
  End If
End Sub

Private Sub Class_Terminate()
  ' Close Query
  If IsNT Then PdhCloseQuery hQuery
End Sub

' NT-System?
Private Function IsNT() As Boolean
  Static VerInfo As OSVERSIONINFOEX, bOsVersionInfoEx As Long
  Static Flag As Boolean, NT As Boolean

  If Not Flag Then
    VerInfo.dwOSVersionInfoSize = Len(VerInfo)
    bOsVersionInfoEx = GetVersionEx(VerInfo)
        
    If bOsVersionInfoEx = 0 Then
      VerInfo.dwOSVersionInfoSize = OSVERSIONINFOSIZE
      GetVersionEx VerInfo
    End If

    NT = (VerInfo.dwPlatformId = VER_PLATFORM_WIN32_NT)
    Flag = True
  End If
        
  IsNT = NT
End Function

Public Property Get StackSize() As Long
  StackSize = m_StackSize
End Property

Public Property Let StackSize(ByRef NewStackSize As Long)
  m_StackSize = NewStackSize
  ReDim Stack(0 To m_StackSize - 1)
  StackPointer = 0
End Property

Private Sub PushBandWidth(ByRef NewBandWidth As Long)
  Static u As Long
    
  u = UBound(Stack)
  If StackPointer <= u Then
    Stack(StackPointer) = NewBandWidth
    m_Sum = m_Sum + NewBandWidth
    StackPointer = StackPointer + 1
        
  Else
    m_Sum = m_Sum - Stack(0) + NewBandWidth
    Call CopyMemory(Stack(0), Stack(1), u * 4)
    Stack(u) = NewBandWidth

  End If
  m_BandWidth = m_Sum / StackPointer
End Sub

Private Function GetValue9x() As Long
  Dim V As Long
  Static hK As Long, sK As String
  Const KDyn& = &H80000006
    
  sK = IIf(hK = 0, "PerfStats\StartStat", "PerfStats\StatData")
  If RegOpenKey(KDyn, sK, hK) Then Exit Function

  Call RegQueryValueEx(hK, "KERNEL\CPUUsage", 0, 4, V, 4)
  Call RegCloseKey(hK)
        
  PushBandWidth V
  GetValue9x = m_BandWidth
End Function

Private Function GetValueNT() As Long
  Dim dblValue As Double
  Dim pdhStatus As Long

  PdhCollectQueryData hQuery
  dblValue = PdhVbGetDoubleCounterValue(hCounter, pdhStatus)
        
  ' Get value of counter
  If (pdhStatus = PDH_CSTATUS_VALID_DATA) Or _
    (pdhStatus = PDH_CSTATUS_NEW_DATA) Then
    
    PushBandWidth CLng(dblValue)
    GetValueNT = m_BandWidth
  End If
End Function

' Returns CPU-Load
Public Property Get Value() As Long
  If IsNT Then
    Value = GetValueNT
  Else
    Value = GetValue9x
  End If
End Property


