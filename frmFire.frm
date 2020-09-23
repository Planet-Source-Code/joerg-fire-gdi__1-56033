VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmFire 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "FIRE! [Gdi+] Press key to quit"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   216
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   440
   StartUpPosition =   2  'Bildschirmmitte
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Unten ausrichten
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   2985
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   1535
            MinWidth        =   1005
            Text            =   "CPU-Load:"
            TextSave        =   "CPU-Load:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   1005
            MinWidth        =   1014
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   9014
            MinWidth        =   9014
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   180
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   318
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "frmFire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Original fire-effect was done by Cicri, using GDI with Setpixel, see:
' http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=55520&lngWId=1
'
' For this great effect i tried to find a faster way to display the effect,
' so here it is.
'
' Required:     GDIplus.dll
'
' Using GdipBitmapLockBits seems to be the fastest way to realize this because we are
' allowed to modify the memory of the bitmap directly using PutMem4 or even CopyMemory.
'
' BE SURE TO COMPILE IT TO SEE HOW FAST THIS WORKS! I needed to add an API-Timer
' because in compiled mode this works too fast (about 7-8 ms to calculate and display,
' this means about 125 fps on my Athlon 1800+ @ 440 Pixels width!).
'
' Thanks to all who did great work i learned from, its impossible to mention all!
' Special thanks go to: Avery (GDIplusAPI-module),
' Benjamin Kunz (clsCPULoad) - http://www.vbarchiv.net/archiv/tipp_1080.html
' http://vdev.net/vbprofiler/perfcount1.htm (Build a perfect counter - modPerfCount)
' http://www.vbarchiv.net - a lot of tips for my modCCBar (german)
' Michel Rutten - PutMem-function - http://www.xbeat.net/vbspeed/i_VBVM6Lib.html#PutMem
'
' Hope you like it all. Comments, sugestions and of course, votes are welcome!


Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias _
        "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal _
        ByteLen As Long)
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

Private Const lWidth As Long = 440  ' ScaleWidth of form
Private Const lHeight As Long = 200 ' ScaleHeight of form

Private Type GraphicLine
    PixelX(lWidth) As Integer
End Type
Dim PixelY(lHeight) As GraphicLine

Private CPULoad As New clsCPULoad

Dim lToken As Long          ' Needed to close GDI+
Dim CPalRGB(255) As Long    ' Holds colors
Dim lngWidth As Long        ' Bitmap-width
Dim lngHeight As Long       ' Bitmap-height
Dim graphics As Long        ' GDI+ graphic class
Dim bitmap As Long          ' GDI+ bitmap class
Dim bmpData As BitmapData   ' BitmapData-Structure, for Lockbits
Dim rctL As RECTL           ' Rect, for Lockbits
Public strCap As String     ' Holds Form-Capture

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    ' only one instane allowed
    If App.PrevInstance Then
        MsgBox "FireGDI+ is running already!"
        AppActivate App.Title
        End
    End If
    
    ' Test for running mode
    If InVBDesignEnvironment Then
        strCap = "FIRE! [Gdi+]/[Design Mode] - Click on form to quit! "
    Else
        strCap = "FIRE! [Gdi+]/[Compiled Mode] - Click on form to quit! "
    End If
    
    ' Ensure forms parameters
    With Me
        .Caption = strCap
        .ScaleMode = vbPixels
        .ScaleWidth = lWidth
        .ScaleHeight = lHeight
    End With
    
    Randomize Timer
    
    ' Stacksize festlegen
    CPULoad.StackSize = 10
    
    ' Set the Progressbars parent
    With ProgressBar1
        .Visible = False
        .value = 0
        SetProgressBarToStatusBar .hwnd, StatusBar1.hwnd, 3
    End With

    ' Set initial forecolor and backcolor of Progressbar
    SetProgressbarColor ProgressBar1.hwnd, &HFF00, vbWhite
    ProgressBar1.Visible = True
    
    ' Load the GDI+ Dll
    Dim GPInput As GdiplusStartupInput

    GPInput.GdiplusVersion = 1
    If GdiplusStartup(lToken, GPInput) <> Ok Then
        MsgBox "Error loading GDI+!", vbCritical
        Unload Me
    End If
    
    ' Initializations for GDI+
    Call GdipCreateFromHDC(Me.hdc, graphics)  ' Initialize the graphics class - required for all drawing
    Call GdipCreateBitmapFromGraphics(Me.ScaleWidth, Me.ScaleHeight, graphics, bitmap)
    
    ' Get the image height and width
    Call GdipGetImageHeight(bitmap, lngHeight)
    Call GdipGetImageWidth(bitmap, lngWidth)
    
    ' initialize rectL for LockBits
    rctL.Right = 0
    rctL.Top = 0
    rctL.Right = lngWidth
    rctL.Bottom = lngHeight
    
    ' Create colors for fire...
    CreateColorTable
    ' Show timer setting
    Label2.Caption = "TimerSetting: " & format(lElapse, "0#.0") & " ms"
    ' we ain't seen nothing yet
    Me.Show
    ' Setup API timer
    SetTimer Me.hwnd, 0, lElapse, AddressOf TimerProc

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'first kill API timer
    KillTimer Me.hwnd, 0
    
    ' GDI+ CleanUp
    Call GdipDisposeImage(bitmap)
    Call GdipDeleteGraphics(graphics)
    
    ' Unload the GDI+ Dll
    Call GdiplusShutdown(lToken)
    
    ' General cleanup
    Set frmFire = Nothing
    Set CPULoad = Nothing
    
End Sub

' Calculate and draw fire pixels
' Pixels calculation done by Cicri,
' added/modified GDIplus-drawing and time measurement

Public Sub ModifyPixels()

    Dim x As Long, y As Long    ' coordinates
    Dim tempPixel As Long       ' temporary var holds index of colortable
    Dim memPtr As Long          ' hold memory-address to write to
    
    'initialise and start time counter
    PerfInit
    PerfStart
    
    'Fill 3 bottom lignes with random values
    For x = 0 To lngWidth - 1
        PixelY(lngHeight - 3).PixelX(x) = 25 * Rnd() + 80
        PixelY(lngHeight - 2).PixelX(x) = 25 * Rnd() + 80
        PixelY(lngHeight - 1).PixelX(x) = 25 * Rnd() + 80
    Next
    'Add random hot spots, i.e. 3x3 pure white "pixels"
    For x = 0 To 40 * Rnd() 'Number of hot spots is random too, can be changed
        y = (lngWidth - 2) * Rnd() + 1
        
        PixelY(lngHeight - 3).PixelX(y - 1) = 255
        PixelY(lngHeight - 3).PixelX(y) = 255
        PixelY(lngHeight - 3).PixelX(y + 1) = 255
        PixelY(lngHeight - 2).PixelX(y - 1) = 255
        PixelY(lngHeight - 2).PixelX(y) = 255
        PixelY(lngHeight - 2).PixelX(y + 1) = 255
        PixelY(lngHeight - 1).PixelX(y - 1) = 255
        PixelY(lngHeight - 1).PixelX(y) = 255
        PixelY(lngHeight - 1).PixelX(y + 1) = 255
    Next
    
    ' The Bitmap class provides the LockBits and corresponding UnlockBits methods
    ' which enable you to fix a portion of the bitmap pixel data array in memory,
    ' access it directly and finally replace the bits in the bitmap with the modified data.
    ' LockBits returns a BitmapData class that describes the layout and position of the data in the locked array.
    Call GdipBitmapLockBits(bitmap, rctL, ImageLockModeWrite, PixelFormat32bppARGB, bmpData)
    
    'Compute each pixel based on 8 neighbours
    For x = 1 To bmpData.Width - 2
        'Note that y loops from 90 because flames never go higher. You might want or need to change that, but limiting j this way makes the code faster :)
        For y = 90 To bmpData.Height - 3
            'Use of a temp variable to avoid redundant accesses to the table
            'Note that we don't use the neighbours of the current pixel but of the pixel below, by using x-2, x-1 and x as opposed to x-1, x, x+1
            'This way, we combine both the fact that the flames go up AND the computation of the pixel based on neighbours in one shot!
            'I don't know why other people don't seem to do that...
            tempPixel = (PixelY(y + 2).PixelX(x - 1) + PixelY(y + 2).PixelX(x) + PixelY(y + 2).PixelX(x + 1) + PixelY(y + 1).PixelX(x - 1) + PixelY(y + 1).PixelX(x) + PixelY(y + 1).PixelX(x + 1) + PixelY(y).PixelX(x - 1) + PixelY(y).PixelX(x + 1)) \ 8
            
            ' This makes "very hot" flames live a little bit longer...sometimes (price: 10 ms in IDE)
            If tempPixel > 106 Then 'Hot pixel?
                If y < 170 Then     'High enough?
                    If Round(Rnd()) = 0 Then    'Do it random...
                        If tempPixel < 250 Then tempPixel = tempPixel + 2 ' and make the pixel just a little bit hotter..
                    End If
                End If
            End If
            
            'As flames go up, they cool down, hence the need to decrease the color, i.e. temperature
            'You could get rid of that but flames would go higher and not seem as realistic
            If tempPixel > 0 Then tempPixel = tempPixel - 1
            PixelY(y).PixelX(x) = tempPixel
            
            ' Get memory address of bitmap to write to [Scan0+(y * stride)+(x*4)]
            memPtr = bmpData.scan0 + (y * bmpData.stride) + (x * 4) 'row = Y; blue, green, red, alpha
            ' write value directly to memory (don't need to check for permission with IsBadWritePtr)
            ' PutMem4 is a little bit faster than copymemory, works only with vb6
            ' further information: http://www.xbeat.net/vbspeed/i_VBVM6Lib.html#PutMem
            PutMem4 memPtr, CPalRGB(tempPixel)
            ' for vb5, you should use this
            'CopyMemory ByVal memPtr, CPalRGB(tempPixel), 4
        Next y
    Next x
    ' Write back the modified bitmap
    Call GdipBitmapUnlockBits(bitmap, bmpData)
    ' Draw the modified image
    Call GdipDrawImageRect(graphics, bitmap, 0, 0, lngWidth, lngHeight)
    ' Stop the time counter
    PerfFinish
    ' Show time needed to calculate and modify pixels
    Label1.Caption = "ModifyPixels: " & format(PerfElapsed, "0#.0") & " ms"
End Sub

' calculate a 256-color-table, done by Cicri.
' added translation of colors for drirect memory access.

Sub CreateColorTable()

    Dim n As Long

    'Set palette
    '0-255=rouge, 256-65280=green, 65536-16711680=blue
    For n = 0 To 10
        'Replace with ColourPalette(i) = 0 if you do not like the blue gradient on top of flames
        CPalRGB(n) = 65536 * n * 8 '0-7:bleu 0-64
        'Comment this line if you do not like the blue gradient on top of flames
        CPalRGB(n + 10) = 65536 * 80 - 65536 * n * 8 '8-15:bleu 64-0
    Next
    'Red gradient 0-256 for cold pixels
    For n = 10 To 41
        CPalRGB(n) = CPalRGB(n) + (n - 10) * 8
    Next
    'Yellow gradient 0-256 for warm pixels, with yellow = red+green in equal proportions
    For n = 42 To 73
        CPalRGB(n) = (n - 42) * 8 * 256 + 255
    Next
    'Yellow to white gradient for hot pixels, start with plain yellow and add blue to get white
    For n = 74 To 105
        CPalRGB(n) = (n - 74) * 8 * 65536 + 65280 + 255
    Next
    'Fill remaining palette with pure white for hotest pixels
    For n = 106 To 255
        CPalRGB(n) = vbWhite
    Next
    ' change the order of colors to copy them direct to memory
    For n = 0 To 255
        CPalRGB(n) = GetRGB_VB2GDIP(CPalRGB(n))
    Next
End Sub

' Desc: This function will return  whether you are running
'       your program or DLL from within the IDE, or compiled.
Private Function InVBDesignEnvironment() As Boolean
'Please see http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=11615&lngWId=1
    Dim strFileName As String
    Dim lngCount As Long
    
    strFileName = String(255, 0)
    lngCount = GetModuleFileName(App.hInstance, strFileName, 255)
    strFileName = Left(strFileName, lngCount)
    
    InVBDesignEnvironment = False

    If UCase(Right(strFileName, 7)) = "VB5.EXE" Then
        InVBDesignEnvironment = True
    ElseIf UCase(Right(strFileName, 7)) = "VB6.EXE" Then
        InVBDesignEnvironment = True
    End If
End Function

' This will draw the CPU-Load in our Progressbar
' the color is green at 0%, fading to yellow at 50%, finally to red at 100%

Public Sub DrawPGBar()

    Dim lValue As Long  ' CPU-Load (percent)
    Dim lColor As Long  ' Forecolor for ProgressBar
    
    On Error Resume Next
    
    ' get CPU-load
    lValue = CPULoad.value
    ' show percent
    StatusBar1.Panels(2).Text = lValue & " %"
    ' calculate color - nice, isn't it?
    Select Case lValue
        Case Is <= 50
            ' 0% = green to 50% = yellow
            lColor = vbGreen Or (lValue / 50 * vbRed)
        Case Else
            ' 50% = yellow to 100% = red
            lColor = vbRed Or (65280 - ((lValue - 50) / 50 * 65280))
    End Select
    ' Set forecolor and value
    SetProgressbarForeColor ProgressBar1.hwnd, lColor
    ProgressBar1.value = lValue
    
End Sub

