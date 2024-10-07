Attribute VB_Name = "mdlStuff"
Option Explicit

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NOSTOP = &H10

Type PointAPI
    x As Integer
    y As Integer
End Type

Enum States
    Blank
    Flag
    Question
    Number
    Mine
    SelMine
    NoMine
End Enum

Enum Faces
    MouseDwn
    Won
    Die
    Suspense
    Smile
End Enum

Public BoxesInRow As Integer
Public BoxesInColumn As Integer
Public MineCount As Integer
Public AllowQMarks As Boolean
Public AllowColor As Boolean
Public AllowSound As Boolean

Global Const BOX_SIZE = 16
Global Const SIDE_SPACE = 12

Const SHADOW = &H808080
Const FACE = &HC0C0C0
Const HIGHLIGHT = &HFFFFFF
Const TAG_INSET = "\"
Const TAG_OUTSET = "/"

Type Rect
    Pt1 As PointAPI
    Pt2 As PointAPI
End Type

Enum BoxTypes
    Inset
    outset
End Enum

Enum Levels
    Beginner
    Intermediate
    Expert
    Custom
End Enum

Public AppPath As String
Const FDR_SND = "Sound\"

Global Const INI_FILE = "WinMine.INI"
Global Const SECTION = "Minesweeper"
Const H_WIDTH = "Width"
Const H_HEIGHT = "Height"
Const H_MINES = "Mines"
Const H_MARK = "Mark"
Const H_COL = "Color"
Const H_SND = "Sound"
Global Const H_TIME = "Time"
Global Const H_NAME = "Name"

Enum PicTypes
    ResBitmap
    ResIcon
End Enum

Const SND_STEP = "Step"
Global Const SND_WIN = "Win"
Global Const SND_CLICK = "Click"
Global Const SND_BLAST = "Blast"

Sub PlaySound(SoundName As String)
    If Not AllowSound Then Exit Sub
    sndPlaySound AppPath & FDR_SND & SoundName & ".wav", SND_ASYNC Or SND_NOSTOP
End Sub

Sub MakeStep()
Static FootNum As Integer
    FootNum = Abs(FootNum - 1)
    PlaySound SND_STEP & FootNum
End Sub

Function GetResPicture(ID As Integer, PicType As PicTypes) As IPictureDisp
    Set GetResPicture = LoadResPicture(ID, PicType)
End Function

Sub GetSettings()
    BoxesInRow = CInt(GetProfileString(H_WIDTH, 8))
    BoxesInColumn = CInt(GetProfileString(H_HEIGHT, 8))
    MineCount = CInt(GetProfileString(H_MINES, 10))
    AllowQMarks = CBool(GetProfileString(H_MARK, "True"))
    AllowColor = CBool(GetProfileString(H_COL, "True"))
    AllowSound = CBool(GetProfileString(H_SND, "True"))
End Sub

Sub WriteSettings()
    WriteProfileString H_WIDTH, CStr(BoxesInRow)
    WriteProfileString H_HEIGHT, CStr(BoxesInColumn)
    WriteProfileString H_MINES, CStr(MineCount)
    WriteProfileString H_MARK, CStr(AllowQMarks)
    WriteProfileString H_COL, CStr(AllowColor)
    WriteProfileString H_SND, CStr(AllowSound)
End Sub

Function GetProfileString(Heading As String, Default As String) As String
Dim RetStr As String * 100, RetVal As Long
    RetVal = GetPrivateProfileString(SECTION, Heading, Default, RetStr, 100, AppPath & INI_FILE)
    GetProfileString = Left(RetStr, RetVal)
End Function

Sub WriteProfileString(Heading As String, Data As String)
    WritePrivateProfileString SECTION, Heading, Data, AppPath & INI_FILE
End Sub

' ----------------------------------------------------------------------
' 3D Wrap-Around Affect
' ----------------------------------------------------------------------

Sub DrawLine(Source As Object, Pt1 As PointAPI, Pt2 As PointAPI)
    Source.Line (Pt1.x, Pt1.y)-(Pt2.x, Pt2.y)
End Sub

Sub MakeControls3D(Source As Object)
Dim i As Integer
Dim BorderWidth As Integer, BoxType As BoxTypes, BoxTypeLetter As String
    On Error Resume Next
    Source.ForeColor = FACE
    Source.Line (0, 0)-(Source.Width, Source.Height), , B
    For i = 0 To Source.Controls.Count - 1
        BoxTypeLetter = Left(Source.Controls(i).Tag, 1)
        Select Case BoxTypeLetter
            Case TAG_INSET
                BoxType = Inset
            Case TAG_OUTSET
                BoxType = outset
        End Select
        BorderWidth = Val(Mid(Source.Controls(i).Tag, 2))
        Make3D Source.Controls(i).Container, Source.Controls(i), BorderWidth, BoxType
    Next i
End Sub

Function GetRect(Pt1 As PointAPI, Pt2 As PointAPI) As Rect
    GetRect.Pt1 = Pt1
    GetRect.Pt2 = Pt2
End Function

Sub Make3D(Source As Object, SourceObj As Control, BorderWidth As Integer, BoxType As BoxTypes)
Dim i As Integer, InTwips As Long
    For i = 1 To BorderWidth
        DrawBox Source, GetRect(GetPoint(SourceObj.Left - i, SourceObj.Top - i), GetPoint(SourceObj.Left + SourceObj.Width + i - 1, SourceObj.Top + SourceObj.Height + i - 1)), BoxType
    Next i
End Sub

Sub DrawBox(Source As Object, Box As Rect, BoxType As BoxTypes)
Dim Col1 As Long, Col2 As Long
    'If Source.Name = "pctPanel" Then Stop
    Select Case BoxType
        Case Inset
            Col1 = HIGHLIGHT
            Col2 = SHADOW
        Case outset
            Col1 = SHADOW
            Col2 = HIGHLIGHT
    End Select
    Source.ForeColor = Col1
    Source.Line (Box.Pt1.x, Box.Pt1.y)-(Box.Pt1.x, Box.Pt2.y)
    Source.Line (Box.Pt1.x, Box.Pt1.y)-(Box.Pt2.x, Box.Pt1.y)
    Source.ForeColor = Col2
    Source.Line (Box.Pt2.x, Box.Pt2.y)-(Box.Pt1.x, Box.Pt2.y)
    Source.Line (Box.Pt2.x, Box.Pt2.y)-(Box.Pt2.x, Box.Pt1.y)
End Sub

' ----------------------------------------------------------------------
' End 3D Wrap-Around Affect
' ----------------------------------------------------------------------

Function GetPoint(x As Integer, y As Integer) As PointAPI
    GetPoint.x = x
    GetPoint.y = y
End Function
