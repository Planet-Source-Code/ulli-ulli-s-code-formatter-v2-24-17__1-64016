VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z 'we're 32 bits

'Mousepointer Class
'One advantage of using this wrapper is that it will automatically reset the mousepointer
'on terminate (whatever caused the termination)

Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Public Enum CursorName

    MpNone = 0

    'standard
    Arrow = 32512
    TextPos = 32513
    Sand = 32514
    CrossHair = 32615
    ArrowUp = 32616
    Pen = 32631
    SquareIcon = 32641
    SizeNWSE = 32642
    SizeNESW = 32643
    SizeWE = 32644
    SizeNS = 32645
    SizeNSWE = 32646
    NoDrop = 32648
    RightHand = 32649
    ArrowSand = 32650
    ArrowQuestion = 32651

    'wheel mouse
    NorthDotSouth = 32652
    WestDotEast = 32653
    NSDotWE = 32654
    NorthDot = 32655
    SouthDot = 32656
    WestDot = 32657
    EastDot = 32658
    NorthWestDot = 32659
    NorthEastDot = 32660
    SouthWestDot = 32661
    SouthEastDot = 32662

    'special
    ArrowCD = 32663

    'probably there are more....

End Enum
#If False Then 'Spoof to preserve Enum capitalization
Private MpNone, Arrow, TextPos, Sand, CrossHair, ArrowUp, Pen, SquareIcon, SizeNWSE, SizeNESW, SizeWE, SizeNS, SizeNSWE, NoDrop, RightHand, ArrowSand, ArrowQuestion, NorthDotSouth, WestDotEast, NSDotWE, NorthDot, SouthDot, WestDot, EastDot, NorthWestDot, NorthEastDot, SouthWestDot, SouthEastDot, ArrowCD
#End If

Private Sub Class_Terminate()

    Do
    Loop Until ShowCursor(True) >= 0
    'default property
    SetPointerIcon Arrow
    Screen.MousePointer = vbNormal

End Sub

Public Property Let MP(Pointer As MousePointerConstants)

    Screen.MousePointer = Pointer

End Property

Public Property Get MP() As MousePointerConstants
Attribute MP.VB_UserMemId = 0

    MP = Screen.MousePointer

End Property

Public Sub SetPointerIcon(Icon As CursorName)

    If IsWindowsSuitable Then
        'prevent discrimination of left handed people
        'we have no LeftHand resource so we just leave the cursor as it is
        If Icon <> RightHand Or MouseButtonsSwapped = False Then
            SetCursor LoadCursor(0, Icon)
        End If
    End If

End Sub

':) Ulli's VB Code Formatter V2.24.11 (2008-Apr-11 10:26)  Decl: 55  Code: 37  Total: 92 Lines
':) CommentOnly: 12 (13%)  Commented: 2 (2,2%)  Filled: 70 (76,1%)  Empty: 22 (23,9%)  Max Logic Depth: 3
