VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Custom Tooltip Class

'This class was inspired by code by Eidos (found at PSC some time ago) and others.

'Jan02 2003   UMG

'Three new options have been added - display tooltip always / only if parent form is active / None
'see TTStyle.

'Added missing Style private property variable.
'Rearranged code a little.

Private Declare Sub InitCommonControls Lib "comctl32" ()

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private TThWnd                      As Long 'Tooltip window handle
Private Const ToolTipWindowClassName As String = "Tooltips_Class32"
Private Const CW_USEDEFAULT         As Long = &H80000000
Private Const TTS_STANDARD          As Long = 0
Private Const TTS_BALLOON           As Long = &H40
Private Const TTS_ALWAYSTIP         As Long = 1 'display even if parent window is inactive
Private Const TTS_NOPREFIX          As Long = 2 'does not remove "&" from text
Private Const TTDT_AUTOPOP          As Long = 2
Private Const TTDT_INITIAL          As Long = 3

Public Enum TTStyle
    TTStandardIfActive = TTS_STANDARD                   'suppress if parent form is not active
    TTBalloonIfActive = TTS_BALLOON                     'suppress if parent form is not active
    TTStandardAlways = TTS_STANDARD Or TTS_ALWAYSTIP    'display even if parent form is not active
    TTBalloonAlways = TTS_BALLOON Or TTS_ALWAYSTIP      'display even if parent form is not active
    TTNone = -1                                         'kill tooltip (this is simply treated as illegal)
End Enum
#If False Then 'Spoof to preserve Enum capitalization
Private TTStandardIfActive, TTBalloonIfActive, TTStandardAlways, TTBalloonAlways, TTNone
#End If
Public Enum TTIcon
    TTIconNone = 0
    TTIconInfo = 1         'i in white balloon
    TTIconWarning = 2      '! in yellow triangle
    TTIconError = 3        'x in red circle
    'all have a light gray shadow so be careful when selecting the ToolTip BackColor
End Enum
#If False Then 'Spoof to preserve Enum capitalization
Private TTIconNone, TTIconInfo, TTIconWarning, TTIconError
#End If

'my properties
Private myStyle                     As TTStyle
Private myIcon                      As TTIcon
Private myForeColor                 As Long
Private myBackColor                 As Long
Private myTitle                     As String 'has the current title
Private myHoverTime                 As Long 'time im millisecs (-1 = use default)
Private myPopupTime                 As Long 'time im millisecs (-1 = use default)
Private myInitialText               As Variant 'has the initial text
Private myInitialTitle              As Variant 'has the initial title

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Const WM_USER               As Long = &H400
Private Const TTM_SETDELAYTIME      As Long = WM_USER + 3
Private Const TTM_ADDTOOL           As Long = WM_USER + 4
Private Const TTM_SETTIPBKCOLOR     As Long = WM_USER + 19
Private Const TTM_SETTIPTEXTCOLOR   As Long = WM_USER + 20
Private Const TTM_SETTITLE          As Long = WM_USER + 32

Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As Rectangle) As Long
Private Type Rectangle
    Left        As Long
    Top         As Long
    Right       As Long
    Bottom      As Long
End Type

Private Type ToolInfo
    ttSize      As Long
    myFlags     As Long
    ttParhWnd   As Long
    ttId        As Long
    ParentRect  As Rectangle
    hInstance   As Long
    myText      As String
    lParam      As Long
End Type
Private ToolInfo                    As ToolInfo

'tool property flag bits                             meaning
Private Const TTF_CENTERTIP         As Long = 2     'center tool on parent
Private Const TTF_SUBCLASS          As Long = &H10  'use implicit subclassing

Public Property Get BackCol() As Long

  'this returns the current tooltip backcolor

    BackCol = myBackColor

End Property

Public Property Get Centered() As Boolean

  'this returns the current tooltip alignment

    Centered = CBool(ToolInfo.myFlags And TTF_CENTERTIP)

End Property

Private Sub Class_Initialize()

    InitCommonControls 'doesn't matter that this is called for every class instance
    myStyle = TTNone

End Sub

Private Sub Class_Terminate()

  'kill tooltip window if one exists

    If TThWnd Then
        DestroyWindow TThWnd
        TThWnd = 0
    End If
    myStyle = TTNone

End Sub

Public Function Create(Parent As Control, _
                       Text As String, _
                       Optional ByVal Style As TTStyle = TTBalloonAlways, _
                       Optional ByVal Centered As Boolean = False, _
                       Optional ByVal Icon As TTIcon = TTIconNone, _
                       Optional Title As String = NullStr, _
                       Optional ByVal ForeColor As Long = vbButtonText, _
                       Optional ByVal BackColor As Long = vbInfoBackground, _
                       Optional ByVal HoverTime As Long = -1, _
                       Optional ByVal PopupTime As Long = -1) As Long

  'Create the tooltip window for parent control
  'This cannot create custom tooltips for hWnd-less controls

    Class_Terminate 'kill tooltip window if one exists
    With ToolInfo
        On Error Resume Next
            .ttParhWnd = Parent.hWnd
            If (Err = 0) And _
                (Style = TTBalloonAlways Or Style = TTStandardAlways Or Style = TTBalloonIfActive Or Style = TTStandardIfActive) And _
                (Icon = TTIconError Or Icon = TTIconInfo Or Icon = TTIconNone Or Icon = TTIconWarning) Then
                'the tooltip parent control has an hWnd and the params are acceptable
                .ttSize = Len(ToolInfo)
                .myFlags = TTF_SUBCLASS Or IIf(Centered, TTF_CENTERTIP, 0&)
                GetClientRect .ttParhWnd, .ParentRect
                .hInstance = App.hInstance
                myTitle = Title
                If myInitialTitle = Empty Then
                    myInitialTitle = myTitle
                End If
                .myText = Replace$(Text, "|", vbCrLf) 'the vertical bar is used as line break character
                If Len(myTitle) = 0 Then
                    .myText = Replace$(.myText, vbCrLf, Spce)
                End If
                If myInitialText = Empty Then
                    myInitialText = .myText
                End If
                If ForeColor < 0 Then
                    ForeColor = GetSysColor(ForeColor And &H7FFFFFFF)
                End If
                If BackColor < 0 Then
                    BackColor = GetSysColor(BackColor And &H7FFFFFFF)
                End If
                If ForeColor = BackColor Then
                    ForeColor = vbButtonText
                    BackColor = vbInfoBackground
                End If
                myForeColor = ForeColor
                myBackColor = BackColor
                myStyle = Style
                myIcon = Icon
                myHoverTime = HoverTime
                myPopupTime = PopupTime
                'create tooltip window and set it's properties
                TThWnd = CreateWindowEx(0&, ToolTipWindowClassName, NullStr, TTS_NOPREFIX Or Style, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, .ttParhWnd, 0&, .hInstance, 0&)
                SetWindowPos TThWnd, SWP_TOPMOST, 0&, 0&, 0&, 0&, SWP_COMBINED
                SendMessage TThWnd, TTM_ADDTOOL, 0&, ToolInfo
                SendMessage TThWnd, TTM_SETTITLE, Icon, ByVal myTitle
                SendMessage TThWnd, TTM_SETTIPTEXTCOLOR, myForeColor, ByVal 0&
                SendMessage TThWnd, TTM_SETTIPBKCOLOR, myBackColor, ByVal 0&
                SendMessage TThWnd, TTM_SETDELAYTIME, TTDT_INITIAL, ByVal myHoverTime
                SendMessage TThWnd, TTM_SETDELAYTIME, TTDT_AUTOPOP, ByVal myPopupTime
                Create = TThWnd
            End If
        On Error GoTo 0
    End With 'ToolInfo

End Function

Public Property Get ForeCol() As Long

  'this returns the current tooltip forecolor

    ForeCol = myForeColor

End Property

Public Property Get HoverTime() As Long

  'this returns the current mouse hover time time in millicecs (-1 for default)

    HoverTime = myHoverTime

End Property

Public Property Get Icon() As TTIcon

  'this returns the current tooltip icon

    Icon = myIcon

End Property

Public Property Get InitialText() As String

  'this returns the inital tooltip text, ie the one that was supplied on creation

    InitialText = myInitialText

End Property

Public Property Get InitialTitle() As String

  'this returns the inital tooltip title, ie the one that was supplied on creation

    InitialTitle = myInitialTitle

End Property

Public Property Get PopupTime() As Long

  'this returns the current max PopupTime time in millisecs (-1 for default)

    PopupTime = myPopupTime

End Property

Public Property Get Style() As TTStyle

  'this returns the current tooltip style

    Style = myStyle

End Property

Public Property Get Text() As String

  'this returns the current tooltip text

    Text = ToolInfo.myText

End Property

Public Property Get Title() As String

  'this returns the current tooltip Title

    Title = myTitle

End Property

':) Ulli's VB Code Formatter V2.24.11 (2008-Apr-11 10:26)  Decl: 91  Code: 179  Total: 270 Lines
':) CommentOnly: 28 (10,4%)  Commented: 24 (8,9%)  Filled: 201 (74,4%)  Empty: 69 (25,6%)  Max Logic Depth: 4
