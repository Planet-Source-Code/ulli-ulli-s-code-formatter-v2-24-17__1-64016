Attribute VB_Name = "mAPI"
Option Explicit
DefLng A-Z 'we're 32 bit

Private Declare Function WinVersion Lib "kernel32" Alias "GetVersion" () As Long 'Windows Version
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As Rect, ByVal wFormat As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As Any) As Long
Public Declare Function IsCharAlphaNumeric Lib "user32" Alias "IsCharAlphaNumericA" (ByVal cChar As Byte) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd, ByVal wMsg, ByVal wParam, lParam As Any) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function MessageBeep Lib "user32" (ByVal wType As VbMsgBoxStyle) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As String) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Integer, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long 'used to get the 'other button
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey, ByVal lpSubKey As String, ByVal ulOptions, ByVal samDesired, phkResult) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey, ByVal lpValueName As String, ByVal lpReserved, lpType, lpData As Any, lpcbData) As Long
Public Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Sub RegCloseKey Lib "advapi32.dll" (ByVal hKey)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'For Find context
Public Type POINTAPI
    X   As Long
    Y   As Long
End Type

Public Type Rect
    Left        As Long
    Top         As Long
    Right       As Long
    Bottom      As Long
End Type

Private Type DRAWITEMSTRUCT
    CtlType     As Long
    CtlID       As Long
    ItemID      As Long
    ItemAction  As Long
    ItemState   As Long
    hWndItem    As Long
    hDC         As Long
    rcItem      As Rect
    ItemData    As Long
End Type

Public Enum PublAPIConstants
    SW_MINIMIZE = 6
    SW_RESTORE = 9
    SWP_TOPMOST = -1
    SWP_NOACTIVATE = &H10
    SWP_NOMOVE = 2
    SWP_NOSIZE = 1
    SWP_COMBINED = SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
    HKEY_CURRENT_USER = &H80000001
    KEY_QUERY_VALUE = 1
    REG_OPTION_RESERVED = 0
    ERROR_NONE = 0
    SleepTime = 555 'Splash duration
    LB_GETTOPINDEX = &H18E
    LB_SETTOPINDEX = &H197
    HTCAPTION = 2
    WM_NCLBUTTONDOWN = &HA1
    WM_LBUTTONDBLCLK = &H203
    SM_SWAPBUTTON = 23
End Enum

Private Enum PrivApiConstants
    RequiredWinVersion = 5 'WinXP
    CS_DROPSHADOW = &H20000
    GCL_STYLE = -26
    BK_TRANSPARENT = 1
    GWL_WNDPROC = -4
    ODT_BUTTON = 4
    ODS_SELECTED = &H1
    WM_DESTROY = &H2
    WM_DRAWITEM = &H2B
    DT_HCENTER = &H1
    DT_TOP = &H0
    DT_VCENTER = &H4
    DT_BOTTOM = &H8
    DT_SINGLELINE = &H20
    SW_SHOWNORMAL = 1
    SE_NO_ERROR = 33 'Values below 33 are error returns
End Enum
#If False Then
Private CS_DROPSHADOW, GCL_STYLE, BK_TRANSPARENT, GWL_WNDPROC, ODT_BUTTON, ODS_SELECTED, WM_DESTROY, WM_DRAWITEM, DT_HCENTER, DT_TOP, DT_VCENTER, DT_BOTTOM, DT_SINGLELINE, SW_SHOWNORMAL, SE_NO_ERROR
#End If

'Stop conditions
Public Enum StopWhen
    Never
    IfNecessary
    Always
End Enum
#If False Then 'Spoof to preserve Enum capitalization
Private Never, IfNecessary, Always
#End If

'Win XP Look
Public Const XPLookAPIProto   As String = "Private Declare Sub InitCommonControls Lib ""comctl32"" ()"
Public Const XPLookAPICall    As String = "InitCommonControls"
'standard manifest file
Public Const XPLookXML        As String = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbCrLf & _
                                          "<assembly xmlns=""urn:schemas-microsoft-com:asm.v1"" manifestVersion=""1.0"">" & vbCrLf & _
                                          "<assemblyIdentity version=""1.0.0.0"" processorArchitecture=""X86"" name=""°"" type=""win32"" />" & vbCrLf & _
                                          "<description>XP-Look - Created by Ulli's Code Formatter</description>" & vbCrLf & _
                                          "<dependency>" & vbCrLf & _
                                          "<dependentAssembly>" & vbCrLf & _
                                          "<assemblyIdentity type=""win32"" name=""Microsoft.Windows.Common-Controls"" version=""6.0.0.0"" processorArchitecture=""X86"" publicKeyToken=""6595b64144ccf1df"" language=""*"" />" & vbCrLf & _
                                          "</dependentAssembly>" & vbCrLf & _
                                          "</dependency>" & vbCrLf & _
                                          "</assembly>"

Public Enum AlignText
    AlignTop = DT_TOP
    AlignCenter = DT_VCENTER
    AlignBottom = DT_BOTTOM
    AlignThreeD = DT_VCENTER Or DT_BOTTOM
End Enum
#If False Then 'Spoof to preserve Enum capitalization
Private AlignTop, AlignCenter, AlignBottom, AlignThreeD
#End If

Public Type PageBoundaries
    Left         As Long
    Right        As Long
    Top          As Long
    Bottom       As Long
    PunchX       As Long
    PunchY       As Long
End Type

'custom property names
Private Const PropCustom            As String = "UMGCustom"
Private Const PropForeColor         As String = "UMGForeColor"
Private Const PropAlign             As String = "UMGAlign"
Private Const PropSubclass          As String = "UMGDrawProc"

'Misc
Private hWndParent                  As Long
Private OldProcPtr                  As Long
Public VBInstance                   As VBIDE.VBE
Public Copyright                    As String
Public Proj                         As VBProject
Public Compo                        As VBComponent
Public Pane                         As VBIDE.CodePane
Public Pt                           As POINTAPI
Public MPIcon                       As New cMP
Public MouseButtonsSwapped          As Boolean
Public UserName                     As String
Public UndoBuffers()                As String
Public UndoTitles()                 As String
Public NumPanels                    As Long
Public PrintLineLen                 As Long
Public Const LnLen                  As Long = 4 'Linenumber length 4 digits
Public Const PrintMark              As String = ">>>> " 'same length as LnLen + 1 space
Public Const LenPunchMark           As Long = 40 'pixels
Public Const ElipsisChar            As String = "…"
Public Const Spce                   As String = " "
Public Const NullStr                As String = ""
Public PBEven                       As PageBoundaries
Public PBOdd                        As PageBoundaries
Public PrintLineHeight              As Long
Public PrintCharWidth               As Long
Public PrinterItalEnabled           As Boolean
Public PrinterBoldEnabled           As Boolean
Public StrucRequested               As Boolean
Public StrFnctsRequested            As Boolean
Public StringCboListIndex           As Long
Public SortRequested                As Boolean
Public XPLookRequested              As Long
Public ColorRequested               As Boolean
Public BookRequested                As Boolean
Public TypeSuffRequested            As Boolean
Public StopRequested                As Boolean
Public WithStationary               As Boolean
Public BreakLoop                    As Boolean
Public HasPrinter                   As Boolean
Public PrintingOK                   As Boolean
Public KillDoc                      As Boolean
Public IDEFontName                  As String
Public IDEFontSize                  As Long
Public MyFontName                   As String
Public MyFontSize                   As Single
Public FullTabWidth                 As Long
Public StF                          As Integer
Public PaS                          As Integer
Public Isc                          As Integer
Public InM                          As Integer
Public Sep                          As Integer
Public HfI                          As Integer
Public bStrFncts()                  As Boolean 'there are 28 modifiable String functions
Public NumSelected                  As Long
Public InsertComments               As Boolean
Public PauseAfterScan               As StopWhen

Public Sub AddShadow(Frm As Form)

    With Frm
        SetClassLong .hWnd, GCL_STYLE, GetClassLong(.hWnd, GCL_STYLE) Or CS_DROPSHADOW
    End With 'FRM

End Sub

Public Function AllocString(ByVal Size As Long) As String

    AllocString = String$(Size, 0)

End Function

Public Function AppDetails() As String

    With App
        AppDetails = .ProductName & " V" & .Major & "." & .Minor & "." & .Revision
    End With 'APP

End Function

Public Function CreateTooltips(Frm As Form) As Collection

  'called on form_load from each individual form to create the custom tooltips

  Dim TTColl    As Collection
  Dim Tooltip   As cToolTip
  Dim Control   As Control
  Dim TtTxt     As String
  Dim CollKey   As String

    Set TTColl = New Collection
    For Each Control In Frm.Controls 'cycle thru all controls
        With Control
            On Error Resume Next 'in case the control has no tooltiptext property
                TtTxt = Trim$(.ToolTipText) 'try to access that property
            On Error GoTo 0
            If Len(TtTxt) Then 'use that to create the custom tooltip
                CollKey = .Name
                On Error Resume Next 'in case control is not in an array of controls and therefore has no index property
                    CollKey = CollKey & "(" & .Index & ")"
                On Error GoTo 0
                Set Tooltip = New cToolTip
                If Tooltip.Create(Control, TtTxt, TTBalloonAlways, (TypeName(Control) = "TextBox"), , , &HB00000, &HFFF0F0) Then
                    TTColl.Add Tooltip, CollKey 'to keep a reference to the current tool tip class instance (prevent it from being destroyed)
                    .ToolTipText = NullStr 'kill tooltiptext so we don't get two tips
                End If
            End If
        End With 'CONTROL
    Next Control
    Set CreateTooltips = TTColl

End Function

Public Sub Dec(What As Long, Optional By As Long = 1)

    What = What - By

End Sub

Private Function DrawButtonProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

  Dim DrawSpec    As DRAWITEMSTRUCT
  Dim ButtonText  As String
  Dim Align       As AlignText

    OldProcPtr = GetProp(hWnd, PropSubclass)
    DrawButtonProc = CallWindowProc(OldProcPtr, hWnd, wMsg, wParam, lParam)
    Select Case wMsg
      Case WM_DRAWITEM
        CopyMemory DrawSpec, ByVal lParam, Len(DrawSpec) 'get draw specification
        With DrawSpec
            If .CtlType = ODT_BUTTON Then 'it's a button
                If GetProp(.hWndItem, PropCustom) Then 'and it needs drawing
                    Align = GetProp(.hWndItem, PropAlign) 'so get the text alignment
                    With .rcItem
                        Select Case Align
                          Case DT_TOP 'should only be used when the button also has a pic
                            Inc .Top, 4
                          Case DT_BOTTOM 'should only be used when the button also has a pic
                            Dec .Bottom, 4
                          Case AlignThreeD
                            'offset the draw rect
                            Dec .Left
                            Dec .Top
                            Dec .Right
                            Dec .Bottom
                            Align = AlignCenter
                        End Select
                        If (DrawSpec.ItemState And ODS_SELECTED) = ODS_SELECTED Then
                            'Button is in down state - offset the draw rect
                            Inc .Left
                            Inc .Top
                            Inc .Right
                            Inc .Bottom
                        End If
                    End With '.RCITEM
                    ButtonText = AllocString(255)
                    GetWindowText .hWndItem, ButtonText, Len(ButtonText) 'get the text
                    ButtonText = Left$(ButtonText, InStr(ButtonText, Chr$(0)) - 1) 'and trim
                    SetBkMode .hDC, BK_TRANSPARENT
                    SetTextColor .hDC, GetProp(.hWndItem, PropForeColor) 'set textcolor from property
                    DrawText .hDC, ButtonText, Len(ButtonText), .rcItem, DT_SINGLELINE Or DT_HCENTER Or Align
                End If
            End If
        End With 'DRAWSPEC
      Case WM_DESTROY
        SetWindowLong hWnd, GWL_WNDPROC, OldProcPtr
        RemoveProp hWnd, PropSubclass
    End Select

End Function

Public Function DriveSerialNumber(ByVal Drive As String) As Long

  'usage: SN = DriveSerialNumber("C")

  Dim s As String * 32

    GetVolumeInformation Drive & ":\", s, Len(s), DriveSerialNumber, 0, 0, Len(s), 16

End Function

Public Function GetModuleCount() As Long

    For Each Proj In VBInstance.VBProjects 'count components is this (these) project(s)
        For Each Compo In Proj.VBComponents
            Dec GetModuleCount, (Compo.Type <> vbext_ct_RelatedDocument And Compo.Type <> vbext_ct_ResFile)
    Next Compo, Proj

End Function

Public Sub Inc(What As Long, Optional By As Long = 1)

    What = What + By

End Sub

Public Function IsWindowsSuitable() As Boolean

    IsWindowsSuitable = ((WinVersion And &HFF&) >= RequiredWinVersion)

End Function

Public Function OneOrMany(Word As String, Num As Long) As String

    If Num = 1 Then
        OneOrMany = Word
      Else 'NOT NUM...
        OneOrMany = Word & "s"
    End If

End Function

Public Sub RemoveShadow(Frm As Form)

    With Frm
        SetClassLong .hWnd, GCL_STYLE, GetClassLong(.hWnd, GCL_STYLE) And Not CS_DROPSHADOW
    End With 'FRM

End Sub

Public Sub SendMeMail(ByVal FromhWnd As Long, Subject As String)

    If ShellExecute(FromhWnd, NullStr, "mailto:UMGEDV@Yahoo.COM?subject=" & Subject & " &body=<b>Hi Ulli,</b><br><br>{your message}<br><br>Best regards from <br><b>" & UserName, NullStr, App.Path, SW_SHOWNORMAL) < SE_NO_ERROR Then
        MessageBeep vbCritical
        MsgBox "Cannot send Mail from this System.", vbCritical, "Mail disabled/not installed"
    End If

End Sub

Public Sub SetButtonForeColor(Button As CommandButton, ByVal ForeColor As Long, Optional ByVal Alignment As AlignText = AlignCenter)

    With Button
        hWndParent = GetParent(.hWnd)
        If GetProp(hWndParent, PropSubclass) = 0 Then 'parent window not yet subclassed
            SetProp hWndParent, PropSubclass, GetWindowLong(hWndParent, GWL_WNDPROC)
            SetWindowLong hWndParent, GWL_WNDPROC, AddressOf DrawButtonProc
        End If
        SetProp .hWnd, PropCustom, True
        SetProp .hWnd, PropForeColor, ForeColor
        SetProp .hWnd, PropAlign, Alignment
        .Refresh
    End With 'BUTTON

End Sub

Public Sub StoreSettings(pStF As Integer, pPaS As Integer, pIsC As Integer, pInM As Integer, pSep As Integer, pHfi As Integer)

    StF = pStF 'StringFuncts requested
    PaS = pPaS 'Pause after Scan
    Isc = pIsC 'Insert comments
    InM = pInM 'Insert marks
    Sep = pSep 'Separate compound lines at colon
    HfI = pHfi 'HalfIndent

End Sub

Public Sub UnsetButtonForeColor(Button As CommandButton)

    With Button
        RemoveProp .hWnd, PropCustom
        RemoveProp .hWnd, PropForeColor
        RemoveProp .hWnd, PropAlign
        '...but the parent window having this button remains subclassed until it is destroyed
        .Refresh
    End With 'BUTTON

End Sub

Public Sub UpdateTooltip(Col As Collection, Cntrl As Control)

  'Update Custom Tooltip for Cntrl

  Dim Indx     As String

    On Error Resume Next
        Indx = "(" & Cntrl.Index & ")"
    On Error GoTo 0
    With Col(Cntrl.Name & Indx) 'finds reference to  tooltip class instance
        .Create Cntrl, Trim$(Cntrl.ToolTipText), .Style, .Centered, , , .ForeCol, .BackCol
    End With 'COL(CNTRL.NAME
    Cntrl.ToolTipText = NullStr

End Sub

':) Ulli's VB Code Formatter V2.24.11 (2008-Apr-11 10:26)  Decl: 211  Code: 230  Total: 441 Lines
':) CommentOnly: 14 (3,2%)  Commented: 47 (10,7%)  Filled: 370 (83,9%)  Empty: 71 (16,1%)  Max Logic Depth: 7
