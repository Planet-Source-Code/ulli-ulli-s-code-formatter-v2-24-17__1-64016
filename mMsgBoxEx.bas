Attribute VB_Name = "mMsgBoxEx"

'Copied from C:\Programme\Microsoft Visual Studio\VB98\Vb Sources\VBCompanion\mMsgBoxEx.bas
'MsgBoxEx  Extended Message Box

'How to:

'   You can use this message box just like you would use a standard VB MsgBox. However, the
'   following extras can be set and support named parameters:

'   TimeOut     If this value is positive then it is taken as timeout in milliseconds.
'               A zero value will calculate the timeout depending on content length.
'               A negative value will disable timeout.

'   PosX        A non-negative value will be used to center the box on this position.
'   PosY        (-1) will center the box on the parent form.
'               (-2) will center the box on the cursor hotspot.
'               In all cases positioning is limited to be within the screen.

'   OffsetX     Number of pixels to offset the box from the cursor hotspot, this is only
'   OffsetY     valid with PosX or PosY = (-2) and allows you to center a messagebox button
'               on the cursor; you will have to find the values for these params empirically.

'   Icon        You can supply an icon to be used instead of the standard message box icons.

'   OCapt       If you want to alter the button captions you must supply the original captions,
'               separated by a vertical bar character for example "&Yes|&No|Cancel". Don't
'               forget the ampersand if the original caption contains a shortcut character.

'       NEW     If you know the resource identifier strings for the captions you can also use
'               those; the example above then becomes "805|806|801", these three numbers being
'               the identifiers for the button caption text resources which are located in user32
'               (see Function Substitute). The advantage of using resource identifiers is
'               that you can be sure that you get the correct localized text.

'               "800"   OK
'               "801"   Cancel without shortcut
'               "802"   Cancel with shortcut
'               "803"   Repeat without shortcut
'               "804"   Ignore
'               "805"   Yes
'               "806"   No
'               "807"   Close
'               "808"   Help
'               "809"   Repeat with shortcut
'               "810"   Continue

'   NCapt       The replacement captions in a similar format as above. There will be a one to
'               one relationship between OCapt and NCapt.

'               If you only want to replace some, not all, captions then you need not supply
'               the others, for example:

'               OCapt:="Cancel"
'               NCapt:="Break"

'               will only replace the 'Cancel'-caption on the third button.
'       NEW

'   Sound       Use the MsgBox style constants to output the corresponding windows sound or
'               use negative values of the format -(frequency.duration) to beep
'               the computer speaker; eg -440.03 will beep @ 440 Hz for 0.03 secs

'   The return values are identical to those returned by the standard message box and are not
'   affected by a button caption replacement. A zero return value indicates that the box was
'   timed out without user interaction.

'   vbYesNo-message boxes do not react to WM_CLOSE so we try and send it the shortcut char
'   for 'No' instead of sending WM_CLOSE

Option Explicit

Private Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
Private Declare Function EndDialog Lib "user32" (ByVal hDlg As Long, ByVal nResult As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Long) As Long
Private Declare Function FindCursorPos Lib "user32" Alias "GetCursorPos" (lpPoint As typPT) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal ParenthWnd As Long, ByVal ChildhWnd As Long, ByVal ClassName As String, ByVal Caption As String) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function SetWindowText Lib "user32.dll" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, ByVal wID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
'Copied from C:\Programme\Microsoft Visual Studio\VB98\Vb Sources\VBCompanion\mSubclass.bas
Public Declare Function Beeper Lib "kernel32" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
'End Copy

Private Const HCBT_ACTIVATE     As Long = 5
Private Const STM_SETICON       As Long = &H170
Private Const SWP_NOSIZE        As Long = 1
Private Const SWP_NOZORDER      As Long = 4
Private Const SWP_NOACTIVATE    As Long = 16
Private Const SWP_COMBINED      As Long = SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
Private Const WH_CBT            As Long = 5
Public Const EM_GETLINECOUNT    As Long = &HBA
Public Const EM_UNDO            As Long = &HC7
Public Const EM_EMPTYUNDOBUFFER As Long = &HCD

Private Const ButtonClassName   As String = "Button"
Private Const MsgBoxClassName   As String = "#32770"
Private Const IconClassName     As String = "Static"
Private Const IconBits          As Long = &H70
Private Const Zero              As Long = 0
Private Const Nil               As Long = Zero

Private CurrMsgBoxTitle         As String

Private Type typRC
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private Type typPT
    X       As Long
    Y       As Long
End Type

'message box button caption resource ids (underline stands for an ampersand)
Public Const OK                 As String = "800" 'OK
Public Const Abbrechen          As String = "801" 'Cancel without shortcut
Public Const A_bbrechen         As String = "802" 'Cancel with shortcut
Public Const Wiederholen        As String = "803" 'Repeat without shortcut
Public Const Ignorieren         As String = "804" 'Ignore
Public Const Ja                 As String = "805" 'Yes
Public Const Nein               As String = "806" 'No
Public Const Schliessen         As String = "807" 'Close
Public Const Hilfe              As String = "808" 'Help
Public Const Wie_derholen       As String = "809" 'Repeat with shortcut
Public Const Weiter             As String = "810" 'Continue

Private Const ScreenMargin      As Long = 5
Private fpNextHook              As Long 'far pointer to next hook
Private hWndBox                 As Long 'window handle of msgbox
Private myIcon                  As Long 'object variable having user icon
Private IsYesNoBox              As Boolean
Private BoxPos                  As typPT
Private Offset                  As typPT
Private myOrigButtonCaptions()  As String 'original button captions or resource id's
Private myNewButtonCaptions()   As String 'replacement button captions

Public Function GetResourceString(ModuleName As String, ByVal StringNumber As Long) As String

    GetResourceString = Space$(1024)
    GetResourceString = Left$(GetResourceString, LoadString(GetModuleHandle(ModuleName), StringNumber, GetResourceString, Len(GetResourceString) - 1))

End Function

Private Function MakePt(X As Long, Y As Long) As typPT

    With MakePt
        .X = X
        .Y = Y
    End With 'MAKEPT

End Function

Private Function MsgBoxCallback(ByVal nCode As Long, ByVal wParam As Long, lParam As Long) As Long

  'cbt call back - modifies the message box

  Dim ClassName     As String * 64
  Dim ParentRect    As typRC
  Dim MsgBoxRect    As typRC
  Dim CurPosn       As typPT
  Dim TmpPosn       As typPT

    MsgBoxCallback = CallNextHookEx(fpNextHook, nCode, wParam, lParam)
    If nCode = HCBT_ACTIVATE Then
        If Left$(ClassName, GetClassName(wParam, ClassName, Len(ClassName))) = MsgBoxClassName Then 'called by message box just opening
            hWndBox = wParam 'save hWnd for timer callback, so we don't have to find it again
            If myIcon Then 'replace icon
                SendMessage FindWindowEx(hWndBox, Nil, IconClassName, vbNullString), STM_SETICON, myIcon, ByVal Nil
            End If 'NOT MYOPACITY...
            GetWindowRect hWndBox, MsgBoxRect
            With MsgBoxRect
                TmpPosn = MakePt((.Right - .Left) / 2, (.Bottom - .Top) / 2) 'center of box
                With TmpPosn
                    If BoxPos.X = -1 Or BoxPos.Y = 1 Then 'center on parent form 'NOT NWOPACITY...
                        GetWindowRect GetParent(hWndBox), ParentRect 'NOT RIGHT$(STMP,...
                        BoxPos = MakePt((ParentRect.Left + ParentRect.Right) / 2 - .X, (ParentRect.Top + ParentRect.Bottom) / 2 - .Y)
                      ElseIf BoxPos.X = -2 Or BoxPos.Y = -2 Then 'center on cursor 'NOT BOXPOS.X...
                        FindCursorPos CurPosn
                        BoxPos = MakePt(CurPosn.X - .X + Offset.X, CurPosn.Y - .Y + Offset.Y)
                      Else 'use user settings 'NOT BOXPOS.X...
                        BoxPos = MakePt(BoxPos.X - .X, BoxPos.Y - .Y)
                    End If
                End With 'TMPPOSN
                'limit within screen
                TmpPosn = MakePt(Screen.Width / Screen.TwipsPerPixelX - (.Right - .Left) - ScreenMargin, Screen.Height / Screen.TwipsPerPixelY - (.Bottom - .Top) - ScreenMargin) 'max bottom right that would just fit the box
                With BoxPos
                    Select Case .X
                      Case Is < ScreenMargin
                        .X = ScreenMargin
                      Case Is > TmpPosn.X
                        .X = TmpPosn.X
                    End Select

                    Select Case .Y
                      Case Is < ScreenMargin
                        .Y = ScreenMargin
                      Case Is > TmpPosn.Y
                        .Y = TmpPosn.Y
                    End Select
                    SetWindowPos hWndBox, Nil, .X, .Y, Nil, Nil, SWP_COMBINED 'move box to proper position
                    For .X = Zero To UBound(myNewButtonCaptions) 'replace button captions
                        If myNewButtonCaptions(.X) <> vbNullString And myOrigButtonCaptions(.X) <> vbNullString Then
                            SetWindowText FindWindowEx(hWndBox, Nil, ButtonClassName, Substitute(myOrigButtonCaptions(.X))), myNewButtonCaptions(.X)
                        End If
                    Next .X
                End With 'BOXPOS
            End With 'MSGBOXRECT
        End If
    End If

End Function

Public Function MsgBoxEx(ByVal Prompt As String, _
                         Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
                         Optional ByVal Title As String = vbNullString, _
                         Optional ByVal TimeOut As Long = Zero, _
                         Optional ByVal PosX As Long = -1, _
                         Optional ByVal PosY As Long = -1, _
                         Optional ByVal OffsetX As Long = 0, _
                         Optional ByVal OffsetY As Long = 0, _
                         Optional ByVal Icon As Long = Zero, _
                         Optional ByVal OCapt As String = vbNullString, _
                         Optional ByVal NCapt As String = vbNullString, _
                         Optional ByVal Sound As Single = 0) _
                         As VbMsgBoxResult

  Dim TimerId       As Long

    myOrigButtonCaptions = Split(OCapt, "|") 'split original captions
    If NCapt = vbNullString Then
        ReDim myNewButtonCaptions(Zero) 'make one empty element
      Else 'NOT NCAPT...
        myNewButtonCaptions = Split(NCapt, "|") 'split the new captions
    End If
    ReDim Preserve myOrigButtonCaptions(Zero To UBound(myNewButtonCaptions)) 'make'm both equal number of elements
    BoxPos = MakePt(PosX, PosY)
    Offset = MakePt(OffsetX, OffsetY)
    myIcon = Icon

    If Title = vbNullString Then
        CurrMsgBoxTitle = App.Title
      Else 'NOT TITLE...
        CurrMsgBoxTitle = Title
    End If
    If TimeOut = Zero Then
        TimeOut = (Len(Trim$(Prompt)) + Len(Trim$(CurrMsgBoxTitle)) + 25) * 40 'adjust timeout depending on prompt length
    End If
    If myIcon And ((Buttons And IconBits) = Zero) Then 'wants a custom icon but has none in the button bits
        Buttons = Buttons Or vbInformation 'so give him one to replace
    End If

    'play sound if any
    Select Case Sound
      Case Is > Zero
        MessageBeep CLng(Fix(Sound))
      Case Is < Zero
        Beeper CLng(Fix(Abs(Sound))), CLng(Fix(Abs(Sound - Fix(Sound)) * 1000))
    End Select

    'and now set the hook and a timer, display the box, and tidy up
    fpNextHook = SetWindowsHookEx(WH_CBT, AddressOf MsgBoxCallback, App.hInstance, GetCurrentThreadId)
    IsYesNoBox = ((Buttons And 15) = vbYesNo)
    If TimeOut > Zero Then 'user wants automatic timeout
        TimerId = SetTimer(Nil, Nil, TimeOut, AddressOf TimerCallback)
    End If
    MsgBoxEx = MsgBox(Prompt, Buttons, CurrMsgBoxTitle)
    If TimerId Then 'the timer was started so we must kill it
        KillTimer Nil, TimerId
    End If
    UnhookWindowsHookEx fpNextHook 'and finally unhook the callback

    If hWndBox = Zero Then 'closed by timer
        MsgBoxEx = Zero 'return zero
    End If

End Function

Private Function Substitute(BtnCaption As String) As String

    If IsNumeric(BtnCaption) Then
        Substitute = GetResourceString("user32", Val(BtnCaption))
    End If
    If Len(Substitute) = 0 Then 'resource not found or text not numeric - return original text
        Substitute = BtnCaption
    End If

End Function

Private Sub TimerCallback(hWnd As Long, uMsg As Long, idEvent As Long, dwTime As Long)

  'Copied from C:\Programme\Microsoft Visual Studio\VB98\Vb Sources\MsgBoxEx\mMsgBoxEx.bas
  'close timed message box

    If hWndBox Then
        EndDialog hWndBox, 0
        hWndBox = Zero 'indication that the timer has ticked
    End If
    'End Copy

End Sub

'End Copy

':) Ulli's VB Code Formatter V2.24.11 (2008-Apr-11 10:25)  Decl: 145  Code: 170  Total: 315 Lines
':) CommentOnly: 64 (20,3%)  Commented: 48 (15,2%)  Filled: 260 (82,5%)  Empty: 55 (17,5%)  Max Logic Depth: 7
