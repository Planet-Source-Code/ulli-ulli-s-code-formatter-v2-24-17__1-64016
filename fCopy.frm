VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form fCopy 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Copy Facility"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8940
   ForeColor       =   &H00E0E0E0&
   Icon            =   "fCopy.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5970
   ScaleWidth      =   8940
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox picFontMeasure 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'Kein
      Height          =   165
      Left            =   150
      ScaleHeight     =   165
      ScaleWidth      =   285
      TabIndex        =   3
      Top             =   465
      Visible         =   0   'False
      Width           =   285
   End
   Begin MSComDlg.CommonDialog CDlg 
      Left            =   4305
      Top             =   2610
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Select Copy File"
      Filter          =   "Visual Basic Source Files|*.frm;*.cls;*.bas;*.ctl;*.pag;*.dsr;*.dob|All Files|*.*"
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   3630
      Top             =   2565
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fCopy.frx":0442
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fCopy.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fCopy.frx":0ECA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fCopy.frx":140E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fCopy.frx":1952
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fCopy.frx":1E96
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fCopy.frx":23DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fCopy.frx":26FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fCopy.frx":2C42
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtfSource 
      Height          =   5145
      Left            =   90
      TabIndex        =   0
      Top             =   735
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   9075
      _Version        =   393217
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"fCopy.frx":3186
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar tb 
      Height          =   345
      Left            =   90
      TabIndex        =   1
      Top             =   0
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open Copy File"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Copy"
            Object.ToolTipText     =   "Add to Clipboard   [Ctrl+V/C]"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "View"
            Object.ToolTipText     =   "View Clipboard"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Reset"
            Object.ToolTipText     =   "Clear Clipboard"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Find   [Ctrl+F/F3]"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste and exit"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancel"
            Object.ToolTipText     =   "Cancel and exit"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label lblFilename 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   2
      Top             =   420
      Width           =   8760
   End
End
Attribute VB_Name = "fCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z 'we're 32 bit

Public FileName             As String
Public DirName              As String
Public PrevPath             As String

Public TextToPaste          As String

Private SourceRTF           As String
Private Const ElipsisChar   As String = "…"

Public sFontName            As String
Public lFontSize

Private FileNum
Private SearchStart
Private xDiff
Private yDiff
Private SourcePos
'Copied from C:\Programme\Microsoft Visual Studio\VB98\Vb Sources\VBCompanion\fCopy.frm
Private Tooltip             As cToolTip
'End Copy

Private Crsr                As cMP

Private Attri               As Boolean
Private Loading             As Boolean
Private ViewingCopy         As Boolean
Private bCase               As Boolean
Private bWhole              As Boolean
Private bSelected           As Boolean
Private SearchCancelled     As Boolean
Private JustSearched        As Boolean

Private SearchString        As String
Private AppendText          As String
Private Words()             As String

Private Const NoName        As String = "<None>"
Private Const Clipbrd       As String = "<Clipboard>"
Private Const MaxLen        As Long = 32
'Copied from C:\Programme\Microsoft Visual Studio\VB98\Vb Sources\VBCompanion\fCopy.frm
Private Const Sel           As String = "Select text and add to clipboard"
'End Copy
Private Const WM_LBUTTONDBLCLK As Long = &H203

'Copied from C:\Programme\Microsoft Visual Studio\VB98\Vb Sources\VBCompanion\fCopy.frm
Private Sub AddToClipboard()

  Dim SelCol        As Variant
  Dim CopyPermit    As Boolean
  Dim AppendLen     As Long

    SelCol = rtfSource.SelColor
    If IsNull(SelCol) Then
        SelCol = vbWhite
    End If
    CopyPermit = (SelCol = vbBlack)
    If Not CopyPermit Then
        CopyPermit = (MsgBoxEx("You have already copied " & IIf(SelCol = vbWhite, "part", "all") & " of this text." & vbCrLf & "Copy again ?", vbYesNo + vbQuestion, , -1, -2) = vbYes)
        'quirk in Rich Text Box - if the very last char is selected then SelColor returns Null
    End If
    If CopyPermit Then
        rtfSource.SelColor = &H800080 'dark magenta
        AppendText = rtfSource.SelText
        Do
            AppendLen = Len(AppendText)

            Do Until Right$(AppendText, 1) <> vbCr
                AppendText = Left$(AppendText, Len(AppendText) - 1)
            Loop
            Do Until Left$(AppendText, 1) <> vbCr
                AppendText = Mid$(AppendText, 2)
            Loop
            AppendText = RTrim$(AppendText)

        Loop Until Len(AppendText) = AppendLen

        If Len(AppendText) Then
            TextToPaste = TextToPaste & IIf(Len(TextToPaste), vbCrLf & vbCrLf, vbNullString) & "'Copied" & Chr$(160) & "from " & FileName & vbCrLf & AppendText & vbCrLf & "'End Copy" 'there is a chr$(160) in End Copy !!
            With tb
                .Buttons("Paste").Enabled = True
                .Buttons("View").Enabled = True
                .Buttons("Copy").Enabled = False
            End With 'TB
        End If
    End If

End Sub
'End Copy

'Copied from C:\Programme\Microsoft Visual Studio\VB98\Vb Sources\VBCompanion\fCopy.frm
Private Sub CreateTooltip()

    Set Tooltip = New cToolTip
    Tooltip.Create rtfSource, Sel, TTBalloonIfActive, True, TTIconInfo, "Code", vbBlack, &HA0FFFF, 500, 5000

End Sub
'End Copy

Private Sub ElipsisFile()

  Dim i, j

    lblFilename = lblFilename.Tag
    i = Len(lblFilename) / 2
    j = lblFilename.Width - 120
    Do Until picFontMeasure.TextWidth(lblFilename) < j 'border
        Dec i
        If i < 1 Then
            Exit Do 'loop 
          Else 'NOT I...
            lblFilename = RTrim$(Left$(lblFilename.Tag, i)) & ElipsisChar & LTrim$(Right$(lblFilename.Tag, i))
        End If
    Loop

End Sub

Private Sub Form_Activate()

  Dim tmp   As String
  Dim Attr  As VbFileAttribute
  Dim Oops  As Long

    On Error Resume Next
        Attr = GetAttr(FileName)
        Oops = Err
    On Error GoTo 0
    If Oops Then 'nothing of that name
        FileName = NullStr
        Initialize False
      Else 'name exists 'OOPS = FALSE/0
        If Attr And vbDirectory Then 'it is a directory
            DirName = FileName & IIf(Right$(FileName, 1) = "\", NullStr, "\")
            tmp = Dir$(DirName) 'look into it
            If Len(tmp) = 0 Then 'no files in it
                FileName = NullStr
                Initialize True
              Else 'it is a directory with something in it 'NOT LEN(TMP)...
                If Len(Dir$) = 0 Then 'a directory with only one file in it
                    FileName = DirName & tmp
                    Initialize False
                  Else 'more than one file in it 'NOT LEN(DIR$)...
                    FileName = NullStr
                    Initialize True
                End If
            End If
          Else 'it is a file 'NOT ATTR...
            Initialize False
        End If
    End If

End Sub

Private Sub Form_Load()

    Loading = True
    xDiff = Width - rtfSource.Width
    yDiff = Height - rtfSource.Height
    picFontMeasure.Fontname = lblFilename.Fontname
    picFontMeasure.Fontsize = lblFilename.Fontsize
    picFontMeasure.FontBold = lblFilename.FontBold
    SearchCancelled = True
    PrevPath = CurDir
    ChDir App.Path
    Set Crsr = New cMP
    'Copied from C:\Programme\Microsoft Visual Studio\VB98\Vb Sources\VBCompanion\fCopy.frm
    Set Tooltip = New cToolTip
    'End Copy

End Sub

'Copied from C:\Programme\Microsoft Visual Studio\VB98\Vb Sources\VBCompanion\fCopy.frm
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Cancel = (UnloadMode <> vbFormCode)
    If Cancel Then
        If Len(TextToPaste) Then
            Select Case MsgBoxEx("The clipboard contains code." & vbCrLf & vbCrLf & "What do you want to do?", vbYesNoCancel Or vbQuestion Or vbDefaultButton2, Caption, -1, -2, -2, 0, -40, , Ja & "|" & Nein & "|" & Abbrechen, "Discard it|Paste it|Return")
              Case vbYes ' discard
                TextToPaste = vbNullString
                Hide
              Case vbNo 'paste
                Hide
              Case vbCancel 'better not
                'do nothing
            End Select
          ElseIf rtfSource.SelLength Then 'LEN(TEXTTOPASTE) = FALSE/0
            Select Case MsgBoxEx("You have not yet added the selected code to the clipboard." & vbCrLf & vbCrLf & "What do you want to do?", vbYesNoCancel Or vbQuestion Or vbDefaultButton2, Caption, -1, -2, -2, 0, -40, , Ja & "|" & Nein & "|" & Abbrechen, "Exit anyway|Paste it|Return")
              Case vbYes ' discard
                TextToPaste = vbNullString
                Hide
              Case vbNo 'paste
                AddToClipboard
                Hide
              Case vbCancel 'better not
                'do nothing
            End Select
          Else 'RTFSOURCE.SELLENGTH = FALSE/0
            Cancel = False
        End If
    End If

End Sub
'End Copy

Private Sub Form_Resize()

    On Error Resume Next
        rtfSource.Width = Width - xDiff
        tb.Width = Width - xDiff
        lblFilename.Width = Width - xDiff
        rtfSource.Height = Height - yDiff
    On Error GoTo 0
    'Copied from C:\Programme\Microsoft Visual Studio\VB98\Vb Sources\VBCompanion\fCopy.frm
    If Len(rtfSource.Text) Then
        CreateTooltip
    End If
    'End Copy
    ElipsisFile

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ChDir PrevPath

End Sub

Private Sub Initialize(ByVal WithOpen As Boolean)

  Dim CodeLine          As String
  Dim Temp              As String
  Dim Bold              As Boolean

    If Loading Then
        Loading = False
    End If
    If JustSearched Then
        JustSearched = False
      Else 'JUSTSEARCHED = FALSE/0
        Fontsize = lFontSize
        Fontname = sFontName
        With rtfSource
            .Font.Size = lFontSize
            .Font.Name = sFontName
            If WithOpen Then
                On Error Resume Next
                    With CDlg
                        Err.Clear
                        .InitDir = DirName
                        .FileName = FileName
                        .ShowOpen
                        If Err Then
                            Err.Clear
                            FileName = NullStr
                          Else 'ERR = FALSE/0
                            FileName = .FileName
                        End If
                    End With 'CDLG
                On Error GoTo 0
            End If
            ViewingCopy = False
            Attri = False
            .Text = vbNullString
            .RightMargin = 0
            DoEvents
            If FileName <> vbNullString Then
                FileNum = FreeFile
                lblFilename.Tag = FileName
                ElipsisFile
                Screen.MousePointer = vbHourglass
                Open FileName For Input As FileNum
                .Visible = False 'speed up loading
                Do Until EOF(FileNum)
                    Line Input #FileNum, CodeLine
                    Temp = Replace$(CodeLine, "(", " ", , 1)
                    Words = Split(LCase$(Temp), , 3)
                    If UBound(Words) < 2 Then
                        ReDim Preserve Words(2)
                    End If
                    If Words(0) = "attribute" Then
                        Attri = True
                      Else 'NOT WORDS(0)...
                        If Attri Then
                            Bold = (Words(0) = "sub" Or Words(0) = "function" Or Words(0) = "property" Or Words(0) = "enum" Or Words(0) = "type")
                            Bold = Bold Or ((Words(0) = "public" Or Words(0) = "private" Or Words(0) = "static" Or Words(0) = "friend") And (Words(1) = "sub" Or Words(1) = "function" Or Words(1) = "property" Or Words(1) = "enum" Or Words(1) = "type"))
                            Bold = Bold Or ((Words(0) = "public" Or Words(0) = "private" Or Words(0) = "friend") And (Words(1) = "static") And (Words(2) = "sub" Or Words(2) = "function" Or Words(2) = "property"))
                            Bold = Bold Or (Words(0) = "end" And (Words(1) = "sub" Or Words(1) = "function" Or Words(1) = "property" Or Words(1) = "enum" Or Words(1) = "type"))
                            FontBold = Bold 'for textwidth measurement
                            If TextWidth(CodeLine) > .RightMargin Then
                                .RightMargin = TextWidth(CodeLine)
                            End If
                            .SelColor = vbBlack
                            .SelBold = Bold
                            .SelText = CodeLine & vbCr
                        End If
                    End If
                Loop
                .Visible = True
                Close FileNum
                .SelStart = 0
                Screen.MousePointer = vbDefault
                'Copied from C:\Programme\Microsoft Visual Studio\VB98\Vb Sources\VBCompanion\fCopy.frm
                CreateTooltip
              Else 'NOT FILENAME...
                lblFilename.Tag = NoName
                ElipsisFile
                Set Tooltip = Nothing
                'End Copy
            End If
        End With 'RTFSOURCE
    End If

End Sub

Private Sub rtfSource_KeyDown(KeyCode As Integer, Shift As Integer)

  Dim CaretPos As Long
  Const Forb = ".,()"" "
  Dim TrimIt As Boolean
  Dim Pt As POINTAPI

    If KeyCode = vbKeyF3 Or KeyCode = vbKeyF Then
        If Shift And vbCtrlMask Then
            JustSearched = True 'prevents Form_Activate when the Search window closes
            With rtfSource
                If .SelLength Then
                    bSelected = True
                    TrimIt = False
                  Else '.SELLENGTH = FALSE/0
                    'no selection, try to find something near the caret
                    GetCaretPos Pt
                    CaretPos = 0
                    Select Case 0
                      Case InStr(vbLf & Forb, Mid$(" " & .Text, .SelStart + 1, 1))
                        'char left of caret is not in Forb
                        CaretPos = -8
                      Case InStr(vbLf & Forb, Mid$(.Text & " ", .SelStart + 1, 1))
                        'char right of caret is not in Forb
                        CaretPos = 8
                    End Select
                    If CaretPos Then
                        CaretPos = Pt.Y * 65536 + Pt.X + CaretPos
                        'select the text
                        SendMessage .hWnd, WM_LBUTTONDBLCLK, 0&, ByVal CaretPos
                    End If
                    bSelected = False
                    TrimIt = True
                End If
                If .SelLength And .SelLength <= MaxLen Then
                    SearchString = .SelText
                    If TrimIt Then
                        SearchString = Trim$(SearchString)
                    End If
                    Do Until Right$(SearchString, 1) <> vbLf
                        SearchString = Left$(SearchString, Len(SearchString) - 1)
                    Loop
                End If
            End With 'RTFSOURCE
            With fSearch
                .txtSearch = SearchString
                .txtSearch.SelStart = 0
                .txtSearch.SelLength = Len(SearchString)
                .ckCase = IIf(bCase, vbChecked, vbUnchecked)
                .ckWhole = IIf(bWhole, vbChecked, vbUnchecked)
                .ckSelected.Enabled = bSelected
                .txtSearch.MaxLength = MaxLen
                .Move Left, Top
                .Show vbModal
                SearchString = .txtSearch
                bCase = (.ckCase = vbChecked)
                bWhole = (.ckWhole = vbChecked)
                bSelected = (.ckSelected = vbChecked)
                SearchCancelled = (.Tag = "C")
            End With 'FSEARCH
            Unload fSearch
        End If
        If Not SearchCancelled Then
            If bSelected Then
                SearchStart = rtfSource.Find(SearchString, , , IIf(bWhole, rtfWholeWord, 0) + IIf(bCase, rtfMatchCase, 0))
                bSelected = False
              Else 'BSELECTED = FALSE/0
                SearchStart = rtfSource.Find(SearchString, SearchStart + 1, , IIf(bWhole, rtfWholeWord, 0) + IIf(bCase, rtfMatchCase, 0))
            End If
            If SearchStart < 0 Then
                If MsgBox("End of text reached." & vbCrLf & vbCrLf & "Would you like to resume at the beginning?", vbQuestion + vbYesNo, "Search [" & SearchString & "]") = vbYes Then
                    SendKeys "{F3}"
                End If
            End If
        End If
        SearchCancelled = (Len(SearchString) = 0)
    End If

End Sub

Private Sub rtfSource_KeyPress(KeyAscii As Integer)

    KeyAscii = 0

End Sub

Private Sub rtfSource_KeyUp(KeyCode As Integer, Shift As Integer)

    If Shift And vbCtrlMask And (KeyCode = vbKeyV Or KeyCode = vbKeyC) Then
        If tb.Buttons("Copy").Enabled Then
            'catch the clipboard text after Ctrl+V/C
            AppendText = Clipboard.GetText
            tb_ButtonClick tb.Buttons("Copy")
        End If
        Clipboard.Clear
    End If
    tb.Buttons("Copy").Enabled = rtfSource.SelLength And Not ViewingCopy

End Sub

Private Sub rtfSource_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Crsr.SetPointerIcon TextPos

End Sub

Private Sub rtfSource_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    tb.Buttons("Copy").Enabled = rtfSource.SelLength And Not ViewingCopy

End Sub

Private Sub rtfSource_SelChange()

    SearchStart = rtfSource.SelStart - 1

End Sub

Private Sub tb_ButtonClick(ByVal Button As MSComctlLib.Button)

  Dim AppendLen     As Long
  Dim CopyPermit    As Boolean
  Dim SelCol        As Variant

    Select Case Button.Key
      Case "Open"
        Initialize True
      Case "Copy"
        AddToClipboard
      Case "View"
        If ViewingCopy Then
            ViewingCopy = False
            'Copied from C:\Programme\Microsoft Visual Studio\VB98\Vb Sources\VBCompanion\fCopy.frm
            With rtfSource
                .Enabled = True
                .TextRTF = SourceRTF
                .SelStart = SourcePos
            End With 'RTFSOURCE
            'End Copy
            tb.Buttons("Copy").Enabled = True
            Button.ToolTipText = "View Clipboard"
          Else 'VIEWINGCOPY = FALSE/0
            ViewingCopy = True
            FontBold = False 'no bold text in view clipboard box
            With rtfSource
                SourceRTF = .TextRTF
                SourcePos = .SelStart
                .Text = vbNullString
                .SelColor = &HC0& 'dark red
                .SelText = TextToPaste
                .Enabled = False
            End With 'RTFSOURCE
            tb.Buttons("Copy").Enabled = False
            lblFilename.Tag = Clipbrd
            ElipsisFile
            Button.ToolTipText = "Back to Code"
        End If
      Case "Reset"
        TextToPaste = vbNullString
        tb.Buttons("Paste").Enabled = False
        tb.Buttons("View").Enabled = ViewingCopy
        If ViewingCopy Then
            rtfSource = vbNullString
        End If
        Initialize False
      Case "Paste"
        Hide
      Case "Cancel"
        TextToPaste = vbNullString
        Hide
      Case "Find"
        rtfSource_KeyDown vbKeyF, vbCtrlMask
    End Select

End Sub

Private Sub tb_Change()

    tb.Buttons("Paste").Enabled = Len(rtfSource.Text)

End Sub

Private Sub tb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Crsr.SetPointerIcon RightHand

End Sub

':) Ulli's VB Code Formatter V2.24.11 (2008-Apr-11 10:26)  Decl: 46  Code: 461  Total: 507 Lines
':) CommentOnly: 28 (5,5%)  Commented: 41 (8,1%)  Filled: 434 (85,6%)  Empty: 73 (14,4%)  Max Logic Depth: 8
