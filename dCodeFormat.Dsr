VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} dCodeFormatter 
   ClientHeight    =   14010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13755
   _ExtentX        =   24262
   _ExtentY        =   24712
   _Version        =   393216
   Description     =   $"dCodeFormat.dsx":0000
   DisplayName     =   "Ulli's Code Formatter"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "dCodeFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*********************************************************************************
'© 2000/2008 UMGEDV GmbH   (umgedv@yahoo.com)

'Author      UMG (Ulli K. Muehlenweg)

'Title       VB6 Code Formatter Add-In

'Purpose     This Add-In will format your code for a generally accepted
'            indentation - one tab per structure level. It will also check
'            your code for some of the most common ommissions and traps.
'            There is an option to mark the offending lines so you can find
'            them easily to rectify. The result is a clean and uniform code
'            appearance.

'            Includes a Structure Viewer, a Copy Facility and a Code Printer

'            A few run time options exist to modify your code:
'               *   Module sorting
'               *   Replacement of type suffixes
'               *   Replacement of variant-returning string functions by their
'                   string-returning counterparts
'               *   Insertion of Enum Case preserving code
'               *   Splitting of compound lines with more than one statement
'               *   Conversion of single line IF's into multiline If's
'               *   Insertion of Code to use .Manifest files
'               *   Remove Call verbs
'               *   Remove Line numbers

'            Compile the DLL into your VB directory and then use the Add-Ins Manager
'            to load the Formatter into VB.

'            For customizing permanent options at compile time see
'            conditional compilation constants below

#Const LinNumComplain = True       'Set to False to accept line numbers
#Const StdLocalDimIndent = True    'Set to False to indent Dim same as procedural statements
#Const MarkOnErrorGoTo = True      'Set to False to prevent marking of 'On Error GoTo
#Const CallComplain = True         'Set to False to prevent complaints about unnecessary 'Call' verbs
#Const OrphanedComment = True      'Set to False to prevent complaints about orphaned comments
'                                   Note: Orphaned comments may not be sorted as you might expect in some cases
'                                         see V2.19.3

'*********************************************************************************

'#################################################################################
'Known Bugs and Quirks (currently two)
'#################################################################################

' 1) When a module is completely surrounded by a conditional compilation bracket,
'    like:

'       #If Some_Condition Then

'       Private Sub Some_Sub()
'           :
'           :
'       End Sub

'       #End If

'    then that module will be lost during module sorting.

'    This is due to the fact that VB apparently does not add such modules to it's
'    members collection. The suggested work-around is to put the conditional compilation
'    bracket inside the module.

'    Also if you plan to sort, it is generally a good idea to put remarks and comments
'    inside modules to make association unambiguous.

'    A comment outside a module (orphaned comment) is associated with the module which
'    follows the comment; if however an orphaned comment is not separated from the preceeding
'    module by at least one blank line then it is associated with the module which
'    preceeds it; see also V2.19.3.

' 2) When only the Let/Set part of a Property is present and an Enum exists in the same
'    module with the same name as the Property, like:

'       Private Enum Same_Name
'           :        =========
'           :
'       End Enum

'       Private Property Let Same_Name(...)
'           :                =========
'           :
'       End Property

'    then the Property will be lost during module sorting.

'    This is due to the fact that VB apparently does not add such modules to it's
'    members collection. The obvious work-around of course is to rename the Enum.

'Currently no check is made whether any of the above two conditions exists. Check out
'Roger's Code Fixer which circumvents the 2nd bug by adding a dummy Property Get.

'#################################################################################
'Development History
'#################################################################################

'11Apr2008 V2.24.17 - UMG

'Fixed some irregularities with For - Next
'Added ToDo29 (Combining Next)
'Added manual Indenting (Rem Indent Begin - Rem Indent End)
'Fixed bug during printing which did not permit optimizations to be switched on
'Fixed bug during print preparation of NaD's (did not always print them all)

'---------------------------------------------------------------------------------

'10Mar2007 V2.23.12 - UMG

'Bug Fix with Rem (tnx to Nightwolfe whose code made me crash)
'Added "Rem Interface"-option which defines a module as interface to be used for
'Implements and therefore does not complain about empty Subs/Functions.

'---------------------------------------------------------------------------------

'27Jan2007 V2.22.14 - UMG

'Made some more adjustments to the Copy Facility and the Formatter to accomodate
'copy insertions made by VB Companion.

'---------------------------------------------------------------------------------

'23Jan2007 V2.22.11 - UMG

'Fixed Rem behavior.

'Fixed some Copy quirks; Copy option still does not work if it is the first and only
'line in a module.

'Applied a few changes and cosmetics suggested by Roger Gilchrist.

'---------------------------------------------------------------------------------

'20Jan2007 V2.22.5 - UMG

'Fixed Bug in Suppressed count; was counting dupl names while in Cc mode.

'Added Line Number Handling

'#Const LinNumComplain governs the behavior:

'If True:  will complain about line number but nevertheless format them correctly.
'          complaints may be suppressed by switching runtime option 'Insert Comments' off.
'          Line Numbers can also be removed in this mode, leaving a remark unless
'          runtime option 'Insert Comments' is off.

'If False: will not complain about Line Numbers but just format them.
'          Line Numbers cannot be removed in this mode.

'---------------------------------------------------------------------------------

'19Jan2007 V2.21.8 - UMG

'Fixed bug in Call bracket removal.
'Modified Help window

'---------------------------------------------------------------------------------

'23Feb2006 V2.21.6 - UMG

'Added option to remove Call verbs and brackets.
'Added Rem Mark Off Silent to be able to suppress the comment about unchecked lines.
'Fixed bug with orphaned comments during skip.
'Fixed forecolor for printing final comment.

'---------------------------------------------------------------------------------

'09Jan2006 V2.19.5 - UMG

'Fixed bug with new Rem recognition in V2.19.3 which caused wrong summary insertion or
'even a crash.

'Fixed quirk with For-variable insertion when insertion ist switched off.

'Fixed bug with single line "If"
'   A "Then" followed by a colon was mistaken for a "Then" without the colon and therefore
'   assumed to start a multiline "If" when in fact it is a single line "If". See following
'   example:

'   If SomeCondition Then Rem optional comment   <-- this is a single line If

' however:

'   If SomeCondition Then 'optional comment      <-- is starting a multi-line If

' whereas:

'   If SomeCondition Then: 'optional comment     <-- is a single line If (oh well *sigh*)

'   Now correctly flagged as single line "If" and also correctly expanded to

'   If SomeCondition Then
'       'optional comment
'   End If

'Altered handling of comment insertion and erasure of "replaced by" in connection
'with "If"-expansion.

'Fixed Enum Case Preservation insertion for cases where the contents of square brackets
'is not a legal VB data name - eg [3D] or [MM-DD-YYYY]

'Changed displayed text on code pane selection.

'Fixed bug with fProgress not counting empty code panes.

'Removed message box complaining about empty code panes and some dead code that was
'never executed.

'---------------------------------------------------------------------------------

'08April2005 V2.19.3 - UMG

'Added check for orphaned comments. Orphaned comments are outside Sub's and therefore
'it may be unclear where they belong - also they may cause problems when you attempt
'to sort the code. An orphaned comment is associated with the module which follows the
'comment; if however an orphaned comment is not separated from the preceeding module
'by at least one blank line then it is associated with the module which preceeds it.

'Added silent If-expansion - right-click on that option. Left clicking will comment
'out the If and add new lines to the code with the If expanded; right clicking will
'remove the single line If and repace it with the expanded version.

'Added Debug recognition.

'Clarified Rem recognition - Rem & somechars is still a comment (like Rem,).

'Added conditional compilation for 'Call' complaints.

'Added conditional compilation for 'Orphaned Comments' complaints.

'---------------------------------------------------------------------------------

'23Mar2005 V2.18.4   - UMG

'Added Enum Case Preservation.
'Added rudimentary If-Expansion; known quirk - compound lines conditioned by a
'                                Single-Line-If will not be exanded correctly.
'                                This is due to the colon ambiguity of VB:
'                                a word followed by a colon may either be a
'                                Sub/Function call  or a GoTo label and this can
'                                only be determined by context, not by syntax
'                                (see also notes on V2.5.12 further down)

'---------------------------------------------------------------------------------

'02Dec2004 V2.17.11  - UMG

'Bug Fix: Corrected Insertion of blank lines; this would go into an endless loop
'         in certain situations.

'Removed a few unused variables.

'Clarified On Error GoTo 0 marking.

'Added check for empty modules.

'---------------------------------------------------------------------------------

'08Nov2004 V2.17.8   - UMG

'Bug Fix: Corrected print bug introduced in V2.17.7
'Now using high speed timer for fQuestion Options expand/collapse

'---------------------------------------------------------------------------------

'04Nov2004 V2.17.7   - UMG

'Bug Fix: Corrected several modules for no-printer condition.
'Replaced all "" literals by NullStr.
'Clarified orphaned comments.
'Removed Fader.

'---------------------------------------------------------------------------------

'30Jul2004 V2.17.4   - UMG

'Bug Fix

'An Open/Close Bracket inside a "Literal" could upset the routine to determine the type
'of a function, for example the line:

'   Private Function SomeFunction(Optional Byval Param As String = "(some text)") As Long

'was marked as untyped ("As Variant?").

'This has been fixed by not only replacing spaces but also brackets inside a literal
'by low-values Chr$(0).

'---------------------------------------------------------------------------------

'10Jul2004 V2.17.3   - UMG

'New continuation line indenting algorithm.

'New Exit For/Do/Sub/Function/Property marking (will need some manual work because
'the original markings are not removed).

'---------------------------------------------------------------------------------

'24Mar2004 V2.16.15  - UMG

'Fixed bug with un-scoped Property Let/Set (used to complain about missing/default
'variant type) (see V2.16.11 where the same bug was fixed for scoped properties).

'---------------------------------------------------------------------------------

'28Jan2004 V2.16.14  - UMG

'Fixed bug with GoTo-Label recognition and compound line separation.
'A comment like in the following line was under certain circumstances separated
'because it thought that  'Thanks:   was a label.
'Thanks: Go to Roger G. for using such a line in one of his submissions *smile*

'Extended dead code recognition: A condition which contains 'False' alone or in
'connection with 'And' will now be complained about. Check the condition carefully
'and try to re-arrange it if it is correct. For example

'    If xyz = False And abc = True
'             =========
'could be re-arranged

'    If abc = True And xyz = False

'or use brackets.

'If all fails switch off 'Insert Comments' in the options panel.

'---------------------------------------------------------------------------------

'29Nov2003 V2.16.13  - UMG

'Changed inserted marks - see BlnkApos.
'Improved OE handling (OELevel).
'Fixed bug with Structure Error not being reset at End Sub/Function/Property.
'Removed separate line limitation for GoTo-labels; they may now contain more
'                                                              code or notes.
'Added OK button in fNoUndo.

'---------------------------------------------------------------------------------

'23Feb2003 V2.16.11  - UMG

'Fixed bug with Property Let/Set (used to complain about missing/default variant type)
'Fixed options checkbox behavior
'Made "GoTo" an error
'Added Toolip class
'Added #conditional compilation for  Local Dim indentation
'                                    On Error GoTo marking
'Fixed erroneous printing on Undo
'Some Code cosmetics

'This is here tnx to a suggestion by Roger Gilchrist:
'----------------------------------------------------
'Added option to stop scan on error. You can alter the code while in stopped state,
'however the scan will continue as if nothing had happened during that state, and
'therefore may find errors which do in fact not exist, for example:
'When the Formatter found a Single-Line-If and you modify that to a Structured-If
'while in stopped state, then the Formatter will still think the 'If' is done with
'when you click Resume and therefore will not recognize that the next line is now
'an If-dependent line

'Now recognizes  {#If False Then --- #End If}  conditional compilation bracket
'which is used by some developers to preserve Enum capitalizations

'Extended Type Suffix funtionality: type suffixes are now optionally replaced by
'their corresponding As [type] keywords; however code with

'  (1) nested brackets  or
'  (2) type suffixes on literals

'is NOT treated correctly

'Examples for (1)  [Scope] Function a&(b&(),c%)
'             (2)  [Scope] Const x& = 1&

'This kind of coding is very rarely used and so I will leave it at that

'---------------------------------------------------------------------------------

'03Sep2002 V2.15.4   - UMG

'Fixed default Procedure-Id for MP in cMP

'Added btShowAll to open the code panels for all components in all project(s)

'Added some counting to make sure that XPLook will find the startup object

'MessageBeeps are consistent now

'Added compile option; checking this option will attempt to compile the active
'project after the format scan if no summary was inserted into the code --
'either because none was necessary or because it was suppressed by unchecking
'the Insert Summary Checkbox

'Allow comments among local Dims/Consts in proc header; the first comment is assumed
'to be descriptive for the Sub/Function/Property and is therefore separated from the
'remainder by one blank line

'Killed little quirks in active For-Variable recognition and in fStrFncts

'Fixed cleaning of last codeline

'Tool tips support Swapped Mouse Buttons now

'---------------------------------------------------------------------------------

'25Aug2002 V2.14.7   - UMG

'Added fading effect (WinXP only) plus bugfix reported by Gonchuki

'Added check for printer availability

'Added check for OS version in cMP

'Modified indenting with On Error

'Fixed some minor bugs in fCopy

'Added Undo function (one level)

'   Saves the state of your source in Undo Buffers before a formatting scan
'   and keeps it until:

'     1 - you undo
'         any manual alterations you may have made after the scan are also undone

'  or 2 - you format again
'         this overwrites the corresponding Undo Buffer(s)

'  or 3 - you close VB

'  Modify code if you prefer no complaints about "On Error GoTo [ErrorHandler]"
'  Find [ErrorHandler] to see how it's done (two places). This has been modified
'  to conditional compilation, see MarkOnErrorGoTo

'---------------------------------------------------------------------------------

'17Aug2002 V2.13.7   - UMG

'New feature: dead code recognition (now also after GoTo and Return)
'Fixed bug with Private Static Sub ()
'Fixed bug with CurrProcType and Static Keyword
'Fixed duplicate name clash while sorting components
'Altered GoTo reaction
'Fixed startup component type recognition for XP insertion
'Added SaveSetting & GetSetting for Printfont & -size
'Made half indenting an option
'Some code cleanup

'14Apr2002 V2.12.7   - UMG

'Fixed string-fct replacement when 'As String$()' is encountered

'Added option to create Win XP Look

'Check "Create WinXP Look" in Option Window, format all components, and compile
'That's it. Easy, ain't it?

'It may be necessary to rearrange some code lines either in Form_Initialize or
'in Sub Main if they already existed and have Local Dims. This is because the
'Formatter will insert the necessary API Call immediatey after the Procedure
'Header disregarding any contents which may already be present. I may fix that
'at some later time; at the moment the Formatter may complain about a self-created
'error :-(

'If the Formatter complains about "No XPLook created" then this is either due to
'the fact that your project type does not permit skinning (I fixed that), or the
'path pointing to the .exe file is invalid. The .manifest file is created in the
'same directory into which the .exe file will also be compiled

'Palette Colors are disabled, only System Colors are available; if you don't like
'what you've got then simply delete or rename the .manifest file and you're back
'to "normal"

'Known Comctl32(?) bug: Radio Buttons, in particular if located in a Frame, cause
'problems

'---------------------------------------------------------------------------------

'31Mar2002 V2.11.3   - UMG

'Modified line width to 3 for frames when printing and clarified a few references to
'Printer.ScaleWidth/Height such that the now wider lines are drawm completely within
'the printable area
'Added comment color for printing and moved some colors to constants
'Modified KillDoc option logic

'Fixed bug with recognizing write protected components (check of a maximized code pane
'for write protection was incorrect)

'Added Max Structure Depth

'Fixed bug with continued Single Line If separation by colon (bug reported by Luis)
'This is now correctly treated as a Single Line If. Same mechanism is used for _
      continued comments

'Fixed bug with printing and no pane selected (printed a single empty sheet)

'---------------------------------------------------------------------------------

'04Mar2002 V2.10.8   - UMG

'Functions are now checked for type and for type suffix character
'Code lines containing space characters only are trimmed to zero length: VB apparently
'trims lines containing code but not lines containing no code

'---------------------------------------------------------------------------------

'14Jan2002 V2.9.4    - UMG

'Added Possible Structure Violation Detection

'    For example:
'    ------------

'    Do
'        :
'        With Something
'            :
'            If SomeCondition Then

'                Exit Do  'This Exits out of the With-End With bracket and may
'                         'cause errors which are very difficult to trace down
'            End If
'            :
'            :
'        End With 'Something
'        :
'        :
'    Loop Until SomeOtherCondition

'This detection only works for unconditional Exits and for conditional Exits within a
'structured If-End If. Exits out of a Sub/Function/Property are not harmful and are
'therefore not checked (VB executes a tidy up routine on exit from a Sub I think)

'Moved most error marking and flag setting to MarkLine (renamed from CountMarks)

'Added fStrFcts
'List of modifiable String Functions moved to fStrFncts.lstStrFncts
'Added option to skip modifying variant-returning string functions

'Modified ResetFocus in fQuestion to try and cure a (non-reproducable) misbehavior
'reported by Akbar

'Modified Selected Panels Count in fQuestion (didn't notice there is a SelCount
'property and counted them myself before :-(  --  oops!)

'---------------------------------------------------------------------------------

'27Dec2001 V2.8.9    - UMG (Christmas Edition)

'Known quirk: The CodeLocation returned by VB is wrong sometimes (generally too
'             small), so there is a rudimentary Search to find the correct loc;
'             however this Search may find a wrong line still or may find nothing
'             at all. The Line Numbers printed in the NaD could therefore be wrong

'Fixed quirk with printing a Continued _
       Comment

'VB-Bug(?) circumvented:
'The Attributes-Properties of VBInstance.ActiveCodePane.CodeModule.Members()
'sometimes fail after a program has been running in the IDE. So when this happens we
'simply don't store them

'Fixed bug in line separation with date/time literals containing colons
'like #12:23:56 PM# and named parameters like "Collection.Add Item, Before := 4"

'Speeded up restoring Member Attributes

'---------------------------------------------------------------------------------

'15Dec2001 V2.8.7    - UMG

'Fixed problem with write-protected source files; checkmarking them is disabled
'Added NaD and Quicksort (unfortunately VB's Members collection is not sorted)
'Now also using Quicksort for ordering Sub's
'Fixed a few quirks
'Added ToDo16 ("Remove Line Number"); however this is not shown in summary box
'Added font selection box for printing
'Removed SaveSetting and GetSetting; these values can also be saved in mAPI
'                                    (discovered that by accident)
'Removed some dead variables

'Known Quirks: 1 When the printer is changed while this DLL is active it will still
'                print to the old printer. So close and re-open the code formatter
'                to switch it to the new printer

'              2 When [part of] a line is selected and the line is subsequently
'                shifted by formatting then the selection does not shift with it

'The member attributes are killed thru formatting if the 'MemberInfo=...' comment
'is not present, so this AddIn saves and restores them now. As restoring takes a
'while the app seems to freeze for a moment at the end of the scan

'---------------------------------------------------------------------------------

'05Dec2001 V2.6.10   - UMG

'Added Print Option
'Made fProgress topmost and added 2nd progress bar
'Altered colors of progress bar to indicate type of processing
'Fixed empty Component quirk

'---------------------------------------------------------------------------------

'21Oct2001 V2.5.12   - UMG

'Added option to order Subs, Functions, and Properties alphabetically. Uncurable
'feature *g*: when a comment is written AFTER an End Sub or -Function (but belonging
'to it; i.e. with no blank line between, then this Comment is detached and not
'correctly ordered. A comment before a Sub or Function is handeled correctly

'Added option to separate compound lines having colons between statements
'Each colon which is not part of a Goto-label or a literal will cause a line break
'A Goto_label: must be the only statement in a line, no comment or other statement
'is allowed
'A line starting with 'If' will not be separated to avoid problems caused by
'the possibly missing 'End If' in a single line If-statement. It will however be
'complained about during parsing
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Note:

'   VB allows a colon after 'Else' and removes a colon before 'Else';
'   however a colon before or after 'Then' is not allowed

'   A Sub-Call without the Call keyword and followed by a colon is
'   mistaken by VB (and by this AddIn) as a Goto-label

'   Example
'   -------
'    MarkLine:

'    'Will not execute MarkLine, thinks it's a label !!

'    Call MarkLine:

'      or

'    MarkLine

'    'Will execute MarkLine

' Private Sub MarkLine()
'    ....

'   VERY STRANGE BEHAVIOR INDEED, Billyboy!
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'Fixed Bug in Dim WithEvents processing ('WithEvents' was taken as a variable name)
'Some Code Cosmetics

'---------------------------------------------------------------------------------

'14Oct2001 V2.3.18   - UMG

'Fixed bug with numeric labels
'If Else ElseIf EndIf on/from stack
'Ignores conditional compilation statements now during structure build
'Rephrased "On Error Resume/Goto"
'Added single line conditional Exit... detection
'Added check for Active For-Variable modification
'    this is treated as an error although maybe it's not
'    if you're sure you know what you're doing you may want to replace
'    the marks inserted by the Formatter with an explanatory comment
'Added check for Missing Scope Declaration
'Altered ErrorLineSelection to select continuation lines also
'Elimination of dead vars and other tidying up
'Code Cosmetics

'Corrected bug with repetition of For-Vars in Next-Stmt
'Added Exit Marking
'fAllDone added
'Changed fQuestion PauseWhen to Combo Box
'Code Cosmetics

'---------------------------------------------------------------------------------

'03Mar2001 V2.3      - UMG

'Added SaveSetting and GetSetting
'Code Cosmetics
'Know quirk: If the code pane is switched to single procedure view then the
'            positioning function from the Structure Window does not work

'08Jan2001 V2.2      - UMG

'Added code to deal with line numbers (partially)
'Added WordStack for For..Next loops and others
'Fixed bug:  With..End With in same line
'Fixed bug:  Formating Dim after Proc Header without Public / Private
'Added multiline formatting for Declares

'Known bugs: If the last line in a Component is a continued _ line
'            then it will duplicate that line leading to a syntax error
'            A Comments Only Component will crash the Formatter

'---------------------------------------------------------------------------------

'25Nov2000 V2.1      - UMG

'Added option for Summary Box Pause
'Corrected quirk with reset selection after scan
'Made ViewStructure Button Graphical and positioned it a bottom of Summary Box
'Known qirk:
'    The reposition function from the structure box may be off a few lines
'    if a summary was inserted into the code or if vertical formatting has
'    added empty lines

'---------------------------------------------------------------------------------

'22Oct2000 V2.0      - UMG

'Added Structure Analyzer
'    I am a bit too late with this, Roger Willcocks, rogerw@dtl.co.nz has
'    already uploaded a structure analyzer; this coding however is completely
'    different from his. I have 'stolen' his positioning idea though

'29Sep2000 V1.9      - UMG

'Added Copy Facility
'    to activate enter Rem Copy [[path\]filename.ext]
'                               no colons or spaces accepted at this time
'                               needs fix
'Added Help and About Box
'Added some tool tips
'Added Variant String Function Replacement
'Added error summary insertion
'   this occurs at end of module if Mark On is still active
'   ie you can turn it off by inserting 'Rem Mark Off' as last line
'Fixed a minor qirk in repositioning after scan
'Removed blank line insertion after End Sub/Function/Property if the next line
'    is a #End If
'Added Bottom Summary Line
'Altered continuation line indentation
'Added [external name] recognition
'Added 'Static' Sub Function Property
'Changed main loop from 'For' to 'Do' for vertical formatting
'Added vertical formating (phooh - at last)
'Relaxed Const type checking for all Const's but Integers
'Included Declared Function names in duplicated names detection
'Added Marks counters

'---------------------------------------------------------------------------------

'08Sep2000 V1.8      - UMG

'Added fPreparing & animation
'Fixed bug in 'Rem' recognition (thanks to Eric who reported that and supplied
'                                                  the corresponding test data)
'Added Module Select Box
'Replaced a few MsgBoxes by Forms
'Made fQuestion with a faked Caption Bar to accomodate buttons for Help and About
'Added check for obsolete type suffix
'Moved some declarations into mAPI
'Corrected DefType for multiple entries
'Corrected single line structure recognition: Next Loop Wend
'Corrected grammar (singular / plural) in some messages

'---------------------------------------------------------------------------------

'04Sep2000 V1.7      - UMG

'Added  'Rem Skip On'  and  'Rem Skip Off'  options
'    Use to protect part of your code from formatting and checking
'    Both options MUST be placed at the SAME structure level
'    Skip On causes the Formatter to ignore all lines until Skip Off is found
'Corrected DefType A-Z emit chars
'Corrected def Const type detection (DefType doesn't apply to Const but Suffix does)
'Corrected pleonasm detection
'Added lenient check for 'GoTo
'Added separator and picture to Add-Ins menu
'Added splash form and Company Logo
'Added multiple CodePanes processing
'    that fixed ActiveWindow VB quirk:
'    sometimes the ActiveWindow was gone after saving
'Added FirstCodeLine detection for 'Option Explicit
'Put selection on (last) offending line
'Fixed bug in dupl name detection
'More comments
'Some optimization and clarification

'---------------------------------------------------------------------------------

'28Aug2000 V1.6      - UMG

'Abandoned Version 1.6
'This was real sh!t, had to re-install VB - thank you very much, Bill G. :·(

'---------------------------------------------------------------------------------

'26Aug2000 V1.5      - UMG

'Added further checks on:
'    untyped variables which would default to Variant
'    duplicated variable names local - modul-wide
'    structural aspects and
'    a common type of pleonasm (If bolCondition = True ...)

'A few more comments have also been inserted

'---------------------------------------------------------------------------------

'27Jul2000 V1.3      - UMG

'The Formatter will now detect local Dim's, Const's etc which are not at the
'top of the procedure, a missing Option Explicit, unstructured If's,
'empty Next's, and unclosed On Error Resume's. I feel that these are the most
'common reasons for unexpected program behaviour

'You have the option between  'Rem Mark On'  and  'Rem Mark Off
'Rem Mark On    causes the Formatter to mark offending lines with ':(' and these
'               may then be edited manually
'Rem Mark Off   does not mark these lines

'In any case will the Formatter output a comment in the final MsgBox
'Both options may be placed anywhere in the code on a separate line and take
'effect starting at the next line. The default is Mark Off
'A few minor quirks have also been corrected

'The Formatter will now honor the Tab Width value from VB Options

'---------------------------------------------------------------------------------

'20Jul2000 Prototype - UMG

'---------------------------------------------------------------------------------

Option Explicit
DefLng A-Z 'we're 32 bit

'Constants
Private Const MenuName            As String = "Add-Ins" 'you may need to localize "Add-Ins"

Private Const VBSettings          As String = "Software\Microsoft\VBA\Microsoft Visual Basic"
Private Const TabWidth            As String = "TabWidth"
Private Const Fontheight          As String = "Fontheight"
Private Const Fontface            As String = "Fontface"
Private Const DefaultTabWidth     As Long = 4
Private Const DefaultFontSize     As Long = 9
Private Const DefaultFontName     As String = "Fixedsys"
Private Const MarkOn              As String = "rem mark on"
Private Const MarkOff             As String = "rem mark off"
Private Const IndentOn            As String = "rem indent begin"
Private Const IndentOff           As String = "rem indent end"
Private Const Interface           As String = "rem interface"
Private Const MarkOffSilent       As String = MarkOff & " silent"
Private Const sThen               As String = " Then "
Private Const sThenD              As String = " Then:"
Private Const sElseB              As String = " Else "
Private Const sElseC              As String = " Else:"
Private Const sEndIf              As String = "end if"
Private Const CcIfBeg             As String = "#if false then"
Private Const CcIfElseBeg         As String = "#elseif false then"
Private Const CcIfEnds            As String = "#" & sEndIf
Private Const Mo                  As Long = 1 'reason codes: Mark Off
Private Const Si                  As Long = 4 '            : Mark Off Silent
Private Const Cc                  As Long = 2 '            : Conditional Compilation
Private Const MyErrMark           As String = " " 'Chr$(160)!!
Private Const MySignature         As String = "':" & "(" & MyErrMark
Private Const sRem                As String = "rem"
Private Const SkipOn              As String = sRem & " skip on"
Private Const SkipOff             As String = sRem & " skip off"
Private Const Copy                As String = "copy"
Private Const CopiedFrom          As String = "'Copied" & MyErrMark & "from"
Private Const EndCopy             As String = "'end" & MyErrMark & Copy
Private Const SkipMark            As String = " >> "
Private Const Smiley              As String = "':) "
Private Const ToDo1               As String = MySignature & "" 'not used
Private Const ToDo2               As String = MySignature & "Move line to top of current "
Private Const ToDo3               As String = MySignature & "Expand Structure"
Private Const ToDo4               As String = MyErrMark & "or consider reversing Condition"
Private Const ToDo5               As String = MySignature & "Repeat For-Variable:"
Private Const ToDo6               As String = MySignature & "On Error Resume still active"
Private Const ToDo7               As String = MySignature & "Remove Pleonasm"
Private Const ToDo8               As String = MySignature & "Structure Error"
Private Const ToDo9               As String = MySignature & "As Variant ?"
Private Const ToDo10              As String = MySignature & "Duplicated Name"
Private Const ToDo11              As String = MySignature & "As 16-bit Integer ?"
Private Const ToDo12              As String = MySignature & "Type Suffixes are obsolete"
Private Const ToDo13              As String = MySignature & "Type Suffix replaced"
Private Const ToDo14              As String = MySignature & "Modifies active For-Variable"
Private Const ToDo15              As String = MySignature & "Missing Scope"
Private Const ToDo16              As String = MySignature & "Remove Line Number"
Private Const Done16              As String = MySignature & "Line Number Removed"
Private Const ToDo17              As String = MySignature & "Possible Structure Violation"
Private Const ToDo18              As String = MySignature & "Dead Code?"
Private Const ToDo19              As String = MySignature & "Use ""Do...Loop"""
Private Const ToDo20              As String = MySignature & "Avoid ""GoTo"""
Private Const ToDo21              As String = MySignature & "Remove ""Call"" verb"
Private Const ToDo22              As String = " and brackets"
Private Const ToDo23              As String = MySignature & "Check Error Handling Structure"
Private Const ToDo24              As String = MySignature & "No executable Code"
Private Const ToDo25              As String = MySignature & "Not suitable for Case Preservation"
Private Const ToDo26              As String = MyErrMark & MySignature & "Move Comment inside Sub/Function/Property"
Private Const ToDo27              As String = MySignature & "Temporary Debugging Code"
Private Const ToDo28              As String = MySignature & "There are better ways to terminate"
Private Const ToDo29              As String = MySignature & "Combine with previous ""Next"""
Private Const InsertedBy          As String = Smiley & "Line inserted by Formatter"
Private Const ReplacedBy          As String = MySignature & "--> replaced by:"
Private Const OptExpl             As String = "Option Explicit " & InsertedBy
Private Const sNone               As String = "[None]"
Private Const StackUnderflow      As String = MySignature & sNone
Private Const ExitDo              As String = " 'loop" & MyErrMark
Private Const ExitFor             As String = ExitDo & "varying "
Private Const ExitSFP             As String = " '--->" & MyErrMark & "Bottom"
Private Const vb2CrLf             As String = vbCrLf & vbCrLf
Private Const Quote               As String = """"
Private Const HashChar            As String = "#"
Private Const Apostrophe          As String = "'"
Private Const BlnkApos            As String = " '"
Private Const Colon               As String = ":"
Private Const Comma               As String = ","
Private Const ContMark            As String = " _"
Private Const ListTitle           As String = "Source Code Listing For Project"
Private Const NaD                 As String = "Names and Definitions"
Private Const PntHyphen           As String = "¬" 'hyphen used for printing
Private Const GridPrintColor      As Long = &HE8E8E8 'some shades of gray for code printer
Private Const FramePrintColor     As Long = &H707070
Private Const NaDHeadColor        As Long = &HE0E0E0
Private Const CommentPrintColor   As Long = &H808080
Private Const AlmostInfinity      As Long = 999999 'used for dead code recognition
Private Const WordWrap            As Long = 12 'max len for wordwrap during printing
Private Const EnumLineLen         As Long = 128 'max length of an enum capitalization line

'other addins
Private Const CopyFacTtl          As String = "Select Text To Copy"
Private Const XRefTtl             As String = "Ulli's VB Cross Reference"
Private Const ExplTtl             As String = "Ulli's VB Project Explorer - "
Private Enum MemAttrPtrs
    MemName = 0
    MemBind = 1
    MemBrws = 2
    MemCate = 3
    MemDfbd = 4
    MemDesc = 5
    MemDbnd = 6
    MemHelp = 7
    MemHidd = 8
    MemProp = 9
    MemRqEd = 10
    MemStMe = 11
    MemUiDe = 12
End Enum
#If False Then
Private MemName, MemBind, MemBrws, MemCate, MemDfbd, MemDesc, MemDbnd, MemHelp, MemHidd, MemProp, MemRqEd, MemStMe, MemUiDe
#End If

'Menu
Private CommandBarMenu            As CommandBar
Private MenuItem                  As CommandBarControl
Private WithEvents MenuEvents     As CommandBarEvents
Attribute MenuEvents.VB_VarHelpID = -1

'Special Variables
Private SelCoord                  As Rect
Private VarNames                  As Collection
Private OtherAddIns               As Variant
Private AvailablePrinter          As Printer
Private NodeStack                 As cStack
Private WordStack()               As String
Private StrucStack()              As String
Private RootNode                  As MSComctlLib.Node
Private CurrParentNode            As MSComctlLib.Node
Private CurrChildNode             As MSComctlLib.Node
Private PB                        As PageBoundaries
Private Member                    As Member
Private MemberAttributes()        As Variant 'vbArray of 13 values
Private TimePrinted               As Date

'Variables (Longs thru DefLng)
Private i, j, k
Private PanelNum
Private NumFormatted
Private LineIndex
Private NumDeclLines
Private NumCodeLines
Private NumCommentOnlyLines
Private NumCommentedLines
Private NumEmptyLines
Private XPLookLine
Private FromLine
Private ToLine
Private NextEnumBreak              'inserts line break into enum capitalization line
Private ErrLineFrom
Private ErrLineTo
Private NodeKey
Private Skipped
Private Suppressed
Private Inserted
Private TopLine
Private FirstCodeLine
Private BracketCount               'used for replacing type suffixes
Private PrintLineNumber
Private HalfTabWidth
Private Indent
Private OELevel
Private CodeIsDead                 'set to -(Indent) on keyword causing dead code, and positive afterwards
Private NumCodeLinesInProc
Private MaxIndent
Private IndentColor
Private Colored
Private FrameTop
Private IndentNext
Private LastNext                   'used to mark Next Next
Private RemIndent
Private RemIndentNext
Private LastWordIndex
Private Pleonasm
Private MarkingIsOff               '1=inside markoff/on 2=inside #If False

'Options and other stuff
Private MarkChecked               As Boolean
Private SepChecked                As Boolean
Private IfExpChecked              As Boolean
Private IfExpSilent               As Boolean
Private EnumChecked               As Boolean

Private CallChecked               As Boolean
Private LinNumChecked             As Boolean
Private EmptyLinesChecked         As Boolean
Private DefConstType              As Boolean
Private DuplNameDetected          As Boolean
Private UndoRequested             As Boolean
Private TypeSuffixFound           As Boolean
Private FinalReq                  As Boolean

'Program States
Private ProcessingLastPanel       As Boolean
Private NowPrintingComment        As Boolean
Private InEnum                    As Boolean 'true while we're inside an enumeration
Private LocalDim                  As Boolean 'true when local 'Dim' etc is found
Private Skipping                  As Boolean
Private NewProcStarting           As Boolean 'true when a proc is starting and kept alive by Rem's after newproc
Private InProcHeader              As Boolean 'takes over NewProcStarting after procs first line

'Complaints
Private Dummy                     As Boolean 'dummy, just to have a param for MarkLine
Private NoOptExpl                 As Boolean
Private FoundGoTo                 As Boolean
Private MissingScope              As Boolean
Private NonProc                   As Boolean 'found a non-procedural line among procedural lines
Private VaCoFuDcl                 As Boolean 'Complain about variable type or definition
Private SlIf                      As Boolean 'Single line If found
Private LNsObs                    As Boolean 'Line Numbers obsolete
Private CallUnnec                 As Boolean 'Call unnecessary
Private ContLine                  As Boolean 'Continued Single-Line-If
Private Pleo                      As Boolean 'Pleonasm found
Private ForVarMod                 As Boolean 'found a For-Variable Modification
Private Dupl                      As Boolean 'Complain about name duplicates
Private EmptyNext                 As Boolean
Private EmptyComplain             As Boolean 'complain about empty procedures
Private NoOEClose                 As Boolean 'Complain about an unclosed "On Error.."
Private StrucErr                  As Boolean
Private OrphCom                   As Boolean 'Orphaned comment found
Private PossVio                   As Boolean 'possible structure violation
Private NoCode                    As Boolean
Private DeadCode                  As Boolean
Private NoXPLook                  As Boolean
Private XPDone                    As Boolean
Private Complain                  As Boolean 'true when we have something to complain
Private AnyComplain               As Boolean

Private CtB()                     As Byte

'Assortment of Strings
Private ProjectName               As String
Private ProjectFile               As String
Private WindowTitle               As String
Private ModuleName                As String
Private StartUpCompoName          As String
Private StartUpProcName           As String
Private EXEName                   As String
Private CodeLine                  As String
Private CodeLineWith              As String
Private CodeLineWithout           As String
Private FirstWord                 As String
Private Words()                   As String
Private EnumMembers               As String
Private HalfIndent                As String
Private OEIndent                  As String
Private OEIndentNext              As String
Private DefTypeChars              As String
Private TmpString1                As String
Private TmpString2                As String
Private LNum                      As String
Private CurrProcType              As String
Private SrcFileName               As String
Private ActiveForVars             As String
Private PendingTSRepl             As String 'deferred type suffix replacement
Private PendingMark               As String
Private ToDo(1 To 29)             As String 'error strings for marking

Private Sub AddinInstance_OnBeginShutdown(custom() As Variant)

    Erase UndoBuffers

End Sub

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)

    Set VBInstance = Application
    If ConnectMode = ext_cm_External Then
        FormatAll
      Else 'NOT CONNECTMODE...
        On Error Resume Next
            Set CommandBarMenu = VBInstance.CommandBars(MenuName)
        On Error GoTo 0
        If CommandBarMenu Is Nothing Then
            MsgBox "Code Formatter was loaded but could not be connected to the " & MenuName & " menu.", vbCritical, AppDetails()
          Else 'NOT COMMANDBARMENU...
            fSplash.imgUMG.Picture = fProgress.img.Picture
            fSplash.Show
            DoEvents
            With CommandBarMenu
                Set MenuItem = .Controls.Add(msoControlButton)
                i = .Controls.Count - 1
                If .Controls(i).BeginGroup And Not .Controls(i - 1).BeginGroup Then
                    'menu separator required  
                    MenuItem.BeginGroup = True
                End If
            End With 'COMMANDBARMENU
            'set menu caption  
            MenuItem.Caption = "&" & AppDetails & "..."
            With Clipboard
                'set menu picture  
                TmpString1 = .GetText
                .SetData fProgress.picMenu.Image
                MenuItem.PasteFace
                .Clear
                .SetText TmpString1
            End With 'CLIPBOARD
            'set event handler  
            Set MenuEvents = VBInstance.Events.CommandBarEvents(MenuItem)
            'done connecting  
            Sleep SleepTime
            Unload fSplash
            Load fStrFncts
            For i = 0 To UBound(bStrFncts)
                bStrFncts(i) = (Right$(fStrFncts.lstStrFncts.List(i), 1) = Spce)
            Next i
            Unload fStrFncts

            'StringFunctsReq,'PauseAfterScan,InsertComments,InsertMarks,SplitLines,HalfIndent  
            StoreSettings vbUnchecked, IfNecessary, vbChecked, vbChecked, vbChecked, vbChecked

            On Error Resume Next 'try to switch the printer to color
                Printer.ColorMode = vbPRCMColor
                ColorRequested = (Printer.ColorMode = vbPRCMColor) 'if success then preset to print in color
            On Error GoTo 0
            BookRequested = True
            WithStationary = True
        End If
        Unload fProgress
    End If

End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)

    On Error Resume Next
        MenuItem.Delete
    On Error GoTo 0

End Sub

Private Function BuildExitFor() As String

    On Error Resume Next 'coz this may be called from an IIF
        BuildExitFor = Mid$(ActiveForVars, InStrRev(ActiveForVars, "=", Len(ActiveForVars) - 1) + 1)
        BuildExitFor = ExitFor & Left$(BuildExitFor, Len(BuildExitFor) - 1)
    On Error GoTo 0

End Function

Private Sub CheckGoto(Extra As String)

  'Checks whether there is a GoTo in the codeline and marks it as an error

    If InStr(Extra & CodeLine, " goto ") Then
        MarkLine ToDo(20), FoundGoTo
    End If

End Sub

Private Function CheckVardefs(ByVal WordIndex As Long, LocalDef As Boolean, IsConst As Boolean, IsFunction As Boolean) As Boolean

  'Returns True if line containes untyped variable  
  'May also set the DuplNameDetected, TypeSuffixFound, and/or DefConstType flags  

    DuplNameDetected = False
    DefConstType = False
    TypeSuffixFound = False
    Do
        If Len(Words(WordIndex)) Then
            TmpString1 = Replace$(Words(WordIndex), Comma, NullStr) 'remove comma
            i = InStr(TmpString1, "(")
            If i Then
                TmpString1 = Left$(TmpString1, i - 1) 'remove open bracket and rest
                Do Until InStr(Words(WordIndex), ")") 'find close bracket
                    Inc WordIndex
                Loop
            End If
            'TmpString1 has variable/const name  
            If TypeIsUndefined(TmpString1, IsConst) Then
                If IsConst Then
                    If WordIndex = LastWordIndex Then 'don't stumble into possible syntax error
                        DefConstType = True
                      Else 'NOT WORDINDEX...
                        If Words(WordIndex + 1) = "=" Then
                            If Val(Words(WordIndex + 2)) >= -32768 And Val(Words(WordIndex + 2)) <= 32767 And Val(Words(WordIndex + 2)) = Int(Val(Words(WordIndex + 2))) Then
                                DefConstType = Not (HasTypeSuffix(Words(WordIndex + 2)) Or Left$(Words(WordIndex + 2), 1) = Quote Or (Left$(Words(WordIndex + 2), 1) = "&" And Len(Words(WordIndex + 2)) > 6))
                            End If
                        End If
                    End If
                  Else 'ISCONST = FALSE/0
                    If WordIndex = LastWordIndex Or Right$(Words(WordIndex), 1) = Comma Then
                        'no DefType and no type suffix and no 'As'  
                        CheckVardefs = True 'return True
                    End If
                End If
            End If
            If HasTypeSuffix(TmpString1) Then
                TmpString1 = Left$(TmpString1, Len(TmpString1) - 1) 'remove type suffix
                TypeSuffixFound = True
            End If
            With VarNames
                On Error Resume Next
                    If LocalDef Then
                        TmpString1 = .Item(TmpString1)
                        DuplNameDetected = DuplNameDetected Or (Err = 0)
                      Else 'LOCALDEF = FALSE/0
                        .Add True, TmpString1
                        DuplNameDetected = DuplNameDetected Or (Err <> 0)
                    End If
                On Error GoTo 0
            End With 'VARNAMES
            Do Until Right$(Words(WordIndex), 1) = Comma
                If WordIndex = LastWordIndex Then
                    Exit Do 'loop 
                End If
                Inc WordIndex
            Loop
        End If
        Inc WordIndex
    Loop Until WordIndex > LastWordIndex Or IsFunction

End Function

Private Sub ClearToDo(ByVal Reason As Long)

    MarkingIsOff = MarkingIsOff Or Reason
    For i = LBound(ToDo) To UBound(ToDo)
        ToDo(i) = NullStr
    Next i

End Sub

Private Sub CompileProject()

  'Compiles the current project

  Dim BuildFile     As String
  Dim ErrNmbr       As Long
  Dim ErrText       As String
  Dim DidntExist    As Boolean
  Dim Attr          As Long 'stores original .EXE file attributes
  Const CF          As String = "Compilation failed"

    fCompile.Show
    StrucRequested = False
    PrintLineLen = 0
    With fSummary
        .lblComplaints = vbCrLf
        .lblComplaints.ToolTipText = "Compile results"
        .StopButtonVisible = False
        .ForCompiling = True
    End With 'FSUMMARY
    DoEvents
    Sleep 555
    With VBInstance.ActiveVBProject
        If Len(.BuildFileName) Then 'VB knows a file name
            BuildFile = Dir$(.BuildFileName) 'locate that file
            If Len(BuildFile) And LCase$(Right$(.BuildFileName, Len(BuildFile))) = LCase$(BuildFile) Then 'file exists
                ErrNmbr = 0
              Else 'NOT LEN(BUILDFILE)...
                i = FreeFile
                Err.Clear
                On Error Resume Next
                    Open .BuildFileName For Output As i '...so create a file to compile into
                    ErrNmbr = Err
                    Close i
                On Error GoTo 0
                DidntExist = True
            End If
            If ErrNmbr = 0 Then 'file was present or successfully created
                Attr = GetAttr(.BuildFileName)
                SetAttr .BuildFileName, vbNormal 'reset .exe file attributes

                Err.Clear
                On Error Resume Next 'try a compilation and ignore any errors
                    .MakeCompiledFile
                    ErrNmbr = Err
                    ErrText = Err.Description
                On Error GoTo 0 'no more errors, please

                If GetAttr(.BuildFileName) = vbArchive Then '.exe file should now have the archive attribut (because it was newly written by VB, or rather by the Linker)
                    CompileResults "Compilation successful", False, .BuildFileName & " successfully " & IIf(DidntExist, "created.", "updated."), vbGreen
                    SetAttr .BuildFileName, Attr
                  Else 'NOT GETATTR(.BUILDFILENAME)...
                    CompileResults CF, True, CF & " with Error " & Hex$(ErrNmbr) & vb2CrLf & Replace$(Replace$(ErrText, "~", "MakeCompiledFile", , 1), "~", .Name, , 1) & vb2CrLf & .BuildFileName & " not " & IIf(DidntExist, "created.", "updated."), vbRed
                    SetAttr .BuildFileName, vbNormal
                    If DidntExist Then
                        Kill .BuildFileName
                    End If
                End If
              Else 'NOT ERRNMBR...
                CompileResults CF, True, "Error " & Hex$(ErrNmbr) & vb2CrLf & App.ProductName & " could not create " & .BuildFileName, vbRed
            End If
          Else 'LEN(.BUILDFILENAME) = FALSE/0
            CompileResults CF, True, "The initial compilation must be made using the IDE.", vbRed
        End If
    End With 'VBINSTANCE.ACTIVEVBPROJECT
    Sleep 888
    DoEvents
    Unload fCompile
    fSummary.Show
    SetWindowPos fSummary.hWnd, SWP_TOPMOST, 0, 0, 0, 0, SWP_COMBINED
    Do
        DoEvents
    Loop While fSummary.Visible
    Unload fSummary
    FinalReq = False

End Sub

Private Sub CompileResults(Capt As String, IsSerious As Boolean, Comment As String, ByVal Colored As Long)

    With fSummary
        .lblSummary = Capt
        .Serious = IsSerious
        .lblComplaints = Comment
    End With 'FSUMMARY
    fCompile.BackColor = Colored

End Sub

Private Sub EmitChars()

  'Emit deftype range chars  

    For i = 1 To LastWordIndex
        j = Asc(Left$(Words(i), 1))
        If Len(Words(i)) < 3 Then
            k = j
          Else 'NOT LEN(WORDS(I))...
            k = Asc(Mid$(Words(i), 3, 1))
        End If
        If j = Asc("a") And k = Asc("z") Then
            For j = 0 To 255 'emit all deftype characters
                DefTypeChars = DefTypeChars & Chr$(j)
            Next j
          Else 'NOT J...
            Do 'emit deftype range characters
                DefTypeChars = DefTypeChars & Chr$(j)
                Inc j
            Loop Until j > k
        End If
    Next i

End Sub

Private Sub FormatAll()

  Dim TmpLong      As Long
  Dim ActivePane   As CodePane
  Dim ActiveState  As Long

    MouseButtonsSwapped = GetSystemMetrics(SM_SWAPBUTTON)
    If VBInstance.ActiveVBProject Is Nothing Then
        ProjectName = "No Project"
        OtherAddIns = Empty
      Else 'NOT VBINSTANCE.ACTIVEVBPROJECT...
        With VBInstance.ActiveVBProject
            NoXPLook = False
            XPDone = False
            ProjectName = .Name
            ProjectFile = .FileName
            If .Type = vbext_pt_StandardExe Or .Type = vbext_pt_ActiveXExe Then
                StartUpCompoName = NullStr 'dont know yet
                StartUpProcName = "Sub Main" 'just guessing
                On Error Resume Next
                    With .VBComponents.StartUpObject
                        StartUpCompoName = .Name
                        If Err = 0 Then 'no error - startup is a form
                            If .Type = vbext_ct_VBForm Then
                                StartUpProcName = "Sub Form_Initialize" 'so now we know
                              Else 'NOT .TYPE...
                                StartUpProcName = "Sub MDIForm_Initialize"
                            End If
                        End If
                    End With '.VBCOMPONENTS.STARTUPOBJECT
                On Error GoTo 0
              Else 'NOT .TYPE...
                NoXPLook = True
            End If
        End With 'VBINSTANCE.ACTIVEVBPROJECT
        OtherAddIns = Array(CopyFacTtl, XRefTtl, ExplTtl & VBInstance.ActiveVBProject.Name)
        For i = LBound(OtherAddIns) To UBound(OtherAddIns)
            j = FindWindow(ByVal 0&, OtherAddIns(i))
            If j Then
                ShowWindow j, SW_MINIMIZE
            End If
        Next i
        DoEvents
    End If
    If ProjectFile = NullStr Then
        ProjectFile = "Unknown File Name"
    End If
    If Len(ProjectFile) > 60 Then
        ProjectFile = Trim$(Left$(ProjectFile, 30)) & ElipsisChar & Trim$(Right$(ProjectFile, 30))
    End If
    NumPanels = VBInstance.CodePanes.Count
    If NumPanels Then
        'get current user  
        k = 128
        UserName = AllocString(k)
        GetUserName UserName, k
        UserName = Left$(UserName, k + (Asc(Mid$(UserName, k, 1)) = 0))
        'get VB tabwidth and font properties from registry  
        If RegOpenKeyEx(HKEY_CURRENT_USER, VBSettings, REG_OPTION_RESERVED, KEY_QUERY_VALUE, i) <> ERROR_NONE Then
            FullTabWidth = DefaultTabWidth
            MyFontName = DefaultFontName
            MyFontSize = DefaultFontSize
          Else 'NOT REGOPENKEYEX(HKEY_CURRENT_USER,...
            k = Len(FullTabWidth)
            If RegQueryValueEx(i, TabWidth, REG_OPTION_RESERVED, j, FullTabWidth, k) <> ERROR_NONE Then
                FullTabWidth = DefaultTabWidth
            End If
            HalfTabWidth = FullTabWidth \ 2
            TmpString1 = AllocString(128)
            k = Len(TmpString1)
            If RegQueryValueEx(i, Fontface, REG_OPTION_RESERVED, j, ByVal TmpString1, k) <> ERROR_NONE Then
                TmpString1 = DefaultFontName
              Else 'NOT REGQUERYVALUEEX(I,...
                TmpString1 = Left$(TmpString1, k + (Asc(Mid$(TmpString1, k, 1)) = 0))
            End If
            If TmpString1 <> IDEFontName Then
                IDEFontName = TmpString1
                MyFontName = TmpString1
            End If
            k = Len(IDEFontSize)
            If RegQueryValueEx(i, Fontheight, REG_OPTION_RESERVED, j, IDEFontSize, k) <> ERROR_NONE Then
                IDEFontSize = DefaultFontSize
            End If
            If IDEFontName = MyFontName Then
                MyFontSize = IDEFontSize
            End If
            RegCloseKey i
        End If

        'see if a printer is available (this is here thanks to Tom Law)  
        HasPrinter = False
        TmpString1 = "No Printer"
        On Error Resume Next
            TmpString1 = Printer.DeviceName
        On Error GoTo 0
        For Each AvailablePrinter In Printers
            If AvailablePrinter.DeviceName = TmpString1 Then
                HasPrinter = True
                Exit For 'loop varying availableprinter
            End If
        Next AvailablePrinter

        If HasPrinter Then
            With Printer
                .FontBold = False
                MyFontName = GetSetting(App.Title, "Print", "Fontname", MyFontName) 'replace MyFontName if it's in registry
                MyFontSize = GetSetting(App.Title, "Print", "Fontsize", MyFontSize) 'replace MyFontSize if it's in registry
                .Fontname = MyFontName
                .Fontsize = MyFontSize
                PrintingOK = (.Fontname = MyFontName) And (.TextWidth("W") = .TextWidth("I"))
                'save current values  
                TmpString2 = MyFontName
                k = MyFontSize
                Do Until PrintingOK Or k < 0
                    With fPrintFont
                        .Show vbModal
                        MyFontName = .PrintFontName
                        MyFontSize = .PrintFontSize
                        i = .chkSave
                    End With 'FPRINTFONT
                    If MyFontName = NullStr Then 'user canceled font selection; restore
                        MyFontName = TmpString2
                        MyFontSize = k
                        k = -1
                    End If
                    DoEvents
                    Unload fPrintFont
                    .Fontname = MyFontName
                    .Fontsize = MyFontSize
                    PrintingOK = (.Fontname = MyFontName) And (.TextWidth("W") = .TextWidth("I"))
                Loop
                If PrintingOK And (i = vbChecked) Then
                    SaveSetting App.Title, "Print", "Fontname", .Fontname
                    SaveSetting App.Title, "Print", "Fontsize", .Fontsize
                End If
                PrintLineHeight = .TextHeight("A")
                PrintCharWidth = .TextWidth("A")
                .FontBold = True
                j = .TextWidth("A")
                'some fonts have different glyph dimensions bold/thin  
                PrinterBoldEnabled = (j = PrintCharWidth)
                .FontBold = False
                .FontItalic = True
                i = .TextHeight("A")
                j = .TextWidth("A")
                .FontItalic = False
                'some fonts have different glyph height straight/italic, get max of them  
                PrintLineHeight = IIf(i > PrintLineHeight, i, PrintLineHeight)
                'some fonts have different glyph dimensions straight/italic  
                PrinterItalEnabled = (j = PrintCharWidth)
                .FontTransparent = True
                PrintLineLen = 0
            End With 'PRINTER
        End If 'HasPrinter
        With fQuestion
            .LoadListbox

#If CallComplain = False Then
            .ckCall.Enabled = False
#End If
#If Not LinNumComplain Then
            .lbl(0) = Replace$(.lbl(0), "must  not", "may now")
            .ckLinNum.Enabled = False
#End If
            XPLookRequested = 0
            j = GetModuleCount
            .ckWinXPLook.Enabled = (IsWindowsSuitable And Not NoXPLook _
                                   And VBInstance.CodePanes.Count = j) 'all code panels are open so we enable WinXP because the startup object is definitely among them
            .btShowAll.Enabled = (VBInstance.CodePanes.Count <> j) 'not all code panels are open so we enable btShowAll
            .Show vbModal
            If .Reply = Continue Or .Reply = Undo Then
                UndoRequested = (.Reply = Undo)
                If UndoRequested Then
                    fUndoList.lstUndone.Clear
                End If
                On Error Resume Next 'allow error because Ubound() does not work with un-redimmed UndoBuffers
                    If UBound(UndoBuffers) < .lstModNames.ListCount Then
                        ReDim Preserve UndoBuffers(0 To .lstModNames.ListCount)
                        ReDim Preserve UndoTitles(1 To .lstModNames.ListCount)
                    End If
                On Error GoTo 0
                MarkChecked = (.ckMark = vbChecked)
                SepChecked = (.ckSep = vbChecked)
                EnumChecked = (.ckEnum = vbChecked)
                IfExpChecked = (.ckIfExp = vbChecked)
                IfExpSilent = CBool(Val(.ckIfExp.Tag))
                CallChecked = (.ckCall = vbChecked)
                LinNumChecked = (.ckLinNum = vbChecked)
                EmptyLinesChecked = (.ckEmptyLines = vbChecked)
                If .ckHalfIndent = vbUnchecked Then
                    HalfTabWidth = 0
                End If
                PanelNum = 0
                BreakLoop = False
                KillDoc = False
                NumFormatted = 0
                If PrintLineLen And NumSelected <> 0 Then
                    'Print Cover Sheet  
                    With Printer
                        .CurrentY = .ScaleHeight / 3
                        .Fontname = "Arial"
                        .FontBold = True
                        .Fontsize = 12
                        .CurrentX = (.ScaleWidth - .TextWidth(ListTitle) + PBOdd.Left) / 2
                        Printer.Print ListTitle; vbCrLf
                        .Fontsize = 24
                        .CurrentX = (.ScaleWidth - .TextWidth(ProjectName) + PBOdd.Left) / 2
                        Printer.Print ProjectName; vbCrLf
                        .FontBold = False
                        .Fontsize = 12
                        TmpString1 = VBInstance.ActiveVBProject.Description
                        .CurrentX = (.ScaleWidth - .TextWidth(TmpString1) + PBOdd.Left) / 2
                        Printer.Print TmpString1; vbCrLf
                        Select Case VBInstance.ActiveVBProject.Type
                          Case vbext_pt_ActiveXControl
                            TmpString1 = "ActiveX OCX"
                          Case vbext_pt_ActiveXDll
                            TmpString1 = "ActiveX DLL"
                          Case vbext_pt_ActiveXExe
                            TmpString1 = "ActiveX EXE"
                          Case vbext_pt_StandardExe
                            TmpString1 = "Standard EXE"
                        End Select
                        .Fontsize = 14
                        .CurrentX = (.ScaleWidth - .TextWidth(TmpString1) + PBOdd.Left) / 2
                        Printer.Print TmpString1; vbCrLf
                        .CurrentX = (.ScaleWidth - .TextWidth(ProjectFile & "()") + PBOdd.Left) / 2
                        Printer.Print "("; ProjectFile; ")"; vb2CrLf; vb2CrLf
                        TimePrinted = Now
                        TmpString1 = "Printed on " & Format$(TimePrinted, "Long Date") & ", at " & Format$(TimePrinted, "Short Time")
                        .Fontsize = 9
                        .CurrentX = (.ScaleWidth - .TextWidth(TmpString1) + PBOdd.Left) / 2
                        Printer.Print TmpString1
                        TmpString1 = " by " & App.ProductName
                        .CurrentX = (.ScaleWidth - .TextWidth(TmpString1) + PBOdd.Left) / 2
                        Printer.Print TmpString1
                        Printer.Print
                        TmpString1 = "Printing initiated through " & UserName
                        .CurrentX = (.ScaleWidth - .TextWidth(TmpString1) + PBOdd.Left) / 2
                        Printer.Print TmpString1
                        Printer.Line (PBOdd.Left, PBOdd.Top)-(PBOdd.Right, PBOdd.Bottom), vbBlack, B
                        If .Orientation = vbPRORLandscape Then
                            Printer.Line (PBOdd.PunchX, PBOdd.PunchY)-(PBOdd.PunchX, PBOdd.PunchY + LenPunchMark * .TwipsPerPixelY), vbBlack
                          Else 'NOT .ORIENTATION...
                            Printer.Line (PBOdd.PunchX, PBOdd.PunchY)-(PBOdd.PunchX + LenPunchMark * .TwipsPerPixelX, PBOdd.PunchY), vbBlack
                        End If
                        .NewPage
                        .Fontname = MyFontName
                        .Fontsize = MyFontSize
                    End With 'PRINTER
                  Else 'NOT PRINTLINELEN...
                    KillDoc = True
                End If
                'create manifest file  
                EXEName = VBInstance.ActiveVBProject.BuildFileName
                If XPLookRequested Then 'any bit is on
                    If Not NoXPLook Then
                        If Len(Dir$(Left$(EXEName, InStrRev(EXEName, "\")), vbDirectory)) Then 'valid directory; create manifest file
                            i = FreeFile
                            Open EXEName & ".Manifest" For Output As i
                            Print #i, Replace$(XPLookXML, "°", UserName & "." & Replace$(Mid$(EXEName, InStrRev(EXEName, "\") + 1), ".exe", NullStr, , , vbTextCompare))
                            Close i
                          Else 'NOT LEN(DIR$(LEFT$(EXENAME,...
                            KillManifest
                            XPLookRequested = 0
                            NoXPLook = True
                        End If
                    End If
                  Else 'XPLOOKREQUESTED = FALSE/0
                    NoXPLook = False
                    KillManifest
                End If
                AnyComplain = False
                Set ActivePane = VBInstance.ActiveCodePane
                ActiveState = ActivePane.Window.WindowState
                For Each Pane In VBInstance.CodePanes
                    Inc PanelNum
                    If .lstModNames.Selected(PanelNum - 1) Then
                        With Pane.Window
                            TmpLong = .WindowState
                            .WindowState = vbext_ws_Maximize
                            Pane.Show
                            DoEvents
                            Inc NumFormatted
                            ProcessingLastPanel = (NumFormatted = NumSelected)
                            If HasPrinter Then
                                Printer.CurrentY = 0
                            End If
                            FormatCode
                            .WindowState = TmpLong
                            .SetFocus
                            DoEvents
                        End With 'PANE.WINDOW
                    End If
                    If BreakLoop Then
                        KillDoc = PrintLineLen
                        ProcessingLastPanel = True
                        Exit For 'loop varying pane
                    End If
                    If Not ProcessingLastPanel And PrintLineLen Then
                        If Printer.CurrentY > PB.Top Then
                            Printer.NewPage 'page break for next code pane
                        End If
                    End If
                Next Pane
                Unload fProgress
                If UndoRequested And fUndoList.lstUndone.ListCount Then
                    MessageBeep vbInformation
                    DoEvents
                    fUndoList.Show vbModal
                    Unload fUndoList
                  ElseIf .ckCompile = vbChecked And (Not AnyComplain Or .ckSumma = vbUnchecked) Then 'NOT UNDOREQUESTED...
                    CompileProject
                    If FinalReq Then
                        fAllDone.Show vbModal
                        Unload fAllDone
                    End If
                  ElseIf FinalReq Then 'NOT .CKCOMPILE...
                    MessageBeep vbInformation
                    DoEvents
                    fAllDone.Show vbModal
                    Unload fAllDone
                End If
                ActivePane.Window.WindowState = ActiveState
                ActivePane.Show
            End If
        End With 'FQUESTION
        Unload fQuestion
      Else 'NUMPANELS = FALSE/0
        MsgBox "Cannot see any code - you must open one or more Code Panels first.", vbExclamation, AppDetails
    End If
    If KillDoc Then
        On Error Resume Next 'not all printer drivers support KillDoc
            Printer.KillDoc
        On Error GoTo 0
      Else 'KILLDOC = FALSE/0
        If Printer.Page > 1 Then
            Printer.EndDoc
        End If
    End If
    If Not IsEmpty(OtherAddIns) Then
        For i = UBound(OtherAddIns) To LBound(OtherAddIns) Step -1
            j = FindWindow(ByVal 0&, OtherAddIns(i))
            If j Then
                ShowWindow j, SW_RESTORE
            End If
        Next i
    End If

End Sub

Private Sub FormatCode()

    Words = Split(VBInstance.ActiveWindow.Caption)
    WindowTitle = Words(0)
    TmpString1 = AppDetails & " [" & WindowTitle & "]"
    With VBInstance.ActiveCodePane
        With fProgress
            .Show
            .picProgress.Cls
            .Total = 100 * NumFormatted / NumSelected
            .lblXofY = NumFormatted & " of " & NumSelected & OneOrMany(" Pane", NumSelected)
        End With 'FPROGRESS
        DoEvents
        If .CodeModule.CountOfLines < 2 Then 'not enough code
            SrcFileName = NullStr
            ModuleName = WindowTitle
            If PauseAfterScan <> Never Then
                fProgress.Percent = 100
                FinalReq = (NumSelected > 1)
            End If
            If PrintLineLen Then
                PageSetup 1
                Printer.Print
                Printer.CurrentX = PB.Left
                Printer.FontBold = PrinterBoldEnabled
                Printer.FontItalic = False
                Printer.Print Space$(LnLen) & "     +++ This component (" & ModuleName & ") is not printed; there is not enough code +++"
            End If
          Else 'NOT .CODEMODULE.COUNTOFLINES...
            If UndoRequested Then
                'search for the correct undo buffer  
                For j = 1 To UBound(UndoTitles)
                    If UndoTitles(j) = .Window.Caption Then
                        Exit For 'loop varying j
                    End If
                Next j
                If j > UBound(UndoTitles) Then 'no buffer found
                    j = 0 'Undo Buffer 0 is always empty
                End If
                If Len(UndoBuffers(j)) Then
                    With .CodeModule
                        SaveMemberAttributes .Members
                        .DeleteLines 1, .CountOfLines
                        .AddFromString UndoBuffers(j)
                        RestoreMemberAttributes .Members
                    End With '.CODEMODULE
                    UndoBuffers(j) = NullStr 'clear undo buffer
                    fUndoList.lstUndone.AddItem .Window.Caption 'save name in undo list
                  Else 'no undo buffer 'LEN(UNDOBUFFERS(J)) = FALSE/0
                    fNoUndo.lbName = "Undo [" & .Window.Caption & "]"
                    With fNoUndo
                        .IsLastPanel = ProcessingLastPanel
                        .Show vbModal
                        If .Tag = "1" Then
                            BreakLoop = True
                        End If
                    End With 'FNOUNDO
                    Unload fNoUndo
                    FinalReq = False
                End If
              Else 'UNDOREQUESTED = FALSE/0
                'save current contents for undo  
                UndoBuffers(PanelNum) = .CodeModule.Lines(1, .CodeModule.CountOfLines)
                UndoTitles(PanelNum) = .Window.Caption
                'save the current position and selection  
                TopLine = .TopLine
                .GetSelection SelCoord.Top, SelCoord.Left, SelCoord.Bottom, SelCoord.Right
                With .CodeModule
                    With fPreparing
                        .lbl(0) = "Preparing..."
                        .imgSort.Visible = False
                        .imgFormat.Visible = False
                        .imgBroom.Visible = True
                        .Show
                    End With 'FPREPARING
                    DoEvents
                    SrcFileName = .Parent.FileNames(1)
                    If Len(SrcFileName) > 40 Then
                        SrcFileName = Left$(SrcFileName, 20) & ElipsisChar & LTrim$(Right$(SrcFileName, 20))
                    End If
                    SrcFileName = SrcFileName & IIf(.Parent.IsDirty, "(unsaved)", NullStr)
                    ModuleName = .Parent.Name
                    If Trim$(.Lines(.CountOfDeclarationLines + 1, 1)) <> NullStr Then
                        'insert a blank line between declarations and code  
                        If EmptyLinesChecked = False Then
                            .InsertLines .CountOfDeclarationLines + 1, NullStr
                            KillSelection
                        End If
                    End If
                    SaveMemberAttributes .Members
                    'remove bottom empty lines  
                    Do Until .CountOfLines = 0 'delete trailing blank lines
                        CodeLine = Trim$(.Lines(.CountOfLines, 1))
                        If CodeLine = NullStr Or Left$(CodeLine & "   ", Len(Smiley)) = Smiley Then
                            .DeleteLines .CountOfLines
                          Else 'NOT CODELINE...
                            Exit Do 'loop 
                        End If
                    Loop
                    LineIndex = 1
                    Do
                        If EmptyLinesChecked Then
                            If Len(Trim$(.Lines(LineIndex, 1))) = 0 Then
                                .DeleteLines LineIndex
                            End If
                        End If
                        'remove dupl empty lines and error summary lines  
                        If (Left$(.Lines(LineIndex, 1), Len(MyErrMark)) = MyErrMark) Or (Left$(.Lines(LineIndex, 1), Len(Smiley)) = Smiley) Or .Lines(LineIndex, 1) = ContMark Or ((Trim$(.Lines(LineIndex, 1)) = NullStr Or Trim$(.Lines(LineIndex, 1)) = Colon) And Trim$(.Lines(LineIndex + 1, 1)) = NullStr) Then
                            .DeleteLines LineIndex
                            KillSelection
                          Else 'NOT (LEFT$(.LINES(LINEINDEX,...
                            CodeLine = Spce & .Lines(LineIndex, 1)
                            i = Len(CodeLine)
                            If StrFnctsRequested Then
                                'modify selected string functions  
                                With fStrFncts.lstStrFncts
                                    For j = 0 To .ListCount - 1
                                        If bStrFncts(j) And InStr(CodeLine, " As ") = 0 Then
                                            TmpString1 = RTrim$(.List(j))
                                            TmpString2 = "(" & TmpString1
                                            TmpString1 = Spce & TmpString1
                                            CodeLine = Replace$(CodeLine, TmpString1 & "(", TmpString1 & "$(")
                                            CodeLine = Replace$(CodeLine, TmpString2 & "(", TmpString2 & "$(")
                                        End If
                                    Next j
                                End With 'FSTRFNCTS.LSTSTRFNCTS
                            End If

                            'XP Look  
                            If XPLookRequested Then
                                Select Case StartUpCompoName
                                  Case ModuleName, NullStr
                                    If Mid$(CodeLine, 2, Len(XPLookAPIProto)) = XPLookAPIProto Then
                                        XPLookRequested = XPLookRequested And 2 'bit 1 off: prototype is already present
                                    End If
                                    If XPLookRequested And 2 Then 'bit 2 on: api call insertion required
                                        If InStr(1, CodeLine, StartUpProcName, vbTextCompare) Then
                                            XPLookRequested = (XPLookRequested And 1) Or 4 'bit 2 off - 3 on: insert api call before next End Sub
                                            StartUpCompoName = ModuleName
                                            XPLookLine = LineIndex + 1
                                        End If
                                    End If
                                    If Mid$(CodeLine, FullTabWidth + 2, Len(XPLookAPICall)) = XPLookAPICall Then
                                        XPLookRequested = XPLookRequested And 3 'bit 4 off: api call is present
                                    End If
                                End Select
                            End If

                            'remove any leftovers from a previous run  
                            k = InStrRev(CodeLine, ExitFor)
                            If k Then
                                CodeLine = Left$(CodeLine, k - 1)
                            End If
                            CodeLine = Replace$(CodeLine, ExitDo, NullStr)
                            CodeLine = Replace$(CodeLine, ExitSFP, NullStr)
                            If InStrRev(.Lines(LineIndex, 1), MySignature) Then
                                CodeLine = Replace$(CodeLine, ToDo2 & "sub", NullStr, , , vbTextCompare)
                                CodeLine = Replace$(CodeLine, ToDo2 & "function", NullStr, , , vbTextCompare)
                                CodeLine = Replace$(CodeLine, ToDo2 & "property", NullStr, , , vbTextCompare)
                                CodeLine = Replace$(CodeLine, ToDo3, NullStr)
                                CodeLine = Replace$(CodeLine, ToDo4, NullStr)
                                CodeLine = Replace$(CodeLine, ToDo5, NullStr)
                                CodeLine = Replace$(CodeLine, ToDo6, NullStr)
                                CodeLine = Replace$(CodeLine, ToDo8, NullStr)
                                CodeLine = Replace$(CodeLine, ToDo7, NullStr)
                                CodeLine = Replace$(CodeLine, ToDo9, NullStr)
                                CodeLine = Replace$(CodeLine, ToDo10, NullStr)
                                CodeLine = Replace$(CodeLine, ToDo11, NullStr)
                                CodeLine = Replace$(CodeLine, ToDo12, NullStr)
                                CodeLine = Replace$(CodeLine, ToDo13, NullStr)
                                CodeLine = Replace$(CodeLine, ToDo14, NullStr)
                                CodeLine = Replace$(CodeLine, ToDo15, NullStr)
                                CodeLine = Replace$(CodeLine, ToDo16, NullStr)
                                CodeLine = Replace$(CodeLine, Done16, NullStr)
                                CodeLine = Replace$(CodeLine, ToDo17, NullStr)
                                CodeLine = Replace$(CodeLine, ToDo18, NullStr)
                                CodeLine = Replace$(CodeLine, ToDo19, NullStr)
                                CodeLine = Replace$(CodeLine, ToDo20, NullStr)
                                CodeLine = Replace$(CodeLine, ToDo21 & ToDo22, NullStr)
                                CodeLine = Replace$(CodeLine, ToDo21, NullStr)
                                CodeLine = Replace$(CodeLine, ToDo23, NullStr)
                                CodeLine = Replace$(CodeLine, ToDo24, NullStr)
                                CodeLine = Replace$(CodeLine, ToDo25, NullStr)
                                CodeLine = Replace$(CodeLine, ToDo26, NullStr)
                                CodeLine = Replace$(CodeLine, ToDo27, NullStr)
                                CodeLine = Replace$(CodeLine, ToDo28, NullStr)
                                CodeLine = Replace$(CodeLine, ToDo29, NullStr)
                                If Left$(LTrim$(CodeLine), 1) <> Apostrophe Then
                                    CodeLine = Replace$(CodeLine, ReplacedBy, NullStr)
                                End If
                                CodeLine = Replace$(CodeLine, StackUnderflow, NullStr)
                                CodeLine = Replace$(CodeLine, sNone, NullStr)
                            End If
                            Do Until Right$(CodeLine, 1) <> Apostrophe
                                CodeLine = Left$(CodeLine, Len(CodeLine) - 1)
                            Loop

                            If SepChecked Then 'separate compound lines
                                CodeLine = RTrim$(CodeLine)
                                k = InStr(CodeLine, Colon)
                                If k Or InStr(CodeLine, " If ") Or InStr(CodeLine, " '") Or InStr(CodeLine, " Rem ") Or ContLine Then   'found possible instruction separator or a continued line, analyze deeper
                                    j = Len(CodeLine)
                                    For Indent = k To 2 Step -1
                                        If Mid$(CodeLine, Indent, 1) = Spce Then
                                            Exit For 'loop varying indent
                                        End If
                                    Next Indent
                                    If Indent > 1 Or Left$(LTrim$(CodeLine), 1) = Apostrophe Or IsRem(Left$(LTrim$(CodeLine), 4)) Then 'not a goto label
                                        k = 0
                                        Do
                                            If ContLine Then 'a continued line
                                                ContLine = (Right$(CodeLine, 2) = ContMark) 'continued further
                                                k = j 'exit
                                              Else 'CONTLINE = FALSE/0
                                                Inc k
                                                Select Case Mid$(CodeLine, k, 1)
                                                  Case Quote
                                                    Do 'skip literal
                                                        Inc k
                                                    Loop Until k >= j Or Mid$(CodeLine, k, 1) = Quote
                                                  Case HashChar 'possibly a date time literal
                                                    If IsNumeric(Mid$(CodeLine, k + 1, 1)) Then 'next char is numeric so it's not a currency type suffix
                                                        Do 'skip date/time literal; there may be colons in it
                                                            Inc k
                                                        Loop Until k >= j Or Mid$(CodeLine, k, 1) = HashChar
                                                    End If
                                                  Case Apostrophe 'comment
                                                    k = j 'exit
                                                    ContLine = (Right$(CodeLine, 2) = ContMark) 'Continued
                                                  Case "R"
                                                    If Mid$(CodeLine, k - 1, 5) = " Rem " Then
                                                        k = j 'exit
                                                        ContLine = (Right$(CodeLine, 2) = ContMark) 'Continued
                                                    End If
                                                  Case "I"
                                                    If Mid$(CodeLine, k - 1, 4) = " If " Then
                                                        k = j 'exit
                                                        ContLine = (Right$(CodeLine, 2) = ContMark) 'Continued
                                                    End If
                                                  Case Colon
                                                    If Mid$(CodeLine, k + 1, 1) <> "=" Then 'not a named paramater, must be a separator
                                                        Mid$(CodeLine, k, 1) = vbLf 'replace colon by line feed
                                                        i = 0 'to cause .Replaceline (further down)
                                                        KillSelection
                                                    End If 'named parameter
                                                End Select 'current char
                                            End If 'Single Line If
                                        Loop Until k >= j
                                      Else 'this is a GoTo label 'NOT INDENT...
                                        If k And k <> j Then
                                            CodeLine = Replace$(CodeLine, Colon, Colon & vbLf, , 1)
                                            KillSelection
                                            i = 0 'to cause .Replaceline (further down)
                                        End If
                                        ContLine = False
                                    End If 'is this a GoTo label?
                                End If 'Colon or If found in this line
                            End If 'SepChecked
                            If i <> Len(CodeLine) Then
                                .ReplaceLine LineIndex, Mid$(CodeLine, 2)
                            End If
                            If Len(CodeLine) <> 0 And Len(Trim$(CodeLine)) = 0 Then 'trim empty line
                                .ReplaceLine LineIndex, NullStr
                            End If
                            DoEvents 'to give the broom in fPreparing a chance to sweep
                            Inc LineIndex
                        End If
                    Loop Until LineIndex > .CountOfLines
                    LineIndex = .CountOfLines
                    .ReplaceLine .CountOfLines, Replace$(.Lines(LineIndex, 1), ToDo6, NullStr)
                    .ReplaceLine .CountOfLines, Replace$(.Lines(LineIndex, 1), ToDo8, NullStr)
                    .ReplaceLine .CountOfLines, Replace$(.Lines(LineIndex, 1), ToDo16, NullStr)
                    If XPLookRequested And 4 Then 'bit 3 on - insert api call
                        .InsertLines XPLookLine, vbCrLf & XPLookAPICall & InsertedBy
                        XPLookRequested = XPLookRequested And 3 'bit 3 off - call is inserted
                    End If

                    'prepare for scan  
                    Set VarNames = New Collection
                    Set NodeStack = New cStack
                    ReDim WordStack(0), StrucStack(0)
                    Load fStruc
                    fStruc.Caption = "Code Structure for " & WindowTitle
                    Set RootNode = fStruc.tvwStruc.Nodes.Add(, , "Root", WindowTitle)
                    With RootNode
                        .Tag = "1"
                        .ForeColor = vbRed
                        .Expanded = True
                        .Bold = True
                    End With 'ROOTNODE
                    NodeKey = 0
                    Indent = 0
                    CodeIsDead = AlmostInfinity
                    MaxIndent = 0
                    IndentNext = 0
                    OEIndent = NullStr
                    OELevel = -1
                    DefTypeChars = NullStr
                    ActiveForVars = "="
                    FirstCodeLine = 0
                    BracketCount = 0
                    PendingTSRepl = NullStr
                    NoOptExpl = True
                    Complain = False
                    NonProc = False
                    VaCoFuDcl = False
                    SlIf = False
                    LNsObs = False
                    CallUnnec = False
                    ForVarMod = False
                    OrphCom = False
                    FoundGoTo = False
                    MissingScope = False
                    EmptyNext = False
                    EmptyComplain = True
                    InProcHeader = False
                    Dupl = False
                    Skipping = False
                    LocalDim = False
                    FinalReq = True
                    Skipped = 0
                    Suppressed = 0
                    Inserted = 0
                    NumCommentedLines = 0
                    NumCommentOnlyLines = 2
                    NumEmptyLines = 1
                    PrintLineNumber = 0
                    NoOEClose = False
                    StrucErr = False
                    PossVio = False
                    NoCode = False
                    DeadCode = False
                    Pleo = False
                    If MarkChecked Then
                        SetToDo Mo Or Cc
                      Else 'MARKCHECKED = FALSE/0
                        ClearToDo Mo Or Cc
                    End If

                    'order modules alphabetically  
                    If SortRequested And .Members.Count > 1 Then
                        'order modules  
                        With fPreparing
                            .lbl(0) = "Sorting..."
                            .imgSort.Visible = True
                            .imgBroom.Visible = False
                            .imgFormat.Visible = False
                        End With 'FPREPARING
                        .CodePane.TopLine = 1
                        DoEvents
                        ReDim SortElems(0)
                        'collect module descriptions -> (Name, StartingLine, Length)  
                        For j = 1 To .Members.Count
                            With .Members(j)
                                TmpString1 = .Name
                                i = (.Type = vbext_mt_Property Or .Type = vbext_mt_Method)
                            End With '.MEMBERS(J)
                            If i Then
                                For i = 1 To 4
                                    k = Choose(i, vbext_pk_Get, vbext_pk_Let, vbext_pk_Set, vbext_pk_Proc) 'determines seq of equal named modules
                                    TempElem = Null
                                    On Error Resume Next
                                        TempElem = Array(TmpString1, .ProcStartLine(TmpString1, k), .ProcCountLines(TmpString1, k))
                                    On Error GoTo 0
                                    If Not IsNull(TempElem) Then
                                        ReDim Preserve SortElems(UBound(SortElems) + 1)
                                        SortElems(UBound(SortElems)) = TempElem
                                    End If
                                Next i
                            End If
                        Next j
                        QuickSort 1, UBound(SortElems), 0
                        'build sorted component  
                        TmpString1 = NullStr
                        For i = 1 To UBound(SortElems)
                            Select Case i
                              Case 1
                                If SortElems(i)(1) > .CountOfDeclarationLines Then 'Sub or Function
                                    TmpString1 = TmpString1 & .Lines(SortElems(i)(1), SortElems(i)(2)) & vbCrLf
                                End If
                              Case Else
                                'there's a quirk in VB: it returns Events as Methods and if an  
                                'Event has the same name as a Sub/Function then this results in  
                                'duplicates, so here duplicates are filtered out  
                                If SortElems(i)(1) <> SortElems(i - 1)(1) Then
                                    If SortElems(i)(1) > .CountOfDeclarationLines Then 'Sub or Function
                                        TmpString1 = TmpString1 & .Lines(SortElems(i)(1), SortElems(i)(2)) & vbCrLf
                                    End If
                                End If
                            End Select
                        Next i
                        'delete original modules  
                        .DeleteLines .CountOfDeclarationLines + 1, .CountOfLines - .CountOfDeclarationLines
                        'add sorted modules  
                        .AddFromString TmpString1
                        'remove trailing blank lines if any  
                        Do
                            CodeLine = Trim$(.Lines(.CountOfLines, 1))
                            If Len(CodeLine) = 0 Then
                                .DeleteLines .CountOfLines
                            End If
                        Loop Until Len(CodeLine)
                        KillSelection
                    End If 'Sort requested

                    With fPreparing
                        .imgSort.Visible = False
                        .imgBroom.Visible = False
                        .imgFormat.Visible = True
                        .lbl(0) = "Formatting..."
                    End With 'FPREPARING
                    DoEvents
                    LineIndex = 1

                    'Here we go  
                    Do
                        Do
                            CodeLineWith = NullStr
                            CodeLineWithout = NullStr
                            FromLine = LineIndex
                            Do 'concat continuation lines  
                                FirstWord = Trim$(.Lines(LineIndex, 1))

#If Not LinNumComplain Then '----------------------------------------------------------
                                If FromLine = LineIndex Then
                                    Words = Split(FirstWord, Spce)
                                    If UBound(Words) >= 0 Then
                                        FirstWord = Words(0)
                                        If IsNumeric(FirstWord) Then  'it's a line number
                                            Words(0) = NullStr
                                        End If
                                      Else 'NOT UBOUND(WORDS)...
                                        FirstWord = NullStr
                                    End If
                                    FirstWord = LTrim$(Join$(Words, Spce))
                                End If
#End If '------------------------------------------------------------------------------

                                CodeLineWith = CodeLineWith & FirstWord
                                CodeLineWithout = CodeLineWithout & FirstWord
                                If Right$(CodeLineWithout, 2) = ContMark Then
                                    CodeLineWith = CodeLineWithout & vbCrLf
                                    CodeLineWithout = Left$(CodeLineWithout, Len(CodeLineWithout) - 1)
                                    Inc LineIndex
                                  Else 'NOT RIGHT$(CodeLineWithout,...
                                    Exit Do 'loop 
                                End If
                            Loop
                            'get first word in original case and examine it  
                            Words = Split(CodeLineWithout, Spce)
                            If UBound(Words) >= 0 Then
                                FirstWord = Words(0)
                                If IsNumeric(FirstWord) Then 'it's a line number
                                    FirstWord = NullStr
                                End If
                              Else 'NOT UBOUND(WORDS)...
                                FirstWord = NullStr
                            End If
                            If Left$(FirstWord, 1) = "[" And Right$(FirstWord, 1) = "]" And Not IsNumeric(Mid$(FirstWord, 2, 1)) Then 'no space between brackets
                                For i = 2 To Len(FirstWord) - 1
                                    If Mid$(FirstWord, i, 1) <> "_" Then
                                        CtB = Mid$(FirstWord, i, 1)
                                        If IsCharAlphaNumeric(CtB(0)) = 0 Then
                                            Exit For 'loop varying i
                                        End If
                                    End If
                                Next i
                                If i = Len(FirstWord) Then 'it's a legal VB data name
                                    FirstWord = Mid$(FirstWord, 2, Len(FirstWord) - 2) 'extract whats between the brackets
                                End If
                            End If
                            CodeLine = LCase$(CodeLineWithout)
                            ToLine = LineIndex
                            'FromLine and ToLine define the concatted range  
                            For i = 1 To Len(CodeLine)
                                Select Case Mid$(CodeLine, i, 1)
                                  Case Apostrophe 'comment
                                    If i = 1 Then
                                        Inc NumCommentOnlyLines
#If OrphanedComment Then '--------------------------------------------------------------
                                        If IsOutsideSub And Not Skipping Then
                                            If CodeLine <> EndCopy And Left$(CodeLine, 12) <> LCase$(CopiedFrom) And Left$(CodeLine, 3) <> Smiley Then
                                                MarkLine ToDo(26), OrphCom
                                            End If
                                        End If
#End If '-------------------------------------------------------------------------------

                                      Else 'not comment only line 'NOT I...
                                        CodeLine = RTrim$(Left$(CodeLine, i - 1)) 'trim comment
                                        Inc NumCommentedLines
                                    End If
                                    Exit For 'and out 'loop varying i
                                  Case "r" 'rem ?
                                    If i > 1 Then
                                        If Mid$(CodeLine, i - 1, 1) = Spce And IsRem(Mid$(CodeLine, i, 4)) Then
                                            CodeLine = Left$(CodeLine, i - 1) 'trim comment
                                            Inc NumCommentedLines
                                            Exit For 'and out 'loop varying i
                                        End If
                                      Else 'NOT I...
                                        If IsRem(Mid$(CodeLine, 1, 4)) Then 'comment only
                                            Inc NumCommentOnlyLines

#If OrphanedComment Then '--------------------------------------------------------------
                                            If IsOutsideSub And Not Skipping Then
                                                If Left$(CodeLine, 8) <> "rem copy" Then
                                                    MarkLine ToDo(26), OrphCom
                                                End If
                                            End If
#End If '-------------------------------------------------------------------------------

                                            Exit For 'and out 'loop varying i
                                        End If
                                    End If
                                  Case Quote 'literal - replace Spaces by LowValues
                                    Do
                                        Inc i
                                        Select Case Mid$(CodeLine, i, 1)
                                          Case Quote
                                            Exit Do 'loop 
                                          Case Spce, "(", ")"
                                            Mid$(CodeLine, i, 1) = Chr$(0)
                                        End Select
                                    Loop
                                  Case "[" 'external name - replace Spaces by LowValues
                                    Do
                                        Inc i
                                        Select Case Mid$(CodeLine, i, 1)
                                          Case "]"
                                            Exit Do 'loop 
                                          Case Spce
                                            Mid$(CodeLine, i, 1) = Chr$(0)
                                        End Select
                                    Loop
                                End Select
                            Next i
                            If IfExpChecked Then
                                CodeLine = Trim$(CodeLine)
                              Else 'IFEXPCHECKED = FALSE/0
                                CodeLine = Trim$(Replace$(CodeLine, Colon & Spce, Spce)) 'remove separators
                            End If
                            If CodeLine = NullStr Then
                                Inc LineIndex  'nothing left
                                Inc NumEmptyLines
                                If Skipping Then
                                    Inc Skipped
                                    CodeIsDead = -CodeIsDead
                                End If
                              Else 'codeline is ready to be processed 'NOT CODELINE...
                                Exit Do 'loop 
                            End If
                        Loop
                        Select Case CodeLine 'check for options and conditional compile (#If False)
                          Case CcIfBeg, CcIfElseBeg
                            ClearToDo Cc
                          Case CcIfEnds
                            SetToDo Cc
                          Case MarkOff
                            ClearToDo Mo
                          Case MarkOffSilent
                            ClearToDo Mo Or Si
                          Case MarkOn
                            SetToDo Mo
                          Case SkipOff
                            If Skipping Then
                                Skipping = False
                                PrintSkipComment
                                PrintLineNumber = ToLine - 1
                            End If
                          Case IndentOn
                            RemIndentNext = RemIndent + 1
                            If RemIndentNext > 3 Then
                                RemIndentNext = 3
                            End If
                          Case IndentOff
                            RemIndent = RemIndent - 1
                            If RemIndent < 0 Then
                                RemIndent = 0
                            End If
                            RemIndentNext = RemIndent
                          Case Interface
                            EmptyComplain = False
                        End Select
                        Pleonasm = 0
                        Words = Split(CodeLine, Spce) 'break codeline into words
                        LastWordIndex = UBound(Words)
                        If LastWordIndex < 2 Then
                            ReDim Preserve Words(2)
                            Words(2) = Spce ' fix to enable exit of remove line number loop below
                        End If
                        If IsNumeric(Words(0)) Then 'remove line number (formatting will be wrong though)
                            Do
                                For i = 1 To LastWordIndex
                                    Words(i - 1) = Words(i)
                                Next i
                                LastWordIndex = LastWordIndex - 1
                            Loop Until Words(0) <> NullStr
                            MarkLine ToDo(16), LNsObs
                        End If
                        '------------------------------------------------------
                        If IsRem(Words(0)) Then
                            Words(0) = sRem
                          Else 'ISREM(WORDS(0)) = FALSE/0
                            If Left$(CodeLine, 1) = Apostrophe Then
                                Words(0) = sRem
                                Words(1) = vbNullString 'have to kill that to prevent "' skip on" from becoming "Rem skip on" etc
                            End If
                        End If
                        '------------------------------------------------------
                        If Words(0) <> sRem And FirstCodeLine = 0 Then
                            FirstCodeLine = FromLine
                        End If
                        If Skipping Then 'look for option explicit anyway
                            If Words(0) = "option" And Words(1) = "explicit" Then
                                NoOptExpl = False
                            End If
                            Words(0) = SkipMark 'kill first word while skipping
                        End If
                        If Words(0) = "static" Then
                            If Words(1) <> "sub" And Words(1) <> "function" And Words(1) <> "property" Then
                                Words(0) = "dim" 'this is a static variable declaration
                            End If
                        End If
                        Inc NumCodeLinesInProc

                        'now - what have we got here?  
                        Select Case Words(0) 'split on first word of code line
                          Case SkipMark
                            Skipped = Skipped + 1 + ToLine - FromLine
                          Case sRem
                            Dec NumCodeLinesInProc
                            If InProcHeader Then 'this handles comment only lines after a procedure header
                                IndentNext = Indent
                                Indent = 0
                                HalfIndent = Space$(HalfTabWidth)
                                NewProcStarting = True 'keep it alive
                            End If
                            If Words(1) = Copy Then
                                With .CodePane
                                    i = LineIndex - .CountOfVisibleLines / 2
                                    If i > 0 Then
                                        .TopLine = i
                                      Else 'NOT I...
                                        .TopLine = 1
                                    End If
                                End With '.CODEPANE
                                TmpString1 = LTrim$(.Lines(LineIndex, 1))
                                .ReplaceLine LineIndex, ">> Pasted code goes here..."
                                TmpString2 = NullStr
                                For i = 2 To LastWordIndex
                                    TmpString2 = TmpString2 & Words(i) & Spce
                                Next i
                                With fCopy
                                    .FileName = Trim$(TmpString2)
                                    .TextToPaste = NullStr
                                    .sFontName = IDEFontName
                                    .lFontSize = IDEFontSize
                                    .Show vbModal
                                End With 'FCOPY
                                If Len(fCopy.TextToPaste) Then
                                    .ReplaceLine LineIndex, fCopy.TextToPaste
                                  Else 'LEN(FCOPY.TEXTTOPASTE) = FALSE/0
                                    .ReplaceLine LineIndex, TmpString1
                                End If
                                Unload fCopy
                                Set fCopy = Nothing
                                DoEvents
                            End If
                          Case "option"
                            If Words(1) = "explicit" Then
                                NoOptExpl = False
                            End If
                          Case "declare"
                            Select Case Words(1)
                              Case "function"
                                CheckVardefs 2, False, False, True
                                If TypeSuffixFound Then
                                    MarkLine ToDo(12), VaCoFuDcl
                                  ElseIf Words(LastWordIndex - 1) <> "as" Then 'TYPESUFFIXFOUND = FALSE/0
                                    MarkLine ToDo(9), VaCoFuDcl
                                End If
                                MarkLine ToDo(15), MissingScope
                              Case "sub"
                                MarkLine ToDo(15), MissingScope
                            End Select
                          Case "public", "private", "friend", "global", "static"
                            Select Case Words(0)
                              Case "global"
                                .ReplaceLine LineIndex, Replace$(.Lines(LineIndex, 1), "Global ", "Public ", , 1)
                              Case "static"
                                MarkLine ToDo(15), MissingScope
                            End Select
                            If Indent Then
                                .ReplaceLine LineIndex, .Lines(LineIndex, 1) & ToDo(8)
                            End If
                            Indent = 0
                            Select Case Words(1) 'second word
                              Case "static"
                                Select Case Words(2) 'third word
                                  Case "sub", "function", "property"
                                    IndentNext = 1
                                    NumCodeLinesInProc = 0
                                    LastNext = -1
                                    CurrProcType = UCase$(Left$(Words(2), 1)) & Mid$(Words(2), 2)
                                    If Trim$(.Lines(ToLine + 1, 1)) <> NullStr Then
                                        If EmptyLinesChecked = False Then
                                            .InsertLines ToLine + 1, NullStr
                                            KillSelection
                                        End If
                                    End If
                                    If Trim$(.Lines(FromLine - 1, 1)) <> NullStr Then
                                        InsertBlankLine
                                    End If
                                    NewProcStarting = True
                                    CodeIsDead = AlmostInfinity
                                    LocalDim = False
                                    If Words(2) <> "sub" Then
                                        If CheckVardefs(3 - (Words(2) = "property"), False, False, True) Then
                                            If Words(2) <> "property" Or Words(3) = "get" Then
                                                MarkLine ToDo(9), VaCoFuDcl
                                            End If
                                        End If
                                        If TypeSuffixFound Then
                                            MarkLine ToDo(12), VaCoFuDcl
                                        End If
                                    End If
                                End Select
                              Case "sub", "function", "property", "enum", "type"
                                IndentNext = 1
                                NumCodeLinesInProc = 0
                                If Words(1) = "enum" Then
                                    InEnum = True
                                    EnumMembers = NullStr
                                    NextEnumBreak = EnumLineLen
                                  Else 'NOT WORDS(1)...
                                    InEnum = False
                                End If
                                CurrProcType = UCase$(Left$(Words(1), 1)) & Mid$(Words(1), 2)
                                If Words(1) <> "enum" And Words(1) <> "type" Then
                                    LastNext = -1
                                    If Trim$(.Lines(ToLine + 1, 1)) <> NullStr Then
                                        If EmptyLinesChecked = False Then
                                            .InsertLines ToLine + 1, NullStr
                                            KillSelection
                                        End If
                                    End If
                                    If Trim$(.Lines(FromLine - 1, 1)) <> NullStr And Left$(.Lines(FromLine - 1, 1), 12) <> CopiedFrom Then
                                        InsertBlankLine
                                    End If
                                    NewProcStarting = True
                                    CodeIsDead = AlmostInfinity
                                    LocalDim = False
                                    If Words(1) <> "sub" Then
                                        If CheckVardefs(2 - (Words(1) = "property"), False, False, True) Then
                                            If Words(1) <> "property" Or Words(2) = "get" Then
                                                MarkLine ToDo(9), VaCoFuDcl
                                            End If
                                        End If
                                        If TypeSuffixFound Then
                                            MarkLine ToDo(12), VaCoFuDcl
                                        End If
                                    End If
                                End If
                              Case "declare"
                                If Words(2) = "function" Then
                                    CheckVardefs 3, False, False, True
                                    If TypeSuffixFound Then
                                        MarkLine ToDo(12), VaCoFuDcl
                                      ElseIf Words(LastWordIndex - 1) <> "as" Then 'TYPESUFFIXFOUND = FALSE/0
                                        MarkLine ToDo(9), VaCoFuDcl
                                    End If
                                End If
                              Case "event", "withevents"
                                Rem do nothing  
                              Case Else 'variable/constant declaration
                                IndentNext = 0
                                If CheckVardefs(1 - (Words(1) = "const"), False, Words(1) = "const", False) Then
                                    MarkLine ToDo(9), VaCoFuDcl
                                End If
                                If DuplNameDetected Then
                                    MarkLine ToDo(10), Dupl
                                End If
                                If DefConstType Then
                                    MarkLine ToDo(11), VaCoFuDcl
                                End If
                                If TypeSuffixFound Then
                                    MarkLine ToDo(12), VaCoFuDcl
                                End If
                            End Select
                          Case "dim", "const" 'static was alredy replaced by dim
                            If CheckVardefs(1 - (Words(1) = "withevents"), Indent <> 0, Words(0) = "const", False) Then
                                MarkLine ToDo(9), VaCoFuDcl
                            End If
                            If DuplNameDetected Then
                                MarkLine ToDo(10), Dupl
                            End If
                            If DefConstType Then
                                MarkLine ToDo(11), VaCoFuDcl
                            End If
                            If TypeSuffixFound Then
                                MarkLine ToDo(12), VaCoFuDcl
                            End If
                            If Indent Then 'local 'Dim' within procedure
                                Dec NumCodeLinesInProc
                                If LocalDim = False And InProcHeader And Trim$(.Lines(FromLine - 1, 1)) <> NullStr Then
                                    InsertBlankLine
                                End If
                                IndentNext = Indent
#If StdLocalDimIndent Then '--------------------------------------------------------
                                Indent = 0
                                HalfIndent = Space$(HalfTabWidth)
#End If '---------------------------------------------------------------------------
                                If InProcHeader Then
                                    NewProcStarting = True 'keep it alive
#If StdLocalDimIndent Then '--------------------------------------------------------
                                  Else 'INPROCHEADER = FALSE/0
                                    MarkLine ToDo(2) & IIf(Len(ToDo(2)), CurrProcType, NullStr), NonProc
#End If '---------------------------------------------------------------------------
                                End If
                                LocalDim = True
                              Else 'INDENT = FALSE/0
                                MarkLine ToDo(15), MissingScope
                            End If
                          Case "event"
                            MarkLine ToDo(15), MissingScope
#If CallComplain Then '-------------------------------------------------------------
                          Case "call" 'call unnecessary
                            If Right$(Words(LastWordIndex), 1) = ")" Then
                                If CallChecked Then
                                    For i = 1 To Len(CodeLineWith)
                                        If Mid$(CodeLineWith, i, 1) = Apostrophe Then
                                            Exit For 'loop varying i
                                        End If
                                    Next i
                                    i = i - 1
                                    j = 0
                                    For i = i To 1 Step -1
                                        Select Case Mid$(CodeLineWith, i, 1)
                                          Case ")"
                                            If j = 0 Then
                                                Mid$(CodeLineWith, i, 1) = " "
                                            End If
                                            Inc j
                                          Case "("
                                            Dec j
                                            If j = 0 Then
                                                Exit For 'loop varying i
                                            End If
                                          Case Quote
                                            Do
                                                Dec i
                                            Loop Until Mid$(CodeLineWith, i, 1) = Quote
                                        End Select
                                    Next i
                                    Mid$(CodeLineWith, i, 1) = " "
                                    .DeleteLines FromLine, ToLine - FromLine + 1
                                    .InsertLines FromLine, Replace$(CodeLineWith, "Call ", "")
                                  Else 'CALLCHECKED = FALSE/0
                                    MarkLine ToDo(21) & ToDo(22), CallUnnec
                                End If
                              Else 'NOT RIGHT$(WORDS(LASTWORDINDEX),...
                                If CallChecked Then
                                    .DeleteLines FromLine, ToLine - FromLine + 1
                                    .InsertLines FromLine, Replace$(CodeLineWith, "Call ", "")
                                  Else 'CALLCHECKED = FALSE/0
                                    MarkLine ToDo(21), CallUnnec
                                End If
                            End If
#End If '---------------------------------------------------------------------------
                          Case "debug.print"
                            If MarkChecked Then
                                SuspendIndent
                                .ReplaceLine LineIndex, .Lines(LineIndex, 1) & ToDo(27)
                            End If
                          Case "exit"
                            SubsequentCodeIsDead
                            Select Case Words(1) 'second word
                              Case "sub", "function", "property"
                                NumCodeLinesInProc = 0
                                LastNext = -1
                                If Indent = 1 Then
                                    Indent = 0
                                    IndentNext = 1
                                    If EmptyLinesChecked = False Then
                                        If Trim$(.Lines(ToLine + 1, 1)) <> NullStr Then
                                            .InsertLines ToLine + 1, NullStr
                                            KillSelection
                                            LineIndex = ToLine + 1
                                        End If
                                    End If
                                    If Trim$(.Lines(FromLine - 1, 1)) <> NullStr Then
                                        InsertBlankLine
                                    End If
                                  Else 'NOT INDENT...
                                    If MarkChecked Then
                                        .ReplaceLine LineIndex, .Lines(LineIndex, 1) & ExitSFP
                                    End If
                                End If
                              Case "for", "do" 'Exit For/Do
                                If GetStruc(False) <> Words(1) Then 'the current structure bracket type is not For/Do
                                    MarkLine ToDo(17), PossVio
                                End If
                                If MarkChecked Then
                                    .ReplaceLine LineIndex, .Lines(LineIndex, 1) & IIf(Words(1) = "do", ExitDo, BuildExitFor)
                                End If
                            End Select
                          Case "sub", "function", "property", "enum", "type"
                            Indent = 0
                            CodeIsDead = AlmostInfinity
                            IndentNext = 1
                            NumCodeLinesInProc = 0
                            If Words(0) = "enum" Then
                                InEnum = True
                                EnumMembers = NullStr
                                NextEnumBreak = EnumLineLen
                              Else 'NOT WORDS(0)...
                                InEnum = False
                            End If
                            CurrProcType = UCase$(Left$(Words(0), 1)) & Mid$(Words(0), 2)
                            If Words(0) <> "enum" And Words(0) <> "type" Then
                                LastNext = -1
                                If EmptyLinesChecked = False Then
                                    If Trim$(.Lines(ToLine + 1, 1)) <> NullStr Then
                                        .InsertLines ToLine + 1, NullStr
                                        KillSelection
                                    End If
                                End If
                                If Trim$(.Lines(FromLine - 1, 1)) <> NullStr And Left$(.Lines(FromLine - 1, 1), 12) <> CopiedFrom Then
                                    InsertBlankLine
                                End If
                                NewProcStarting = True
                                LocalDim = False
                                If Words(0) <> "sub" Then
                                    If CheckVardefs(1 - (Words(0) = "property"), False, False, True) Then
                                        If Words(0) <> "property" Or Words(1) = "get" Then
                                            MarkLine ToDo(9), VaCoFuDcl
                                        End If
                                    End If
                                    If TypeSuffixFound Then
                                        MarkLine ToDo(12), VaCoFuDcl
                                    End If
                                End If
                            End If
                            MarkLine ToDo(15), MissingScope
                          Case "if"
                            If Words(LastWordIndex) = "then" Then 'structured If
                                IndentNext = Indent + 1
                                Pleonasm = InStr(CodeLine, "= true ") + InStr(CodeLine, " true =") + InStr(CodeLine, "= true)") + InStr(CodeLine, "(true =")
                                PushWord BlnkApos & IIf(Words(2) = "then", NullStr, "NOT ") & Words(1) & IIf(Words(2) = "then", " = FALSE/0", "...")
                              Else 'single line If 'NOT WORDS(LASTWORDINDEX)...
                                If IfExpSilent = False Then
                                    MarkLine ToDo(3), SlIf
                                End If
                                Select Case Words(LastWordIndex)
                                  Case "sub", "function", "property", "for", "do"
                                    .ReplaceLine LineIndex, .Lines(LineIndex, 1) & ToDo(4)
                                End Select
                                CheckGoto NullStr
                                If IfExpChecked Then
                                    i = 10
                                    k = 0
                                    Do
                                        j = InStr(i, CodeLineWithout, sElseB)
                                        If j = 0 Then
                                            j = InStr(i, CodeLineWithout, sElseC)
                                        End If
                                        If j Then
                                            Inc k
                                            i = j + 7
                                        End If
                                    Loop While j
                                    If k < 2 Then 'max one else can be expanded
                                        i = 10
                                        TmpString2 = vbCrLf & sEndIf
                                        Do
                                            j = InStr(i, CodeLine, " if ")
                                            If j Then
                                                TmpString2 = TmpString2 & vbCrLf & sEndIf
                                                i = j + 10
                                            End If
                                        Loop While j
                                        If IfExpSilent Then 'silent if expansion
                                            .DeleteLines FromLine, ToLine - FromLine + 1
                                            j = FromLine
                                            IndentNext = Indent + 1
                                            PushWord BlnkApos & IIf(Words(2) = "then", NullStr, "NOT ") & Words(1) & IIf(Words(2) = "then", " = FALSE/0", "...")
                                          Else 'IFEXPSILENT = FALSE/0
                                            .ReplaceLine FromLine, Apostrophe & LTrim$(.Lines(FromLine, 1))
                                            .ReplaceLine ToLine, .Lines(ToLine, 1) & ReplacedBy
                                            SetComplain
                                            j = ToLine + 1
                                        End If
                                        .InsertLines j, Replace$(Replace$(Replace$(Replace$(Replace$(Replace$(CodeLineWithout, ReplacedBy, NullStr), sElseB, vbCrLf & sElseB & vbCrLf), sElseC, vbCrLf & sElseB & vbCrLf), sThen, sThen & vbCrLf), sThenD, sThen & vbCrLf), Chr$(0), Spce) & TmpString2
                                    End If
                                End If
                            End If
                            If (Words(1) = "false" And (Words(2) = "then" Or Words(2) = "and")) Or InStr(CodeLine, " and false ") Then
                                SubsequentCodeIsDead 1
                            End If
                          Case "#if", "#else", "#elseif", "#end", "#const"
                            SuspendIndent
                          Case "do", "for", "while", "select", "with"
                            IndentNext = Indent + 1
                            If Words(0) <> "select" Then 'push words(0) on structure stack
                                StrucStack(UBound(StrucStack)) = Words(0)
                                ReDim Preserve StrucStack(UBound(StrucStack) + 1)
                            End If
                            Select Case Words(0)
                              Case "do"
                                If InStr(CodeLine & Spce, " loop ") Then
                                    IndentNext = Indent
                                    MarkLine ToDo(3), StrucErr
                                End If
                              Case "for"
                                If InStr(CodeLine & Spce, " next ") Then
                                    IndentNext = Indent
                                    MarkLine ToDo(3), StrucErr
                                  Else 'NOT INSTR(ACTIVEFORVARS,... 'NOT INSTR(CODELINE...
                                    If Words(1) = "each" Then
                                        If InStr(ActiveForVars, "=" & Words(2) & "=") Then
                                            MarkLine ToDo(14), ForVarMod
                                        End If
                                        PushWord Words(2)
                                        ActiveForVars = ActiveForVars & Words(2) & "="
                                      Else 'NOT WORDS(1)...
                                        If HasTypeSuffix(Words(1)) Then
                                            Words(1) = Left$(Words(1), Len(Words(1)) - 1)
                                        End If
                                        If InStr(ActiveForVars, "=" & Words(1) & Words(2)) Then
                                            MarkLine ToDo(14), ForVarMod
                                        End If
                                        PushWord Words(1)
                                        ActiveForVars = ActiveForVars & Words(1) & "="
                                    End If
                                End If
                              Case "while"
                                If InStr(CodeLine & Spce, " wend ") Then
                                    IndentNext = Indent
                                    MarkLine ToDo(3), StrucErr
                                End If
                                MarkLine ToDo(19), Dummy
                              Case "with"
                                If InStr(CodeLine & Spce, " end with ") Then
                                    IndentNext = Indent
                                    MarkLine ToDo(3), StrucErr
                                  Else 'NOT INSTR(CODELINE...
                                    PushWord Words(1)
                                End If
                                If Words(1) = "me" Then
                                    MarkLine ToDo(7), Pleo
                                End If
                            End Select
                            Pleonasm = InStr(CodeLine & Spce, "= true ") + InStr(CodeLine, " true =") + InStr(CodeLine, "= true)") + InStr(CodeLine, "(true =")
                          Case "on"
                            For i = 1 To 2
                                If Words(i) = "error" Then 'On Error/On Local Error
                                    Select Case Words(i + 1)
                                      Case "resume"
                                        OEIndentNext = Space$(FullTabWidth)
                                        OELevel = Indent
                                      Case "goto"
                                        If Words(i + 2) = "0" Then
                                            If OELevel = Indent Then
                                                OEIndentNext = NullStr
                                                OEIndent = NullStr
                                                OELevel = -1
                                              Else 'NOT OELEVEL...
                                                MarkLine ToDo(23), FoundGoTo
                                            End If
#If MarkOnErrorGoTo Then '--------------------------------------------------------------
                                          Else 'NOT WORDS(I...
                                            CheckGoto NullStr
#End If '-------------------------------------------------------------------------------
                                        End If
                                    End Select
                                End If
                            Next i
                          Case "case", "else", "elseif"
                            IndentNext = Indent
                            Dec Indent
                            If Indent < 0 Then
                                Indent = 0
                                MarkLine ToDo(8), StrucErr
                            End If
                            Pleonasm = InStr(CodeLine & Spce, "= true ") + InStr(CodeLine, " true =") + InStr(CodeLine, "= true)") + InStr(CodeLine, "(true =")
                            HalfIndent = Space$(HalfTabWidth)
                            Select Case Words(0)
                              Case "else"
                                PopPush CodeLine = "else", NullStr, NullStr
                              Case "elseif"
                                PopPush Words(LastWordIndex) = "then", Words(1), Words(2)
                            End Select
                            CheckGoto NullStr
                          Case "end"
                            If CodeLine = "end" Then 'it's the 'End' statement
                                Words(0) = "END"
                                MarkLine ToDo(28), FoundGoTo
                                SubsequentCodeIsDead
                              Else 'NOT CODELINE...
                                Dec Indent
                                InEnum = False
                                If Indent < 0 Then
                                    Indent = 0
                                    MarkLine ToDo(8), StrucErr
                                End If
                                If Len(EnumMembers) Then
                                    If LCase$(Left$(.Lines(LineIndex + 1, 1), Len(CcIfBeg))) <> CcIfBeg Then
                                        .InsertLines LineIndex + 1, _
                                                     CcIfBeg & InsertedBy & vbCrLf & _
                                                     "Private " & Left$(EnumMembers, Len(EnumMembers) - 2) & InsertedBy & vbCrLf & _
                                                     CcIfEnds & InsertedBy
                                        SetComplain 1
                                    End If
                                    EnumMembers = NullStr 'reset
                                    NextEnumBreak = EnumLineLen 'reset
                                End If

                                IndentNext = Indent
                                Select Case Words(1)
                                  Case "sub", "function", "property"
                                    If Indent Then
                                        MarkLine ToDo(8), StrucErr
                                      ElseIf Len(OEIndent) Then 'INDENT = FALSE/0
                                        MarkLine ToDo(6), NoOEClose
                                    End If
                                    If NumCodeLinesInProc = 1 And EmptyComplain Then
                                        MarkLine ToDo(24), NoCode
                                    End If
                                    If EmptyLinesChecked = False Then
                                        If Trim$(.Lines(ToLine + 1, 1)) <> NullStr And _
                                                 Left$(LTrim$(.Lines(ToLine + 1, 1)), 2) <> "#E" And _
                                                 Left$(LTrim$(.Lines(ToLine + 1, 1)), 1) <> "'" And _
                                                 Not IsRem(Left$(LTrim$(.Lines(ToLine + 1, 1)), 4)) Then
                                            'the next line is not blank and not #EndIf nor #ElseIf nor End Copy 
                                            .InsertLines ToLine + 1, NullStr
                                            KillSelection
                                        End If
                                    End If
                                    If Trim$(.Lines(FromLine - 1, 1)) <> NullStr Then
                                        InsertBlankLine
                                    End If
                                    IndentNext = 0
                                    Indent = 0
                                    OEIndentNext = NullStr
                                    OEIndent = NullStr
                                    OELevel = -1
                                    ReDim WordStack(0), StrucStack(0)
                                  Case "with"
                                    GetStruc True
                                    TmpString1 = BlnkApos & PopWord()
                                    If MarkChecked Then
                                        If InStr(1, .Lines(LineIndex, 1), TmpString1, vbTextCompare) = 0 Then
                                            .ReplaceLine LineIndex, .Lines(LineIndex, 1) & TmpString1
                                            Inc NumCommentedLines
                                        End If
                                      Else 'MARKCHECKED = FALSE/0
                                        .ReplaceLine LineIndex, Replace$(.Lines(LineIndex, 1), TmpString1, NullStr)
                                    End If
                                  Case "if"
                                    PopWord
                                End Select
                            End If
                          Case "loop", "wend"
                            Dec Indent
                            GetStruc True
                            If Indent < 0 Then
                                Indent = 0
                                MarkLine ToDo(8), StrucErr
                            End If
                            IndentNext = Indent
                            Pleonasm = InStr(CodeLine & Spce, "= true ") + InStr(CodeLine, " true =") + InStr(CodeLine, "= true)") + InStr(CodeLine, "(true =")
                            If Words(0) = "wend" Then
                                MarkLine ToDo(19), Dummy
                            End If
                          Case "next"
                            Dec Indent
                            GetStruc True
                            i = 4
                            Do
                                i = InStr(i + 1, CodeLine, Comma)
                                If i Then
                                    Dec Indent
                                    GetStruc True
                                    ActiveForVars = Replace$(ActiveForVars, "=" & PopWord & "=", "=", , , vbTextCompare)
                                  Else 'I = FALSE/0
                                    Exit Do 'loop 
                                End If
                            Loop
                            If Indent < 0 Then
                                Indent = 0
                                MarkLine ToDo(8), StrucErr
                            End If
                            IndentNext = Indent
                            TmpString1 = PopWord
                            ActiveForVars = Replace$(ActiveForVars, "=" & TmpString1 & "=", "=", , , vbTextCompare)
                            If CodeLine = "next" And MarkChecked Then
                                MarkLine ToDo(5) & Spce & TmpString1, EmptyNext
                            End If
                            If LastNext = NumCodeLinesInProc - 1 Then
                                MarkLine ToDo(29), EmptyNext
                            End If
                            LastNext = NumCodeLinesInProc
                          Case "goto"
                            SubsequentCodeIsDead
                            CheckGoto Spce
                          Case "return" 'this is here thanks to Sebastian
                            SubsequentCodeIsDead
                          Case "let"
                            If HasTypeSuffix(Words(1)) Then
                                Words(1) = Left$(Words(1), Len(Words(1)) - 1)
                            End If
                            If InStr(ActiveForVars, "=" & Words(1) & Words(2)) Then
                                If LastWordIndex = 3 Then
                                    If Words(3) <> Words(1) Then
                                        MarkLine ToDo(14), ForVarMod
                                    End If
                                  Else 'NOT LASTWORDINDEX...
                                    MarkLine ToDo(14), ForVarMod
                                End If
                            End If
                          Case "deflng", "defbool", "defbyte", "defint", "defcur", _
                               "defsng", "defdbl", "defdec", "defdate", "defstr", _
                               "defobj"  'defvar is missing on purpose
                            EmitChars
                          Case Else
                            If InEnum Then 'inside enumeration
                                CheckVardefs 0, False, False, False
                                If DuplNameDetected Then
                                    MarkLine ToDo(10), Dupl
                                End If
                                If EnumChecked Then
                                    If Left$(Words(0), 1) = "[" Then 'case preserve only for non-bracketed words
                                        MarkLine ToDo(25), VaCoFuDcl
                                      Else 'NOT LEFT$(WORDS(0),...
                                        If Len(EnumMembers) >= NextEnumBreak Then
                                            EnumMembers = EnumMembers & " _" & vbCrLf 'insert line break
                                            Inc NextEnumBreak, EnumLineLen 'and wait for next overflow
                                        End If
                                        EnumMembers = EnumMembers & FirstWord & ", "
                                    End If
                                End If
                              Else 'INENUM = FALSE/0
                                If HasTypeSuffix(Words(0)) Then
                                    Words(0) = Left$(Words(0), Len(Words(0)) - 1)
                                End If
                                If InStr(ActiveForVars, "=" & Words(0) & Words(1)) Then
                                    If LastWordIndex = 2 Then
                                        If Words(2) <> Words(0) Then
                                            MarkLine ToDo(14), ForVarMod
                                        End If
                                      Else 'NOT LASTWORDINDEX...
                                        MarkLine ToDo(14), ForVarMod
                                    End If
                                End If
                                If Right$(Words(0), 1) = ":" Then 'this is a GoTo label
                                    CodeIsDead = AlmostInfinity
                                End If
                            End If
                        End Select 'first word
                        If Indent < CodeIsDead Then 'out of dead structure level
                            CodeIsDead = AlmostInfinity
                          Else 'NOT INDENT...
                            If CodeIsDead >= 0 And Not IsRem(Words(0)) Then 'this line is dead code
                                MarkLine ToDo(18), DeadCode
                            End If
                        End If
                        If Not Skipping Then
                            If InProcHeader And Not NewProcStarting Then
                                If Trim$(.Lines(FromLine - 1, 1)) <> NullStr Then
                                    InsertBlankLine
                                End If
                            End If
                            If Pleonasm Then
                                MarkLine ToDo(7), Pleo
                            End If
                            IndentColor = Indent Mod 9
                            If IndentColor > 6 Then 'skip light gray
                                Inc IndentColor
                            End If
                            IndentColor = QBColor(IndentColor)
                            If StrucRequested And Not IsRem(Words(0)) And Words(0) <> "dim" And Words(0) <> "const" And Words(0) <> "static" And Left$(Words(0), 1) <> HashChar Then
                                With NodeStack
                                    Select Case Indent
                                      Case 0
                                        Set CurrParentNode = RootNode
                                        .Reset 5
                                      Case Is > .StackSize
                                        .Push CurrParentNode
                                        Set CurrParentNode = CurrChildNode
                                      Case Is < .StackSize
                                        Do Until .StackSize = Indent
                                            Set CurrParentNode = .Pop
                                        Loop
                                        If CurrParentNode Is Nothing Then
                                            Set CurrParentNode = RootNode
                                        End If
                                    End Select
                                End With 'NODESTACK
                                'add child node under current parent  
                                TmpString1 = Trim$(.Lines(FromLine, 1))
                                If TmpString1 = NullStr Then 'vert formatted - next line
                                    TmpString1 = Trim$(.Lines(FromLine + 1, 1))
                                End If
                                i = Len(TmpString1)
                                For j = 1 To i
                                    If Mid$(TmpString1, j, 1) = Quote Then
                                        Do
                                            Inc j 'skip over quote
                                            Select Case Mid$(TmpString1, j, 1)
                                              Case Spce
                                                Mid$(TmpString1, j, 1) = MyErrMark 'to prevent literal compression
                                              Case Quote
                                                Exit Do 'loop 
                                            End Select
                                        Loop Until j >= i
                                    End If
                                Next j
                                Do 'shift left and compress comment if any is present
                                    TmpString1 = Replace$(TmpString1, "  ", Spce)
                                    j = Len(TmpString1)
                                    If i = j Then 'no (more) compression
                                        Exit Do 'loop 
                                      Else 'NOT I...
                                        i = j 'try again
                                    End If
                                Loop
                                If i > 72 Then
                                    TmpString1 = Left$(TmpString1, 70) & ElipsisChar
                                End If
                                If Right$(TmpString1, 2) = ContMark Then 'continued line
                                    TmpString1 = Left$(TmpString1, i - 2) & ElipsisChar
                                End If
                                Inc NodeKey
                                Set CurrChildNode = fStruc.tvwStruc.Nodes.Add(CurrParentNode.Key, tvwChild, "C" & Format$(NodeKey), TmpString1)
                                CurrChildNode.Tag = FromLine 'stolen from Roger
                                CurrChildNode.Bold = ((Words(0) = "#const") Or (FromLine <= .CountOfDeclarationLines And (Words(0) = "type" Or Words(0) = "enum" Or Words(1) = "type" Or Words(1) = "enum")) Or (FromLine > .CountOfDeclarationLines And (Indent = 0 Or Right$(TmpString1, 1) = Colon Or InStr(TmpString1, ": '") Or InStr(TmpString1, ": Rem"))))
                                CurrChildNode.ForeColor = IndentColor
                                If InStr(.Lines(FromLine, 1), MySignature) Then
                                    'something is wrong: mark this node..  
                                    CurrChildNode.BackColor = &HC0C0FF
                                    '..and make it visible  
                                    CurrChildNode.EnsureVisible
                                End If
                            End If
                            For i = FromLine To ToLine
                                If NewProcStarting Or i <= .CountOfDeclarationLines Then
                                    TmpString1 = .Lines(i, 1)
                                    Select Case ReplaceTypeSuffix(TmpString1)
                                      Case 1
                                        .ReplaceLine i, TmpString1
                                      Case 2
                                        .ReplaceLine i, TmpString1 & ToDo(13)
                                        SetComplain
                                    End Select
                                End If
                                VBInstance.ActiveCodePane.SetSelection i, 1, i, 1
                                FirstWord = Trim$(.Lines(i, 1))
                                If i = FromLine Then
                                    Words = Split(FirstWord, Spce)
                                    If UBound(Words) >= 0 Then
                                        If IsNumeric(Words(0)) Then
                                            TmpString1 = Words(0)
                                            If Len(TmpString1) < FullTabWidth Or Indent = 0 Then
                                                TmpString1 = TmpString1 & vbTab
                                            End If
                                            Words(0) = NullStr
                                            If LinNumChecked Then
                                                TmpString1 = NullStr
                                            End If
                                            FirstWord = LTrim$(Join$(Words, Spce))
                                          Else 'ISNUMERIC(WORDS(0)) = FALSE/0
                                            TmpString1 = NullStr
                                        End If
                                    End If
                                  Else 'NOT I...
                                    TmpString1 = NullStr
                                End If
                                If Len(FirstWord) < 800 Then
                                    .ReplaceLine i, TmpString1 & String$(Indent + RemIndent, vbTab) & HalfIndent & OEIndent & FirstWord
                                End If
                                If Indent > MaxIndent Then
                                    MaxIndent = Indent
                                End If
                                If Len(TmpString1) Then 'this has a line numebr
                                    Indent = 0 ' so we rely on halfindent only
                                End If

                                'printing
                                If PrintLineLen Then 'print requested
                                    Inc PrintLineNumber
                                    PageSetup i
                                    With Printer
                                        Do Until PrintLineNumber = i
                                            Inc PrintLineNumber
                                            .CurrentY = .CurrentY + PrintLineHeight / 2
                                            PageSetup i
                                        Loop
                                    End With 'PRINTER
                                    k = Indent * FullTabWidth + Len(HalfIndent) + Len(OEIndent)
                                    LNum = Right$(Space$(LnLen) & Format$(PrintLineNumber), LnLen)
                                    If k < PrintLineLen - 4 * FullTabWidth - Len(PntHyphen) Then
                                        TmpString1 = Trim$(.Lines(i, 1))
                                        If IsNumeric(Left$(TmpString1, 1)) Then 'probably a line number
                                            k = k And (i <> FromLine)
                                        End If
                                        With Printer
                                            If i = FromLine Then
                                                .FontItalic = False
                                                NowPrintingComment = False
                                            End If
                                            j = 1
                                            Do 'find comment if any
                                                Select Case Mid$(TmpString1, j, 1)
                                                  Case Quote 'skip literal
                                                    Do
                                                        Inc j
                                                    Loop Until Mid$(TmpString1, j, 1) = Quote Or j = Len(TmpString1)
                                                  Case Apostrophe
                                                    Exit Do 'loop 
                                                  Case "R"
                                                    If Mid$(TmpString1, j + 1, 3) = "em " Then
                                                        Select Case True
                                                          Case j = 1
                                                            Exit Do 'loop 
                                                          Case Mid$(TmpString1, j - 1, 1) = Spce
                                                            Exit Do 'loop 
                                                        End Select
                                                    End If
                                                End Select
                                                Inc j
                                            Loop Until j > Len(TmpString1)
                                            'separate Code from Comment 
                                            TmpString2 = Mid$(TmpString1, j)
                                            TmpString1 = Left$(TmpString1, j - 1)
                                            Do While Len(TmpString1) + Len(TmpString2)
                                                If NewProcStarting And Not InProcHeader And .CurrentY > PB.Bottom - 4 * PrintLineHeight Then
                                                    'almost at page bottom and new sub starting, triggers new page  
                                                    .CurrentY = PB.Bottom
                                                End If
                                                PageSetup i
                                                .ForeColor = IIf(NowPrintingComment, CommentPrintColor, IndentColor And ColorRequested)
                                                .CurrentX = PB.Left
                                                .FontBold = False
                                                .FontItalic = False
                                                Printer.Print LNum; Spce; Space$(k); 'line number and indent
                                                .FontItalic = NowPrintingComment And PrinterItalEnabled
                                                If UBound(Words) = 0 Then
                                                    ReDim Preserve Words(0 To 1)
                                                End If
                                                .FontBold = PrinterBoldEnabled And (k = 0 And NowPrintingComment = False And (FromLine > VBInstance.ActiveCodePane.CodeModule.CountOfDeclarationLines Or Words(0) = "#const" Or Words(0) = "type" Or Words(0) = "enum" Or Words(1) = "type" Or Words(1) = "enum"))
                                                j = Len(TmpString2)
                                                If Len(TmpString1) <= PrintLineLen - k + (j <> 0) Then 'line fits
                                                    Printer.Print TmpString1; 'print it
                                                    If j Then 'comment follows
                                                        .FontItalic = PrinterItalEnabled
                                                        NowPrintingComment = True
                                                        .FontBold = False
                                                        .ForeColor = CommentPrintColor
                                                        j = Len(TmpString1)
                                                        If Len(TmpString2) > PrintLineLen - k - j Then
                                                            j = j + Len(PntHyphen)
                                                            Printer.Print Left$(TmpString2, PrintLineLen - k - j); PntHyphen  'print part of comment and cont-hyphen
                                                            TmpString1 = Mid$(TmpString2, PrintLineLen - k - j + 1)  'put rest of comment into TmpString1 to continue normally
                                                          Else 'NOT LEN(TMPSTRING2)...
                                                            Printer.Print Left$(TmpString2, PrintLineLen - k - j) 'print comment in same line
                                                            TmpString1 = NullStr
                                                        End If
                                                        TmpString2 = NullStr
                                                      Else 'no comment to follow 'J = FALSE/0
                                                        Printer.Print
                                                        TmpString1 = NullStr
                                                    End If
                                                  Else 'line does not fit 'NOT LEN(TMPSTRING1)...
                                                    j = PrintLineLen - k - Len(PntHyphen)
                                                    Printer.Print Left$(TmpString1, j) & PntHyphen 'print that part which fits plus PntHyphen
                                                    TmpString1 = Mid$(TmpString1, j + 1) 'shift text appropriately
                                                End If
                                                LNum = Space$(LnLen) 'kill line number for continuation line
                                            Loop
                                        End With 'PRINTER
                                      Else 'too much indentation 'NOT K...
                                        PageSetup i
                                        With Printer
                                            .CurrentX = PB.Left
                                            .FontBold = False
                                            .ForeColor = vbRed And ColorRequested
                                        End With 'PRINTER
                                        Printer.Print LNum; Left$(" Cannot accomodate code line within page bounds", PrintLineLen)
                                    End If
                                    Printer.CurrentY = Printer.CurrentY + PrintLineHeight / 2.75 'extra advance after printing
                                End If 'printing requested

                                'indentation for continuation lines  
                                If i = FromLine Then 'that's the first, possibly continued line
                                    TmpString1 = LCase$(Trim$(.Lines(i, 1)))
                                    j = InStr(TmpString1, " = ") + 2    'try equal sign first
                                    If j = 2 Then                       'no equal sign
                                        j = InStr(TmpString1, "(")      'so try open bracket
                                        If j = 0 Then                   'and if there is neither
                                            j = InStr(TmpString1, Spce)  'finally try space
                                        End If
                                      Else ' there is an equal sign 'NOT J...
                                        If Left$(TmpString1, 3) = "if " Then 'however its the equal sign in a condition
                                            j = 3 - (Mid$(TmpString1, 4, 1) = "(") - (Mid$(TmpString1, 5, 1) = "(")
                                          ElseIf Left$(TmpString1, 6) = "while " Then 'NOT LEFT$(TMPSTRING1,...
                                            j = 6 - (Mid$(TmpString1, 7, 1) = "(") - (Mid$(TmpString1, 8, 1) = "(")
                                          ElseIf Left$(TmpString1, 7) = "elseif " Then 'NOT LEFT$(TMPSTRING1,...
                                            j = 7 - (Mid$(TmpString1, 8, 1) = "(") - (Mid$(TmpString1, 9, 1) = "(")
                                          ElseIf Left$(TmpString1, 9) = "do while " Then 'NOT LEFT$(TMPSTRING1,...
                                            j = 9 - (Mid$(TmpString1, 10, 1) = "(") - (Mid$(TmpString1, 11, 1) = "(")
                                          ElseIf Left$(TmpString1, 9) = "do until " Then 'NOT LEFT$(TMPSTRING1,...
                                            j = 9 - (Mid$(TmpString1, 10, 1) = "(") - (Mid$(TmpString1, 11, 1) = "(")
                                          ElseIf Left$(TmpString1, 11) = "loop while " Then 'NOT LEFT$(TMPSTRING1,...
                                            j = 11 - (Mid$(TmpString1, 12, 1) = "(") - (Mid$(TmpString1, 13, 1) = "(")
                                          ElseIf Left$(TmpString1, 11) = "loop until " Then 'NOT LEFT$(TMPSTRING1,...
                                            j = 11 - (Mid$(TmpString1, 12, 1) = "(") - (Mid$(TmpString1, 13, 1) = "(")
                                        End If
                                    End If
                                    HalfIndent = HalfIndent & Space$(j)
                                End If
                            Next i
                            Indent = IndentNext
                            RemIndent = RemIndentNext
                            OEIndent = OEIndentNext
                            HalfIndent = NullStr
                            InProcHeader = NewProcStarting
                            NewProcStarting = False
                        End If 'Not Skipping
                        CodeIsDead = Abs(CodeIsDead) 'make positive so subsequent lines are marked
                        If CodeLine = SkipOn Then
                            Skipping = True 'starting from next line
                            CodeIsDead = -CodeIsDead 'make negative - we're skipping and not marking
                        End If
                        'update progress bar  
                        fProgress.Percent = LineIndex * 100 / .CountOfLines
                        DoEvents
                        Inc LineIndex
                    Loop Until LineIndex > .CountOfLines

                    If Skipping Then
                        PrintSkipComment
                    End If

                    If StartUpCompoName = ModuleName Then
                        If XPLookRequested And 1 Then 'bit 1 is still on: prototype decl not yet found
                            .InsertLines .CountOfDeclarationLines + 1, XPLookAPIProto & InsertedBy
                        End If
                        If XPLookRequested And 2 Then 'bit 2 still on: no Proc found to insert
                            .InsertLines .CountOfDeclarationLines + 1, vbCrLf & _
                                         "Private " & StartUpProcName & InsertedBy & vb2CrLf & _
                                         vbTab & XPLookAPICall & InsertedBy & vb2CrLf & _
                                         "End Sub" & InsertedBy
                        End If
                        XPDone = True
                    End If
                    NumDeclLines = .CountOfDeclarationLines
                    NumCodeLines = .CountOfLines - NumDeclLines
                    If NumCodeLines = 0 Then
                        Inc NumDeclLines, 3
                      Else 'NOT NUMCODELINES...
                        Inc NumCodeLines, 3
                    End If
                    i = NumDeclLines + NumCodeLines
                    .InsertLines .CountOfLines + 1, vbCrLf & Smiley & AppDetails & " (" & Format$(Now, "YYYY\-MMM\-DD HH\:MM") & ")  Decl: " & NumDeclLines & "  Code: " & NumCodeLines & "  Total: " & i & " Lines" & IIf(Skipped, " (Skipped: " & Skipped & ")", NullStr)
                    .InsertLines .CountOfLines + 1, Smiley & "CommentOnly: " & NumCommentOnlyLines & " (" & Round(NumCommentOnlyLines / i * 100, 1) & "%)  Commented: " & NumCommentedLines & " (" & Round(NumCommentedLines / i * 100, 1) & "%)  Filled: " & i - NumEmptyLines & " (" & Round((i - NumEmptyLines) / i * 100, 1) & "%)  Empty: " & NumEmptyLines & " (" & Round(NumEmptyLines / i * 100, 1) & "%)  Max Logic Depth: " & MaxIndent
                    RestoreMemberAttributes .Members
                    If PrintLineLen Then
                        PageSetup PrintLineNumber
                        With Printer
                            .CurrentX = PB.Left + LnLen * PrintCharWidth + PrintCharWidth
                            .FontBold = PrinterBoldEnabled
                            .FontItalic = False
                            .ForeColor = vbBlack
                            Printer.Print Left$("Printed by " & AppDetails, PrintLineLen)
                            .CurrentY = PB.Bottom
                            'Sort NaD  
                            j = VBInstance.ActiveCodePane.CodeModule.Members.Count
                            ReDim SortElems(0 To j) 'SortElem 0 is for members.count = 0 -> Redim (0 to 0)
                            i = 0
                            For Each Member In VBInstance.ActiveCodePane.CodeModule.Members
                                Inc i
                                With Member
                                    TmpString1 = Spce & .Name
                                    k = .CodeLocation - 1 'CodeLoc may be too small (VB Bug?)
                                    Do
                                        Inc k
                                        If InStr(VBInstance.ActiveCodePane.CodeModule.Lines(k, 1), TmpString1) Then
                                            Exit Do 'loop 
                                        End If
                                    Loop Until k >= VBInstance.ActiveCodePane.CodeModule.CountOfLines
                                    If k >= VBInstance.ActiveCodePane.CodeModule.CountOfLines Then 'not found
                                        k = 0
                                    End If
                                    SortElems(i) = Array(.Name, k, .Scope, .Static, .Type, Replace$(.Description, vbCrLf, NullStr))
                                End With 'MEMBER
                            Next Member
                            QuickSort 1, j, 0 'sort by name
                            'Print NaD  
                            For i = 1 To j
                                If (.CurrentY > PB.Bottom - PrintLineHeight * 1.25 + PrintLineHeight * (Len(SortElems(i)(5)) <> 0)) Then
                                    .NewPage
                                    PageSetup PrintLineNumber
                                    If WithStationary Then
                                        With PB
                                            Printer.Line (.Left, FrameTop + PrintLineHeight * 6)-(.Right, FrameTop), NaDHeadColor, BF
                                            Printer.Line (.Left + (LnLen + 0.4) * PrintCharWidth, FrameTop + PrintLineHeight * 4)-(.Left + (LnLen + 0.4) * PrintCharWidth, FrameTop + PrintLineHeight * 6), FramePrintColor
                                            Printer.Line (.Left, FrameTop + PrintLineHeight * 4)-(.Right, FrameTop), FramePrintColor, B
                                            Printer.Line (.Left, FrameTop + PrintLineHeight * 6)-(.Right, FrameTop), FramePrintColor, B
                                        End With 'PB
                                    End If
                                    .Fontname = "Arial"
                                    .FontBold = True
                                    .Fontsize = 11
                                    .FontItalic = False
                                    Printer.Print vbCrLf;
                                    .CurrentX = (PB.Right - .TextWidth(NaD) + PB.Left) / 2
                                    .FontUnderline = True
                                    Printer.Print NaD; vbCrLf
                                    .Fontsize = MyFontSize
                                    .Fontname = MyFontName
                                    .FontUnderline = False
                                    .CurrentY = FrameTop + PrintLineHeight * 4
                                    .CurrentX = PB.Left + PrintCharWidth / 4
                                    .FontBold = PrinterBoldEnabled
                                    Printer.Print "Line"; Space$(LnLen - 3); "Member Name"
                                    .CurrentX = PB.Left + PrintCharWidth / 4
                                    Printer.Print "Nmbr"; Space$(LnLen - 3);
                                    .FontItalic = PrinterItalEnabled
                                    Printer.Print " Description"; Tab(PrintLineLen * 2 / 3); "Member Type"
                                    .FontBold = False
                                End If
                                Select Case SortElems(i)(4) 'member type
                                  Case vbext_mt_Method
                                    TmpString2 = "Sub or Function"
                                    Colored = &H40A0& 'red
                                  Case vbext_mt_Property
                                    TmpString2 = "Property"
                                    Colored = &HC00000 'blue
                                  Case vbext_mt_Event
                                    TmpString2 = "Event"
                                    Colored = &H80A000 'cyan
                                  Case vbext_mt_Variable
                                    TmpString2 = "Variable"
                                    Colored = &HA00080 'magenta
                                  Case vbext_mt_Const
                                    TmpString2 = "Constant"
                                    Colored = &H8000& 'green
                                End Select
                                .CurrentX = PB.Left
                                .CurrentY = .CurrentY + PrintLineHeight / 4
                                .FontItalic = False
                                .ForeColor = Colored And ColorRequested
                                Printer.Print Right$(Space$(LnLen) & Format$(SortElems(i)(1)), LnLen); Spce; SortElems(i)(0); Tab(PrintLineLen * 2 / 3);
                                .FontItalic = PrinterItalEnabled
                                Printer.Print Choose(SortElems(i)(2), "Private ", "Public ", "Friend ") & IIf(SortElems(i)(3) And (SortElems(i)(4) = vbext_mt_Method Or SortElems(i)(4) = vbext_mt_Property), "Static ", NullStr); TmpString2
                                TmpString2 = SortElems(i)(5)
                                .ForeColor = CommentPrintColor
                                Do While Len(TmpString2)
                                    k = InStrRev(Left$(TmpString2, PrintLineLen - 1), Spce)
                                    If k < PrintLineLen - WordWrap Then
                                        k = PrintLineLen - 1
                                    End If
                                    .CurrentX = PB.Left
                                    Printer.Print Space$(LnLen + 2); Left$(TmpString2, k)
                                    TmpString2 = LTrim$(Mid$(TmpString2, k + 1))
                                Loop
                                .ForeColor = vbBlack
                                If Len(SortElems(i)(5)) Then
                                    .CurrentY = .CurrentY + PrintLineHeight / 7 'a little extra advance if a desription was printed
                                End If
                            Next i
                        End With 'PRINTER
                    End If
                    Set VarNames = Nothing
                End With ' 'NOT ... '.CODEMODULE
                NoOptExpl = NoOptExpl And (NumCodeLines <> 0)
                Select Case True
                  Case Complain
                    .SetSelection ErrLineFrom, 1, ErrLineTo, Len(.CodeModule.Lines(ErrLineTo, 1)) + 1
                    If NoOptExpl Then
                        .CodeModule.InsertLines FirstCodeLine, OptExpl
                        MarkLine "*", Dummy
                        TopLine = FirstCodeLine
                      Else 'NOOPTEXPL = FALSE/0
                        .SetSelection ErrLineFrom, 1, ErrLineTo, Len(.CodeModule.Lines(ErrLineTo, 1)) + 1
                    End If
                  Case NoOptExpl
                    .CodeModule.InsertLines FirstCodeLine, OptExpl
                    .SetSelection FirstCodeLine, 1, FirstCodeLine, Len(.CodeModule.Lines(FirstCodeLine, FirstCodeLine)) + 1
                    MarkLine "*", Dummy
                    TopLine = FirstCodeLine
                  Case SelCoord.Top <> 0
                    .SetSelection SelCoord.Top, SelCoord.Left, SelCoord.Bottom, SelCoord.Right
                  Case Else
                    .SetSelection TopLine, 1, TopLine, 1
                End Select
                .TopLine = TopLine
                CodeLine = IIf(NoOptExpl, vb2CrLf & """Option Explicit"" is good for catching typos.", NullStr) & _
                           IIf(FoundGoTo, vb2CrLf & """GoTo"" or ""End"" may not be compatible with structured" & vbCrLf & "programming concepts.", NullStr) & _
                           IIf(NonProc, vb2CrLf & """Dim"", ""Static"", or ""Const"" are non-procedural" & vbCrLf & "statements and should come before procedural statements.", NullStr) & _
                           IIf(Dupl, vb2CrLf & "Avoid duplicating module-wide variable names.", NullStr) & _
                           IIf(LNsObs, vb2CrLf & "Line numbers are obsolete.", NullStr) & _
                           IIf(CallUnnec, vb2CrLf & """Call"" is unnecessary.", NullStr) & _
                           IIf(VaCoFuDcl, vb2CrLf & "Check your function-, variable-, or constant-names" & vbCrLf & "and type definitions.", NullStr) & _
                           IIf(DeadCode, vb2CrLf & "Found possibly dead code.", NullStr) & _
                           IIf(MissingScope, vb2CrLf & "Check your scope declarations.", NullStr) & _
                           IIf(SlIf, vb2CrLf & "The structured ""If ... End If"" may be a better choice.", NullStr) & _
                           IIf(Pleo, vb2CrLf & "Using pleonasms is bad coding style and superfluous.", NullStr) & _
                           IIf(ForVarMod, vb2CrLf & "An active For-Variable is modified within a For-Loop." & vbCrLf & "Consider using Do...Loop.", NullStr) & _
                           IIf(NoOEClose, vb2CrLf & """On Error GoTo 0"" should come before procedure end.", NullStr) & _
                           IIf(EmptyNext, vb2CrLf & "Consider combining or repeating ""For"" variable(s) in ""Next"" statement.", NullStr) & _
                           IIf(StrucErr, vb2CrLf & "Place structure delimiters in separate lines.", NullStr) & _
                           IIf(PossVio, vb2CrLf & "Possible structure violation by Exit statement.", NullStr) & _
                           IIf(NoCode, vb2CrLf & "Empty Procedure or superfluous Exit statement.", NullStr) & _
                           IIf(OrphCom, vb2CrLf & "Found orphaned comment(s).", NullStr) & _
                           IIf(NoXPLook Or (ProcessingLastPanel And XPLookRequested And XPDone = False), vb2CrLf & "No WinXP look created. No startup point found" & vbCrLf & " or .EXE-directory invalid.", NullStr) & _
                           IIf(Skipped, vb2CrLf & Skipped & IIf(Skipped = 1, " Line was", " Lines were") & " not examined due to Skip option.", NullStr) & _
                           IIf(Inserted, vb2CrLf & Inserted & OneOrMany(" Line", Inserted) & IIf(Inserted = 1, " was", " were") & " marked" & IIf(Suppressed <> 0 And Inserted, " and", "."), NullStr) & _
                           IIf(Suppressed, vb2CrLf & Suppressed & OneOrMany(" Mark", Suppressed) & IIf(Suppressed = 1, " was", " were") & " suppressed.", NullStr)
                If Len(CodeLine) <> 0 And Len(ToDo(2)) <> 0 And InsertComments Then
                    .CodeModule.InsertLines FirstCodeLine, MyErrMark & Replace$(Mid$(CodeLine, 5), vbCrLf, vbCrLf & MyErrMark) & IIf(Inserted, vbCrLf & MyErrMark & vbCrLf & MyErrMark & "Search for  " & MySignature & "  to locate.", NullStr)
                End If
                If PauseAfterScan = Always Or (PauseAfterScan = IfNecessary And Len(CodeLine)) Then
                    With fSummary
                        .lblSummary = "Summary for " & WindowTitle
                        .lblComplaints = NumDeclLines & OneOrMany(" Declaration line", NumDeclLines) & " and " & NumCodeLines & " Code lines have" & vbCrLf & _
                                         "been" & IIf(SortRequested, " sorted and", NullStr) & " formatted totalling " & NumDeclLines + NumCodeLines & " lines." & CodeLine & vb2CrLf & _
                                         "Maximum Structure Depth is " & MaxIndent
                        .Serious = Complain Or NoOptExpl Or Skipped Or NoXPLook
                        .StopButtonVisible = Not ProcessingLastPanel
                        .ForCompiling = False
                        DoEvents
                        .Show vbModal
                    End With 'FSUMMARY
                    Unload fSummary
                    FinalReq = False
                  Else 'NOT PAUSEAFTERSCAN...
                    Unload fStruc 'normally unloaded by fSummary
                End If
                If NoXPLook Then
                    KillManifest
                    XPLookRequested = 0
                    NoXPLook = False
                End If
                Set NodeStack = Nothing
            End If 'Undo
        End If
    End With 'VBINSTANCE.ACTIVECODEPANE

End Sub

Private Function GetStruc(PopIt As Boolean) As String

  'Stack keeps track of the current stucture level type, ie  
  'whether we are inside a for, do, while, with &c structure bracket  
  'This stack is used to detect possible structure violations by exit statements 

    k = UBound(StrucStack) - 1
    If k < 0 Then
        GetStruc = Mid$(StackUnderflow, 2)
      Else 'NOT K...
        If PopIt Then
            ReDim Preserve StrucStack(k)
        End If
        GetStruc = StrucStack(k)
    End If

End Function

Private Function GetSubstForTS(TypeSuffix As String) As String

    Select Case TypeSuffix
      Case "&"
        GetSubstForTS = " As Long"
      Case "%"
        GetSubstForTS = " As Integer"
      Case "$"
        GetSubstForTS = " As String"
      Case "@"
        GetSubstForTS = " As Currency"
      Case "!"
        GetSubstForTS = " As Single"
      Case "#"
        GetSubstForTS = " As Double"
    End Select

End Function

Private Function HasTypeSuffix(VarName As String) As Boolean

  'Returns True if VarName has type suffix: int% long& single! double# currency@ string$  

    HasTypeSuffix = CBool(InStr("%&@!#$", Right$(VarName, 1)))

End Function

Private Sub InsertBlankLine()

    If EmptyLinesChecked = False Then
        VBInstance.ActiveCodePane.CodeModule.InsertLines FromLine, NullStr
        KillSelection
        Inc FromLine
        Inc ToLine
        Inc LineIndex
    End If

End Sub

Private Function IsOutsideSub() As Boolean

    IsOutsideSub = IndentNext = 0 And FromLine > VBInstance.ActiveCodePane.CodeModule.CountOfDeclarationLines

End Function

Private Function IsRem(Chars As String) As Boolean

    If LCase$(Left$(Chars, 3)) = sRem Then
        IsRem = True
        If Len(Chars) >= 4 Then
            CtB = Mid$(Chars, 4, 1)
            If IsCharAlphaNumeric(CtB(0)) Or CtB(0) = 95 Then 'underline
                IsRem = False
            End If
        End If
    End If

End Function

Private Sub KillManifest()

    On Error Resume Next
        Kill EXEName & ".Manifest"
    On Error GoTo 0

End Sub

Private Sub KillSelection()

    SelCoord.Top = 0
    TopLine = 1

End Sub

Private Sub MarkLine(MarkText As String, ByRef ErrFlag As Boolean)

    ErrFlag = ErrFlag Or (MarkingIsOff = 0)
    Select Case Len(MarkText)
      Case 0
        If (MarkingIsOff And Mo) = Mo Then
            If (MarkingIsOff And Si) = 0 Then
                Inc Suppressed, -((MarkingIsOff And Cc) <> Cc)
            End If
        End If
      Case 1
        'Continue  
      Case Else
        If Not (TypeSuffRequested And VarPtr(ErrFlag) = VarPtr(VaCoFuDcl)) Then
            With VBInstance.ActiveCodePane.CodeModule
                .ReplaceLine LineIndex, .Lines(LineIndex, 1) & MarkText
            End With 'VBINSTANCE.ACTIVECODEPANE.CODEMODULE
        End If
        ShowHalted Mid$(MarkText, Len(MySignature) + 1)
        Inc Inserted
    End Select
    SetComplain

End Sub

Private Sub MenuEvents_Click(ByVal CommandBarControl As Object, Handled As Boolean, CancelDefault As Boolean)

    FormatAll

End Sub

Private Sub PageSetup(ByVal CurrLineNumber As Long)

  Dim Fontname      As String
  Dim Fontsize      As Single
  Dim FontItalic    As Boolean
  Dim DateAndPage   As String
  Dim lp            As Long

    With Printer
        If .CurrentY = 0 Or .CurrentY > PB.Bottom - PrintLineHeight Then
            'next line won't fit  
            Fontname = .Fontname
            Fontsize = .Fontsize
            FontItalic = .FontItalic
            .Fontname = "Arial"
            .FontBold = True
            .Fontsize = 11
            .FontItalic = False
            .ForeColor = vbBlack
            If (.CurrentY > PB.Bottom - PrintLineHeight) And PB.Bottom Then
                .NewPage
            End If
            DateAndPage = Format$(TimePrinted, "Medium Date") & " / Page " & .Page - 1
            If (.Page And 1) Or (BookRequested = False) Then
                'gutter left/top  
                PB = PBOdd
              Else 'NOT (.PAGE...
                'gutter right/bottom  
                PB = PBEven
            End If
            .CurrentX = PB.Left
            .CurrentY = PB.Top
            Printer.Print SrcFileName; " ["; ModuleName; "]";
            .FontBold = False
            .Fontsize = 9
            .CurrentX = PB.Right - .TextWidth(DateAndPage)
            Printer.Print DateAndPage
            .Fontname = Fontname
            .Fontsize = Fontsize
            .FontItalic = FontItalic
            .CurrentY = .CurrentY + PrintLineHeight / 2
            If WithStationary Then
                lp = .CurrentY
                'reading lines  
                Do Until .CurrentY > PB.Bottom
                    Printer.Line (PB.Left, .CurrentY)-(PB.Right, .CurrentY + PrintLineHeight * 3), GridPrintColor, BF 'very light gray
                    .CurrentY = .CurrentY + PrintLineHeight * 3
                Loop
                'erase overshoot  
                Printer.Line (PB.Left, PB.Bottom)-(PB.Right, .ScaleHeight), vbWhite, BF
                'center mark for punching  
                If .Orientation = vbPRORLandscape Then
                    Printer.Line (PB.PunchX, PB.PunchY)-(PB.PunchX, PB.PunchY + LenPunchMark * .TwipsPerPixelY), vbBlack
                  Else 'NOT .ORIENTATION...
                    Printer.Line (PB.PunchX, PB.PunchY)-(PB.PunchX + LenPunchMark * .TwipsPerPixelX, PB.PunchY), vbBlack
                End If
                .CurrentY = lp
                FrameTop = lp
                .DrawMode = vbNotXorPen
                lp = PB.Left + LnLen * PrintCharWidth + PrintCharWidth
                Do
                    lp = lp + IIf(FullTabWidth < 3, 4, FullTabWidth) * PrintCharWidth
                    If lp >= PB.Right Then
                        Exit Do 'loop 
                    End If
                    'tab stop vertical lines  
                    Printer.Line (lp, PB.Bottom)-(lp, .CurrentY), GridPrintColor
                Loop
                .DrawMode = vbCopyPen
                'frame  
                Printer.Line (PB.Left, PB.Bottom)-(PB.Right, .CurrentY), FramePrintColor, B
                Printer.Line (PB.Left + (LnLen + 0.4) * PrintCharWidth, PB.Bottom)-(PB.Left + (LnLen + 0.4) * PrintCharWidth, .CurrentY), FramePrintColor
            End If
            .CurrentY = .CurrentY + PrintLineHeight / 4
            PrintLineNumber = CurrLineNumber
        End If
    End With 'PRINTER

End Sub

Private Sub PopPush(Condition As Boolean, Word1 As String, Word2 As String)

    With VBInstance.ActiveCodePane.CodeModule
        If Condition Then
            TmpString1 = PopWord()
            If MarkChecked Then
                If InStr(1, .Lines(LineIndex, 1), TmpString1, vbTextCompare) = 0 Then
                    .ReplaceLine LineIndex, .Lines(LineIndex, 1) & TmpString1
                    Inc NumCommentedLines
                End If
              Else 'no marking 'MARKCHECKED = FALSE/0
                .ReplaceLine LineIndex, Replace$(.Lines(LineIndex, 1), TmpString1, NullStr)
            End If
            PushWord BlnkApos & IIf(Word2 = "then", NullStr, "NOT ") & Word1 & IIf(Word2 = "then", " = FALSE/0", "...")
          Else 'CONDITION = FALSE/0
            MarkLine ToDo(3), SlIf
            Select Case Words(LastWordIndex)
              Case "sub", "function", "property", "for", "do"
                .ReplaceLine LineIndex, .Lines(LineIndex, 1) & ToDo(4)
            End Select
        End If
    End With 'VBINSTANCE.ACTIVECODEPANE.CODEMODULE

End Sub

Private Function PopWord() As String

  'Retrieve Word from Stack  

    k = UBound(WordStack) - 1
    If k < 0 Then
        PopWord = sNone
      Else 'NOT K...
        ReDim Preserve WordStack(k)
        PopWord = UCase$(WordStack(k))
    End If

End Function

Private Sub PrintSkipComment()

    If PrintLineLen Then
        PageSetup PrintLineNumber
        With Printer
            .CurrentX = PB.Left
            .FontBold = PrinterBoldEnabled
            .FontItalic = False
            .ForeColor = vbRed And ColorRequested
        End With 'PRINTER
        Printer.Print PrintMark; Format$(ToLine - PrintLineNumber - 1); " Line(s) skipped"
    End If

End Sub

Private Sub PushWord(Word As String)

  'Save (modified) Word on Stack  

    WordStack(UBound(WordStack)) = Replace$(Word, Chr$(0), Spce)
    ReDim Preserve WordStack(UBound(WordStack) + 1)

End Sub

Private Function ReplaceTypeSuffix(InText As String) As Long

  'Replaces Type Suffixes in InText and returns 0 if nothing has happened  
  '                                             1 if type suffic replacement only  
  '                                             2 if type suffic replacement and error mark insertion necessary  

  Dim Ptr       As Long
  Dim Replaced  As String

    If TypeSuffRequested Then
        Replaced = Spce & InText & Spce 'add a space before and after
        Ptr = 1
        Do Until Ptr > Len(Replaced)
            If Mid$(Replaced, Ptr, 1) = Quote Then
                Do
                    Inc Ptr
                Loop Until Mid$(Replaced, Ptr, 1) = Quote
            End If
            If Mid$(Replaced, Ptr, 2) = " '" Or Mid$(Replaced, Ptr, 5) = " rem " Then
                Exit Do 'loop 
            End If
            Select Case True
              Case Mid$(Replaced, Ptr, 3) = " & " 'special concatenation &
                Inc Ptr, 2 'skip that
              Case Mid$(Replaced, Ptr, 2) Like "[%@!#$&](" And BracketCount = 0 'thats a TS followed by an open bracket
                PendingTSRepl = GetSubstForTS(Mid$(Replaced, Ptr, 1)) 'so we have to defer replacement
                Replaced = Left$(Replaced, Ptr - 1) & Mid$(Replaced, Ptr + 1) 'and just remove the TS
                BracketCount = 1 'we're inside a bracket pair
                ReplaceTypeSuffix = 1 'make it replace that line
              Case Mid$(Replaced, Ptr, 2) Like "[%@!#$&][,) ]" 'thats a TS followed by a separation char
                Replaced = Left$(Replaced, Ptr - 1) & GetSubstForTS(Mid$(Replaced, Ptr, 1)) & Mid$(Replaced, Ptr + 1) 'replace TS
                If BracketCount Then 'we're inside a bracket pair
                    ReplaceTypeSuffix = 1 'make it replace that line only
                    PendingMark = ToDo(13) 'but defer marking
                  Else 'not inside a bracket pair 'BRACKETCOUNT = FALSE/0
                    ReplaceTypeSuffix = 2 'make it replace that line and mark it
                End If
              Case Mid$(Replaced, Ptr, 1) = "(" 'another open bracket
                'Note: the open bracket detected in the second Case has been skipped  
                '      due to removing the TS and shifting the rest of InText one place  
                '      to the left  
                Inc BracketCount
              Case Mid$(Replaced, Ptr, 1) = ")" 'a closing bracket
                Dec BracketCount
                If BracketCount = 0 Then 'outside bracket
                    Replaced = Replaced & PendingTSRepl & PendingMark
                    ReplaceTypeSuffix = IIf(Len(PendingTSRepl) And Len(PendingMark) = 0, 2, 1)
                    If Len(PendingMark) Then
                        MarkLine ToDo(13), VaCoFuDcl
                    End If
                    PendingTSRepl = NullStr 'kill defered insertions
                    PendingMark = NullStr
                End If
            End Select
            Inc Ptr
        Loop
        InText = Mid$(Replaced, 2)
    End If

End Function

Private Sub RestoreMemberAttributes(Membs As Members)

  'restore the member attributes  

    With fPreparing
        .lbl(0) = "Tidying up..."
        .imgSort.Visible = False
        .imgFormat.Visible = False
        .imgBroom.Visible = True
        .Show
    End With 'FPREPARING

    For i = 1 To UBound(MemberAttributes)
        Err.Clear
        On Error Resume Next
            'may produce an error on undo when member attributes cannot be restored  
            'because a new member was created after the last format scan and thats  
            'now missing in the undo buffer but it's attributes have been saved  
            With Membs(MemberAttributes(i)(MemName))
                If Err = 0 Then
                    If Len(MemberAttributes(i)(MemCate)) Then
                        .Category = MemberAttributes(i)(MemCate)
                    End If
                    If Len(MemberAttributes(i)(MemDesc)) Then
                        .Description = MemberAttributes(i)(MemDesc)
                    End If
                    If MemberAttributes(i)(MemHelp) Then
                        .HelpContextID = MemberAttributes(i)(MemHelp)
                    End If
                    If Len(MemberAttributes(i)(MemProp)) Then
                        .PropertyPage = MemberAttributes(i)(MemProp)
                    End If
                    If MemberAttributes(i)(MemStMe) <= 0 Then
                        .StandardMethod = MemberAttributes(i)(MemStMe)
                    End If
                    If MemberAttributes(i)(MemBind) Then
                        .Bindable = True
                    End If
                    If MemberAttributes(i)(MemBrws) Then
                        .Browsable = True
                    End If
                    If MemberAttributes(i)(MemDfbd) Then
                        .DefaultBind = True
                    End If
                    If MemberAttributes(i)(MemDbnd) Then
                        .DisplayBind = True
                    End If
                    If MemberAttributes(i)(MemHidd) Then
                        .Hidden = True
                    End If
                    If MemberAttributes(i)(MemRqEd) Then
                        .RequestEdit = True
                    End If
                    If MemberAttributes(i)(MemUiDe) Then
                        .UIDefault = True
                    End If
                End If
            End With 'MEMBS(MEMBERATTRIBUTES(I)(MEMNAME))
        On Error GoTo 0
        DoEvents
    Next i
    Unload fPreparing

End Sub

Private Sub SaveMemberAttributes(Membs As Members)

    i = 0
    ReDim MemberAttributes(0 To Membs.Count)
    On Error Resume Next
        For Each Member In Membs
            Inc i
            Rem Something strange happens here: _
                Getting the Description Attribute sometimes fails _
                after the source under examination has run in the IDE. _
                It never fails when the source has not been running yet. _
                Maybe it's a VB-Bug(?)  
            Err.Clear
            With Member '                   MemName MemBind    MemBrws     MemCate    MemDfbd       MemDesc       MemDbnd       MemHelp         MemHidd  MemProp        MemRqed       MemStme          MemUide
                MemberAttributes(i) = Array(.Name, .Bindable, .Browsable, .Category, .DefaultBind, .Description, .DisplayBind, .HelpContextID, .Hidden, .PropertyPage, .RequestEdit, .StandardMethod, .UIDefault)
                If Err Then 'getting the attributes has failed; so don't store them
                    Dec i
                End If
            End With 'MEMBER
            DoEvents
        Next Member
    On Error GoTo 0
    ReDim Preserve MemberAttributes(0 To i) 'Redim to reflect the real number of stored attributes

End Sub

Private Sub SetComplain(Optional ByVal Offset As Long = 0)

    If MarkingIsOff = Mo Or MarkingIsOff = 0 Then
        Complain = True
        AnyComplain = True
        TopLine = FromLine + Offset
        ErrLineFrom = FromLine + Offset
        ErrLineTo = ToLine + Offset
    End If

End Sub

Private Sub SetToDo(ByVal Reason As Long)

    MarkingIsOff = MarkingIsOff And Not Reason And Not Si
    If MarkingIsOff = 0 Then
        ToDo(1) = ToDo1
        ToDo(2) = ToDo2
        ToDo(3) = ToDo3
        ToDo(4) = ToDo4
        ToDo(5) = ToDo5
        ToDo(6) = ToDo6
        ToDo(7) = ToDo7
        ToDo(8) = ToDo8
        ToDo(9) = ToDo9
        ToDo(10) = ToDo10
        ToDo(11) = ToDo11
        ToDo(12) = ToDo12
        ToDo(13) = ToDo13
        ToDo(14) = ToDo14
        ToDo(15) = ToDo15
        If LinNumChecked Then
            ToDo(16) = Done16
          Else 'LINNUMCHECKED = FALSE/0
            ToDo(16) = ToDo16
        End If
        ToDo(17) = ToDo17
        ToDo(18) = ToDo18
        ToDo(19) = ToDo19
        ToDo(20) = ToDo20
        ToDo(21) = ToDo21
        ToDo(22) = ToDo22
        ToDo(23) = ToDo23
        ToDo(24) = ToDo24
        ToDo(25) = ToDo25
        ToDo(26) = ToDo26
        ToDo(27) = ToDo27
        ToDo(28) = ToDo28
        ToDo(29) = ToDo29
    End If

End Sub

Private Sub ShowHalted(ErrText As String)

  'shows fHalted and waits till it's gone again  

    If StopRequested Then
        With fHalted
            .lbErrText = Spce & ErrText & Spce
            .Show
            With VBInstance.ActiveCodePane
                .Show
                .SetSelection LineIndex, 1, LineIndex, 1023
            End With 'VBINSTANCE.ACTIVECODEPANE
            Do While .Visible 'fHalted hides itself
                DoEvents
            Loop
        End With 'FHALTED
        Unload fHalted
        DoEvents
    End If

End Sub

Private Sub SubsequentCodeIsDead(Optional ByVal Level As Long = 0)

  'Prepares for Dead Code Recognition  

    If CodeIsDead = AlmostInfinity Then
        CodeIsDead = -Indent - Level
    End If

End Sub

Private Sub SuspendIndent()

    IndentNext = Indent
    Indent = 0
    RemIndentNext = RemIndent
    RemIndent = 0
    OEIndentNext = OEIndent
    OEIndent = NullStr

End Sub

Private Function TypeIsUndefined(VarName As String, IsConst As Boolean) As Boolean

  'Returns True if type of variable or constant is undefined  

    If Not HasTypeSuffix(VarName) Then
        TypeIsUndefined = IsConst Or (InStr(DefTypeChars, Left$(VarName, 1)) = 0)
    End If

End Function

':) Ulli's VB Code Formatter V2.24.11 (2008-Apr-11 10:26)  Decl: 1092  Code: 2908  Total: 4000 Lines
':) CommentOnly: 668 (16,7%)  Commented: 409 (10,2%)  Filled: 3582 (89,6%)  Empty: 418 (10,4%)  Max Logic Depth: 15
