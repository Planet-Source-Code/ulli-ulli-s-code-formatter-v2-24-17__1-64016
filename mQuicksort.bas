Attribute VB_Name = "mQuicksort"
Option Explicit
DefLng A-Z 'we're 32 bits

Public SortElems()  As Variant 'table of vbArrays
Public TempElem     As Variant 'one element as temporary
Private TempKey     As String

Public Sub QuickSort(ByVal xFrom As Long, ByVal xThru As Long, ByVal yKey As Long)

  'Sorts a table of vbArrays

  Dim xLeft As Long, xRite As Long

    Do While xFrom < xThru  'we have something to sort (@ least two elements)
        xLeft = xFrom
        xRite = xThru
        TempElem = SortElems(xLeft) 'get ref element and make room
        TempKey = LCase$(TempElem(yKey))
        Do
            Do Until xRite = xLeft
                If LCase$(SortElems(xRite)(yKey)) >= TempKey Then
                    Dec xRite
                  Else 'is smaller than ref so move it to the left... 'NOT LCASE$(SORTELEMS(XRITE)(YKEY))...
                    SortElems(xLeft) = SortElems(xRite)
                    Inc xLeft '...and leave the item just moved alone for now
                    Exit Do 'loop 
                End If
            Loop
            Do Until xLeft = xRite
                If LCase$(SortElems(xLeft)(yKey)) <= TempKey Then
                    Inc xLeft
                  Else 'is greater than ref so move it to the right... 'NOT LCASE$(SORTELEMS(XLEFT)(YKEY))...
                    SortElems(xRite) = SortElems(xLeft)
                    Dec xRite '...and leave the item just moved alone for now
                    Exit Do 'loop 
                End If
            Loop
        Loop Until xLeft = xRite
        'now the indexes have met and all bigger items are to the right and all smaller items are left
        SortElems(xRite) = TempElem 'insert ref elem in proper place and sort the two areas left and right of it
        If xLeft - xFrom < xThru - xRite Then 'smaller part 1st to reduce recursion depth
            xLeft = xFrom
            xFrom = xRite + 1
            xRite = xRite - 1
          Else 'NOT XLEFT...
            xRite = xThru
            xThru = xLeft - 1
            xLeft = xLeft + 1
        End If
        If xLeft < xRite Then 'smaller part is not empty...
            QuickSort xLeft, xRite, yKey '...so sort it
        End If
    Loop

End Sub

':) Ulli's VB Code Formatter V2.24.11 (2008-Apr-11 10:26)  Decl: 6  Code: 52  Total: 58 Lines
':) CommentOnly: 4 (6,9%)  Commented: 14 (24,1%)  Filled: 51 (87,9%)  Empty: 7 (12,1%)  Max Logic Depth: 5
