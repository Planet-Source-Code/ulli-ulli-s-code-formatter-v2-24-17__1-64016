VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z 'we're 32 bit

'stack to keep track of the treeview nodes
Private NodeStack() As MSComctlLib.Node
Private Index       As Long
Private Capacity    As Long

Private Sub Class_Initialize()

    Reset 5

End Sub

Private Sub Class_Terminate()

    Reset 0

End Sub

Public Function Pop() As MSComctlLib.Node

    If Index Then
        Dec Index
        Set Pop = NodeStack(Index)
      Else 'INDEX = FALSE/0
        Set Pop = Nothing
    End If

End Function

Public Sub Push(NodeToPush As MSComctlLib.Node)

    If Index > Capacity Then
        ReDim Preserve NodeStack(Index + 5)
        Capacity = Index + 5
    End If
    Set NodeStack(Index) = NodeToPush
    Inc Index

End Sub

Public Sub Reset(ByVal Cap As Long)

    Index = 0
    ReDim NodeStack(Cap)
    Capacity = Cap

End Sub

Public Static Property Get StackSize() As Long

    StackSize = Index

End Property

':) Ulli's VB Code Formatter V2.24.11 (2008-Apr-11 10:26)  Decl: 7  Code: 51  Total: 58 Lines
':) CommentOnly: 3 (5,2%)  Commented: 2 (3,4%)  Filled: 38 (65,5%)  Empty: 20 (34,5%)  Max Logic Depth: 2
