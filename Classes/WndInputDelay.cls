VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WndInputDelay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IWndInputSender
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliSeconds As Long)
Private m_Name    As String
Private m_Milsecs As Long

Friend Sub New_(Name As String, Milliseconds As Long)
    m_Milsecs = Milliseconds
End Sub

Public Sub NewC(other As WndInputDelay)
    With other
        m_Name = other.Name
        m_Milsecs = other.Milliseconds
    End With
End Sub

Public Function Clone() As WndInputDelay
    Set Clone = New WndInputDelay: Clone.NewC Me
End Function

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Get Key() As String
    Key = m_Name
End Property

Public Property Get Milliseconds() As Long
    Milliseconds = m_Milsecs
End Property
'Public Property Let Milliseconds(ByVal value As Long)
'    m_Milsecs = value
'End Property

Public Sub Send()
    Sleep m_Milsecs
End Sub

' v ' ############################## ' v '    Implements IWndInputSender    ' v ' ############################## ' v '
Private Sub IWndInputSender_Send()
    Me.Send
End Sub

Private Property Get IWndInputSender_Key() As String
    IWndInputSender_Key = Me.Key
End Property

