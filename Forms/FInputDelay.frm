VERSION 5.00
Begin VB.Form FInputDelay 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Edit Input Delay"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5055
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text1 
      Alignment       =   1  'Rechts
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton BtnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox TxtDelaytime 
      Alignment       =   1  'Rechts
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Delaytime in milliseconds:"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "[ms]"
      Height          =   255
      Left            =   4560
      TabIndex        =   5
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Name/Index/Identifier:"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "FInputDelay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Result As VbMsgBoxResult
Private m_Object As WndInputDelay

Public Function ShowDialog(obj As WndInputDelay, Owner As Form) As VbMsgBoxResult
    Set m_Object = obj.Clone
    UpdateView
    Me.Show vbModal, Owner
    ShowDialog = m_Result
    If ShowDialog = vbCancel Then Exit Function
    obj.NewC m_Object
End Function

Sub UpdateView()
    TxtDelaytime.Text = m_Object.Milliseconds
End Sub

Function UpdateData() As Boolean
    Dim s As String: s = TxtDelaytime.Text
    If Not IsNumeric(s) Then
        Dim mr As VbMsgBoxResult: mr = MsgBox("Please give a numeric value: " & vbCrLf & s, vbOKCancel)
        If mr = vbCancel Then
            UpdateView
        Else
            TxtDelaytime.SetFocus
        End If
        Exit Function
    End If
    Dim dt As Long: dt = CLng(TxtDelaytime.Text)
    Set m_Object = MNew.WndInputDelay(dt)
    UpdateData = True
End Function

Private Sub BtnCancel_Click()
    m_Result = VbMsgBoxResult.vbCancel
    Unload Me
End Sub

Private Sub BtnOK_Click()
    If Not UpdateData Then Exit Sub
    m_Result = VbMsgBoxResult.vbOK
    Unload Me
End Sub

Private Sub TxtDelaytime_LostFocus()
    UpdateData
End Sub
