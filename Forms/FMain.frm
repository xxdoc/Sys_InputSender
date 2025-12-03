VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Input Actions"
   ClientHeight    =   4695
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7455
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnMoveDown 
      Caption         =   "v"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      ToolTipText     =   "Move down"
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton BtnMoveUp 
      Caption         =   "^"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      ToolTipText     =   "Move up"
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton BtnDelete 
      Caption         =   "-"
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      ToolTipText     =   "Delete the selected object"
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton BtnClone 
      Caption         =   "++"
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      ToolTipText     =   "Clone the selected object"
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton BtnClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      ToolTipText     =   "Delete the list"
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton BtnSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      ToolTipText     =   "Send all inputs"
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton BtnNewText 
      Caption         =   "+Text"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton BtnNewDelay 
      Caption         =   "+Delay"
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      ToolTipText     =   "Add New Mouse-Input"
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton BtnNewInputs 
      Caption         =   "+Inputs"
      Height          =   375
      Left            =   0
      TabIndex        =   8
      ToolTipText     =   "Add New Keyboard-Input"
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox TxtToStr 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   10
      Top             =   360
      Width           =   3855
   End
   Begin VB.ListBox LstWndInputs 
      Height          =   4140
      ItemData        =   "FMain.frx":0CCA
      Left            =   0
      List            =   "FMain.frx":0CCC
      TabIndex        =   9
      ToolTipText     =   "Select to view; doubleclick to edit"
      Top             =   360
      Width           =   3135
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   " &? "
      Begin VB.Menu mnuHelpInfo 
         Caption         =   "&Info"
      End
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BtnNewDelay_Click()
    Dim idt As New WndInputDelay
    If FInputDelay.ShowDialog(idt, Me) = vbCancel Then Exit Sub
    MApp.CurDocument.Add idt
End Sub

Private Sub BtnNewInputs_Click()
    Dim ipts As New WndInputs
    If FInputs.ShowDialog(ipts, Me) = vbCancel Then Exit Sub
    MApp.CurDocument.Add ipts
End Sub

Private Sub BtnSend_Click()
    MApp.CurDocument.Send
End Sub

Private Sub BtnClear_Click()
    MApp.CurDocument.Clear
End Sub

Private Sub Form_Resize()
    Dim l As Single, t As Single, W As Single, h As Single
    l = TxtToStr.Left:     t = TxtToStr.Top
    W = Me.ScaleWidth - l: h = Me.ScaleHeight - t
    If W > 0 And h > 0 Then TxtToStr.Move l, t, W, h
End Sub
