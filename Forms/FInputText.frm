VERSION 5.00
Begin VB.Form FInputText 
   Caption         =   "Edit Input Text"
   ClientHeight    =   4455
   ClientLeft      =   120
   ClientTop       =   465
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
   ScaleHeight     =   4455
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CheckBox ChkDelayed 
      Caption         =   "Delayed"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox TxtDelaytime 
      Height          =   3135
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   2
      Text            =   "FInputText.frx":0000
      Top             =   1320
      Width           =   5055
   End
   Begin VB.CommandButton BtnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Name/Identifier"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "[ms]"
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Text:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "FInputText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

End Sub
