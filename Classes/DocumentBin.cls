VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DocumentBin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Implements IWndInputSender
'todo:
'Klassen für Makros:
'* Makro Text      : erzeugt eine Sequenz von Zeichen mit Groß/kleinschreibung optional mit oder ohne Delay zwischen den Zeichen
'* Makro MausClick : erzeugt einen Maussprung zu einer absoluten Position mit einem anschließenden Einfach- oder Doppelklick
'* Makro Wnd-Handle: ermittelt das Handle des Fensters unter der Maus und setzt es als das aktuelles Fensterhandle
'
'hier:
'in Datei speichern, von Datei lesen

Private m_Inputs As Collection 'Of IWndInputSender


Private Sub Class_Initialize()
    Set m_Inputs = New Collection
End Sub

Public Sub Clear()
    Class_Initialize
End Sub

Public Sub Add(aInputsender As IWndInputSender)
    m_Inputs.Add aInputsender
End Sub

Public Sub Send()
    Dim inp As IWndInputSender
    For Each inp In m_Inputs
        inp.Send
    Next
End Sub

Private Sub IWndInputSender_Send()
    Me.Send
End Sub
