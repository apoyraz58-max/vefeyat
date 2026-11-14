VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "VefeyatOlaylar0341"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mWord As Word.Application
Private mRibbon As Object

Public Sub Attach(ByVal host As Word.Application, ByVal ribbon As Object)
    Detach
    Set mWord = host
    Set mRibbon = ribbon
End Sub

Public Sub Detach()
    Set mWord = Nothing
    Set mRibbon = Nothing
End Sub

Private Sub mWord_DocumentChange()
    On Error Resume Next
    VefeyatEklenti0341.Vefeyat0341UiMemoryPrune
    If mWord.Documents.Count > 0 Then VefeyatEklenti0341.Vefeyat0341UiMemoryActivate mWord.ActiveDocument
    On Error GoTo 0
    NotifyRibbon
End Sub

Private Sub mWord_WindowSelectionChange(ByVal Sel As Word.Selection)
    VefeyatEklenti0341.Vefeyat0341UiMemorySelection Sel
    NotifyRibbon
End Sub

Private Sub mWord_WindowActivate(ByVal Doc As Word.Document, ByVal Wn As Word.Window)
    VefeyatEklenti0341.Vefeyat0341UiMemoryActivate Doc
    NotifyRibbon
End Sub

Private Sub mWord_DocumentOpen(ByVal Doc As Word.Document)
    VefeyatEklenti0341.Vefeyat0341UiMemoryActivate Doc
End Sub

Private Sub mWord_NewDocument(ByVal Doc As Word.Document)
    VefeyatEklenti0341.Vefeyat0341UiMemoryActivate Doc
End Sub

Private Sub mWord_DocumentBeforeClose(ByVal Doc As Word.Document, Cancel As Boolean)
    VefeyatEklenti0341.Vefeyat0341UiInventoryDocumentClose Doc
    VefeyatEklenti0341.Vefeyat0341UiToolDocumentClose Doc
    VefeyatEklenti0341.Vefeyat0341UiReviewDocumentClose Doc
    VefeyatEklenti0341.Vefeyat0341UiMemoryClose Doc
End Sub

Private Sub NotifyRibbon()
    ' Memory callbacks read local blocks only; event handlers never edit text.
    On Error Resume Next
    If Not mRibbon Is Nothing Then mRibbon.Invalidate
    On Error GoTo 0
End Sub

Private Sub Class_Terminate()
    Detach
End Sub
