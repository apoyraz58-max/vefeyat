VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VefeyatKisiSec0341
   Caption         =   "Vefeyat - Kisi secimi"
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9000
   OleObjectBlob   =   "VefeyatKisiSec0341.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "VefeyatKisiSec0341"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function VfkActiveWindow Lib "user32" Alias "GetActiveWindow" () As LongPtr
Private Declare PtrSafe Function VfkClassNameW Lib "user32" Alias "GetClassNameW" (ByVal hwnd As LongPtr, ByVal buffer As LongPtr, ByVal size As Long) As Long
Private Declare PtrSafe Function VfkSetWindowTextW Lib "user32" Alias "SetWindowTextW" (ByVal hwnd As LongPtr, ByVal title As LongPtr) As Long
#Else
Private Declare Function VfkActiveWindow Lib "user32" Alias "GetActiveWindow" () As Long
Private Declare Function VfkClassNameW Lib "user32" Alias "GetClassNameW" (ByVal hwnd As Long, ByVal buffer As Long, ByVal size As Long) As Long
Private Declare Function VfkSetWindowTextW Lib "user32" Alias "SetWindowTextW" (ByVal hwnd As Long, ByVal title As Long) As Long
#End If


' Empty designer binary: Copyright (c) 2026 Samy Haliem, MIT License.
' See THIRD-PARTY-NOTICES.txt. All code below is Vefeyat source.

Private WithEvents mResults As MSForms.ListBox
Private WithEvents mChoose As MSForms.CommandButton
Private WithEvents mCancel As MSForms.CommandButton
Private WithEvents mSearch As MSForms.CommandButton
Private WithEvents mQuery As MSForms.TextBox
Private WithEvents mDetails As MSForms.TextBox
Private mPrompt As MSForms.Label
Private mRows As Collection
Private mBrowse As Boolean
Private mPlan As Boolean
Private mRegistryMode As Boolean
Private mPlanCoverage As String
Public PlanAction As String
Private mWorkbench As Boolean
Private mInventory As Boolean
Private mInventoryUpdating As Boolean
Private mCatalogChoice As Boolean
Private mToolChanges As Boolean
Private mIdentityProof As Object
Private mExplainQuery As String
Private mToolHeading As String
Private mEarlier As Boolean
Private mDia As Boolean
Private mLiveReview As Boolean
Private mLiveClosing As Boolean
Private mLiveCollapsed As Boolean
Private mLiveHeight As Single
Private mLiveVisibility As Object
Private WithEvents mContext As MSForms.TextBox
Private mLiveNote As MSForms.Label
Private WithEvents mBack As MSForms.CommandButton
Private WithEvents mLocate As MSForms.CommandButton
Private WithEvents mToggle As MSForms.CommandButton
Private WithEvents mPause As MSForms.CommandButton
Private WithEvents mIgnore As MSForms.CommandButton
Private mSpelling As Boolean
Private mBatch As Boolean
Private mNameMode As Boolean
Private mExclusionMode As Boolean
Private mReview As Boolean
Private mProgress As Boolean
Private mProgressStarted As Boolean
Private mAllowCancel As Boolean
Private mTrack As MSForms.Label
Private mFill As MSForms.Label
Private mNotice As MSForms.Label
Private mReviewExcluded As Object
Private WithEvents mStop As MSForms.CommandButton
Public StopRequested As Boolean


' Decorative hit layers are disabled: native buttons own pointer and keyboard events.
Private Const VfkFlatSurfacesEnabled As Boolean = False
Private WithEvents mChooseFace As MSForms.Label
Private WithEvents mChooseText As MSForms.Label
Private WithEvents mCancelFace As MSForms.Label
Private WithEvents mCancelText As MSForms.Label
Private WithEvents mSearchFace As MSForms.Label
Private WithEvents mSearchText As MSForms.Label
Private WithEvents mBackFace As MSForms.Label
Private WithEvents mBackText As MSForms.Label
Private WithEvents mLocateFace As MSForms.Label
Private WithEvents mLocateText As MSForms.Label
Private WithEvents mToggleFace As MSForms.Label
Private WithEvents mToggleText As MSForms.Label
Private WithEvents mPauseFace As MSForms.Label
Private WithEvents mPauseText As MSForms.Label
Private WithEvents mIgnoreFace As MSForms.Label
Private WithEvents mIgnoreText As MSForms.Label
Private WithEvents mStopFace As MSForms.Label
Private WithEvents mStopText As MSForms.Label
Private mVfkFocus As String
Private mVfkArmed As String
Private mVfkPressed As String
Private mVfkClickBusy As Boolean

' One-based index in the candidates collection; 0 means cancellation.
Public ChosenIndex As Long

Public Function ConfirmationVersion() As Long
    ConfirmationVersion = 15
End Function

' Every record is a Variant array indexed 0..5:
' short name, full name, original name, death date, description, source link.
Public Sub InitCandidates(ByVal people As Collection, ByVal query As String)
    Dim i As Long, person As Variant
    Set mIdentityProof = Nothing
    Set mRows = people
    mExplainQuery = query
    ChosenIndex = 0
    StopRequested = False
    mBatch = False
    mEarlier = False
    mDia = False
    mNameMode = False
    mReview = False
    mStop.Visible = False
    mResults.Clear
    mBrowse = False
    mChoose.Visible = True
    mCancel.Caption = VfkText("\u0130ptal")
    mPrompt.Caption = VfkText("Se\u00e7ili ad: ") & Left$(query, 100) & vbCrLf & CStr(people.Count) & _
        VfkText(" aday bulundu. Tan\u0131t\u0131m\u0131 kontrol edip eklemek istedi\u011finiz \u015fahs\u0131 do\u011frulay\u0131n.")
    For i = 1 To people.Count
        person = people(i)
        mResults.AddItem CStr(person(0))
        mResults.List(i - 1, 1) = CStr(person(3))
        mResults.List(i - 1, 2) = CStr(person(1))
    Next i
    mChoose.Enabled = False
    If people.Count > 0 Then
        mResults.ListIndex = 0
        mResults_Change
    Else
        mDetails.Text = VfkText("E\u015fle\u015fen \u015fah\u0131s bulunamad\u0131.")
    End If
End Sub

Public Sub InitWorkbench(ByVal people As Collection, ByVal heading As String, ByVal canChange As Boolean)
    Dim control As Object
    mWorkbench = True: mToolChanges = canChange: mToolHeading = heading
    InitCandidates people, heading
    For Each control In Me.Controls
        If control.Tag = "VefeyatHeading" Then control.Visible = False
    Next control
    mPrompt.Caption = heading: mPrompt.Height = 50
    mResults.ColumnCount = 1: mResults.ColumnWidths = "570 pt"
    mChoose.Caption = VfkText("Onar / ba\u011f\u0131 yenile")
    mChoose.Left = 16: mChoose.Width = 168
    If people.Count = 0 Then mChoose.Enabled = False
    mStop.Caption = VfkText("Ba\u011f\u0131 kald\u0131r"): mStop.Visible = canChange
    mStop.Left = 196: mStop.Width = 132
    mCancel.Caption = "Kapat"
    mCancel.Left = 496: mCancel.Width = 120
    If mLocate Is Nothing Then
        Set mLocate = Me.Controls.Add("Forms.CommandButton.1", "ToolLocate", True)
        mLocate.Left = 340: mLocate.Top = 416: mLocate.Width = 144: mLocate.Height = 28
        mLocate.Caption = VfkText("Metinde g\u00f6ster")
    End If
    mLocate.Enabled = people.Count > 0
    mQuery.Visible = False: mSearch.Visible = False
    If people.Count = 0 Then mDetails.Text = VfkText("G\u00f6sterilecek kay\u0131t yok. Bu pencerede belge de\u011fi\u015ftirilmedi.")
    mChoose.Accelerator = "O": mCancel.Accelerator = "K"
    If people.Count > 0 Then mResults_Change
    VfkRefreshButtons
End Sub

Public Sub InitLiveReview()
    Dim control As Object
    mLiveReview = True: mLiveClosing = False
    mProgress = False: mBrowse = False
    Me.Width = 340: Me.Height = 520
    Me.ScrollBars = fmScrollBarsVertical: Me.ScrollHeight = 494
    mPrompt.Left = 12: mPrompt.Top = 10: mPrompt.Width = 308: mPrompt.Height = 46
    For Each control In Me.Controls
        If control.Tag = "VefeyatHeading" Then control.Visible = False
    Next control
    Set mContext = Me.Controls.Add("Forms.TextBox.1", "LiveContext", True)
    With mContext
        .Left = 12: .Top = 62: .Width = 308: .Height = 80
        .MultiLine = True: .WordWrap = True: .Locked = True
        .ScrollBars = fmScrollBarsVertical
    End With
    mResults.Left = 12: mResults.Top = 152: mResults.Width = 308: mResults.Height = 82
    mResults.ColumnCount = 1: mResults.ColumnWidths = "288 pt"
    mDetails.Left = 12: mDetails.Top = 242: mDetails.Width = 308: mDetails.Height = 100
    Set mBack = Me.Controls.Add("Forms.CommandButton.1", "LiveBack", True)
    With mBack
        .Caption = VfkText("\u00d6nceki"): .Left = 12: .Top = 350: .Width = 94: .Height = 26
    End With
    Set mLocate = Me.Controls.Add("Forms.CommandButton.1", "LiveLocate", True)
    With mLocate
        .Caption = VfkText("Metinde g\u00f6ster"): .Left = 114: .Top = 350: .Width = 100: .Height = 26
    End With
    Set mToggle = Me.Controls.Add("Forms.CommandButton.1", "LiveToggle", True)
    With mToggle
        .Caption = VfkText("K\u00fc\u00e7\u00fclt"): .Left = 222: .Top = 350: .Width = 98: .Height = 26
    End With
    mChoose.Left = 12: mChoose.Top = 384: mChoose.Width = 148: mChoose.Height = 28
    mChoose.Caption = "Onayla ve sonraki": mChoose.Default = False
    mStop.Left = 170: mStop.Top = 384: mStop.Width = 150: mStop.Height = 28
    mStop.Caption = "Atla / Sonraki": mStop.Visible = True
    mCancel.Left = 222: mCancel.Top = 420: mCancel.Width = 98: mCancel.Height = 26
    mCancel.Caption = "Bitir": mCancel.Default = False
    Set mPause = Me.Controls.Add("Forms.CommandButton.1", "LivePause", True)
    With mPause
        .Caption = "Duraklat": .Left = 12: .Top = 420: .Width = 94: .Height = 26
        .ControlTipText = VfkText("Bu Word oturumunda Denetime devam et ile kald\u0131\u011f\u0131n\u0131z yerden devam edin.")
    End With
    Set mIgnore = Me.Controls.Add("Forms.CommandButton.1", "LiveIgnore", True)
    With mIgnore
        .Caption = VfkText("\u0130fadeyi d\u0131\u015fla"): .Left = 114: .Top = 420: .Width = 100: .Height = 26
        .ControlTipText = VfkText("Ayn\u0131 yaz\u0131m\u0131 bu belgenin sonraki toplu denetimlerinde sorma. Kaydetmeden \u00f6nce onay istenir.")
    End With
    Set mLiveNote = Me.Controls.Add("Forms.Label.1", "LiveNote", True)
    With mLiveNote
        .Left = 12: .Top = 452: .Width = 308: .Height = 38: .WordWrap = True: .Font.Size = 8
    End With
End Sub

Public Sub ShowLiveReview(ByVal people As Collection, ByVal heading As String, ByVal context As String, Optional ByVal choice As Long = 1)
    Dim i As Long, row As Variant
    If mLiveCollapsed Then mToggle_Click
    Set mRows = people
    mExplainQuery = ""
    mResults.Clear
    mPrompt.Caption = heading: mContext.Text = context
    For i = 1 To people.Count
        row = people(i)
        mResults.AddItem CStr(row(0)) & " | " & CStr(row(3))
    Next i
    mStop.Enabled = True: mLocate.Enabled = True
    If people.Count > 0 Then
        mResults.ListIndex = choice - 1
        mResults_Change
    Else
        mDetails.Text = VfkText("Bu ge\u00e7i\u015f i\u00e7in aday yok.")
        mChoose.Enabled = False
    End If
    VfkRefreshButtons
End Sub

Public Function LiveReviewChoice() As Long
    If mResults.ListIndex >= 0 Then LiveReviewChoice = mResults.ListIndex + 1
End Function

Public Sub LiveReviewStatus(ByVal canApply As Boolean, ByVal note As String, ByVal canBack As Boolean, Optional ByVal canPause As Boolean = True)
    mChoose.Enabled = canApply: mLiveNote.Caption = note: mBack.Enabled = canBack
    mIgnore.Enabled = canApply
    mPause.Enabled = canPause And mLocate.Enabled
    If Not canPause Then mStop.Enabled = False: mLocate.Enabled = False
    VfkRefreshButtons
End Sub

Public Sub LiveReviewDone(ByVal message As String)
    mPrompt.Caption = VfkText("Belge sonuna kadar inceleme tamamland\u0131.")
    mContext.Text = message: mResults.Clear
    mDetails.Text = VfkText("Yap\u0131lan eklemeler belgede kal\u0131r. Geri almak i\u00e7in Word\u0027e t\u0131klay\u0131p Ctrl+Z kullan\u0131n. Bitir ile pencereyi kapatabilirsiniz.")
    mChoose.Enabled = False: mStop.Enabled = False: mLocate.Enabled = False
    mPause.Enabled = False: mIgnore.Enabled = False
    mLiveNote.Caption = ""
    VfkRefreshButtons
End Sub

Public Sub SetSideWarning()
    Me.Width = 340: Me.Height = 460
    Me.ScrollBars = fmScrollBarsVertical: Me.ScrollHeight = 432
    mPrompt.Width = 308: mPrompt.Height = 50
    mDetails.Left = 12: mDetails.Width = 308: mDetails.Top = 70: mDetails.Height = 320
    mStop.Left = 12: mStop.Top = 398: mStop.Width = 98
    mChoose.Left = 116: mChoose.Top = 398: mChoose.Width = 98
    mCancel.Left = 222: mCancel.Top = 398: mCancel.Width = 98
End Sub

Private Sub mPause_Click()
    If mInventory Then
        VefeyatEklenti0341.Vefeyat0341UiInventoryPage 1
        Exit Sub
    End If
    VefeyatEklenti0341.Vefeyat0341UiReviewPause
End Sub

Private Sub mIgnore_Click()
    If mInventory Then
        VefeyatEklenti0341.Vefeyat0341UiInventoryIgnore LiveReviewChoice()
        Exit Sub
    End If
    VefeyatEklenti0341.Vefeyat0341UiReviewIgnore
End Sub

Private Sub mBack_Click()
    If mPlan Then PlanSelectAll False: Exit Sub
    If mInventory Then
        VefeyatEklenti0341.Vefeyat0341UiInventoryPage -1
        Exit Sub
    End If
    VefeyatEklenti0341.Vefeyat0341UiReviewMove -1
End Sub
Private Sub mLocate_Click()
    If mPlan Then
        ChosenIndex = mResults.ListIndex + 1: PlanAction = "choose": Me.Hide
        Exit Sub
    End If
    If mInventory Then
        VefeyatEklenti0341.Vefeyat0341UiInventoryLocate LiveReviewChoice()
        Exit Sub
    End If
    If mWorkbench Then
        VefeyatEklenti0341.Vefeyat0341UiToolLocate LiveReviewChoice()
        Exit Sub
    End If
    VefeyatEklenti0341.Vefeyat0341UiReviewLocate
End Sub
Private Sub mToggle_Click()
    If mInventory Then
        VfkToggleInventory
        Exit Sub
    End If
    Dim control As Object
    If Not mLiveCollapsed Then
        Set mLiveVisibility = CreateObject("Scripting.Dictionary")
        For Each control In Me.Controls
            mLiveVisibility(control.Name) = control.Visible: control.Visible = False
        Next control
        mLiveHeight = Me.Height: Me.ScrollTop = 0: Me.Height = 64
        mToggle.Left = 12: mToggle.Top = 6: mToggle.Width = 308
        mToggle.Caption = VfkText("Denetim penceresini geni\u015flet"): mToggle.Visible = True
        mLiveCollapsed = True
    Else
        Me.Height = mLiveHeight
        For Each control In Me.Controls
            If mLiveVisibility.Exists(control.Name) Then control.Visible = CBool(mLiveVisibility(control.Name))
        Next control
        mToggle.Left = 222: mToggle.Top = 350: mToggle.Width = 98
        mToggle.Caption = VfkText("K\u00fc\u00e7\u00fclt"): mLiveCollapsed = False
    End If
    VfkRefreshButtons
End Sub

Public Sub InitEarlier(ByVal message As String)
    Dim control As Object
    mEarlier = True
    mBrowse = False
    mProgress = False
    ChosenIndex = 0
    mResults.Visible = False
    For Each control In Me.Controls
        If control.Tag = "VefeyatHeading" Then control.Visible = False
    Next control
    mPrompt.Caption = VfkText("Bu ad daha \u00f6nce ge\u00e7ti. Yapmak istedi\u011finiz i\u015flemi se\u00e7in.")
    mDetails.Top = 68
    mDetails.Height = 328
    mDetails.Text = message
    mChoose.Visible = True
    mChoose.Enabled = True
    mChoose.Default = False
    mChoose.Caption = "Yine de ekle"
    mStop.Visible = True
    mStop.Caption = VfkText("\u00d6nceki ge\u00e7i\u015fe git")
    mCancel.Caption = VfkText("Vazge\u00e7")
    mCancel.Default = True
End Sub

Public Sub SetDiaMode()
    mDia = True
    mChoose.Caption = VfkText("D\u0130A maddesini a\u00e7")
    mPrompt.Caption = VfkText("D\u0130A maddesini a\u00e7mak istedi\u011finiz \u015fahs\u0131 se\u00e7in. Belge metni de\u011fi\u015ftirilmez.")
End Sub

Public Sub InitSearch()
    Dim control As Object
    Set mRows = New Collection
    ChosenIndex = 0
    SetBrowseOnly
    mPrompt.Caption = VfkText("Yerel verilerde aranacak ad\u0131 yaz\u0131n ve Ara'ya bas\u0131n.")
    Me.Height = 520
    For Each control In Me.Controls
        If control.Tag = "VefeyatHeading" Then control.Top = 92
    Next control
    mResults.Top = 112
    mDetails.Top = 280
    mCancel.Top = 446
    mChoose.Top = 446
    mQuery.Visible = True
    mSearch.Visible = True
    mChoose.Default = False
    mSearch.Default = True
End Sub

Private Sub mSearch_Click()
    If mPlan Then PlanSelectAll True: Exit Sub
    If mInventory Then
        VefeyatEklenti0341.Vefeyat0341UiInventoryRefresh
        Exit Sub
    End If
    Dim people As Collection, query As String
    On Error GoTo Failed
    query = Trim$(mQuery.Text)
    If Len(query) = 0 Then Exit Sub
    If Len(query) > 200 Then
        mPrompt.Caption = VfkText("En fazla 200 karakterlik bir ad aray\u0131n.")
        Exit Sub
    End If
    If mCatalogChoice Then
        Set people = VefeyatEklenti0341.Vefeyat0341UiFindInventoryPeople(query)
    Else
        Set people = VefeyatEklenti0341.Vefeyat0341UiFindPeople(query)
    End If
    InitCandidates people, query
    If mCatalogChoice Then
        mBrowse = False: mChoose.Visible = True: mChoose.Caption = VfkText("\u015eahs\u0131 se\u00e7")
        mChoose.Default = False: mSearch.Default = True
        VfkRefreshButtons
    Else
        SetBrowseOnly
    End If
    If people.Count = 0 Then mPrompt.Caption = VfkText("Bu aramada kay\u0131t bulunamad\u0131.")
    Exit Sub
Failed:
    mPrompt.Caption = VfkText("Arama tamamlanamad\u0131: ") & Err.Description
End Sub

Public Sub SetSpellingMode()
    mSpelling = True
    mChoose.Caption = VfkText("\u015eahs\u0131 se\u00e7")
    mPrompt.Caption = VfkText("Se\u00e7ili ad\u0131n hangi \u015fahsa ait oldu\u011funu do\u011frulay\u0131n. Sonraki ad\u0131mda kaynak yaz\u0131m\u0131n\u0131 \u00f6nerece\u011fiz; hen\u00fcz metin de\u011fi\u015ftirilmez.")
End Sub

Public Sub SetBrowseOnly()
    mBrowse = True
    mChoose.Visible = False
    mCancel.Caption = VfkText("Kapat")
    mPrompt.Caption = VfkText("Yerel kay\u0131tlar: ayr\u0131nt\u0131lar i\u00e7in bir ki\u015fi se\u00e7in.")
    VfkRefreshButtons
End Sub

Public Sub SetScanMode(ByVal index As Long, ByVal total As Long)
    mBatch = True
    mStop.Caption = VfkText("Taramay\u0131 durdur")
    mCancel.Caption = VfkText("Bu ad\u0131 atla")
    mStop.Visible = True
    mPrompt.Caption = CStr(index) & " / " & CStr(total) & VfkText(" farkl\u0131 ad: ") & mPrompt.Caption
End Sub

Public Sub SetReviewMode()
    mReview = True
    Set mReviewExcluded = CreateObject("Scripting.Dictionary")
    mChoose.Caption = VfkText("Listeyi uygula")
    mStop.Caption = VfkText("D\u0131\u015fla / geri al")
    mStop.Visible = True
    mPrompt.Caption = VfkText("Son kontrol: listede kalan adlara belge ayar\u0131na g\u00f6re tarih eklenecek. Bir ad\u0131 \u00e7ift t\u0131klayarak d\u0131\u015flay\u0131n veya geri al\u0131n. \u0130ptal ederseniz belge de\u011fi\u015fmez.")
    mResults_Change
End Sub

Public Function ReviewIncludes(ByVal index As Long) As Boolean
    ReviewIncludes = Not mReviewExcluded.Exists(CStr(index))
End Function

Public Sub SetExclusionMode()
    Dim control As Object
    mExclusionMode = True
    mChoose.Default = False
    For Each control In Me.Controls
        If control.Tag = "VefeyatHeading" Then control.Visible = False
    Next control
    mResults.ColumnCount = 1: mResults.ColumnWidths = "570 pt"
    mChoose.Caption = VfkText("Listeden \u00e7\u0131kar")
    mPrompt.Caption = VfkText("Yeniden sorulmas\u0131n\u0131 istedi\u011finiz ifadeyi se\u00e7ip Listeden \u00e7\u0131kar'a bas\u0131n. \u0130ptal listeyi de\u011fi\u015ftirmez.")
    mResults_Change
End Sub

Public Sub SetNameMode()
    mNameMode = True
    mChoose.Caption = VfkText("Bu yaz\u0131m\u0131 se\u00e7")
    mPrompt.Caption = VfkText("\u0130stedi\u011finiz ad bi\u00e7imini se\u00e7in: k\u0131sa yaz\u0131m, kaynak k\u0131sa ad\u0131, tam/uzun ad veya orijinal ad. Sonraki ad\u0131mda de\u011fi\u015fikli\u011fi onaylayabilirsiniz.")
    mResults_Change
End Sub

Private Sub mStop_Click()
    If mPlan Then
        If mResults.ListIndex >= 0 Then PlanSetSelection mResults.ListIndex + 1, Not PlanIncludes(mResults.ListIndex + 1)
        Exit Sub
    End If
    If mInventory Then
        VefeyatEklenti0341.Vefeyat0341UiInventoryAction LiveReviewChoice(), False
        Exit Sub
    End If
    If mWorkbench Then
        VefeyatEklenti0341.Vefeyat0341UiToolApply LiveReviewChoice(), True
        Exit Sub
    End If
    Dim key As String, row As Variant
    If mLiveReview Then
        VefeyatEklenti0341.Vefeyat0341UiReviewMove 1
        Exit Sub
    End If
    If mEarlier Then
        ChosenIndex = 2
        Me.Hide
        Exit Sub
    End If
    If mReview Then
        If mResults.ListIndex < 0 Then Exit Sub
        key = CStr(mResults.ListIndex + 1)
        row = mRows(mResults.ListIndex + 1)
        If mReviewExcluded.Exists(key) Then
            mReviewExcluded.Remove key
            mResults.List(mResults.ListIndex, 0) = CStr(row(0))
        Else
            mReviewExcluded.Add key, True
            mResults.List(mResults.ListIndex, 0) = "[-] " & CStr(row(0))
        End If
        Exit Sub
    End If
    StopRequested = True
    ChosenIndex = 0
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "Vefeyat"
    Me.Width = 640
    Me.Height = 490
    Me.Font.Name = "Segoe UI"
    Me.Font.Size = 10

    Set mPrompt = Me.Controls.Add("Forms.Label.1", "ResultPrompt", True)
    With mPrompt
        .Left = 16: .Top = 12: .Width = 600: .Height = 44
        .WordWrap = True
    End With

    Set mQuery = Me.Controls.Add("Forms.TextBox.1", "SearchQuery", True)
    With mQuery
        .Left = 16: .Top = 54: .Width = 480: .Height = 24
        .Visible = False
    End With
    Set mSearch = Me.Controls.Add("Forms.CommandButton.1", "SearchButton", True)
    With mSearch
        .Caption = "Ara"
        .Left = 510: .Top = 52: .Width = 106: .Height = 28
        .Visible = False
    End With

    AddHeading VfkText("K\u0131sa ad"), 16, 160
    AddHeading "Vefat tarihi", 176, 90
    AddHeading "Tam ad", 266, 350

    Set mResults = Me.Controls.Add("Forms.ListBox.1", "Results", True)
    With mResults
        .Left = 16: .Top = 82: .Width = 600: .Height = 155
        .ColumnCount = 3
        .ColumnWidths = "160 pt;90 pt;330 pt"
        .MultiSelect = fmMultiSelectSingle
        .TabIndex = 0
    End With

    Set mDetails = Me.Controls.Add("Forms.TextBox.1", "Details", True)
    With mDetails
        .Left = 16: .Top = 250: .Width = 600: .Height = 150
        .MultiLine = True
        .WordWrap = True
        .ScrollBars = fmScrollBarsVertical
        .Locked = True
        .TabIndex = 1
    End With

    Set mChoose = Me.Controls.Add("Forms.CommandButton.1", "ChooseButton", True)
    With mChoose
        .Caption = VfkText("Do\u011frula ve ekle")
        .Left = 370: .Top = 416: .Width = 132: .Height = 28
        .Default = True
        .Enabled = False
        .TabIndex = 2
    End With

    Set mStop = Me.Controls.Add("Forms.CommandButton.1", "StopScanButton", True)
    With mStop
        .Caption = VfkText("Taramay\u0131 durdur")
        .Left = 16: .Top = 416: .Width = 145: .Height = 28
        .Visible = False
    End With

    Set mCancel = Me.Controls.Add("Forms.CommandButton.1", "CancelButton", True)
    With mCancel
        .Caption = VfkText("\u0130ptal")
        .Left = 514: .Top = 416: .Width = 102: .Height = 28
        .Cancel = True
        .TabIndex = 3
    End With
    VfkApplyTheme
End Sub

Private Sub UserForm_Activate()
    Dim title As String
    title = VfkText("Vefey\u00e2t: \u015fahs\u0131 do\u011frulay\u0131n")
    If mLiveReview Then title = VfkText("Vefey\u00e2t: metin \u00fczerinde denetim")
    If mEarlier Then title = VfkText("Vefey\u00e2t: daha \u00f6nce ge\u00e7ti")
    If mDia Then title = VfkText("Vefey\u00e2t: D\u0130A maddesini a\u00e7")
    If mSpelling Then title = VfkText("Vefey\u00e2t: isim yaz\u0131m\u0131n\u0131 denetle")
    If mReview Then title = VfkText("Vefey\u00e2t: belge taramas\u0131n\u0131 onaylay\u0131n")
    If mExclusionMode Then title = VfkText("Vefey\u00e2t: belge istisnas\u0131n\u0131 kald\u0131r")
    If mNameMode Then title = VfkText("Vefey\u00e2t: ad bi\u00e7imini se\u00e7in")
    If mProgress Then title = VfkText("Vefey\u00e2t: belge taramas\u0131")
    If mWorkbench Then title = VfkText("Vefey\u00e2t: belge ara\u00e7lar\u0131")
    If mInventory Then title = VfkText("Vefey\u00e2t: \u015fah\u0131s d\u00f6k\u00fcm\u00fc")
    If mPlan Then title = VfkText("Vefey\u00e2t: toplu d\u00fczeltme \u00f6nizlemesi")
    If mRegistryMode Then title = VfkText("Vefey\u00e2t: dinamik \u015fah\u0131s kay\u0131tlar\u0131")
    If mCatalogChoice Then title = VfkText("Vefey\u00e2t: katalogdan \u015fah\u0131s se\u00e7in")
    VfkApplyTheme
    VfkSafeTitle title
    If mProgress Then
        If Not mProgressStarted Then
            mProgressStarted = True
            Me.Repaint
            VefeyatEklenti0341.Vefeyat0341UiRunAudit
        End If
        Exit Sub
    End If
    If mQuery.Visible Then mQuery.SetFocus
End Sub

' An active window can be the Word owner. Never rename an arbitrary window.
Private Function VfkFormClass(ByVal value As String) As Boolean
    Select Case LCase$(value)
        Case "thunderdframe", "thunderxframe"
            VfkFormClass = True
    End Select
End Function

Private Sub VfkSafeTitle(ByVal title As String)
    Dim buffer As String, length As Long
#If VBA7 Then
    Dim window As LongPtr
#Else
    Dim window As Long
#End If
    window = VfkActiveWindow()
    buffer = String$(64, vbNullChar)
    length = VfkClassNameW(window, StrPtr(buffer), 64)
    If length = 0 Then Exit Sub
    If Not VfkFormClass(Left$(buffer, length)) Then Exit Sub
    Call VfkSetWindowTextW(window, StrPtr(title))
End Sub

Public Sub InitProgress()
    Dim control As Object
    mProgress = True
    mProgressStarted = False
    mAllowCancel = True
    StopRequested = False
    Me.Width = 490
    Me.Height = 225
    For Each control In Me.Controls
        control.Visible = False
    Next control
    With mPrompt
        .Visible = True
        .Left = 16: .Top = 14: .Width = 448: .Height = 60
        .Caption = VfkText("Tarama haz\u0131rlan\u0131yor. Son onay\u0131n\u0131za kadar belge de\u011fi\u015ftirilmez.")
    End With
    Set mTrack = Me.Controls.Add("Forms.Label.1", "ProgressTrack", True)
    With mTrack
        .Left = 16: .Top = 82: .Width = 448: .Height = 12
        .BackColor = RGB(225, 230, 235)
    End With
    Set mFill = Me.Controls.Add("Forms.Label.1", "ProgressFill", True)
    With mFill
        .Left = 16: .Top = 82: .Width = 0: .Height = 12
        .BackColor = RGB(18, 100, 170)
    End With
    Set mNotice = Me.Controls.Add("Forms.Label.1", "ProgressNotice", True)
    With mNotice
        .Left = 16: .Top = 102: .Width = 448: .Height = 40
        .WordWrap = True
        .Caption = VfkText("Son onay\u0131n\u0131za kadar belgeniz de\u011fi\u015ftirilmez. Durdur ile \u00e7\u0131kabilirsiniz.")
    End With
    With mCancel
        .Visible = True
        .Left = 340: .Top = 158: .Width = 124
        .Caption = VfkText("Taramay\u0131 durdur")
    End With
End Sub

Public Sub UpdateProgress(ByVal stage As String, ByVal completed As Long, ByVal total As Long, ByVal allowCancel As Boolean)
    Dim fraction As Double
    mAllowCancel = allowCancel
    mCancel.Enabled = allowCancel
    If allowCancel Then
        mNotice.Caption = VfkText("Son onay\u0131n\u0131za kadar belgeniz de\u011fi\u015ftirilmez. Durdur ile \u00e7\u0131kabilirsiniz.")
        mCancel.Caption = VfkText("Taramay\u0131 durdur")
    Else
        mNotice.Caption = VfkText("Onaylad\u0131\u011f\u0131n\u0131z tarihler ekleniyor. \u0130\u015flem bitince belgenize d\u00f6nebilirsiniz.")
        mCancel.Caption = VfkText("Uygulan\u0131yor...")
    End If
    mPrompt.Caption = stage
    If total > 0 Then
        fraction = CDbl(completed) / CDbl(total)
        If fraction < 0 Then fraction = 0
        If fraction > 1 Then fraction = 1
        mFill.Width = CSng(fraction * mTrack.Width)
        mPrompt.Caption = stage & vbCrLf & CStr(completed) & " / " & CStr(total) & "  (%" & CStr(CLng(fraction * 100)) & ")"
    End If
    VfkRefreshButtons
End Sub

Private Sub mResults_Change()
    If mPlan Then
        If mResults.ListIndex >= 0 Then
            mDetails.Text = CStr(mRows(mResults.ListIndex + 1)(4)) & vbCrLf & vbCrLf & mPlanCoverage
            mStop.Enabled = (CStr(mRows(mResults.ListIndex + 1)(2)) <> "info")
        End If
        Exit Sub
    End If
    Dim row As Variant, description As String
    If mInventoryUpdating Then Exit Sub
    If mInventory Then
        If mResults.ListIndex >= 0 Then VefeyatEklenti0341.Vefeyat0341UiInventorySelection mResults.ListIndex + 1
        Exit Sub
    End If
    If mRows Is Nothing Then Exit Sub
    If mResults.ListIndex < 0 Then
        mDetails.Text = vbNullString
        mChoose.Enabled = False
        VfkRefreshButtons
        Exit Sub
    End If
    row = mRows(mResults.ListIndex + 1)
    If mWorkbench Then
        mDetails.Text = CStr(row(4))
        mChoose.Enabled = mToolChanges And CStr(row(2)) <> "none"
        mStop.Enabled = mToolChanges And CStr(row(3)) = "unlink"
        VfkRefreshButtons
        Exit Sub
    End If
    description = Trim$(CStr(row(4)))
    If Len(description) = 0 Or description = "---" Then
        description = VfkText("Bu kay\u0131tta k\u0131sa tan\u0131t\u0131m bulunmuyor.")
    End If
    mDetails.Text = VfkText("K\u0131sa tan\u0131t\u0131m: ") & description & vbCrLf & vbCrLf & _
        VfkText("K\u0131sa ad: ") & CStr(row(0)) & vbCrLf & _
        "Tam ad: " & CStr(row(1)) & vbCrLf & _
        "Orijinal ad: " & CStr(row(2)) & vbCrLf & _
        "Vefat tarihi: " & CStr(row(3)) & vbCrLf & vbCrLf & _
        "Kaynak: " & CStr(row(5))
    If mReview Then
        mDetails.Text = VfkText("Belgedeki ad: ") & CStr(row(0)) & vbCrLf & _
            VfkText("Se\u00e7ilen \u015fah\u0131s: ") & CStr(row(1)) & vbCrLf & _
            "Tam ad: " & CStr(row(2)) & vbCrLf & "Vefat tarihi: " & CStr(row(3)) & vbCrLf & _
            description & vbCrLf & "Kaynak: " & CStr(row(5))
    End If
    If mNameMode Then
        mDetails.Text = VfkText("\u00d6nerilen yaz\u0131m: ") & CStr(row(0)) & vbCrLf & _
            VfkText("Ad bi\u00e7imi: ") & CStr(row(1)) & vbCrLf & description & vbCrLf & "Kaynak: " & CStr(row(5))
    End If
    If Not mNameMode And Not mExclusionMode And Not mReview Then mDetails.Text = mDetails.Text & vbCrLf & vbCrLf & VefeyatEklenti0341.Vefeyat0341UiExplain(row, mExplainQuery)
    If mExclusionMode Then mDetails.Text = CStr(row(4))
    If Not mIdentityProof Is Nothing Then
        mDetails.Text = mDetails.Text & vbCrLf & vbCrLf & VefeyatEklenti0341.Vefeyat0341UiContextExplain(mIdentityProof, row)
    End If
    mChoose.Enabled = True
    If mLiveReview Then VefeyatEklenti0341.Vefeyat0341UiReviewChoice
    VfkRefreshButtons
End Sub

Private Sub mResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If mPlan Then mStop_Click: Exit Sub
    If mInventory Then
        mLocate_Click
        Exit Sub
    End If
    If mWorkbench Then
        mLocate_Click
        Exit Sub
    End If
    If mLiveReview Then
        VefeyatEklenti0341.Vefeyat0341UiReviewLocate
        Exit Sub
    End If
    If mReview Then
        mStop_Click
    Else
        mChoose_Click
    End If
End Sub

Private Sub mChoose_Click()
    If mPlan Then
        ChosenIndex = 1: PlanAction = "apply": Me.Hide
        Exit Sub
    End If
    If mInventory Then
        VefeyatEklenti0341.Vefeyat0341UiInventoryAction LiveReviewChoice(), True
        Exit Sub
    End If
    If mWorkbench Then
        VefeyatEklenti0341.Vefeyat0341UiToolApply LiveReviewChoice(), False
        Exit Sub
    End If
    If mLiveReview Then
        VefeyatEklenti0341.Vefeyat0341UiReviewApply LiveReviewChoice()
        Exit Sub
    End If
    If mEarlier Then
        ChosenIndex = 1
        Me.Hide
        Exit Sub
    End If
    If mBrowse Then Exit Sub
    If mResults.ListIndex < 0 Then Exit Sub
    ChosenIndex = mResults.ListIndex + 1
    Me.Hide
End Sub

Private Sub mCancel_Click()
    If mInventory Then
        Unload Me
        Exit Sub
    End If
    If mWorkbench Then
        Unload Me
        Exit Sub
    End If
    If mLiveReview Then
        Unload Me
        Exit Sub
    End If
    If mProgress Then
        If mAllowCancel Then
            StopRequested = True
            mCancel.Enabled = False
            mPrompt.Caption = VfkText("Tarama durduruluyor; belge de\u011fi\u015ftirilmedi.")
            VfkRefreshButtons
        End If
        Exit Sub
    End If
    ChosenIndex = 0
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If mInventory Then
        VefeyatEklenti0341.Vefeyat0341UiInventoryClose True, Me
        Exit Sub
    End If
    If mWorkbench Then
        VefeyatEklenti0341.Vefeyat0341UiToolClose True
        Exit Sub
    End If
    If mLiveReview Then
        If mLiveClosing Then Exit Sub
        mLiveClosing = True
        ' Let MSForms finish its own close; never recursively unload this form.
        VefeyatEklenti0341.Vefeyat0341UiReviewClose True
        Exit Sub
    End If
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        If mProgress Then
            mCancel_Click
            Exit Sub
        End If
        If mBatch Then StopRequested = True
        ChosenIndex = 0
        Me.Hide
    End If
End Sub

Private Sub AddHeading(ByVal text As String, ByVal x As Single, ByVal size As Single)
    Dim heading As MSForms.Label
    Set heading = Me.Controls.Add("Forms.Label.1")
    With heading
        .Caption = text
        .Tag = "VefeyatHeading"
        .Left = x
        .Top = 62
        .Width = size
        .Height = 16
        .Font.Bold = True
    End With
End Sub

' Decode ASCII source labels without depending on another module's helpers.
Private Function VfkText(ByVal escaped As String) As String
    Dim i As Long, j As Long, code As Long, digit As Long
    Dim result As String, valid As Boolean, part As String
    i = 1
    Do While i <= Len(escaped)
        If Mid$(escaped, i, 2) = "\u" And i + 5 <= Len(escaped) Then
            part = Mid$(escaped, i + 2, 4)
            code = 0
            valid = True
            For j = 1 To 4
                digit = InStr(1, "0123456789abcdef", LCase$(Mid$(part, j, 1)), vbBinaryCompare) - 1
                If digit < 0 Then
                    valid = False
                    Exit For
                End If
                code = code * 16 + digit
            Next j
            If valid Then
                If code > 32767 Then code = code - 65536
                result = result & ChrW(code)
                i = i + 6
            Else
                result = result & Mid$(escaped, i, 1)
                i = i + 1
            End If
        Else
            result = result & Mid$(escaped, i, 1)
            i = i + 1
        End If
    Loop
    VfkText = result
End Function


' Presentation only. Never set selection, text, values, enablement, visibility,
' focus, default actions, geometry or document properties from these helpers.
Private Function VfkColor(ByVal token As String) As Long
    Select Case token
        Case "canvas": VfkColor = RGB(245, 247, 250)
        Case "paper": VfkColor = RGB(255, 255, 255)
        Case "context": VfkColor = RGB(248, 250, 252)
        Case "ink": VfkColor = RGB(30, 41, 59)
        Case "muted": VfkColor = RGB(100, 116, 139)
        Case "line": VfkColor = RGB(213, 222, 233)
        Case "brand": VfkColor = RGB(37, 99, 235)
        Case "warning": VfkColor = RGB(255, 247, 237)
        Case "warningInk": VfkColor = RGB(154, 52, 18)
        Case "danger": VfkColor = RGB(255, 241, 242)
        Case "dangerInk": VfkColor = RGB(185, 28, 28)
        Case "disabled": VfkColor = RGB(234, 239, 245)
        Case "disabledInk": VfkColor = RGB(100, 116, 139)
    End Select
End Function

Private Sub VfkStyleControl(ByVal control As Object)
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo CosmeticOnly
    control.Font.Name = "Segoe UI"
    control.Font.Size = 10
    Select Case TypeName(control)
        Case "TextBox", "ListBox"
            control.SpecialEffect = fmSpecialEffectFlat
            control.BorderStyle = fmBorderStyleSingle
            control.BorderColor = VfkColor("line")
            control.BackColor = VfkColor("paper")
            control.ForeColor = VfkColor("ink")
            If control.Name = "LiveContext" Then control.BackColor = VfkColor("context")
        Case "CommandButton"
            control.BackColor = VfkColor("paper")
            control.ForeColor = VfkColor("ink")
        Case "Label"
            control.BackStyle = fmBackStyleTransparent
            control.ForeColor = VfkColor("muted")
            control.Font.Size = 9
            If control.Tag = "VefeyatHeading" Or control.Tag = "VefeyatInventoryHeading" Then control.Font.Bold = True
            Select Case control.Name
                Case "ResultPrompt"
                    control.Font.Size = 10
                    control.Font.Bold = True
                    control.ForeColor = VfkColor("ink")
                Case "LiveNote": control.Font.Size = 8
                Case "ProgressTrack"
                    control.BackStyle = fmBackStyleOpaque
                    control.BackColor = VfkColor("line")
                Case "ProgressFill"
                    control.BackStyle = fmBackStyleOpaque
                    control.BackColor = VfkColor("brand")
            End Select
    End Select
    Exit Sub
CosmeticOnly:
    ' An unavailable native style property must not interrupt a user action.
    Exit Sub
End Sub

Private Sub VfkPaintButton(ByVal button As Object, ByVal role As String)
    If Err.Number <> 0 Then Exit Sub
    If button Is Nothing Then Exit Sub
    On Error GoTo CosmeticOnly
    button.BackColor = VfkColor("paper")
    button.ForeColor = VfkColor("ink")
    button.Font.Bold = False
    If Not button.Enabled Then
        button.BackColor = VfkColor("disabled")
        button.ForeColor = VfkColor("disabledInk")
    Else
        Select Case role
            Case "primary"
                button.BackColor = VfkColor("brand")
                button.ForeColor = VfkColor("paper")
                button.Font.Bold = True
            Case "warning"
                button.BackColor = VfkColor("warning")
                button.ForeColor = VfkColor("warningInk")
                button.Font.Bold = True
            Case "danger"
                button.BackColor = VfkColor("danger")
                button.ForeColor = VfkColor("dangerInk")
        End Select
    End If
    Exit Sub
CosmeticOnly:
    Exit Sub
End Sub

Private Sub VfkRefreshButtons()
    If Err.Number <> 0 Then Exit Sub
    If mEarlier Then
        VfkPaintButton mChoose, "warning"
    Else
        VfkPaintButton mChoose, "primary"
    End If
    VfkPaintButton mSearch, "primary"
    VfkPaintButton mCancel, "secondary"
    If mWorkbench Or mProgress Then
        VfkPaintButton mStop, "danger"
    Else
        VfkPaintButton mStop, "secondary"
    End If
    If mProgress Then VfkPaintButton mCancel, "danger"
    If mInventory Then
        VfkPaintButton mChoose, "secondary"
        VfkPaintButton mStop, "primary"
        VfkPaintButton mSearch, "secondary"
        VfkPaintButton mIgnore, "warning"
    End If
    VfkSyncButtonFaces
End Sub

Private Sub VfkApplyTheme()
    Dim control As Object
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo CosmeticOnly
    Me.BackColor = VfkColor("canvas")
    For Each control In Me.Controls
        VfkStyleControl control
    Next control
    VfkRefreshButtons
    Exit Sub
CosmeticOnly:
    Exit Sub
End Sub

' Two windowless labels draw a flat button; its native command remains visible
' underneath for Tab, accelerators, Default/Cancel and assistive technologies.
' No window subclassing, timers, image files, new references or document access.
Private Sub VfkSyncOneButton(ByVal key As String, ByVal button As MSForms.CommandButton, ByRef face As MSForms.Label, ByRef textFace As MSForms.Label)
    Dim back As Long, ink As Long, edge As Long
    If Err.Number <> 0 Then Exit Sub
    If button Is Nothing Then Exit Sub
    On Error GoTo NativeFallback
    If Not VfkFlatSurfacesEnabled Then GoTo NativeFallback
    If Not button.Visible Then
        If Not face Is Nothing Then face.Visible = False
        If Not textFace Is Nothing Then textFace.Visible = False
        Exit Sub
    End If
    If face Is Nothing Then Set face = Me.Controls.Add("Forms.Label.1", "VfkFace_" & key, False)
    If textFace Is Nothing Then Set textFace = Me.Controls.Add("Forms.Label.1", "VfkText_" & key, False)
    back = button.BackColor: ink = button.ForeColor: edge = VfkColor("line")
    If Not button.Enabled Then
        back = VfkColor("disabled"): ink = VfkColor("disabledInk")
    Else
        If button.Default Then edge = VfkColor("brand")
        If mVfkFocus = key Then
            edge = VfkColor("ink")
            If back = VfkColor("paper") Then back = RGB(239, 246, 255)
        End If
        If mVfkPressed = key Then
            If back = VfkColor("brand") Then
                back = RGB(29, 78, 216)
            Else
                back = RGB(219, 234, 254)
            End If
        End If
    End If
    With face
        .Caption = "": .Tag = "VefeyatButtonSurface"
        .Left = button.Left: .Top = button.Top: .Width = button.Width: .Height = button.Height
        .BackStyle = fmBackStyleOpaque: .BackColor = back
        .SpecialEffect = fmSpecialEffectFlat: .BorderStyle = fmBorderStyleSingle: .BorderColor = edge
        .Enabled = button.Enabled: .ControlTipText = button.ControlTipText
        .Visible = True: .ZOrder 0
    End With
    With textFace
        .Tag = "VefeyatButtonSurface": .Caption = button.Caption
        .Left = button.Left + 2: .Width = button.Width - 4
        .Height = 16: .Top = button.Top + (button.Height - 16) / 2
        .BackStyle = fmBackStyleTransparent: .ForeColor = ink
        .SpecialEffect = fmSpecialEffectFlat: .BorderStyle = fmBorderStyleNone
        .Font.Name = "Segoe UI": .Font.Size = 10: .Font.Bold = button.Font.Bold
        .TextAlign = fmTextAlignCenter: .WordWrap = False
        .Enabled = button.Enabled: .ControlTipText = button.ControlTipText
        .Visible = True: .ZOrder 0
    End With
    Exit Sub
NativeFallback:
    ' A partial decorative layer must never cover a usable native command.
    On Error Resume Next
    If Not face Is Nothing Then face.Visible = False
    If Not textFace Is Nothing Then textFace.Visible = False
End Sub

Private Sub VfkSyncButtonFaces()
    If Err.Number <> 0 Then Exit Sub
    VfkSyncOneButton "choose", mChoose, mChooseFace, mChooseText
    VfkSyncOneButton "cancel", mCancel, mCancelFace, mCancelText
    VfkSyncOneButton "search", mSearch, mSearchFace, mSearchText
    VfkSyncOneButton "back", mBack, mBackFace, mBackText
    VfkSyncOneButton "locate", mLocate, mLocateFace, mLocateText
    VfkSyncOneButton "toggle", mToggle, mToggleFace, mToggleText
    VfkSyncOneButton "pause", mPause, mPauseFace, mPauseText
    VfkSyncOneButton "ignore", mIgnore, mIgnoreFace, mIgnoreText
    VfkSyncOneButton "stop", mStop, mStopFace, mStopText
End Sub

Private Function VfkButton(ByVal key As String) As MSForms.CommandButton
    Select Case key
        Case "choose": Set VfkButton = mChoose
        Case "cancel": Set VfkButton = mCancel
        Case "search": Set VfkButton = mSearch
        Case "back": Set VfkButton = mBack
        Case "locate": Set VfkButton = mLocate
        Case "toggle": Set VfkButton = mToggle
        Case "pause": Set VfkButton = mPause
        Case "ignore": Set VfkButton = mIgnore
        Case "stop": Set VfkButton = mStop
    End Select
End Function

Private Sub VfkButtonFocus(ByVal key As String, ByVal focused As Boolean)
    If Err.Number <> 0 Then Exit Sub
    If focused Then
        mVfkFocus = key
    Else
        If mVfkFocus = key Then mVfkFocus = ""
        If mVfkPressed = key Then mVfkPressed = ""
    End If
    VfkSyncButtonFaces
End Sub

Private Sub VfkButtonPress(ByVal key As String, ByVal pressed As Boolean)
    Dim button As MSForms.CommandButton
    If Err.Number <> 0 Then Exit Sub
    Set button = VfkButton(key)
    If button Is Nothing Then Exit Sub
    If Not button.Enabled Or Not button.Visible Then Exit Sub
    If pressed Then
        mVfkArmed = ""
        mVfkPressed = key
    Else
        mVfkPressed = ""
    End If
    VfkSyncButtonFaces
End Sub

Private Sub VfkButtonRelease(ByVal key As String, ByVal X As Single, ByVal Y As Single, ByVal onText As Boolean)
    Dim button As MSForms.CommandButton, localX As Single, localY As Single, wasPressed As Boolean
    If Err.Number <> 0 Then Exit Sub
    Set button = VfkButton(key)
    wasPressed = mVfkPressed = key
    mVfkPressed = "": mVfkArmed = ""
    If Not button Is Nothing Then
        If wasPressed And button.Enabled And button.Visible Then
            localX = X: localY = Y
            If onText Then localX = X + 2: localY = Y + (button.Height - 16) / 2
            If localX >= 0 And localX < button.Width And localY >= 0 And localY < button.Height Then mVfkArmed = key
        End If
    End If
    VfkSyncButtonFaces
End Sub

Private Sub VfkButtonClick(ByVal key As String)
    Dim button As MSForms.CommandButton, errorCode As Long, errorText As String, errorSource As String
    If mVfkClickBusy Then Exit Sub
    If mVfkArmed <> key Then Exit Sub
    mVfkArmed = ""
    Set button = VfkButton(key)
    If button Is Nothing Then Exit Sub
    If Not Me.Visible Then Exit Sub
    If Not button.Enabled Or Not button.Visible Then Exit Sub
    On Error GoTo Failed
    mVfkClickBusy = True
    button.SetFocus
    VfkButtonFocus key, True
    ' Recheck after focus events; do not bypass the original command gate.
    If button.Enabled And button.Visible And Me.Visible Then
        Select Case key
            Case "choose": mChoose_Click
            Case "cancel": mCancel_Click
            Case "search": mSearch_Click
            Case "back": mBack_Click
            Case "locate": mLocate_Click
            Case "toggle": mToggle_Click
            Case "pause": mPause_Click
            Case "ignore": mIgnore_Click
            Case "stop": mStop_Click
        End Select
    End If
    mVfkClickBusy = False
    Exit Sub
Failed:
    errorCode = Err.Number: errorText = Err.Description: errorSource = Err.Source
    mVfkClickBusy = False
    Err.Raise errorCode, errorSource, errorText
End Sub

Private Sub mChooseFace_Click()
    VfkButtonClick "choose"
End Sub

Private Sub mChooseFace_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonPress "choose", True
End Sub

Private Sub mChooseFace_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonRelease "choose", X, Y, False
End Sub

Private Sub mChooseText_Click()
    VfkButtonClick "choose"
End Sub

Private Sub mChooseText_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonPress "choose", True
End Sub

Private Sub mChooseText_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonRelease "choose", X, Y, True
End Sub

Private Sub mCancelFace_Click()
    VfkButtonClick "cancel"
End Sub

Private Sub mCancelFace_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonPress "cancel", True
End Sub

Private Sub mCancelFace_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonRelease "cancel", X, Y, False
End Sub

Private Sub mCancelText_Click()
    VfkButtonClick "cancel"
End Sub

Private Sub mCancelText_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonPress "cancel", True
End Sub

Private Sub mCancelText_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonRelease "cancel", X, Y, True
End Sub

Private Sub mSearchFace_Click()
    VfkButtonClick "search"
End Sub

Private Sub mSearchFace_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonPress "search", True
End Sub

Private Sub mSearchFace_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonRelease "search", X, Y, False
End Sub

Private Sub mSearchText_Click()
    VfkButtonClick "search"
End Sub

Private Sub mSearchText_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonPress "search", True
End Sub

Private Sub mSearchText_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonRelease "search", X, Y, True
End Sub

Private Sub mBackFace_Click()
    VfkButtonClick "back"
End Sub

Private Sub mBackFace_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonPress "back", True
End Sub

Private Sub mBackFace_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonRelease "back", X, Y, False
End Sub

Private Sub mBackText_Click()
    VfkButtonClick "back"
End Sub

Private Sub mBackText_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonPress "back", True
End Sub

Private Sub mBackText_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonRelease "back", X, Y, True
End Sub

Private Sub mLocateFace_Click()
    VfkButtonClick "locate"
End Sub

Private Sub mLocateFace_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonPress "locate", True
End Sub

Private Sub mLocateFace_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonRelease "locate", X, Y, False
End Sub

Private Sub mLocateText_Click()
    VfkButtonClick "locate"
End Sub

Private Sub mLocateText_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonPress "locate", True
End Sub

Private Sub mLocateText_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonRelease "locate", X, Y, True
End Sub

Private Sub mToggleFace_Click()
    VfkButtonClick "toggle"
End Sub

Private Sub mToggleFace_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonPress "toggle", True
End Sub

Private Sub mToggleFace_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonRelease "toggle", X, Y, False
End Sub

Private Sub mToggleText_Click()
    VfkButtonClick "toggle"
End Sub

Private Sub mToggleText_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonPress "toggle", True
End Sub

Private Sub mToggleText_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonRelease "toggle", X, Y, True
End Sub

Private Sub mPauseFace_Click()
    VfkButtonClick "pause"
End Sub

Private Sub mPauseFace_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonPress "pause", True
End Sub

Private Sub mPauseFace_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonRelease "pause", X, Y, False
End Sub

Private Sub mPauseText_Click()
    VfkButtonClick "pause"
End Sub

Private Sub mPauseText_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonPress "pause", True
End Sub

Private Sub mPauseText_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonRelease "pause", X, Y, True
End Sub

Private Sub mIgnoreFace_Click()
    VfkButtonClick "ignore"
End Sub

Private Sub mIgnoreFace_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonPress "ignore", True
End Sub

Private Sub mIgnoreFace_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonRelease "ignore", X, Y, False
End Sub

Private Sub mIgnoreText_Click()
    VfkButtonClick "ignore"
End Sub

Private Sub mIgnoreText_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonPress "ignore", True
End Sub

Private Sub mIgnoreText_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonRelease "ignore", X, Y, True
End Sub

Private Sub mStopFace_Click()
    VfkButtonClick "stop"
End Sub

Private Sub mStopFace_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonPress "stop", True
End Sub

Private Sub mStopFace_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonRelease "stop", X, Y, False
End Sub

Private Sub mStopText_Click()
    VfkButtonClick "stop"
End Sub

Private Sub mStopText_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonPress "stop", True
End Sub

Private Sub mStopText_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then VfkButtonRelease "stop", X, Y, True
End Sub

Private Sub mChoose_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 9 Then VfkButtonFocus "choose", False
    If KeyCode = 32 Then VfkButtonPress "choose", True
End Sub

Private Sub mChoose_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    VfkButtonPress "choose", False
    VfkButtonFocus "choose", True
End Sub

Private Sub mCancel_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 9 Then VfkButtonFocus "cancel", False
    If KeyCode = 32 Then VfkButtonPress "cancel", True
End Sub

Private Sub mCancel_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    VfkButtonPress "cancel", False
    VfkButtonFocus "cancel", True
End Sub

Private Sub mSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 9 Then VfkButtonFocus "search", False
    If KeyCode = 32 Then VfkButtonPress "search", True
End Sub

Private Sub mSearch_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    VfkButtonPress "search", False
    VfkButtonFocus "search", True
End Sub

Private Sub mBack_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 9 Then VfkButtonFocus "back", False
    If KeyCode = 32 Then VfkButtonPress "back", True
End Sub

Private Sub mBack_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    VfkButtonPress "back", False
    VfkButtonFocus "back", True
End Sub

Private Sub mLocate_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 9 Then VfkButtonFocus "locate", False
    If KeyCode = 32 Then VfkButtonPress "locate", True
End Sub

Private Sub mLocate_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    VfkButtonPress "locate", False
    VfkButtonFocus "locate", True
End Sub

Private Sub mToggle_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 9 Then VfkButtonFocus "toggle", False
    If KeyCode = 32 Then VfkButtonPress "toggle", True
End Sub

Private Sub mToggle_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    VfkButtonPress "toggle", False
    VfkButtonFocus "toggle", True
End Sub

Private Sub mPause_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 9 Then VfkButtonFocus "pause", False
    If KeyCode = 32 Then VfkButtonPress "pause", True
End Sub

Private Sub mPause_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    VfkButtonPress "pause", False
    VfkButtonFocus "pause", True
End Sub

Private Sub mIgnore_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 9 Then VfkButtonFocus "ignore", False
    If KeyCode = 32 Then VfkButtonPress "ignore", True
End Sub

Private Sub mIgnore_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    VfkButtonPress "ignore", False
    VfkButtonFocus "ignore", True
End Sub

Private Sub mStop_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 9 Then VfkButtonFocus "stop", False
    If KeyCode = 32 Then VfkButtonPress "stop", True
End Sub

Private Sub mStop_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    VfkButtonPress "stop", False
    VfkButtonFocus "stop", True
End Sub

Private Sub VfkClearButtonFocus()
    If Err.Number <> 0 Then Exit Sub
    mVfkFocus = "": mVfkPressed = "": mVfkArmed = ""
    VfkSyncButtonFaces
End Sub

Private Sub mQuery_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    VfkClearButtonFocus
End Sub

Private Sub mQuery_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    VfkClearButtonFocus
End Sub

Private Sub mDetails_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    VfkClearButtonFocus
End Sub

Private Sub mDetails_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    VfkClearButtonFocus
End Sub

Private Sub mContext_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    VfkClearButtonFocus
End Sub

Private Sub mContext_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    VfkClearButtonFocus
End Sub

Private Sub mResults_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    VfkClearButtonFocus
End Sub

Private Sub mResults_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    VfkClearButtonFocus
End Sub



Public Sub InitCatalogChoice(ByVal query As String)
    mCatalogChoice = True
    InitSearch
    mQuery.Text = query
    mSearch_Click
End Sub

Public Function ChosenCatalogPerson() As Variant
    If mRows Is Nothing Then Exit Function
    If ChosenIndex >= 1 And ChosenIndex <= mRows.Count Then ChosenCatalogPerson = mRows(ChosenIndex)
End Function

Public Sub InitInventory(ByVal rows As Collection, ByVal heading As String, ByVal canEdit As Boolean, Optional ByVal selected As Long = 1)
    Dim control As Object, row As Variant, i As Long, number As Long, detail As String
    mInventoryUpdating = True
    On Error GoTo Failed
    If Not mInventory Then
        mInventory = True: mWorkbench = False: mLiveReview = False: mProgress = False
        Me.Width = 720: Me.Height = 544
        mPrompt.Left = 12: mPrompt.Top = 12: mPrompt.Width = 680: mPrompt.Height = 46
        For Each control In Me.Controls
            If control.Tag = "VefeyatHeading" Then control.Visible = False
        Next control
        VfkInventoryHeading VfkText("Metindeki ad"), 12, 180
        VfkInventoryHeading VfkText("D\u0130A kar\u015f\u0131l\u0131\u011f\u0131"), 192, 165
        VfkInventoryHeading VfkText("\u00d6l\u00fcm tarihi"), 357, 110
        VfkInventoryHeading "Durum", 467, 175
        VfkInventoryHeading "Yer", 642, 50
        mResults.Left = 12: mResults.Top = 84: mResults.Width = 680: mResults.Height = 240
        mResults.ColumnCount = 5: mResults.ColumnWidths = "180 pt;165 pt;110 pt;175 pt;50 pt"
        mDetails.Left = 12: mDetails.Top = 336: mDetails.Width = 680: mDetails.Height = 82
        mChoose.Left = 12: mChoose.Top = 430: mChoose.Width = 150: mChoose.Height = 30
        mChoose.Caption = VfkText("\u0130smi de\u011fi\u015ftir"): mChoose.Default = False
        mStop.Left = 174: mStop.Top = 430: mStop.Width = 140: mStop.Height = 30
        mStop.Caption = "Tarih ekle": mStop.Visible = True
        Set mLocate = Me.Controls.Add("Forms.CommandButton.1", "InventoryLocate", True)
        mLocate.Left = 326: mLocate.Top = 430: mLocate.Width = 140: mLocate.Height = 30
        mLocate.Caption = VfkText("Metinde g\u00f6ster")
        Set mToggle = Me.Controls.Add("Forms.CommandButton.1", "InventoryToggle", True)
        mToggle.Left = 478: mToggle.Top = 430: mToggle.Width = 98: mToggle.Height = 30
        mToggle.Caption = VfkText("K\u00fc\u00e7\u00fclt")
        mCancel.Left = 588: mCancel.Top = 430: mCancel.Width = 104: mCancel.Height = 30
        mCancel.Caption = "Kapat": mCancel.Default = False
        Set mBack = Me.Controls.Add("Forms.CommandButton.1", "InventoryBack", True)
        mBack.Left = 12: mBack.Top = 474: mBack.Width = 138: mBack.Height = 28
        mBack.Caption = VfkText("\u00d6nceki sayfa")
        Set mPause = Me.Controls.Add("Forms.CommandButton.1", "InventoryNext", True)
        mPause.Left = 162: mPause.Top = 474: mPause.Width = 138: mPause.Height = 28
        mPause.Caption = "Sonraki sayfa"
        mSearch.Left = 312: mSearch.Top = 474: mSearch.Width = 150: mSearch.Height = 28
        mSearch.Caption = VfkText("D\u00f6k\u00fcm\u00fc yenile"): mSearch.Visible = True: mSearch.Default = False
        Set mIgnore = Me.Controls.Add("Forms.CommandButton.1", "InventoryIgnore", True)
        mIgnore.Left = 474: mIgnore.Top = 474: mIgnore.Width = 218: mIgnore.Height = 28
        mIgnore.Caption = VfkText("\u015eah\u0131s de\u011fil")
        mIgnore.ControlTipText = VfkText("Bu yaz\u0131m\u0131 yaln\u0131z bu belgede d\u00f6k\u00fcm ve toplu denetim d\u0131\u015f\u0131 b\u0131rak\u0131r. Kaydetmeden \u00f6nce onay ister.")
        mQuery.Visible = False
    End If
    Set mRows = rows: mToolChanges = canEdit
    mPrompt.Caption = heading
    mResults.Clear
    For i = 1 To rows.Count
        row = rows(i): mResults.AddItem CStr(row(0))
        mResults.List(i - 1, 1) = CStr(row(1)): mResults.List(i - 1, 2) = CStr(row(2))
        mResults.List(i - 1, 3) = CStr(row(3)): mResults.List(i - 1, 4) = CStr(row(7))
    Next i
    mChoose.Enabled = False: mStop.Enabled = False: mIgnore.Enabled = False: mLocate.Enabled = rows.Count > 0
    If rows.Count > 0 Then
        If selected < 1 Or selected > rows.Count Then selected = 1
        mResults.ListIndex = selected - 1
    Else
        mDetails.Text = VfkText("G\u00f6sterilecek ad bulunamad\u0131. Tan\u0131ma kurallar\u0131 b\u00fct\u00fcn \u015fah\u0131s adlar\u0131n\u0131 yakalama garantisi vermez.")
    End If
    mInventoryUpdating = False
    VfkRefreshButtons
    If mResults.ListIndex >= 0 Then mResults_Change
    Exit Sub
Failed:
    number = Err.Number: detail = Err.Description
    mInventoryUpdating = False
    Err.Raise number, "Vefeyat", detail
End Sub

Private Sub VfkInventoryHeading(ByVal text As String, ByVal x As Single, ByVal size As Single)
    Dim heading As MSForms.Label
    Set heading = Me.Controls.Add("Forms.Label.1")
    heading.Caption = text: heading.Tag = "VefeyatInventoryHeading"
    heading.Left = x: heading.Top = 64: heading.Width = size: heading.Height = 18
    heading.Font.Bold = True
End Sub

Public Sub InventoryNavigation(ByVal canBack As Boolean, ByVal canNext As Boolean)
    mBack.Enabled = canBack: mPause.Enabled = canNext
    VfkRefreshButtons
End Sub

Public Sub InventoryState(ByVal index As Long, ByVal aliasText As String, ByVal catalog As String, ByVal dateText As String, _
                          ByVal status As String, ByVal detail As String, ByVal canChange As Boolean, ByVal canAdd As Boolean, _
                          Optional ByVal canIgnore As Boolean = False)
    If index <> mResults.ListIndex + 1 Then Exit Sub
    mInventoryUpdating = True
    On Error GoTo Finished
    mResults.List(index - 1, 0) = aliasText: mResults.List(index - 1, 1) = catalog
    mResults.List(index - 1, 2) = dateText: mResults.List(index - 1, 3) = status
    mDetails.Text = detail
    mChoose.Enabled = canChange: mStop.Enabled = canAdd: mIgnore.Enabled = canIgnore
Finished:
    mInventoryUpdating = False
    VfkRefreshButtons
End Sub

Public Sub CollapseInventory()
    If mInventory And Not mLiveCollapsed Then VfkToggleInventory
End Sub

Private Sub VfkToggleInventory()
    Dim control As Object
    If Not mLiveCollapsed Then
        Set mLiveVisibility = CreateObject("Scripting.Dictionary")
        For Each control In Me.Controls
            mLiveVisibility(control.Name) = control.Visible: control.Visible = False
        Next control
        mLiveHeight = Me.Height: Me.Height = 64
        mToggle.Left = 12: mToggle.Top = 6: mToggle.Width = 680
        mToggle.Caption = VfkText("D\u00f6k\u00fcm\u00fc geni\u015flet"): mToggle.Visible = True
        mLiveCollapsed = True
    Else
        Me.Height = mLiveHeight
        For Each control In Me.Controls
            If mLiveVisibility.Exists(control.Name) Then control.Visible = CBool(mLiveVisibility(control.Name))
        Next control
        mToggle.Left = 478: mToggle.Top = 430: mToggle.Width = 98
        mToggle.Caption = VfkText("K\u00fc\u00e7\u00fclt"): mLiveCollapsed = False
    End If
    VfkRefreshButtons
End Sub

Public Sub SetIdentityEvidence(ByVal proof As Object)
    Set mIdentityProof = proof
    If mResults.ListIndex >= 0 Then mResults_Change
End Sub

Public Sub SetIdentityManageCaption()
    mPrompt.Caption = VfkText("Bu belgenin kimlik onaylar\u0131. Kayd\u0131 se\u00e7ip ge\u00e7i\u015fe gidin; kimli\u011fi yeniden do\u011frulay\u0131n.")
    mChoose.Caption = VfkText("Ge\u00e7i\u015fe git ve d\u00fczelt")
    mChoose.Default = False
    mCancel.Caption = "Kapat"
End Sub

Public Sub SetPlanMode(ByVal coverage As String, Optional ByVal canApply As Boolean = True)
    Dim control As Object
    mPlan = True: mPlanCoverage = coverage: PlanAction = ""
    Set mReviewExcluded = CreateObject("Scripting.Dictionary")
    Me.Height = 560
    mPrompt.Caption = VfkText("Toplu d\u00fczeltme \u00f6nizlemesi. Yaz\u0131m ve tarih ayr\u0131 se\u00e7ilir. Bilgi sat\u0131rlar\u0131 uygulanmaz; onaylamad\u0131klar\u0131n\u0131z de\u011fi\u015ftirilmez.")
    For Each control In Me.Controls
        If control.Tag = "VefeyatHeading" Then control.Top = 102
    Next control
    mResults.Top = 122: mResults.Height = 155
    mDetails.Top = 290: mDetails.Height = 180
    mChoose.Top = 486: mStop.Top = 486: mCancel.Top = 486
    mChoose.Caption = VfkText("Se\u00e7ilenleri uygula"): mChoose.Default = False
    mStop.Caption = VfkText("Se\u00e7imi de\u011fi\u015ftir"): mStop.Visible = True
    mSearch.Visible = True: mSearch.Caption = VfkText("T\u00fcm\u00fcn\u00fc se\u00e7")
    mSearch.Left = 16: mSearch.Top = 60: mSearch.Width = 170
    Set mBack = Me.Controls.Add("Forms.CommandButton.1", "PlanSelectNone", True)
    mBack.Caption = VfkText("Se\u00e7imleri temizle"): mBack.Left = 198: mBack.Top = 60: mBack.Width = 170: mBack.Height = 28
    Set mLocate = Me.Controls.Add("Forms.CommandButton.1", "PlanIdentity", True)
    mLocate.Caption = VfkText("Kimli\u011fi / yaz\u0131m\u0131 se\u00e7"): mLocate.Left = 380: mLocate.Top = 60: mLocate.Width = 236: mLocate.Height = 28
    PlanSelectAll True
    mResults_Change
    mChoose.Enabled = canApply
    VfkApplyTheme
End Sub

Public Function PlanIncludes(ByVal index As Long) As Boolean
    If index < 1 Or index > mRows.Count Then Exit Function
    If CStr(mRows(index)(2)) = "info" Then Exit Function
    PlanIncludes = Not mReviewExcluded.Exists(CStr(index))
End Function

Public Sub PlanSetSelection(ByVal index As Long, ByVal include As Boolean)
    Dim row As Variant, key As String
    If index < 1 Or index > mRows.Count Then Exit Sub
    key = CStr(index): row = mRows(index)
    If CStr(row(2)) = "info" Then include = False
    If include Then
        If mReviewExcluded.Exists(key) Then mReviewExcluded.Remove key
        mResults.List(index - 1, 0) = "[x] " & CStr(row(0))
    Else
        mReviewExcluded(key) = True
        mResults.List(index - 1, 0) = "[ ] " & CStr(row(0))
    End If
End Sub

Private Sub PlanSelectAll(ByVal include As Boolean)
    Dim i As Long
    For i = 1 To mRows.Count: PlanSetSelection i, include: Next i
End Sub

Public Sub SetRegistryMode(ByVal coverage As String, ByVal canApply As Boolean)
    Dim control As Object
    SetPlanMode coverage, canApply
    mRegistryMode = True
    mPrompt.Caption = VfkText("Belgeye kay\u0131tl\u0131 \u015fah\u0131slar. Yaz\u0131m/tarih ayr\u0131 onaylan\u0131r; elle de\u011fi\u015fen ge\u00e7i\u015fler korunur.")
    mLocate.Caption = VfkText("Metinde g\u00f6ster")
    mLocate.Enabled = mRows.Count > 0
    For Each control In Me.Controls
        If control.Tag = "VefeyatHeading" Then
            Select Case CLng(control.Left)
                Case 16: control.Caption = VfkText("Belgedeki ad")
                Case 176: control.Caption = VfkText("Yeni tarih")
                Case 266: control.Caption = VfkText("\u00d6nerilen yaz\u0131m")
            End Select
        End If
    Next control
    If mRows.Count = 0 Then mDetails.Text = VfkText("Bu belgede dinamik kay\u0131t yok. \u00d6zelli\u011fi etkinle\u015ftirip se\u00e7ili \u015fahs\u0131 kayda ba\u011flayabilirsiniz.")
End Sub
