VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Document"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11880
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   510
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   900
      ButtonWidth     =   847
      ButtonHeight    =   741
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   20
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Bold"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Italics"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Underline"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Left"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Center"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Right"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   4
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Bullet"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Font Color"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Find"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Spelling Check"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Left Indent"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Right Indent"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Strike Through"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
      EndProperty
      Begin VB.ComboBox cmbSize 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "Form1.frx":030A
         Left            =   8760
         List            =   "Form1.frx":03D1
         Sorted          =   -1  'True
         TabIndex        =   4
         Text            =   "10"
         Top             =   120
         Width           =   855
      End
      Begin VB.ComboBox cmbFont 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "Form1.frx":04D9
         Left            =   9720
         List            =   "Form1.frx":04DB
         Sorted          =   -1  'True
         TabIndex        =   3
         Text            =   "Arial"
         Top             =   120
         Width           =   2055
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   8055
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   7
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   7434
            MinWidth        =   2116
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "CAPS"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2117
            MinWidth        =   2116
            TextSave        =   "NUM"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   2117
            MinWidth        =   2116
            TextSave        =   "INS"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   4
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "SCRL"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2646
            MinWidth        =   2646
            TextSave        =   "10:03 AM"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "04-Apr-2006"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   7920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RTF 
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12726
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":04DD
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   7920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   25
      ImageHeight     =   22
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   13
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":055F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":0C39
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":124B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":185D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":1E6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":2481
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":2ABB
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":30CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":3797
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":3DA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":44BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":4ACD
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":51DF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File   "
      Begin VB.Menu mnuNew 
         Caption         =   "&New   "
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open   "
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close   "
      End
      Begin VB.Menu mnuSepa1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save   "
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As   "
      End
      Begin VB.Menu mnuSepa2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu sepa5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "Properties..."
      End
      Begin VB.Menu mnuSepa4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Quit "
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit   "
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo Typing   "
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "&Redo Typing   "
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuSepa3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy   "
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t   "
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste   "
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuSepa13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuSepa12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "&Find   "
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuReplace 
         Caption         =   "&Replace"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools   "
      Begin VB.Menu mnuChkSpell 
         Caption         =   "&Check Spelling   "
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuChangeCase 
         Caption         =   "&Change Case   "
      End
      Begin VB.Menu mnuSepa8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCalc 
         Caption         =   "&Calculator"
      End
      Begin VB.Menu mnuSepa10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh   "
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "F&ormat   "
      Begin VB.Menu mnuLeftIndent 
         Caption         =   "&Left Indent   "
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuRightIndent 
         Caption         =   "&Right Indent   "
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuSepa7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddDate 
         Caption         =   "Add &Date   "
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuAddTime 
         Caption         =   "Add &Time   "
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About..."
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strFromCalc As String
Dim boolClose
Dim strBeforCut As String
Dim strAftrCut As String
Dim strForCopy As String
Dim strBeforPaste As String
Dim strAfterPaste As String
Dim strForFind As String

Dim strRTFKeyUp As String
Dim strRTFKeyDown As String
Dim intUndo As Integer
Dim undo1 As String
Dim undo2 As String
Dim undo3 As String
Dim undo4 As String
Dim undo5 As String
Dim undo6 As String
Dim undo7 As String
Dim undo8 As String
Dim undo9 As String
Dim undo10 As String

Dim boolUndo1 As Boolean
Dim boolUndo2 As Boolean
Dim boolUndo3 As Boolean
Dim boolUndo4 As Boolean
Dim boolUndo5 As Boolean
Dim boolUndo6 As Boolean
Dim boolUndo7 As Boolean
Dim boolUndo8 As Boolean
Dim boolUndo9 As Boolean
Dim boolUndo10 As Boolean

Dim strLastSave As String
Dim strNowText As String
Dim fs, f, fc, f1, a

Dim strForInsertion As String
Dim strAfterInsertion As String
Dim saveEditFile1 As String
Dim saveEditFile As String
Dim strText As String

Private Sub cmbFont_Change()
On Error Resume Next
RTF.SelFontName = cmbFont.Text
End Sub

Private Sub cmbFont_Click()
On Error Resume Next
RTF.SelFontName = cmbFont.Text
End Sub

Private Sub cmbFont_LostFocus()
On Error Resume Next
cmbFont.Text = RTF.SelFontName
End Sub

Private Sub cmbSize_Change()
Call cmbSize_Click
End Sub

Private Sub cmbSize_Click()
On Error Resume Next
RTF.SelFontSize = cmbSize.Text
End Sub

Private Sub cmbSize_LostFocus()
On Error Resume Next
cmbSize.Text = RTF.SelFontSize
End Sub

Private Sub Form_Load()
Set fs = CreateObject("Scripting.FileSystemObject")
On Error Resume Next
fs.DeleteFile (App.Path & "\Editor.txt")
Set a = fs.CreateTextFile(App.Path & "\Editor.txt", 1, True)
RTF.LoadFile App.Path & "\Editor.txt"

Form1.Caption = " New File"
strLastSave = ""
strNowText = ""
Call setFont

undo1 = ""
undo2 = ""
undo3 = ""
undo4 = ""
undo5 = ""
undo6 = ""
undo7 = ""
undo8 = ""
undo9 = ""
undo10 = ""
intUndo = 0
strRTFKeyUp = ""
strRTFKeyDown = ""
boolUndo1 = True
boolUndo2 = False
boolUndo3 = False
boolUndo4 = False
boolUndo5 = False
boolUndo6 = False
boolUndo7 = False
boolUndo8 = False
boolUndo9 = False
boolUndo10 = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
strNowText = RTF.Text
If strLastSave <> strNowText Then
    boolClose = MsgBox("Do you want to save the file before quitting?", vbYesNoCancel + vbDefaultButton1)
    If boolClose = vbYes Then
        Call mnuSave_Click
        
        On Error Resume Next
        Unload Form2
        On Error Resume Next
        Unload Form3
     ElseIf boolClose = vbNo Then
        On Error Resume Next
        Unload Form2
        On Error Resume Next
        Unload Form3
        
    ElseIf boolClose = vbCancel Then
        Cancel = 1
    End If
End If
End Sub

Private Sub mnuAbout_Click()
MsgBox "Coded by Bijo Mathew", vbInformation
End Sub

Private Sub mnuAddDate_Click()
    RTF.Text = RTF.Text & vbCrLf & "Date: " & Format(Now(), "dd/mm/yyyy")
End Sub

Private Sub mnuAddTime_Click()
RTF.Text = RTF.Text & vbCrLf & "Time: " & Format(Now(), "hh:mm:ss")
End Sub

Private Sub mnuCalc_Click()
Dim retVal
retVal = Shell("calc", vbNormalFocus)
End Sub

Private Sub mnuChangeCase_Click()
Form3.Show
End Sub

Private Sub mnuChkSpell_Click()
    strText = RTF.Text
    Call MsSpellCheck(strText)
End Sub

Private Sub mnuClose_Click()
RTF.Text = ""
Call Form_Load
End Sub

Private Sub mnuCopy_Click()
Clipboard.Clear
strForCopy = RTF.SelText
End Sub

Private Sub mnuCut_Click()
   strForCopy = RTF.SelText
   RTF.SelText = Replace(RTF.SelText, RTF.SelText, "")
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuFind_Click()
strForFind = RTF.SelText
With Form2
    .Show
    .txtReplace.Visible = False
    .Label2.Visible = False
    .SetFocus
    .txtFind.Text = strForFind
    .txtFind.SelStart = Len(strForFind)
End With
End Sub

Private Sub mnuLeftIndent_Click()
RTF.SelIndent = RTF.SelIndent - 500
End Sub

Private Sub mnuNew_Click()
Call Form_Load
RTF.Text = ""
End Sub

Private Sub mnuOpen_Click()
CD1.Filter = "Text File (*.txt)|*.txt|RTF File (*.rtf)|*.rtf|DOC File (*.doc)|*.doc|HTM File (*.htm)|*.htm|HTML File (*.html)|*.html"
CD1.ShowOpen
On Error GoTo fileOpen
RTF.FileName = CD1.FileName
If CD1.FileName = "" Then
    Form1.Caption = " New File"
Else
    Form1.Caption = " " & CD1.FileName
End If
Exit Sub

fileOpen:
MsgBox "File in use. Close the opened file and try again.", vbExclamation
End Sub

Private Sub mnuPaste_Click()
strBeforPaste = Mid(RTF.Text, 1, RTF.SelStart)
strAfterPaste = Mid(RTF.Text, RTF.SelStart + RTF.SelLength + 1, Len(RTF.Text))
RTF.Text = strBeforPaste & strForCopy & strAfterPaste
RTF.SelStart = Len(strBeforPaste) + Len(strForCopy)
RTF.SelLength = 0
End Sub

Private Sub mnuPrint_Click()
CD1.ShowPrinter
RTF.SelPrint (Printer.hDC)
End Sub

Private Sub mnuProperties_Click()
On Error Resume Next
Form4.Show , Me
End Sub

Private Sub mnuRedo_Click()
    If intUndo < 10 Then
        intUndo = intUndo + 1
    Else
        intUndo = 1
    End If
    
If intUndo = 1 And undo1 <> "" Then
    RTF.Text = ""
    RTF.Text = undo1
ElseIf intUndo = 2 And undo2 <> "" Then
    RTF.Text = ""
    RTF.Text = undo2
ElseIf intUndo = 3 And undo3 <> "" Then
    RTF.Text = ""
    RTF.Text = undo3
ElseIf intUndo = 4 And undo4 <> "" Then
    RTF.Text = ""
    RTF.Text = undo4
ElseIf intUndo = 5 And undo5 <> "" Then
    RTF.Text = ""
    RTF.Text = undo5
ElseIf intUndo = 6 And undo6 <> "" Then
    RTF.Text = ""
    RTF.Text = undo6
ElseIf intUndo = 7 And undo7 <> "" Then
    RTF.Text = ""
    RTF.Text = undo7
ElseIf intUndo = 8 And undo8 <> "" Then
    RTF.Text = ""
    RTF.Text = undo8
ElseIf intUndo = 9 And undo9 <> "" Then
    RTF.Text = ""
    RTF.Text = undo2
ElseIf intUndo = 10 And undo10 <> "" Then
    RTF.Text = ""
    RTF.Text = undo10
End If
End Sub

Private Sub mnuRefresh_Click()
RTF.Refresh
End Sub

Private Sub mnuReplace_Click()
strForFind = RTF.SelText
With Form2
    .Show
    .txtFind.Text = strForFind
    .txtFind.SelStart = Len(strForFind)
    .txtReplace.SetFocus
End With
End Sub

Private Sub mnuRightIndent_Click()
If RTF.SelIndent < 10000 Then
    RTF.SelIndent = RTF.SelIndent + 500
End If
End Sub

Private Sub mnuSave_Click()
saveEditFile = RTF.FileName
saveEditFile1 = App.Path & "\Editor.txt"

If Form1.Caption = " New File" Then
    Call mnuSaveAs_Click
Else
    RTF.SaveFile RTF.FileName, 1
    strLastSave = RTF.Text
End If

If CD1.FileName = "" Then
    Form1.Caption = " New File"
ElseIf CD1.FileTitle <> "Editor.txt" Then
    Form1.Caption = " " & CD1.FileName
End If

End Sub

Private Sub mnuSaveAs_Click()
CD1.Filter = "Text File (*.txt)|*.txt|RTF File (*.rtf)|*.rtf|DOC File (*.doc)|*.doc|HTM File (*.htm)|*.htm|HTML File (*.html)|*.html"
CD1.ShowSave

If CD1.FileTitle = "Editor.txt" Then
    Form1.Caption = " New File"
    MsgBox "The filename cannot be 'Editor.txt'", vbExclamation
    Exit Sub
End If

On Error GoTo errorFound
If Mid(UCase(CD1.FileTitle), Len(CD1.FileTitle) - 2, Len(CD1.FileTitle)) = "TXT" Or Mid(UCase(CD1.FileTitle), Len(CD1.FileTitle) - 3, Len(CD1.FileTitle)) = "HTML" Or Mid(UCase(CD1.FileTitle), Len(CD1.FileTitle) - 2, Len(CD1.FileTitle)) = "HTM" Then
    RTF.SaveFile CD1.FileName, 1
    strLastSave = RTF.Text
Else
    RTF.SaveFile CD1.FileName, 0
    strLastSave = RTF.Text
End If

If CD1.FileName = "" Then
    Form1.Caption = " New File"
Else
    Form1.Caption = " " & CD1.FileName
End If

Exit Sub

errorFound:
Exit Sub
End Sub

Private Sub mnuSelectAll_Click()
RTF.SelStart = 0
RTF.SelLength = Len(RTF)
End Sub

Private Sub mnuUndo_Click()
If intUndo = 1 And undo1 <> "" Then
    RTF.Text = ""
    RTF.Text = undo1
ElseIf intUndo = 2 And undo2 <> "" Then
    RTF.Text = ""
    RTF.Text = undo2
ElseIf intUndo = 3 And undo3 <> "" Then
    RTF.Text = ""
    RTF.Text = undo3
ElseIf intUndo = 4 And undo4 <> "" Then
    RTF.Text = ""
    RTF.Text = undo4
ElseIf intUndo = 5 And undo5 <> "" Then
    RTF.Text = ""
    RTF.Text = undo5
ElseIf intUndo = 6 And undo6 <> "" Then
    RTF.Text = ""
    RTF.Text = undo6
ElseIf intUndo = 7 And undo7 <> "" Then
    RTF.Text = ""
    RTF.Text = undo7
ElseIf intUndo = 8 And undo8 <> "" Then
    RTF.Text = ""
    RTF.Text = undo8
ElseIf intUndo = 9 And undo9 <> "" Then
    RTF.Text = ""
    RTF.Text = undo2
ElseIf intUndo = 10 And undo10 <> "" Then
    RTF.Text = ""
    RTF.Text = undo10
End If
If intUndo > 1 Then
    intUndo = intUndo - 1
End If
End Sub

Private Sub RTF_KeyDown(KeyCode As Integer, Shift As Integer)
strRTFKeyDown = RTF.Text
End Sub

Private Sub RTF_KeyUp(KeyCode As Integer, Shift As Integer)
strRTFKeyUp = ""
strRTFKeyUp = RTF.Text
If strRTFKeyDown <> strRTFKeyUp Then
    If intUndo < 10 Then
        intUndo = intUndo + 1
    Else
        intUndo = 1
    End If

    If boolUndo1 = True Then
        undo1 = ""
        undo1 = RTF.Text
        boolUndo1 = False
        boolUndo2 = True
        Exit Sub
    ElseIf boolUndo2 = True Then
        undo2 = ""
        undo2 = RTF.Text
        boolUndo2 = False
        boolUndo3 = True
        Exit Sub
    ElseIf boolUndo3 = True Then
        undo3 = ""
        undo3 = RTF.Text
        boolUndo3 = False
        boolUndo4 = True
        Exit Sub
    ElseIf boolUndo4 = True Then
        undo4 = ""
        undo4 = RTF.Text
        boolUndo4 = False
        boolUndo5 = True
        Exit Sub
    ElseIf boolUndo5 = True Then
        undo5 = ""
        undo5 = RTF.Text
        boolUndo5 = False
        boolUndo6 = True
        Exit Sub
    ElseIf boolUndo6 = True Then
        undo6 = ""
        undo6 = RTF.Text
        boolUndo6 = False
        boolUndo7 = True
        Exit Sub
    ElseIf boolUndo7 = True Then
        undo7 = ""
        undo7 = RTF.Text
        boolUndo7 = False
        boolUndo8 = True
        Exit Sub
    ElseIf boolUndo8 = True Then
        undo8 = ""
        undo8 = RTF.Text
        boolUndo8 = False
        boolUndo9 = True
        Exit Sub
    ElseIf boolUndo9 = True Then
        undo9 = ""
        undo9 = RTF.Text
        boolUndo9 = False
        boolUndo10 = True
        Exit Sub
    ElseIf boolUndo10 = True Then
        undo10 = ""
        undo10 = RTF.Text
        boolUndo10 = False
        boolUndo1 = True
        Exit Sub
    End If
End If
End Sub

Private Sub RTF_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    PopupMenu mnuEdit
End If
End Sub

Private Sub RTF_SelChange()
If RTF.SelBold = True Then
    Toolbar1.Buttons.Item(1).Value = tbrPressed
Else
    Toolbar1.Buttons.Item(1).Value = tbrUnpressed
End If

If RTF.SelItalic = True Then
    Toolbar1.Buttons.Item(2).Value = tbrPressed
Else
    Toolbar1.Buttons.Item(2).Value = tbrUnpressed
End If

If RTF.SelUnderline = True Then
    Toolbar1.Buttons.Item(3).Value = tbrPressed
Else
    Toolbar1.Buttons.Item(3).Value = tbrUnpressed
End If

If RTF.SelAlignment = rtfLeft Then
    Toolbar1.Buttons.Item(5).Value = tbrPressed
Else
    Toolbar1.Buttons.Item(5).Value = tbrUnpressed
End If

If RTF.SelAlignment = rtfCenter Then
    Toolbar1.Buttons.Item(6).Value = tbrPressed
Else
    Toolbar1.Buttons.Item(6).Value = tbrUnpressed
End If

If RTF.SelAlignment = rtfRight Then
    Toolbar1.Buttons.Item(7).Value = tbrPressed
Else
    Toolbar1.Buttons.Item(7).Value = tbrUnpressed
End If

If RTF.SelBullet = True Then
    Toolbar1.Buttons.Item(9).Value = tbrPressed
Else
    Toolbar1.Buttons.Item(9).Value = tbrUnpressed
End If

If RTF.SelStrikeThru = True Then
    Toolbar1.Buttons.Item(20).Value = tbrPressed
Else
    Toolbar1.Buttons.Item(20).Value = tbrUnpressed
End If

On Error Resume Next
cmbSize.Text = RTF.SelFontSize
On Error Resume Next
cmbFont.Text = RTF.SelFontName

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
If Button.Index = 1 Then
    If RTF.SelBold = True Then
        RTF.SelBold = False
    Else
        RTF.SelBold = True
    End If
End If

If Button.Index = 2 Then
    If RTF.SelItalic = True Then
        RTF.SelItalic = False
    Else
        RTF.SelItalic = True
    End If
End If

If Button.Index = 3 Then
    If RTF.SelUnderline = True Then
        RTF.SelUnderline = False
    Else
        RTF.SelUnderline = True
    End If
End If

If Button.Index = 5 Then
    If RTF.SelAlignment = rtfLeft Then

    Else
        RTF.SelAlignment = rtfLeft
    End If
End If

If Button.Index = 6 Then
    If RTF.SelAlignment = rtfCenter Then

    Else
        RTF.SelAlignment = rtfCenter
    End If
End If

If Button.Index = 7 Then
    If RTF.SelAlignment = rtfRight Then

    Else
        RTF.SelAlignment = rtfRight
    End If
End If

If Button.Index = 9 Then
    If RTF.SelBullet = True Then
        RTF.SelBullet = False
    Else
        RTF.BulletIndent = 250
        RTF.SelBullet = True

    End If
End If

If Button.Index = 11 Then
    CD1.ShowColor
    RTF.SelColor = CD1.Color
End If

If Button.Index = 13 Then
    Call mnuFind_Click
End If

If Button.Index = 15 Then
    strText = RTF.Text
    Call MsSpellCheck(strText)
End If

If Button.Index = 17 Then
    RTF.SelIndent = RTF.SelIndent - 500
End If

If Button.Index = 18 Then
    If RTF.SelIndent < 10000 Then
        RTF.SelIndent = RTF.SelIndent + 500
    End If
End If

If Button.Index = 20 Then
    If RTF.SelStrikeThru = True Then
        RTF.SelStrikeThru = False
    Else
        RTF.SelStrikeThru = True
    End If
End If

Call RTF_SelChange
End Sub

Function MsSpellCheck(strText As String) As String
Screen.MousePointer = vbHourglass
RTF.SelLength = 0 'added
     Dim oWord As Object
     Dim strSelection As String
     On Error GoTo noMSWordInstalled
     Set oWord = CreateObject("Word.Basic")
     oWord.AppMinimize
     MsSpellCheck = strText
     oWord.FileNewDefault
     oWord.EditSelectAll
     oWord.EditCut
     oWord.Insert strText
     oWord.StartOfDocument
     On Error Resume Next
     oWord.ToolsSpelling
     On Error GoTo 0
     oWord.EditSelectAll
     strSelection = oWord.Selection$

 If Mid(strSelection, Len(strSelection), 1) = Chr(13) Then 'added
    strSelection = Mid(strSelection, 1, Len(strSelection) - 1) 'added
 End If 'added

 If Len(strSelection) > 1 Then 'added
    MsSpellCheck = strSelection 'added
    RTF.Text = strSelection 'added
 End If 'added

     oWord.FileCloseAll 2
     oWord.AppClose
     Set oWord = Nothing
 MsgBox "Spell check is over.", vbInformation
 Screen.MousePointer = vbNormal
Exit Function

noMSWordInstalled:
MsgBox "Cannot perform spell check.", vbCritical
Screen.MousePointer = vbNormal
End Function

Sub setFont()
Dim intX As Integer
intX = 0
cmbFont.Clear
While intX <= Screen.FontCount
    cmbFont.AddItem Screen.Fonts(intX)
    intX = intX + 1
Wend
cmbFont.Text = "Times New Roman"
End Sub

