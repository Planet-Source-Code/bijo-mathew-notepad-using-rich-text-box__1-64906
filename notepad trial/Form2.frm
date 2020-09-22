VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Find"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   Icon            =   "Form2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Visible         =   0   'False
      Width           =   3975
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Label3"
         Height          =   195
         Left            =   1080
         TabIndex        =   11
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdReplaceAll 
      Caption         =   "Replace &All"
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      ToolTipText     =   "Replace all"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3975
      Begin VB.TextBox txtReplace 
         Height          =   375
         Left            =   960
         TabIndex        =   3
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox txtFind 
         Height          =   375
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
      Begin VB.CheckBox Check2 
         Caption         =   "&Whole Word"
         Height          =   255
         Left            =   2280
         TabIndex        =   5
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&Match Case"
         Height          =   195
         Left            =   600
         TabIndex        =   4
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Replac&e:"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fin&d:"
         Height          =   195
         Left            =   540
         TabIndex        =   0
         Top             =   360
         Width           =   345
      End
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "&Replace"
      Height          =   375
      Left            =   1610
      TabIndex        =   7
      ToolTipText     =   "Replace with"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Find word to find"
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngStrtChk As Long
Dim lngNoOfReplaces As Long
Dim wordResult
Dim chkMatch
Dim chkWhole

Private Sub Check1_Click()
If Check1.Value = 0 Then
    chkMatch = 0
ElseIf Check1.Value = 1 Then
    chkMatch = 4
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 0 Then
    chkWhole = 0
ElseIf Check2.Value = 1 Then
    chkWhole = 2
End If
End Sub

Private Sub cmdFind_Click()
Screen.MousePointer = vbHourglass

Frame2.Visible = False
With Form1
    wordResult = -1
    wordResult = .RTF.Find(txtFind.Text, lngStrtChk, Len(.RTF.Text), chkWhole Or chkMatch)
    If wordResult = -1 Then
        MsgBox "Word not found !!!", vbInformation
        cmdFind.Caption = "&Find"
        lngStrtChk = 0
    Else
        lngStrtChk = .RTF.SelStart + 1
        cmdFind.Caption = "&Find Next"
    End If
    
End With

Screen.MousePointer = vbNormal
End Sub

Private Sub cmdReplace_Click()
Screen.MousePointer = vbHourglass

Frame2.Visible = False
If Label2.Visible = False And txtReplace.Visible = False Then
    Label2.Visible = True
    txtReplace.Visible = True
    If txtFind.Text = "" Then
        txtFind.SetFocus
    Else
        txtReplace.SetFocus
    End If
Else
With Form1
    If .RTF.SelText <> "" Then
'        .RTF.SelText = Replace(.RTF.SelText, txtFind.Text, txtReplace.Text, , , vbDatabaseCompare)
        .RTF.SelText = txtReplace.Text
        Call cmdFind_Click
    Else
        Call cmdFind_Click
    End If
End With
End If
Screen.MousePointer = vbNormal
End Sub

Private Sub cmdReplaceAll_Click()
Screen.MousePointer = vbHourglass

Call Form_Load
With Form1
    lngNoOfReplaces = 0

If Label2.Visible = False And txtReplace.Visible = False Then
    Label2.Visible = True
    txtReplace.Visible = True
    txtReplace.SetFocus
    If txtFind.Text = "" Then
        txtFind.SetFocus
    Else
        txtReplace.SetFocus
    End If
Else
    Call cmdFind_Click
    wordResult = 0
    lngStrtChk = 0

    If .RTF.SelText <> "" Then
        While wordResult <> -1
            wordResult = .RTF.Find(txtFind.Text, lngStrtChk, Len(.RTF.Text), chkWhole Or chkMatch)
'            .RTF.SelText = Replace(.RTF.SelText, txtFind.Text, txtReplace.Text, , , vbDatabaseCompare)
            .RTF.SelText = txtReplace.Text
            lngStrtChk = .RTF.SelStart + 1
           lngNoOfReplaces = lngNoOfReplaces + 1
        Wend
    End If
End If
End With
If lngNoOfReplaces > 1 Then
    Frame2.Visible = True
    Label3.Caption = "Totally replaced " & lngNoOfReplaces - 1 & " word(s)"
End If
Screen.MousePointer = vbNormal
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
chkMatch = 0
chkWhole = 0
lngStrtChk = 0
Form1.RTF.SelLength = 0
End Sub

