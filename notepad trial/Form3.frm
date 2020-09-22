VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   " Change Case"
   ClientHeight    =   2115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2595
   Icon            =   "Form3.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   2595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      ToolTipText     =   "Change case"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.OptionButton Option3 
         Caption         =   "lower case"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "UPPER CASE"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Title Case"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngSelStrt As Long
Dim lngSelLen As Long

Private Sub cmdOk_Click()
With Form1
lngSelStrt = .RTF.SelStart
lngSelLen = .RTF.SelLength

    If Option1.Value = True Then
        .RTF.SelText = StrConv(.RTF.SelText, vbProperCase)
    ElseIf Option2.Value = True Then
        .RTF.SelText = StrConv(.RTF.SelText, vbUpperCase)
    ElseIf Option3.Value = True Then
        .RTF.SelText = StrConv(.RTF.SelText, vbLowerCase)
    End If
    
.RTF.SelStart = lngSelStrt
.RTF.SelLength = lngSelLen
End With
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub
