VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Properties"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3600
         TabIndex        =   9
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "O&K"
         Height          =   375
         Left            =   2280
         TabIndex        =   8
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label lblSize 
         AutoSize        =   -1  'True
         Caption         =   "Size"
         Height          =   195
         Left            =   1560
         TabIndex        =   18
         Top             =   1440
         Width           =   300
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Size:"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   1440
         Width           =   345
      End
      Begin VB.Label lblDosName 
         AutoSize        =   -1  'True
         Caption         =   "Dos Name"
         Height          =   195
         Left            =   1560
         TabIndex        =   16
         Top             =   2040
         Width           =   750
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "MS-Dos Name:"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   2040
         Width           =   1080
      End
      Begin VB.Label lblAccessed 
         AutoSize        =   -1  'True
         Caption         =   "Accessed"
         Height          =   195
         Left            =   1560
         TabIndex        =   14
         Top             =   3090
         Width           =   705
      End
      Begin VB.Label lblModified 
         AutoSize        =   -1  'True
         Caption         =   "Modified"
         Height          =   195
         Left            =   1560
         TabIndex        =   13
         Top             =   2730
         Width           =   600
      End
      Begin VB.Label lblCreated 
         AutoSize        =   -1  'True
         Caption         =   "Created"
         Height          =   195
         Left            =   1560
         TabIndex        =   12
         Top             =   2370
         Width           =   555
      End
      Begin VB.Label lblLocation 
         AutoSize        =   -1  'True
         Caption         =   "Location"
         Height          =   195
         Left            =   1560
         TabIndex        =   11
         Top             =   1050
         Width           =   615
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         Caption         =   "File Type"
         Height          =   195
         Left            =   1560
         TabIndex        =   10
         Top             =   690
         Width           =   645
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Modified:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   2730
         Width           =   645
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Accessed:"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   3090
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Created:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   2370
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Location:"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   1050
         Width           =   660
      End
      Begin VB.Label label2 
         Caption         =   "Type:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   690
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   330
         Width           =   465
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         Index           =   5
         X1              =   120
         X2              =   4920
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   4
         X1              =   120
         X2              =   4920
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label lblFileName 
         AutoSize        =   -1  'True
         Caption         =   "File Name"
         Height          =   195
         Left            =   1560
         TabIndex        =   1
         Top             =   330
         Width           =   705
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         Index           =   1
         X1              =   120
         X2              =   4920
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   120
         X2              =   4920
         Y1              =   1800
         Y2              =   1800
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fs
Dim f
Dim s
Public strFileName As String

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
strFileName = Trim(Form1.Caption)
Set fs = CreateObject("Scripting.FileSystemObject")
On Error GoTo noFile
Set f = fs.GetFile(strFileName)

lblFileName = StrConv(f.Name, vbProperCase)
lblType = Replace(StrConv(f.Type, vbUpperCase), "FILE", "File")
lblLocation = f.ParentFolder
lblCreated = f.DateCreated
lblModified = f.DateLastModified
lblAccessed = f.DateLastAccessed
lblDosName = f.ShortName
lblSize = Round(f.Size / 1024, 1) & " KB" & " (" & f.Size & " bytes)"
Exit Sub

noFile:
    MsgBox "Cannot display properties now!!!", vbExclamation
    Unload Me
End Sub
