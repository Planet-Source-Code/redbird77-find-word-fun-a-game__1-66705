VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FindWordFun! Settings"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlColor 
      Left            =   120
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   2280
      TabIndex        =   12
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   960
      TabIndex        =   11
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Frame fraBoard 
      Caption         =   "FUN! FUN! FUN!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.CheckBox chkBack 
         Caption         =   "Place Words Backwards"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2640
         Width           =   2295
      End
      Begin VB.TextBox txtCount 
         Height          =   315
         Left            =   2520
         TabIndex        =   4
         Text            =   "15"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtSize 
         Height          =   315
         Left            =   2520
         TabIndex        =   2
         Text            =   "50"
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   8
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   6
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblCap 
         Caption         =   "Solution Color"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label lblCap 
         Caption         =   "Selecting Color"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   3120
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label lblCap 
         Caption         =   "Number of Squares Across"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label lblCap 
         Caption         =   "Square Size (In Pixels)"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Label lblNote 
      Caption         =   "NOTE: Squares Across and Backwards settings do not take effect until a new game is started."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Width           =   3255
   End
End
Attribute VB_Name = "fSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' fSettings.frm \ redbird77@earthlink.net \ 2006.10.02

Option Explicit

Private Sub Form_Load()

' Use Settings user-defined type values as the controls' initial values.

    With Game.Settings
        txtSize.Text = .SquareSize
        txtCount.Text = .SquareCount
        
        lblColor(0).BackColor = .SelectColor
        lblColor(1).BackColor = .SolutionColor
        
        chkBack.Value = .Backwards
    
    End With
    
End Sub

Private Sub cmdOK_Click()

' Fill Settings user-defined type with new user-specfied values and write the
' type to the Settings INI file.

    With Game.Settings
        .SquareSize = txtSize.Text
        .SquareCount = txtCount.Text
        
        .SelectColor = lblColor(0).BackColor
        .SolutionColor = lblColor(1).BackColor
    
        .Backwards = Abs(chkBack.Value)
    End With
    
    Call mINI.PutSettings
    
    ' If game is active, redraw board with possible new size and colors.
    ' TODO: Only redraw if actual change.
    If Game.Active Then
        Call mGame.ResizeBoard
        Call mGame.Draw
    End If
    
    Call Unload(Me)
    
End Sub

Private Sub cmdCancel_Click()

' Leave without saving any changes.

    Call Unload(Me)
    
End Sub

Private Sub lblColor_Click(Index As Integer)
    
On Error GoTo ErrHandler

    cdlColor.ShowColor
    
    lblColor(Index).BackColor = cdlColor.Color
    
    Exit Sub
    
ErrHandler:
    If Err.Number = cdlCancel Then Exit Sub
    
End Sub
