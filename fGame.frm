VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fGame 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FindWordFun!"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   289
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   409
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdlOpen 
      Left            =   5160
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmdSettings 
      Caption         =   "&Settings"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Open &New Puzzle"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.PictureBox pBoard 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   2520
      ScaleHeight     =   233
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   233
      TabIndex        =   3
      Top             =   720
      Width           =   3495
   End
   Begin MSComctlLib.TreeView tvWords 
      Height          =   3495
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   6165
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "fGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' fGame.frm \ redbird77@earthlink.net \ 2006.10.02

' VB is my hobby not my profession, so I decided on scrapping the Hungarian-ish
' notatation for this project.  Just because I was tired of looking at underscores
' and lowercase letters... I don't know ;)

Option Explicit

Private Sub cmdSettings_Click()

    Call fSettings.Show '(vbModal, Me)
    
End Sub

Private Sub cmdPlay_Click()
   
    Call mGame.Play
    
End Sub

Private Sub Form_Load()

    Call mINI.GetSettings
    
    fGame.pBoard.Width = Game.Settings.SquareSize * Game.Settings.SquareCount
    fGame.pBoard.Height = fGame.pBoard.Width
    fGame.tvWords.Height = fGame.pBoard.Height
    
    fGame.Width = (fGame.pBoard.Left + fGame.pBoard.Width + 15) * Screen.TwipsPerPixelX
    fGame.Height = (fGame.tvWords.Top + fGame.tvWords.Height + fGame.cmdPlay.Height) * Screen.TwipsPerPixelY

End Sub

Private Sub pBoard_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call mGame.OnMouseDown(X, Y)

End Sub

Private Sub pBoard_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If (X > -1 And X < pBoard.Width) And (Y > -1 And Y < pBoard.Height) Then
        Call mGame.OnMouseMove(X, Y)
    End If
    
End Sub

Private Sub pBoard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call mGame.OnMouseUp(X, Y)

End Sub
