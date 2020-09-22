VERSION 5.00
Begin VB.Form TicTacToeMainScreen 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000008&
   Caption         =   "Tic Tac Toe Millennium Edition"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   438
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   589
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picTiles 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   600
      Picture         =   "TicTacToeMainScreen.frx":0000
      ScaleHeight     =   329.669
      ScaleMode       =   0  'User
      ScaleWidth      =   492
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   7410
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   3720
      Left            =   0
      ScaleHeight     =   248.008
      ScaleMode       =   0  'User
      ScaleWidth      =   738
      TabIndex        =   0
      Top             =   1440
      Width           =   11100
   End
   Begin VB.PictureBox picDetection 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   3720
      Left            =   240
      Picture         =   "TicTacToeMainScreen.frx":75DA0
      ScaleHeight     =   248.008
      ScaleMode       =   0  'User
      ScaleWidth      =   738
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   11100
   End
End
Attribute VB_Name = "TicTacToeMainScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BBBoard(SquareNo As Long, X As Long, Y As Long, BlueOrWhite As Long)
  BitBlt picMain.hdc, X, Y, TileSizeX, TileSizeY, picTiles.hdc, 0, 246, BBMASK
  BitBlt picMain.hdc, X, Y, TileSizeX, TileSizeY, picTiles.hdc, 246 * BlueOrWhite, 82 * GD.Board(SquareNo), BBPAINT
End Sub
Private Sub DrawBoard()
  picMain.Cls
  BBBoard 1, 246, 0, 0
  BBBoard 2, 369, 41, 1
  BBBoard 3, 492, 82, 0
  BBBoard 4, 123, 41, 1
  BBBoard 5, 246, 82, 0
  BBBoard 6, 369, 123, 1
  BBBoard 7, 0, 82, 0
  BBBoard 8, 123, 123, 1
  BBBoard 9, 246, 164, 0
    
  If GD.CurrentBoardNo > 0 Then
    If GD.Board(GD.CurrentBoardNo) = 0 Then
      BitBlt picMain.hdc, GD.Current.X, GD.Current.Y, TileSizeX, TileSizeY, picTiles.hdc, 0, 246, BBMASK
      BitBlt picMain.hdc, GD.Current.X, GD.Current.Y, TileSizeX, TileSizeY, picTiles.hdc, BlueX, 246, BBPAINT
    End If
  End If
  
  picMain.Refresh
  'DoEvents
End Sub

Private Sub Form_Load()
  Resetboard
  GD.CurrentPlayer = 1
End Sub

Private Sub Form_Resize()
  picMain = picDetection
  If Me.WindowState = 0 Then Me.WindowState = 2
  If Not Me.WindowState = 2 Then Exit Sub
  Me.picMain.Top = Int((Me.ScaleHeight) - (Me.picMain.ScaleHeight)) / 2
  Me.picMain.Left = Int((Me.ScaleWidth) - (Me.picMain.ScaleWidth)) / 2
  
  Me.picDetection.Top = Me.picMain.Top
  Me.picDetection.Left = Me.picMain.Left
  DrawBoard
End Sub



Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If CheckBoardFull Then
      DrawBoard
      MsgBox "No winner"
      Resetboard
    End If
  If GD.CurrentPlayer = 1 Then
    If GD.CurrentBoardNo = 0 Then Exit Sub
    If GD.Board(GD.CurrentBoardNo) > 0 Then Exit Sub
    GD.Board(GD.CurrentBoardNo) = GD.CurrentPlayer
    If GameWon(1) Then
      DrawBoard
      MsgBox "Player " & GD.CurrentPlayer & " Wins"
      Resetboard
    End If
    GD.CurrentPlayer = 2
    picMain_MouseDown 0, 0, 0, 0
  Else
    ComputersGo
    If GameWon(2) Then
      DrawBoard
      MsgBox "Computer Wins"
      Resetboard
    End If
    
    If CheckBoardFull Then
      DrawBoard
      MsgBox "No winner"
      Resetboard
    End If
    GD.CurrentPlayer = 1
  End If
  DrawBoard
End Sub


Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  SetCurrentSquare picDetection.Point(X, Y)
  DrawBoard
End Sub


