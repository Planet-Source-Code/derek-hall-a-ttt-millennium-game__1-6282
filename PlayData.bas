Attribute VB_Name = "PlayData"
Option Explicit

Public Const TileSizeX = 246
Public Const TileSizeY = 81
Public Const WhiteX = 0
Public Const BlueX = 246
Public Const SilverY = 82
Public Const GoldY = 164
Public Const EmptySquare = 0
Public Const Player1Square = 1
Public Const Player2Square = 2

Type XYData
  X As Single
  Y As Single
End Type

Type GameData
  Current As XYData
  CurrentBoardNo As Long
  CurrentPlayer As Long
  PlayerScore(1 To 2) As Long
  PlayerName(1 To 2) As String
  Board(1 To 9) As Long
  
End Type

Public GD As GameData

Public Sub SetCurrentSquare(Selection As Long)
Select Case Selection
    Case 16711935 '1
      GD.Current.X = TileSizeX
      GD.Current.Y = 0
      GD.CurrentBoardNo = 1
    Case 65280 '2
      GD.Current.X = 369
      GD.Current.Y = 41
      GD.CurrentBoardNo = 2
    Case 42495 '3
      GD.Current.X = 492
      GD.Current.Y = 82
      GD.CurrentBoardNo = 3
    Case 29456 '4
      GD.Current.X = 123
      GD.Current.Y = 41
      GD.CurrentBoardNo = 4
    Case 16711680 '5
      GD.Current.X = TileSizeX
      GD.Current.Y = 82
      GD.CurrentBoardNo = 5
    Case 255 '6
      GD.Current.X = 369
      GD.Current.Y = 123
      GD.CurrentBoardNo = 6
    Case 65535 '7
      GD.Current.X = 0
      GD.Current.Y = 82
      GD.CurrentBoardNo = 7
    Case 16776960 '8
      GD.Current.X = 123
      GD.Current.Y = 123
      GD.CurrentBoardNo = 8
    Case 11862181 '9
      GD.Current.X = TileSizeX
      GD.Current.Y = 164
      GD.CurrentBoardNo = 9
    Case Else
      GD.CurrentBoardNo = 0
  End Select
End Sub

Public Sub Resetboard()
  Dim i
  For i = 1 To 9
    GD.Board(i) = 0
  Next i
End Sub
Public Function CheckBoardFull() As Boolean
  Dim i
  CheckBoardFull = True
  For i = 1 To 9
    If GD.Board(i) = 0 Then CheckBoardFull = False
  Next i
End Function
Public Function GameWon(bPlayer As Long) As Boolean
Dim Checking As Boolean
'MsgBox "check"
  Checking = False
  If GD.Board(1) = bPlayer Then
    If GD.Board(2) = bPlayer Then
      If GD.Board(3) = bPlayer Then Checking = True
    Else
      If GD.Board(4) = bPlayer Then
        If GD.Board(7) = bPlayer Then Checking = True
      Else
        If GD.Board(5) = bPlayer Then
          If GD.Board(9) = bPlayer Then Checking = True
        End If
      End If
    End If
  End If
  If GD.Board(4) = bPlayer Then
    If GD.Board(5) = bPlayer Then
      If GD.Board(6) = bPlayer Then Checking = True
    End If
  End If
  If GD.Board(2) = bPlayer Then
    If GD.Board(5) = bPlayer Then
      If GD.Board(8) = bPlayer Then Checking = True
    End If
  End If
  If GD.Board(3) = bPlayer Then
    If GD.Board(6) = bPlayer Then
      If GD.Board(9) = bPlayer Then Checking = True
    Else
      If GD.Board(5) = bPlayer Then
        If GD.Board(7) = bPlayer Then Checking = True
      End If
    End If
  End If
  If GD.Board(7) = bPlayer Then
    If GD.Board(8) = bPlayer Then
      If GD.Board(9) = bPlayer Then Checking = True
    End If
  End If
  GameWon = Checking

End Function
Public Sub ComputersGo()
Dim RanNumber As Long, LoopCounter As Long

'computer checks to see if it can win
  For LoopCounter = 1 To 9
    If GD.Board(LoopCounter) = 0 Then
      GD.Board(LoopCounter) = 2
      If GameWon(2) Then
        Exit Sub
      Else
        GD.Board(LoopCounter) = 0
      End If
    End If
  Next LoopCounter
  
'if Computer can not win then Computer checks to see if player can win on next go,
'if player can win then computer will block the player
  For LoopCounter = 1 To 9
    If GD.Board(LoopCounter) = 0 Then
      GD.Board(LoopCounter) = 1
      If GameWon(1) Then
        GD.Board(LoopCounter) = 2
        Exit Sub
      Else
        GD.Board(LoopCounter) = 0
      End If
    End If
   Next LoopCounter
  
  'if No Player can win then Computer will select a random square
  Do
    Randomize Timer
    RanNumber = Int(Rnd * 9) + 1
  Loop Until CheckBoardFull Or GD.Board(RanNumber) = 0
  GD.Board(RanNumber) = 2
End Sub
