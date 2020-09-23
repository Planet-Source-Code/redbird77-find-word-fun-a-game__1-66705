Attribute VB_Name = "mGame"
' mGame.bas \ redbird77@earthlink.net \ 2006.10.02

Option Explicit

Public Type udtPoint
    X           As Integer
    Y           As Integer
End Type

Private Type udtSquare
    Text        As String
    Used        As Boolean              ' used in a valid solution
End Type

Public Enum DirectionConstants
    ' right-to-left directions
    DirEast
    DirSouth
    DirNorthEast
    DirSouthEast
    
    ' left-to-right directions
    DirWest
    DirNorth
    DirSouthWest
    DirNorthWest
    
    DirInvalid
    DirNone
End Enum

Private Type udtSolution
    X           As Integer              ' board position of first letter
    Y           As Integer
    Direction   As DirectionConstants
    Text        As String
    Found       As Boolean
End Type

Private Type udtSettings
    SquareSize      As Integer          ' size of each square (in pixels)
    SquareCount     As Integer          ' how many squares across board?
    SelectColor     As Long             ' solution candidate color
    SolutionColor   As Long             ' valid solution color
    Backwards       As Integer          ' are soultions placed "backwards"?
End Type

Private Type udtGame
    Active      As Boolean              ' are we having fun?
    Words()     As String               ' holds words to be searched for
    Board()     As udtSquare            ' represents squares in board
    Solutions() As udtSolution          ' holds info about all soultions
    Candidate   As udtSolution          ' currently selected "word"
    Selecting   As Boolean              ' is user currently selecting?
    Settings    As udtSettings
End Type: Public Game As udtGame        ' a little nod to C-syntax ;)

Private FirstMove   As Boolean
Private Initial     As udtPoint
Private Old()       As udtPoint
Private Squares     As Integer          ' Squares = SquareCount to keep Draw & Resize
                                        ' from referencing SquareCount which cannot
                                        ' change during the game

Public Function Draw() As Long

' Draw entire board square by square.

Dim X   As Integer
Dim Y   As Integer

    fGame.pBoard.Cls
    
    For Y = 0 To Squares - 1
        For X = 0 To Squares - 1
            Call DrawSquare(X, Y)
        Next
    Next
    
    fGame.pBoard.Refresh
    
End Function

Public Function Play() As Long

' The start of it all..

    Squares = Game.Settings.SquareCount
    
    If mGame.Initialize Then
        Call mGame.Create
        Call mGame.Draw
        Game.Active = True
    End If
    
End Function

Public Sub ResizeBoard()

    With fGame
        .pBoard.Width = Game.Settings.SquareSize * Squares
        .pBoard.Height = .pBoard.Width
        .tvWords.Height = .pBoard.Height
    
        .Width = (.pBoard.Left + .pBoard.Width + 15) * Screen.TwipsPerPixelX
        .Height = (.tvWords.Top + .tvWords.Height + .cmdPlay.Height) * Screen.TwipsPerPixelY
    
        ' Set font size.
        ' SqaureSize (in pixels) * Screen.TwipsPerPixelX = Twips per Square
        ' Twips per Square / 20 (approx Twips per Point) = new font size.
        ' Used 30 so the letters would not take up the entire square.
        .pBoard.Font.Size = (Game.Settings.SquareSize * Screen.TwipsPerPixelX) / 30
    End With
    
End Sub

Public Function OnMouseDown(ByVal X As Integer, ByVal Y As Integer) As Long

' Start selecting solution candidate.

    If Not Game.Active Then Exit Function
    
    Game.Selecting = True
    FirstMove = True

    ' Get board index from clicked coordinates.
    Initial.X = X \ Game.Settings.SquareSize
    Initial.Y = Y \ Game.Settings.SquareSize
    
End Function

Public Function OnMouseMove(ByVal X As Integer, ByVal Y As Integer) As Long

Dim I           As Integer
Dim X2          As Integer
Dim Y2          As Integer
Dim Count       As Integer
Dim Direction   As DirectionConstants
Dim Point       As udtPoint

    If Not Game.Selecting Then Exit Function
    
    ' Get board index from current mouse coordinates.
    Point.X = X \ Game.Settings.SquareSize
    Point.Y = Y \ Game.Settings.SquareSize
    
    ' Get direction from initial to current.
    Direction = GetDirection(Point, Initial)
    
    If Direction = DirInvalid Then Exit Function
    
    ' Erase old solution candidate.
    If FirstMove Then
        FirstMove = False
    Else
        For I = 0 To UBound(Old)
            Call DrawSquare(Old(I).X, Old(I).Y)
        Next
    End If

    ' Get steps from initial to current.
    Select Case Direction
        Case DirEast: X2 = 1: Y2 = 0
        Case DirSouth: X2 = 0: Y2 = 1
        Case DirNorthEast: X2 = 1: Y2 = -1
        Case DirSouthEast: X2 = 1: Y2 = 1
        Case DirWest: X2 = -1: Y2 = 0
        Case DirNorth: X2 = 0: Y2 = -1
        Case DirSouthWest: X2 = -1: Y2 = 1
        Case DirNorthWest: X2 = -1: Y2 = -1
    End Select
        
    ' Highlight initial.
    Call DrawSquare(Initial.X, Initial.Y, True)
                
    ' Create array to hold solution candidate's points.
    ReDim Old(Count)
    Old(Count) = Initial
    Count = Count + 1
          
    ' Draw solution candidate.
    Do Until (Point.X = Initial.X) And (Point.Y = Initial.Y)
 
        Call DrawSquare(Point.X, Point.Y, True)

        ' Save solution candidate's points for erasing if invalid.
        ReDim Preserve Old(Count)
        Old(Count) = Point
        Count = Count + 1

        Point.X = Point.X - X2
        Point.Y = Point.Y - Y2
    Loop
    
    With Game.Candidate
        .Text = Game.Board(Old(0).X, Old(0).Y).Text
        .X = Old(0).X
        .Y = Old(0).Y
        .Direction = Direction
    End With
    
    ' Save current highlighted as solution candidate.
    For I = UBound(Old) To 1 Step -1
        Game.Candidate.Text = Game.Candidate.Text & Game.Board(Old(I).X, Old(I).Y).Text
    Next

' This is what was causing the freezing-when-compiled behavior.
'ErrHandler:

    ' Array not dimensioned in the first pass, could also use a boolean flag.
    'If Err.Number = 9 Then Resume Next
    
End Function

Public Function OnMouseUp(ByVal X As Integer, ByVal Y As Integer) As Long
   
Dim I       As Integer
Dim J       As Integer
Dim Found   As Boolean

    If Not Game.Active Then Exit Function

    Game.Selecting = False
    
    ' Check candidate solution against all solutions.
    For I = 0 To UBound(Game.Solutions)

        ' TODO : Allow valid words to be selected from end as well.
        
        ' If candidate starts at same coords as a solution..
        If Game.Candidate.X = Game.Solutions(I).X And _
           Game.Candidate.Y = Game.Solutions(I).Y Then
            
            If Game.Candidate.Text = Game.Solutions(I).Text Then
            
                ' Mark candidate solution's points as used (valid).
                For J = 0 To UBound(Old)
                    Game.Board(Old(J).X, Old(J).Y).Used = True
                Next
                
                Found = True
            End If
        
            If Found Then Exit For
        End If
    Next
    
    If Found Then
    
        fGame.tvWords.Nodes(I + 1).Bold = False
        Call fGame.tvWords.Nodes(I + 1).EnsureVisible
        
        Game.Solutions(I).Found = True
        
        ' Check if all solutions are found.
        For I = 0 To UBound(Game.Solutions)
            
            If Not Game.Solutions(I).Found Then
                Exit For
            End If
            
        Next
        
        If I > UBound(Game.Solutions) Then
            MsgBox "You win." & vbCrLf & vbCrLf & _
                   "That's all." & vbCrLf & vbCrLf & _
                   "Thank you, come again.", vbInformation, App.Title
            Game.Active = False
        End If
        
        'Debug.Print "Found " & Game.Candidate.Text & " at (" & _
                    Game.Candidate.X& ", " & Game.Candidate.Y & _
                    ") with direction=" & DirToWord(Game.Candidate.Direction)
    Else
        'Debug.Print "Not Valid"
    End If
    
    Call mGame.Draw

End Function

Private Function Create() As Long

Dim I       As Integer
Dim J       As Integer
Dim K       As Integer
Dim Length  As Integer
Dim Word    As String           ' word being placed
Dim Tries   As Long             ' attempts to place word on board
Dim Count   As Integer
Dim W()     As String           ' holds each letter of word
Dim X       As Integer
Dim Y       As Integer
Dim eDir    As DirectionConstants
Dim Nd      As Node

Const MAX_TRIES As Long = 20000

    For J = 0 To UBound(Game.Words)

' Pure unadulterated eeeevil - GOTO!
StartTrying:
        
        ' Reset try count.
        Tries = 0
        
        ' Get word from list.
        Word = Game.Words(J)
        Length = Len(Word)
        
        ' Transform word to array for easy access.
        ReDim W(Len(Word) - 1)
        For K = 0 To UBound(W)
            W(K) = Mid$(Word, K + 1, 1)
        Next

' More cowbell... oops I mean evil.
PlaceWord:
    
        Tries = Tries + 1
    
        ' If exceed max no. of tries for one word, skip to the next word in list.
        If Tries > MAX_TRIES Or Len(Word) > Squares Then
        
            Debug.Print "Cannot place word: " & Word
        
            If J = UBound(Game.Words) Then
                GoTo FillRandom
            Else
                J = J + 1
                GoTo StartTrying
            End If
            
        End If
         
        ' Pick random starting point and direction.
        X = Int(Rnd * Squares)
        Y = Int(Rnd * Squares)
        
        If Game.Settings.Backwards Then
            eDir = Int(Rnd * DirInvalid)
        Else
            eDir = Int(Rnd * DirWest)
        End If
        
        Select Case eDir
            Case DirEast
            
                ' Can word fit?
                If X + Length <= Squares Then
                
                    ' Check each letter.
                    For I = 0 To Length - 1
                        
                        ' Not valid letter, try another place.
                        If IsBad(W(I), Game.Board(X + I, Y).Text) Then
                            GoTo PlaceWord
                        End If
                        
                    Next
                    
                    ' All letters valid, place word on board.
                    For I = 0 To Length - 1
                        Game.Board(X + I, Y).Text = W(I)
                    Next
                Else
                    ' Word cannot fit, try another place.
                    GoTo PlaceWord
                End If
                
            Case DirSouth
                If Y + Length <= Squares Then
                    For I = 0 To Length - 1
                        If IsBad(W(I), Game.Board(X, Y + I).Text) Then
                            GoTo PlaceWord
                        End If
                    Next
                    
                    For I = 0 To Length - 1
                        Game.Board(X, Y + I).Text = W(I)
                    Next
                Else
                    GoTo PlaceWord
                End If
            
            Case DirNorthEast
                If (X + Length <= Squares) And (Y - Length >= -1) Then
                    For I = 0 To Length - 1
                        If IsBad(W(I), Game.Board(X + I, Y - I).Text) Then
                            GoTo PlaceWord
                        End If
                    Next
                    
                    For I = 0 To Length - 1
                        Game.Board(X + I, Y - I).Text = W(I)
                    Next
                Else
                    GoTo PlaceWord
                End If
                
            Case DirSouthEast
                If (X + Length <= Squares) And (Y + Length <= Squares) Then
                    For I = 0 To Length - 1
                        If IsBad(W(I), Game.Board(X + I, Y + I).Text) Then
                            GoTo PlaceWord
                        End If
                    Next
                    
                    For I = 0 To Length - 1
                        Game.Board(X + I, Y + I).Text = W(I)
                    Next
                Else
                    GoTo PlaceWord
                End If
               
            Case DirWest
                If X - Length >= -1 Then
                    For I = 0 To Length - 1
                        If IsBad(W(I), Game.Board(X - I, Y).Text) Then
                            GoTo PlaceWord
                        End If
                    Next
                    
                    For I = 0 To Length - 1
                        Game.Board(X - I, Y).Text = W(I)
                    Next
                Else
                    GoTo PlaceWord
                End If
        
            Case DirNorth
                If Y - Length >= -1 Then
                    For I = 0 To Length - 1
                        If IsBad(W(I), Game.Board(X, Y - I).Text) Then
                            GoTo PlaceWord
                        End If
                    Next
            
                    For I = 0 To Length - 1
                        Game.Board(X, Y - I).Text = W(I)
                    Next
                Else
                    GoTo PlaceWord
                End If
                
            Case DirSouthWest
                If (X - Length >= -1) And (Y + Length <= Squares) Then
                    For I = 0 To Length - 1
                        If IsBad(W(I), Game.Board(X - I, Y + I).Text) Then
                            GoTo PlaceWord
                        End If
                    Next
                    
                    For I = 0 To Length - 1
                        Game.Board(X - I, Y + I).Text = W(I)
                    Next
                Else
                    GoTo PlaceWord
                End If
                
            Case DirNorthWest
                If (X - Length >= -1) And (Y - Length >= -1) Then
                    For I = 0 To Length - 1
                        If IsBad(W(I), Game.Board(X - I, Y - I).Text) Then
                            GoTo PlaceWord
                        End If
                    Next
                    
                    For I = 0 To Length - 1
                        Game.Board(X - I, Y - I).Text = W(I)
                    Next
                Else
                    GoTo PlaceWord
                End If
        End Select
      
        ' Add word to solution array property.
        ReDim Preserve Game.Solutions(Count)
        With Game.Solutions(Count)
            .X = X
            .Y = Y
            .Direction = eDir
            .Text = Word
        End With
        
        Count = Count + 1
        
        ' Add word to GUI list.
        Set Nd = fGame.tvWords.Nodes.Add(, , Word, Word)
        Nd.Bold = True

        'Debug.Print "Took " & Tries & " tries to place: " & Word
    Next
    
FillRandom:
    
    ' Fill in unused squares with random letters.
    For Y = 0 To Squares - 1
        For X = 0 To Squares - 1
            If Game.Board(X, Y).Text = "" Then
                Game.Board(X, Y).Text = Chr$(Int(Rnd * 26) + 65)
            End If
        Next
    Next

    'Debug.Print "Placed " & Count & " words of " & UBound(Game.Words) + 1
        
End Function

Private Function DirToWord(ByVal Direction As DirectionConstants) As String

' For debugging purposes.

    Select Case Direction
        Case DirEast: DirToWord = "DirEast"
        Case DirSouth: DirToWord = "DirSouth"
        Case DirNorthEast: DirToWord = "DirNorthEast"
        Case DirSouthEast: DirToWord = "DirSouthEast"
        Case DirWest: DirToWord = "DirWest"
        Case DirNorth: DirToWord = "DirNorth"
        Case DirSouthWest: DirToWord = "DirSouthWest"
        Case DirNorthWest: DirToWord = "DirNorthWest"
        Case DirInvalid: DirToWord = "DirInvalid"
    End Select
    
End Function

Private Function DrawSquare(ByVal X As Integer, _
                            ByVal Y As Integer, _
                            Optional Selecting As Boolean) As Long
                            
' Draw the square specified by the Board array's coordinates (which are
' translated into actual coordinates).

Dim Fill    As Long
Dim Char    As String
    
    ' If selecting then draw in SelectColor, else in SolutionColor.
    If Selecting Then
        Fill = Game.Settings.SelectColor
    Else
        Fill = IIf(Game.Board(X, Y).Used, Game.Settings.SolutionColor, vbWhite)
    End If
    
    With Game.Settings
        fGame.pBoard.FillColor = Fill
        
        fGame.pBoard.Line (X * .SquareSize, Y * .SquareSize)-Step _
                          (.SquareSize, .SquareSize), 0, BF
        
        fGame.pBoard.Line (X * .SquareSize, Y * .SquareSize)-Step _
                          (.SquareSize, .SquareSize), 0, B
                 
        Char = Game.Board(X, Y).Text
    
        fGame.pBoard.CurrentX = X * .SquareSize + (.SquareSize \ 2) - _
                                    (fGame.pBoard.TextWidth(Char) \ 2)
        fGame.pBoard.CurrentY = Y * .SquareSize + (.SquareSize \ 2) - _
                                    (fGame.pBoard.TextHeight(Char) \ 2)
        
        fGame.pBoard.Print UCase$(Char)
    End With
    
End Function

Private Function GetDirection(ByRef P1 As udtPoint, _
                              ByRef P2 As udtPoint) As DirectionConstants
                              
' Returns direction of point 1 from point 2.
                                     
Dim X   As Integer
Dim Y   As Integer
Dim D   As DirectionConstants

    GetDirection = DirInvalid

    ' Get X and Y differences.
    X = P1.X - P2.X
    Y = P1.Y - P2.Y
    
    ' There is perhaps (or, of course ;) a more elegant way to do this.
    ' Maybe a little trigonometry, something with the Atn function, right?
    Select Case X
        Case Is < 0
            D = DirWest
    
            If Abs(X) = Abs(Y) Then
                If Y < 0 Then
                    D = DirNorthWest
                ElseIf Y > 0 Then
                    D = DirSouthWest
                End If
            ElseIf Y Then
                D = DirInvalid
            End If
        Case Is > 0
            D = DirEast
    
            If Abs(X) = Abs(Y) Then
                If Y < 0 Then
                    D = DirNorthEast
                ElseIf Y > 0 Then
                    D = DirSouthEast
                End If
            ElseIf Y Then
                D = DirInvalid
            End If
        Case Is = 0
          If Y < 0 Then
              D = DirNorth
          ElseIf Y > 0 Then
              D = DirSouth
          Else
              D = DirNone
          End If
    End Select
    
    GetDirection = D
    
End Function

Private Function Initialize() As Long

Dim F       As Integer
Dim I       As Integer
Dim J       As Integer
Dim Count   As Integer

On Error GoTo ErrHandler
    
    Randomize
    
    ' Read words from vbCrLf delimited file.
    F = FreeFile
    
    fGame.cdlOpen.FileName = ""
    fGame.cdlOpen.InitDir = App.Path
    fGame.cdlOpen.ShowOpen
    
    Erase Game.Board()
    Erase Game.Words()
    Erase Game.Solutions()
    
    Open fGame.cdlOpen.FileName For Input As #F
        While Not EOF(F)
            ReDim Preserve Game.Words(Count)
            Line Input #F, Game.Words(Count)
            Count = Count + 1
        Wend
    Close #F
    
    ' Word cannot be completely contained in any other word.
    For I = 0 To UBound(Game.Words)
        For J = 0 To UBound(Game.Words)
            If I <> J Then
            
                ' If word in word is found set word to a long string of Xs - a
                ' quick way to assure it will be rejected for placement.
                
                If InStr(Game.Words(I), Game.Words(J)) Then
                    'Debug.Print "Rejecting: " & Game.Words(J) & ", Found In: " & Game.Words(I)
                    Game.Words(J) = String$(1000, "x")
                End If
            End If
        Next
    Next
    
    ReDim Game.Board(Squares - 1, Squares - 1)
    ReDim Used(Squares - 1, Squares - 1)
    
    Call mGame.ResizeBoard
    fGame.tvWords.Nodes.Clear
    
    Initialize = 1
    
    Exit Function
    
ErrHandler:
    If Err.Number = cdlCancel Then
        Exit Function
    Else
        Resume Next
    End If
    
End Function

Private Function IsBad(ByVal C1 As String, ByVal C2 As String) As Boolean

' Does the current character of the word match the exisiting character on the
' board (or is the board blank)?

' C1 = word char
' C2 = board char
    
    IsBad = CBool(C2 <> "" And C2 <> C1)
    
End Function
