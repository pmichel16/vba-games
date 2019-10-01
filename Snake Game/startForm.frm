VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} startForm 
   Caption         =   "Snake Game"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "startForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "startForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim direction As Integer, length As Integer

Private Sub StartButton_Click()
Dim head As Range, oldHead As Range, square As Range, apple As Range, item As Range
Dim body As New Collection
Dim pauseTime As Double
Dim start As Single
Dim bCount As Integer
Dim inside As Boolean

'Place the first apple
Set apple = SnakeGame.generateApple
apple.Interior.Color = RGB(255, 0, 0)

'Start the snake's head at A1.
Set head = Sheet1.Range("A1")
pauseTime = 0.25
direction = vbKeyRight
length = 3
i = 1
continuing = True

While continuing
    start = Timer    ' Set start time.
    Do While Timer < start + pauseTime
        DoEvents    ' Yield to other processes.
    Loop
    head.Interior.Color = RGB(0, 0, 0)
    Set oldHead = head
    
    'Color in the rest of the snake's body, which is stored in the Collection.
    If i = 1 Then
        body.Add item:=oldHead
    ElseIf body.Count <= length Then
        body.Add item:=oldHead, Before:=1
    Else
        body.Add item:=oldHead, Before:=1
        body(body.Count).Interior.Color = RGB(255, 255, 255)
        body.Remove (body.Count)
    End If
    
    'Make sure there hasn't been a collision.
    If head.Column > 30 Or (head.Column = 1 And direction = 37) Or head.Row > 30 Or (head.Row = 1 And direction = 38) Then
        continuing = False
        MsgBox ("You lose!" & Chr(10) & "Score: " & length + 1)
        GoTo ending
    End If
    
    'Generate the new snake head
    Set head = SnakeGame.moveHead(oldHead, direction)
    
    'Ensure the head is not colliding with the body.
    For Each item In body
        If item.Row = head.Row And item.Column = head.Column Then
            continuing = False
            MsgBox ("You lose!" & Chr(10) & "Score: " & length + 1)
            GoTo ending
        End If
    Next item

    'Generate a new apple, and lengthen the snake, when the previous apple is eaten
    If (head.Column = apple.Column And head.Row = apple.Row) Then
        length = length + 1
        Set apple = SnakeGame.generateApple
        
        'Make sure the apple is not inside the body, and generate a new one if it is.
        inside = True
        While (inside)
        For Each item In body
            If item.Row = apple.Row And item.Column = apple.Column Then
                Set apple = SnakeGame.generateApple
                GoTo forEnd
            End If
        Next item
        inside = False
forEnd:
        Wend
        apple.Interior.Color = RGB(255, 0, 0)
    End If

i = i + 1
Wend

'Reset the sheet after losing.
ending:
Sheet1.Cells.Interior.ColorIndex = 0
Range("A1:AD30").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
Range("A1").Select
    
End Sub

Private Sub StartButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyRight Or KeyCode = vbKeyDown Then
    direction = KeyCode
    KeyCode = 0
End If
End Sub
