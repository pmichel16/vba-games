Attribute VB_Name = "SnakeGame"
'
'Generates a random integer, and sets the range of the apple there.
'
Function generateApple() As Range
Dim rowz As Integer, col As Integer
    rowz = Int((30) * Rnd + 1)
    col = Int((30) * Rnd + 1)
    Set generateApple = Sheet1.Range(Cells(rowz, col).Address)
    
End Function

'
'Moves the head (selection) in the direction specified.
'
Function moveHead(lastHead As Range, dir As Integer) As Range
Dim head As Range

lastHead.Activate
Select Case dir
    Case vbKeyLeft:
        ActiveCell.Offset(0, -1).Select
        Set moveHead = Selection
    Case vbKeyUp:
        ActiveCell.Offset(-1, 0).Select
        Set moveHead = Selection
    Case vbKeyRight:
        ActiveCell.Offset(0, 1).Select
        Set moveHead = Selection
    Case vbKeyDown:
        ActiveCell.Offset(1, 0).Select
        Set moveHead = Selection
End Select

End Function
