Attribute VB_Name = "Module1"


Global Mouse As New CMouse

 
 
 
 Sub MoveMouse(X As Integer, Y As Integer)
'Move the mouse in locked window

Mouse.X = CLng(CDbl(X))
Mouse.Y = CLng(CDbl(Y))

End Sub
