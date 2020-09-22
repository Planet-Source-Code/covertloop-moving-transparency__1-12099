Attribute VB_Name = "PauseBas"

Function Pause(interval)
current = Timer
Do While Timer - current < Val(interval)
DoEvents
Loop
End Function

