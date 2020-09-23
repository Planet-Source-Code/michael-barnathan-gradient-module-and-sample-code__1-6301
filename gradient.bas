Attribute VB_Name = "Module1"
Private Type ColorCode
    Red As Long
    Green As Long
    Blue As Long
End Type

Sub CreateGradient(TargetControl As Object, Red1 As Integer, Green1 As Integer, Blue1 As Integer, Red2 As Integer, Green2 As Integer, Blue2 As Integer, Pattern As Integer, StartSteps, EndSteps, NoClear)
'Red1 - starting red attribute
'Red2 - ending red attribute
'Same with green and blue

'Pattern 1 is a horizontal gradient
'Pattern 2 is vertical
'Pattern 3 is circular - takes a little more time to generate

'Steps is how many color changes the gradient will take
'The more steps, the better the quality, but the slower
'The gradient will draw
If NoClear = 0 Then TargetControl.Cls
B = Red2 - Red1
C = Green2 - Green1
D = Blue2 - Blue1
E = EndSteps
F = StartSteps
If Pattern = 1 Then
For A = F To E
TargetControl.Line ((TargetControl.Width / E) * (A - 1), 1)-((TargetControl.Width / E) * A, TargetControl.Height), RGB(Abs(Red1 + Int((B / E) * A)), Abs(Green1 + Int((C / E) * A)), Abs(Blue1 + Int((D / E) * A))), BF
Next
ElseIf Pattern = 2 Then
For A = F To E
TargetControl.Line (1, (TargetControl.Height / E) * (A - 1))-(TargetControl.Width, (TargetControl.Height / E) * A), RGB(Abs(Red1 + Int((B / E) * A)), Abs(Green1 + Int((C / E) * A)), Abs(Blue1 + Int((D / E) * A))), BF
Next
ElseIf Pattern = 3 Then
For A = F To E
TargetControl.Circle (TargetControl.Width / 2, TargetControl.Height / 2), ((TargetControl.Width / 2) / E) * A, RGB(Abs(Red1 + Int((B / E) * A)), Abs(Green1 + Int((C / E) * A)), Abs(Blue1 + Int((D / E) * A))), BF
Next
End If
End Sub

Function SeperateColors(ImageControl As Control) As ColorCode
'Seperates the image into red, green, and blue colors
SeperateColors.Blue = ImageControl.BackColor And RGB(0, 0, 255) '&HFF0000 \ 65536
SeperateColors.Blue = SeperateColors.Blue / 65536
SeperateColors.Green = ImageControl.BackColor And RGB(0, 255, 0)
SeperateColors.Green = SeperateColors.Green / 256
SeperateColors.Red = ImageControl.BackColor And RGB(255, 0, 0) '&HFF&
End Function


