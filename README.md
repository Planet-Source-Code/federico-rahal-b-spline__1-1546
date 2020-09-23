<div align="center">

## B\-Spline


</div>

### Description

Draws a B-Spline over a PictureBox while the user inputs a series of points with the mouse.

It also allows to drag the Control Points of the B-Spline to modify it
 
### More Info
 
'Create a new project, Project1 is created by default

'Add a new module and name it modSpline

'

'Change the name of Form1 to frmSpline

'Add a PictureBox, this is where you are going to draw the Spline

'Change the following properties of Picture1:

' Name= picDraw

' ScaleMode = 3 'Pixels

' BackColor = vbWhite

'Add also the following

' Two Option Buttons:

' OpMode(0) and OpMode(1)

' Opmode(0).Caption = "Move"

' Opmode(1).Caption = "Draw"

'

'Add One Command Button named cmdClear

' Caption = "Clear"

'Add Three Labels

'Label1:

' Name=lblT

' Caption="Degree T"

'Label2:

' Name=lblRes

' Caption="Resolution"

'Label3:

' Name=lblLen

' Caption="Spline Length"

'Inside the pictureBox add one label:

' Name=lblGrip

' Index = 0 'Very important

' BackColor = vbRed

' Height = 3

' Width = 3

' Visible = False

'Add a ComboBox

' Name=cboDegree

' Style = 2 'DropDownList

'Add a TextBox Named txtRes:

' Text ="5"

'

'Add a menu ' Edit mnuEdit

'and a subitem Delete mnuDelete

'Set the visible property of mnuEdit to False

'

'Returns the Outp() array filled with the points along the b-Spline


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Federico Rahal](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/federico-rahal.md)
**Level**          |Unknown
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/federico-rahal-b-spline__1-1546/archive/master.zip)





### Source Code

```
'
'Add the following code to modSpline
'
Option Explicit
Public Type POINTAPI
 X As Long
 Y As Long
End Type
Public inp() As POINTAPI
Public outp() As POINTAPI
Public N As Integer
Public T As Integer
Public RESOLUTION As Integer
' Example of how to call the spline functions
' Basically one needs to create the control points, then compute
' the knot positions, then calculate points along the curve.
'
'1. You have to define two arrays of the Type POINTAPI
' 'Dim inp() As POINTAPI, outp() as POINTAPI
'2. Define te array of Knots as integer
' 'Dim knots() As Integer
' Define Three more variables
' N as integer : number of entries in inp()-1 '
' T as integer : The blending factor usually 3
'  a value of 2 draws the polyline
' RESOLUTION as integer : The number of segments in which the whole
'  spline will be divided
'  I prefer to calculate the resolution after the inp() array is filled
'  that's a way to ensure a proper resolution
'   e.g resolution = 10 * N or
'  you can enter a constant resolution regardless of the length of the
'  of the spline e.g RESOLUTION = 200
'
'3. Fill the input array either by code or interactively by clicking
' in the destination form or picturebox
'4. Once you have the filled inp() array, you have to fill the rest of the variables
'
' N = UBound(inp) - 1
' RESOLUTION = 10*n
' T=3
' Redim knots(N + T + 1)
' Redim outp(RESOLUTION)
' Now it's time to call the Functions
'
' Call SplineKnots(knots(), N, T)
' Call SplineCurve(inp(), N, knots(), T, outp(), RESOLUTION)
'
' SplineCurve Returns outp() filled with the points along the Spline
'
' To draw the spline do the following:
'Dim i as integer
'For i = 0 To RESOLUTION
'  Form1.Picture1.Line (outp(i-1).x, outp(i-1).y) - (outp(i).x, outp(i).y)
'Next
'
' That's all to it. Enjoy!
'
'SPLINEPOINT
'This returns the point "output" on the spline curve.
'The parameter "v" indicates the position, it ranges from 0 to n-t+2
Private Function SplinePoint(u() As Integer, N As Integer, T As Integer, v As Single, Control() As POINTAPI, output As POINTAPI)
Dim k As Integer
Dim b As Single
output.X = 0: output.Y = 0 ': output.Z = 0
For k = 0 To N
 b = SplineBlend(k, T, u(), v)
  output.X = output.X + Control(k).X * b
  output.Y = output.Y + Control(k).Y * b
  'for a 3D b-Spline use the following
  ' output.Z = output.Z + Control(k).Z * b
Next
End Function
'SPLINEBLEND
'Calculate the blending value, this is done recursively.
'If the numerator and denominator are 0 the expression is 0.
'If the deonimator is 0 the expression is 0
Private Function SplineBlend(k As Integer, T As Integer, u() As Integer, v As Single) As Single
Dim value As Single
 If T = 1 Then
  If (u(k) <= v And v < u(k + 1)) Then
   value = 1
   Else
   value = 0
  End If
 Else
  If ((u(k + T - 1) = u(k)) And (u(k + T) = u(k + 1))) Then
   value = 0
  ElseIf (u(k + T - 1) = u(k)) Then
   value = (u(k + T) - v) / (u(k + T) - u(k + 1)) * SplineBlend(k + 1, T - 1, u, v)
  ElseIf (u(k + T) = u(k + 1)) Then
   value = (v - u(k)) / (u(k + T - 1) - u(k)) * SplineBlend(k, T - 1, u, v)
  Else
   value = (v - u(k)) / (u(k + T - 1) - u(k)) * SplineBlend(k, T - 1, u, v) + _
     (u(k + T) - v) / (u(k + T) - u(k + 1)) * SplineBlend(k + 1, T - 1, u, v)
  End If
 End If
SplineBlend = value
End Function
'SPLINEKNOTS
' The positions of the subintervals of v and breakpoints, the position
' on the curve are called knots. Breakpoints can be uniformly defined
' by setting u(j) = j, a more useful series of breakpoints are defined
' by the function below. This set of breakpoints localises changes to
' the vicinity of the control point being modified.
Public Sub SplineKnots(u() As Integer, N As Integer, T As Integer)
Dim j As Integer
For j = 0 To N + T
  If j < T Then
   u(j) = 0
  ElseIf (j <= N) Then
   u(j) = j - T + 1
  ElseIf (j > N) Then
   u(j) = N - T + 2
  End If
Next
End Sub
'SPLINECURVE
' Create all the points along a spline curve
' Control points "inp", "n" of them. Knots "knots", degree "t".
' Ouput curve "outp", "res" of them.
Public Sub SplineCurve(inp() As POINTAPI, N As Integer, knots() As Integer, T As Integer, outp() As POINTAPI, res As Integer)
Dim i As Integer
Dim interval As Single, increment As Single
interval = 0
increment = (N - T + 2) / (res - 1)
 For i = 0 To res - 1 '{
  Call SplinePoint(knots(), N, T, interval, inp(), outp(i))
  interval = interval + increment
 Next
  outp(res - 1) = inp(N)
End Sub
'EOF() module modSpline
'
'
'
'The following code goes in frmSpline
'
Option Explicit
Dim selGrip As Label
Dim mode As Integer
Private Sub cboDegree_Click()
If Not Me.Visible Then Exit Sub
 eraseSpline
 DrawSpline
End Sub
Private Sub cmdClear_Click()
Dim i As Integer
lblGrip(0).Visible = False
For i = 1 To lblGrip.UBound
 Unload lblGrip(i)
Next
ReDim inp(0)
N = 0
ReDim outp(RESOLUTION)
PicDraw.Cls
lblLen = "Spline Length: 0"
cboDegree.Enabled = False
txtRes.Enabled = False
End Sub
Private Sub Form_Load()
With cboDegree
 .AddItem "1"
 .AddItem "2"
 .AddItem "3"
 .AddItem "4"
 .AddItem "5"
 .ListIndex = 2
 .Enabled = False
End With
txtRes.Enabled=False
RESOLUTION = 5
End Sub
Private Sub mnuDelete_Click()
delGrip
End Sub
Private Sub OpMode_Click(Index As Integer)
mode = Index
End Sub
Private Sub lblGrip_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Set selGrip = lblGrip(Index)
If Button = vbLeftButton Then
 lblGrip(Index).Drag
Else
 PopupMenu mnuEdit
End If
End Sub
Private Sub PicDraw_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
Source.Move X, Y
eraseSpline
inp(Source.Index).X = X
inp(Source.Index).Y = Y
DrawSpline
End Sub
Private Sub PicDraw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim tmp As Integer
Static sErase As Boolean
If Button = vbRightButton Then Exit Sub
If mode = 1 Then 'Drawing mode
 ReDim Preserve inp(N)
 inp(N).X = X: inp(N).Y = Y
 If N > 0 Then Load lblGrip(N)
 With lblGrip(N)
  .Move X - .Width \ 2, Y - .Height \ 2
  .Visible = True
 End With
 N = N + 1
 If N >= 3 Then
 cboDegree.Enabled = True
 txtRes.Enabled = True
 If sErase Then eraseSpline
  DrawSpline
  sErase = True
 End If
End If
Set selGrip = Nothing
End Sub
Private Sub DrawSpline()
Dim i As Integer
Dim knots() As Integer
Dim sLen As Single
Dim h!, d!
Dim sRes As Integer
sRes = RESOLUTION * N
 T = CInt(cboDegree.ListIndex + 1)
 ReDim knots(N + T) '+ 1)
 ' tmp = UBound(knots)
 ReDim outp(sRes)
 Call SplineKnots(knots(), N - 1, T)
 Call SplineCurve(inp(), N - 1, knots(), T, outp(), sRes)
 'Calculate the length of each segment
 'and draw it
 For i = 1 To (sRes) - 1
  d = Abs(outp(i).X - outp(i - 1).X)
  h = Abs(outp(i).Y - outp(i - 1).Y)
  sLen = sLen + Sqr(d ^ 2 + h ^ 2)
  frmSpline.PicDraw.Line (outp(i - 1).X, outp(i - 1).Y)-(outp(i).X, outp(i).Y), vbBlack
 Next
 lblLen = "Spline Length:" & CInt(sLen) & " Pixels"
End Sub
Private Sub eraseSpline()
On Local Error Resume Next
'If the Outp() array isn't initialized goto error routine
 Dim i As Integer
 Dim aLen As Integer
 aLen = UBound(outp)
 If Err = 0 Then
 For i = 1 To aLen
  frmSpline.PicDraw.Line (outp(i - 1).X, outp(i - 1).Y)-(outp(i).X, outp(i).Y), PicDraw.BackColor
 Next
 End If
errErase:
 Err = 0
 On Local Error GoTo 0
End Sub
Private Sub txtRes_LostFocus()
eraseSpline
 RESOLUTION = CInt(txtRes.Text)
DrawSpline
End Sub
Private Sub delGrip()
Dim newInp() As POINTAPI
Dim i As Integer, apos As Integer
Dim idx As Integer
ReDim newInp(UBound(inp) - 1)
idx = selGrip.Index
For i = 0 To UBound(inp)
 If i <> 0 Then Unload lblGrip(i)
 If i <> idx Then
  newInp(apos) = inp(i)
  apos = apos + 1
 End If
Next
ReDim inp(UBound(newInp))
For i = 0 To UBound(newInp)
 If i <> 0 Then Load lblGrip(i)
 With lblGrip(i)
  .Move newInp(i).X - (.Width \ 2), newInp(i).Y - (.Height \ 2)
  .Visible = True
 End With
 inp(i) = newInp(i)
Next
N = UBound(inp) + 1
eraseSpline
DrawSpline
End Sub
'EOF() frmSpline Code
```

