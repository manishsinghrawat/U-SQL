VERSION 5.00
Begin VB.Form Grid 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Graph"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7590
   FillStyle       =   0  'Solid
   LinkTopic       =   "Grid"
   ScaleHeight     =   6750
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton opens 
      Caption         =   "<<"
      Height          =   375
      Left            =   7080
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   720
      Top             =   1800
   End
End
Attribute VB_Name = "grid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim moving As Boolean
Dim movewhat As Integer
Dim lastx As Long
Dim lasty As Long
Dim funcs(100, 30000, 1) As Double
Dim info(100, 2) As Long
Dim infos(100) As String

Public originx As Double
Public originy As Double
Public unitxf As Double
Public unityf As Double
Public unitx As Double
Public unity As Double

Function st()
originx = 5000
originy = 5000
unitxf = 500
unityf = 500
unitx = 1
unity = 1
num = 0
End Function

Public Function openf(nam As String, col As String)
For i = 0 To 100
If info(i, 0) = 1 Then
If infos(i) = nam Then
err.Raise 512, , "Graph with this name already exists"
End If
End If
Next
openf = newf
infos(newf) = nam
info(newf, 2) = col
info(newf, 1) = 0
info(newf, 0) = 1

End Function

Public Function entdat(fnum As Integer, xxx As Double, yyy As Double)
funcs(fnum, info(fnum, 1), 0) = xxx
funcs(fnum, info(fnum, 1), 1) = yyy
info(fnum, 1) = info(fnum, 1) + 1
End Function
Public Function clearit()
For i = 0 To 100
info(i, 0) = 0
info(i, 1) = 0
Next
End Function
Function delete(fnam As String)
For i = 0 To 100
If info(i, 0) = 1 Then
If infos(i) = nam Then
GoTo repl
End If
End If
Next
err.Raise 512, , "No graph with this name exists"
Exit Function
repl:
info(i, 0) = 0
info(i, 1) = 0

End Function
Public Function match(nam As String)
For i = 0 To 100
If info(i, 0) = 1 Then
If infos(i) = nam Then
match = i
Exit Function
End If
End If
Next
err.Raise 512, , "Undetemined Name variable. Check the graph name entered"
End Function

Public Function entdata(name As Integer, xxx As Double, yyy As Double)

For i = 0 To 100
If info(i, 0) = 1 Then
If infos(i) = name Then
fnum = i
GoTo nexts
End If
End If
Next
Exit Function
nexts:
entdat fnum, xxx, yyy
End Function

Private Sub Form_Load()
st
Form_Resize



End Sub


Function drwbase()
Me.Cls

'set temporary variables
unix = unitxf * marks(unitxf)
uniy = unityf * marks(unityf)
'main process initiated
For temp = 1 To (originx - Me.ScaleLeft) / unix
Line (originx - (unix * temp), Me.ScaleTop)-(originx - (unix * temp), Me.ScaleTop + Height), gridcol
Next
For temp = 1 To ((Width - originx + Me.ScaleLeft) / unix)
Line (originx + (unix * temp), Me.ScaleTop)-(originx + (unix * temp), Me.ScaleTop + Height), gridcol
Next
For temp = 1 To (originy - Me.ScaleTop) / uniy
Line (Me.ScaleLeft, originy - (uniy * temp))-(Me.ScaleLeft + Width, originy - (uniy * temp)), gridcol
Next
For temp = 1 To (Height - originy + Me.ScaleTop) / uniy
Line (Me.ScaleLeft, originy + (uniy * temp))-(Me.ScaleLeft + Width, originy + (uniy * temp)), gridcol
Next

'main process initiated

For temp = 1 To (Width - originx + Me.ScaleLeft) / (unitxf * marks(unitxf))
tempx = originx + (temp * unitxf * marks(unitxf))
Line (tempx, originy - mklen)-(tempx, originy + mklen), markscol
Next
For temp = 1 To (originx - Me.ScaleLeft) / (unitxf * marks(unitxf))
tempx = originx - (temp * unitxf * marks(unitxf))
Line (tempx, originy - mklen)-(tempx, originy + mklen), markscol
Next
For temp = 1 To (originy - Me.ScaleTop) / (unityf * marks(unityf))
tempy = originy - (temp * unityf * marks(unityf))
Line (originx - mklen, tempy)-(originx + mklen, tempy), markscol
Next
For temp = 1 To (Height + Me.ScaleTop - originy) / (unityf * marks(unityf))
tempy = originy + (temp * unityf * marks(unityf))
Line (originx - mklen, tempy)-(originx + mklen, tempy), markscol
Next


FontSize = 6.3
ForeColor = textcol
For temp = 1 To (Width + Me.ScaleLeft - originx) / (marks(unitxf) * unitxf)
CurrentX = originx + temp * unitxf * marks(unitxf) - 0.5 * lengthx
CurrentY = originy + distx

Print (CStr(Round((temp * unitx) * marks(unitxf), 2)) + suffix)
Next

For temp = -1 To -(originx - Me.ScaleLeft) / (marks(unitxf) * unitxf) Step -1
CurrentX = originx + temp * unitxf * marks(unitxf) - 0.5 * lengthx
CurrentY = originy + distx
Print (CStr(Round((temp * unitx) * marks(unitxf), 2)) + suffix)
Next

For temp = 1 To (Height + Me.ScaleTop - originy) / (marks(unityf) * unityf)
CurrentX = originx - disty
CurrentY = originy + temp * unityf * marks(unityf) - lenghty
Print (CStr(-1 * Round((temp * unity) * marks(unityf), 2)) + suffix)
Next

For temp = -1 To -(originy - Me.ScaleTop) / (marks(unityf) * unityf) Step -1
CurrentX = originx - disty
CurrentY = originy + temp * unityf * marks(unityf) - lenghty
Print (CStr(-1 * Round((temp * unity) * marks(unityf), 2)) + suffix)
Next



FillStyle = 0
FillColor = ptcol
ForeColor = vbBlack
Line (originx, Me.ScaleTop)-(originx, Me.ScaleTop + Me.Height), axiscol
Line (Me.ScaleLeft, originy)-(Me.ScaleLeft + Me.Width, originy), axiscol



Me.Circle (originx, originy), ptrad
Me.Circle (originx + unitxf, originy), ptrad
Me.Circle (originx, originy - unityf), ptrad

End Function

Function drwchr(fstyle, fcolor, forecol)
FillStyle = fstyle
FillColor = fcolor
End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If hits(x, Y, originx, originy, hitrad) = True Then
MousePointer = 5
moving = True
movewhat = origin
ElseIf hits(x, Y, originx + unitxf, originy, hitrad) = True Then
MousePointer = 9
moving = True
movewhat = xunit
ElseIf hits(x, Y, originx, originy - unityf, hitrad) = True Then
MousePointer = 7
moving = True
movewhat = yunit
Else
MousePointer = 5
lastx = x
lasty = Y
moving = True
movewhat = scroll

End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If moving = True Then
If movewhat = origin Then
originx = x
originy = Y
ElseIf movewhat = xunit Then
If x - originx >= minval Then
unitxf = x - originx
End If
ElseIf movewhat = yunit Then
If originy - Y >= minval Then
unityf = originy - Y
End If
Else
originx = originx + x - lastx
originy = originy + Y - lasty
lastx = x
lasty = Y
End If
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
MousePointer = 0
moving = False
End Sub

Private Sub Form_Resize()
opens.Left = Width - opens.Width - 200
opens.top = 0
drwbase
redraw
End Sub


Function redraw()
For i = 0 To 100
If info(i, 0) = 1 Then
plotit (i)
End If
Next
End Function

Function plotit(n)

Dim x As Long
Dim lstx As Double
Dim lsty As Double
lstx = funcs(n, 0, 0)
lsty = funcs(n, 0, 1)
res = 0

FillStyle = 0
FillColor = ptcol
ForeColor = vbBlack

For x = 1 To info(n, 1) - 1

Line (originx + conform(lstx, unitxf), originy - conform(lsty, unityf))-(originx + conform(funcs(n, x, 0), unitxf), originy - conform(funcs(n, x, 1), unityf)), info(n, 2)

lstx = funcs(n, x, 0)
lsty = funcs(n, x, 1)
Next

End Function

Function dist(x1, y1, x2, y2) As Double
dist = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
End Function


Function congraph(num, var)
congraph = num / var
End Function

Function conform(num As Double, var As Double)
conform = (num * (var))
End Function



Private Sub opens_Click()
Load manager
Dim i As Integer
For i = 0 To 100
If info(i, 0) = 1 Then
manager.addit infos(i), i
End If
Next
manager.Show vbModal
End Sub

Function newf()
For i = 0 To 100
If info(i, 0) = 0 Then
newf = i
Exit Function
End If
Next
End Function

Function del(funnum As Integer)
info(funnum, 0) = 0
info(funnum, 1) = 0
Form_Resize
End Function

Function getcol(funnum As Integer)
getcol = info(funnum, 2)
End Function

Function setcol(funnum As Integer, color As Long)
info(funnum, 2) = color
Form_Resize
End Function
Function refreshit()
Form_Resize
End Function

Private Sub Timer1_Timer()
Form_Resize
End Sub
