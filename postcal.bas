Attribute VB_Name = "postcal"
Dim opr(100) As String
Dim fnc(100) As String
Dim expression(1000) As String
Const calerr = 512 + 11 'Error in calculation

Function leng(stack() As String)
i = 0
While (Not stack(i) = vbNullString)
i = i + 1
Wend
leng = i
End Function

Function reverse(stack() As String)
Dim newstack(1000) As String

length = leng(stack)

newstack(length) = vbNullString
For i = 0 To length - 1
newstack(length - i - 1) = stack(i)
Next
reverse = newstack
End Function


Function apfnc(o1 As String, char As String) As Double
Dim o11 As Double
o11 = CDbl(o1)
num = Asc(char) - Asc("A")
f = fnc(num)

Select Case f
Case "sin"
apfnc = Sin(o11)
Case "cos"
apfnc = Cos(o11)
Case "tan"
apfnc = Tan(o11)
Case "ln"
If o11 = 0 Then
apfnc = 1000
Else
If Log(Abs(o11)) > 10000 Then
apfnc = 1000
Else
apfnc = Log(Abs(o11))
End If
End If
Case "sqr"
apfnc = Sqr(o11)
Case "abs"
apfnc = Abs(o11)
Case "sgn"
apfnc = Sgn(o11)
Case "grt"
apfnc = Round(o11, 0)
Case "cosec"
apfnc = 1 / Sin(o11)
Case "sec"
apfnc = 1 / Cos(o11)
Case "cot"
apfnc = 1 / Tan(o11)

End Select


End Function

Function isfnc(ch As String) As Boolean
If Asc(ch) <= Asc("Z") And Asc(ch) >= Asc("A") Then
isfnc = True
Else
isfnc = False
End If

End Function
Function apopr(o1 As String, o2 As String, char As String)
If IsNumeric(o1) Then
o11 = CDbl(Round(o1, 10))
o22 = CDbl(o2)
Else
o11 = o1
o22 = o2
End If
Select Case char
Case "#"
apopr = o11 Or o22
apopr = Abs(apopr)
Case "&"
apopr = o11 And o22
apopr = Abs(apopr)
Case "'"
apopr = o11 <> o22
apopr = Abs(apopr)
Case "="
If o11 = o22 Then
apopr = 1
Else
apopr = 0
End If

Case ";"
apopr = o11 >= o22
apopr = Abs(apopr)
Case ">"
apopr = o11 > o22
apopr = Abs(apopr)
Case ":"
apopr = o11 <= o22
apopr = Abs(apopr)
Case "<"
apopr = o11 < o22
apopr = Abs(apopr)
Case "%"
apopr = o11 Mod o22
Case "^"
apopr = o11 ^ o22
Case "/"
apopr = o11 / o22
Case "*"
apopr = o11 * o22
Case "+"
apopr = o11 + o22
Case "-"
apopr = o11 - o22
End Select

End Function
Function replaceall(str As String, find As String, by As String)

Do While Not InStr(1, str, find) = 0
str = Replace(str, find, by, 1)
Loop
replaceall = str
End Function


Function op()
opr(0) = "("
opr(1) = "#"
opr(2) = "&"
opr(3) = "'"
opr(4) = "="
opr(5) = ";"
opr(6) = ">"
opr(7) = ":"
opr(8) = "<"
opr(9) = "-"
opr(10) = "+"
opr(11) = "%"
opr(12) = "*"
opr(13) = "/"
opr(14) = "^"
opr(15) = "!"
opr(16) = vbNullString

fnc(0) = "sin"
fnc(1) = "cos"
fnc(2) = "tan"
fnc(3) = "ln"
fnc(4) = "sqr"
fnc(5) = "abs"
fnc(6) = "sgn"
fnc(7) = "grt"
fnc(8) = "cosec"
fnc(9) = "sec"
fnc(10) = "cot"
fnc(11) = vbNullString



End Function

Function isopr(str As String)
isopr = 0
For i = 0 To 100
If opr(i) = str Then
isopr = 1
ElseIf opr(i) = vbNullString Then
Exit For
End If
Next

If str = "X" Then
Exit Function
End If

For i = 0 To 100
If Chr(i + Asc("A")) = str Then
isopr = 1
ElseIf fnc(i) = vbNullString Then
Exit For
End If
Next

End Function

Function getpr(str As String)
getpr = -1
For i = 0 To 100
If opr(i) = str Then
getpr = i
ElseIf opr(i) = vbNullString Then
Exit For
End If
Next


If isfnc(str) = True Then
getpr = 1000
End If


End Function

Function crpost(eqn As String)
eqn = "(" + LCase(eqn) + ")"


i = 0
While (Not fnc(i) = vbNullString)
eqn = replaceall(eqn, fnc(i), Chr(Asc("A") + i))
i = i + 1
Wend
eqn = replaceall(eqn, "pi", CStr(CDbl(22 / 7)))
eqn = replaceall(eqn, "e", "2.71828")


Dim ostack(1000) As String
Dim otop As Integer
otop = -1

Dim etop As Integer
etop = -1


For start = 1 To Len(eqn)
cr = Mid(eqn, start, 1)

If cr = "(" Then
ostack(otop + 1) = cr
otop = otop + 1

If isfnc(Mid(eqn, start + 1, 1)) = False Or Mid(eqn, start + 1, 1) = "X" Then
If Not Mid(eqn, start + 1, 1) = "(" Then
expression(etop + 1) = Mid(eqn, start + 1, 1)
start = start + 1
etop = etop + 1
End If
End If

ElseIf cr = ")" Then
While Not ostack(otop) = "("
expression(etop + 1) = ostack(otop)
etop = etop + 1
otop = otop - 1
Wend
otop = otop - 1

ElseIf isopr(CStr(cr)) = 1 Then

If getpr(ostack(otop)) > getpr(CStr(cr)) Then
While getpr(ostack(otop)) > getpr(CStr(cr))
expression(etop + 1) = ostack(otop)
etop = etop + 1
otop = otop - 1
Wend
End If
ostack(otop + 1) = cr
otop = otop + 1

If isfnc(Mid(eqn, start + 1, 1)) = False Or Mid(eqn, start + 1, 1) = "X" Then
If Not Mid(eqn, start + 1, 1) = "(" Then
expression(etop + 1) = Mid(eqn, start + 1, 1)
start = start + 1
etop = etop + 1
End If
End If

Else
expression(etop) = expression(etop) + cr
End If



Next
For i = 0 To etop
expr = expr + " " + expression(i)

Next
expression(etop + 1) = vbNullString
etop = etop + 1

End Function

Function calc()

Dim stack(1000) As String

For i = 0 To leng(expression) - 1
stack(i) = expression(i)

Next
stack(i) = vbNullString

Dim ele As String
Dim o1 As String
Dim o2 As String
Dim newstack(1000) As String
topn = -1

For tops = 0 To leng(stack) - 1
ele = stack(tops)
If isopr(ele) = 1 Then
If isfnc(ele) = True Then
o1 = newstack(topn)
newstack(topn) = CStr(apfnc(o1, ele))
ElseIf ele = "!" Then
o1 = newstack(topn)
If o1 = 0 Then
newstack(topn) = "1"
Else
newstack(topn) = "0"
End If

Else
o2 = newstack(topn)
topn = topn - 1
o1 = newstack(topn)
newstack(topn) = CStr(apopr(o1, o2, ele))
End If
Else
newstack(topn + 1) = ele
topn = topn + 1
End If
Next

calc = newstack(0)
Exit Function
errr:
calc = "..."
End Function
Function encode(str As String)
str = replaceall(str, "<=", ":")
str = replaceall(str, ">=", ";")
str = replaceall(str, "<>", "'")
encode = str
End Function

Function calculate(str As String)
str = encode(str)
crpost (str)
inp = calc()
If IsNumeric(inp) Then
calculate = Round(inp, 4)
Else
calculate = inp
End If

If calculate = "..." Then
err.Raise calerr, , "Calculation error found. Check expression"
End If
End Function
