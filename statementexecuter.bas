Attribute VB_Name = "statementexecuter"
Public stack(100) As Integer
Public ststr(100) As String
Public top As Integer
Const looperr = 512 + 12 'wrong loop assignment
Function push(strn)
top = top + 1
ststr(top) = strn
End Function
Function pop()
pop = ststr(top)
top = top - 1
End Function
Function geting(num)
geting = ststr(top - num)
End Function

Function gettop()
gettop = top
End Function
Function chkloop()
top = -1

For i = 0 To mcrorun.gettot() - 1
atline = i
If LCase(getit(mcrorun.getline(i), 0)) = "for" Then
push "for"
ElseIf LCase(getit(mcrorun.getline(i), 0)) = "while" Then
push "while"
ElseIf LCase(getit(mcrorun.getline(i), 0)) = "if" Then
push "if"

ElseIf LCase(mcrorun.getline(i)) = "wend" Then
If top = -1 Then
err.Raise looperr, , "Unexpected Wend"
Exit Function
End If

    If geting(0) <> "while" Then
        err.Raise looperr, , "Misplaced Wend"
        Exit Function
    Else
        pop
    End If

ElseIf LCase(mcrorun.getline(i)) = "next" Then
If top = -1 Then
err.Raise looperr, , "Unexpected Next"
Exit Function
End If
    
    If geting(0) <> "for" Then
        err.Raise looperr, , "Misplaced Next"
        Exit Function
    Else
        pop
    End If
    
ElseIf LCase(getit(mcrorun.getline(i), 0)) = "elseif" Then
    If top = -1 Then
    err.Raise looperr, , "Unexpected Elseif"
    Exit Function
    End If
    If geting(0) <> "if" Then
        err.Raise looperr, , "Misplaced Elseif"
    Exit Function
    End If


ElseIf LCase(mcrorun.getline(i)) = "else" Then
If top = -1 Then
err.Raise looperr, , "Unexpected Else"
Exit Function
End If

    If geting(0) <> "if" Then
        err.Raise looperr, , "Misplaced Else"
    Exit Function
    End If


ElseIf LCase(mcrorun.getline(i)) = "end if" Then
 If top = -1 Then
err.Raise looperr, , "Unexpected If"
Exit Function
End If
    
    If geting(0) <> "if" Then
        err.Raise looperr, , "Misplaced End if"
        Exit Function
    Else
        pop
    End If
    
End If

Next
If top <> -1 Then
err.Raise looperr, , "Ending Statement for " + pop() + " Missing"
End If
End Function
Function getmatch(start, lstart, skip)
lctr = lstart
For i = start To mcrorun.gettot() - 1
If LCase(getit(mcrorun.getline(i), 0)) = "for" Then
    lctr = lctr + 1
ElseIf LCase(getit(mcrorun.getline(i), 0)) = "while" Then
    lctr = lctr + 1
ElseIf LCase(getit(mcrorun.getline(i), 0)) = "if" Then
    lctr = lctr + 1

ElseIf LCase(mcrorun.getline(i)) = "wend" Then
    lctr = lctr - 1
    If lctr = 0 Then
    getmatch = i
    Exit Function
    End If
ElseIf LCase(getit(mcrorun.getline(i), 0)) = "elseif" Then
If skip = False Then
    If lctr = 1 Then
    getmatch = i
    Exit Function
    End If
End If
ElseIf LCase(mcrorun.getline(i)) = "else" Then
If skip = False Then
    If lctr = 1 Then
    getmatch = i
    Exit Function
    End If
End If

ElseIf LCase(mcrorun.getline(i)) = "end if" Then
    lctr = lctr - 1
    If lctr = 0 Then
    getmatch = i
    Exit Function
    End If
ElseIf LCase(mcrorun.getline(i)) = "next" Then
    lctr = lctr - 1
    If lctr = 0 Then
    getmatch = i
    Exit Function
    End If
End If

Next
End Function

Function runit(statement As String)
Dim rs As Recordset
Dim stt As String
stt = LCase(statement)
stt = replvar(stt)


Select Case gets(stt)
Case "if"
top = top + 1
ststr(top) = mcrorun.getline(atline)
stack(top) = 0
If CInt(calculate(getit(stt, 1))) = 0 Then
atline = getmatch(atline, 0, False) - 1
Else
stack(top) = 1
End If
Exit Function

Case "elseif"
If stack(top) = 1 Then
atline = getmatch(atline, 1, True) - 1
Exit Function
End If

If CInt(calculate(getit(stt, 1))) = 0 Then
atline = getmatch(atline + 1, 1, False) - 1
Else
stack(top) = 1
End If
Exit Function

Case "else"
If stack(top) = 1 Then
atline = getmatch(atline, 1, True) - 1
Exit Function
End If
stack(top) = 1

Exit Function

Case "end if"
top = top - 1
Exit Function

Case "for"
If ststr(top) <> mcrorun.getline(atline) Then
top = top + 1
ststr(top) = mcrorun.getline(atline)
stack(top) = atline
modify getit(stt, 1), calculate(getit(stt, 2))
Else
varnam = getit(stt, 1)
modify varnam, calculate(CStr(getval(varnam)) + "+" + getit(stt, 4))
End If

If CDbl(getit(stt, 4)) > 0 Then

If CDbl(getval(getit(stt, 1))) > CDbl(calculate(getit(stt, 3))) Then
top = top - 1
atline = getmatch(atline, 0, True)
End If

Else

If CDbl(getval(getit(stx, 1))) < CDbl(calculate(getit(stx, 3))) Then
top = top - 1
atline = getmatch(atline, 0, True)
End If

End If

Exit Function

Case "while"
If ststr(top) <> mcrorun.getline(atline) Then
top = top + 1
ststr(top) = mcrorun.getline(atline)
stack(top) = atline
End If

If CInt(calculate(getit(CStr(stt), 1))) = 1 Then
atline = getmatch(atline, 0, True)
top = top - 1
End If

Exit Function

Case "wend"
atline = stack(top) - 1
Exit Function

Case "next"
atline = stack(top) - 1
Exit Function

Case "replace"
vals = getval(getit(stt, 1))
modify getit(stt, 1), replaceall(getval(getit(stt, 1)), getit(stt, 2), getit(stt, 3))
Exit Function

Case "asg"
modify getit(stt, 1), getit(stt, 2)
Exit Function

Case "set"
modify getit(stt, 1), calculate(CStr(getit(stt, 2)))
Exit Function

Case "calc"
modify getit(stt, 1), calculate(CStr(getit(stt, 2)))
Exit Function

Case "sql"
db.Execute (getit(stt, 1))
Exit Function

Case "input"
inp = InputBox(getit(stt, 2), getit(stt, 3), "")
If inp = vbNullString Then
err.Raise 516, , "Null Value Encountered or Cancel was Selected"
End If
modify getit(stt, 1), inp
Exit Function

Case "init"
init getit(stt, 1), vbNullString
Exit Function

Case "crt"
Set rs = db.OpenRecordset(getit(stt, 2), dbOpenDynaset)
createtable getit(stt, 1), rs
Exit Function

Case "editit"
Set rs = db.OpenRecordset("select * from " + getit(stt, 1), dbOpenDynaset)
editit rs, getit(stt, 2), getit(stt, 3), getit(stt, 4)
Exit Function

Case "cardinality"
Set rs = db.OpenRecordset("select * from " + getit(stt, 1), dbOpenDynaset)
Load frmshow
frmshow.disrec rs
modify getit(stt, 2), rs.RecordCount


Exit Function

Case "attribute"
Set rs = db.OpenRecordset("select * from " + getit(stt, 1), dbOpenDynaset)
Load frmshow
frmshow.disrec rs
modify getit(stt, 2), rs.Fields.Count
Exit Function

Case "msg"
MsgBox getit(stt, 1), , getit(stt, 2)
Exit Function

Case "show"
Set rs = db.OpenRecordset("select * from " + getit(stt, 1), dbOpenDynaset)
Load frmshow
frmshow.disrec rs
frmshow.Show vbModal
rs.Close
Exit Function

Case "display"
Set rs = db.OpenRecordset(getit(stt, 1), dbOpenDynaset)
Load frmshow
frmshow.disrec rs
frmshow.Show vbModal
rs.Close
Exit Function

Case "edit"
Set rs = db.OpenRecordset(getit(stt, 1), dbOpenDynaset)
Load frmedit
frmedit.disrec rs
frmedit.Show vbModal
rs.Close
Exit Function

Case "newdata"
grid.openf (getit(stt, 1)), getit(stt, 2)
Exit Function

Case "showgraph"
grid.Show vbModal
Exit Function

Case "entdat"
grid.entdat CInt(getit(stt, 1)), getit(stt, 2), getit(stt, 3)
Exit Function

Case "entdata"
grid.entdat grid.match(getit(stt, 1)), getit(stt, 2), getit(stt, 3)
Exit Function

Case "plot"
grid.plotit grid.match(getit(stt, 1))
Exit Function

Case "clrg"
grid.clearit
grid.refreshit
Exit Function

Case "refresh"
grid.refreshit
Exit Function
End Select

err.Raise 513, , "Unknown Command.Check Help for List of supported Commands"
End Function
 

Function gets(str)
On Error GoTo err
arrs = Split(str, "|")
gets = Trim(arrs(0))
Exit Function
err:
err.Raise 513, , "Unknown Command"
End Function

Function getit(str, num)
str = CStr(str)
arrs = Split(str, "|")
getit = arrs(CInt(num))
End Function
