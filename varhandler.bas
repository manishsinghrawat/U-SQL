Attribute VB_Name = "variablehandler"

Dim avar() As String

Const aexist = 512 + 10 'already exist
Const nexist = 512 + 11 'not exist
Function clearall()
ReDim avar(0)
top = 0
atline = 0
End Function


 Function replvar(str)
On Error GoTo err
 mem = str
 For i = 0 To UBound(avar) - 1
 arr1 = Split(avar(i), "<>")

 While (InStr(1, mem, "#" + CStr(arr1(0)) + "#") > 0)
 If arr1(1) = "" Then
 arr1(1) = " "
 End If
 mem = Replace(mem, "#" + CStr(arr1(0)) + "#", CStr(arr1(1)))
  Wend
  
  Next
  replvar = mem
Exit Function
err:
replvar = str
 End Function
 
Function init(nam, val)
For i = 0 To UBound(avar) - 1
arr1 = Split(avar(i), "<>")
If arr1(0) = CStr(nam) Then
err.Raise aexist, , "Variable Already Declared"
Exit Function
End If
Next

avar(UBound(avar)) = CStr(nam) + "<>" + CStr(val)
ReDim Preserve avar(UBound(avar) + 1)
End Function

Function modify(nam, newval)
For i = 0 To UBound(avar) - 1
arr1 = Split(avar(i), "<>")
If arr1(0) = CStr(nam) Then
avar(i) = CStr(nam) + "<>" + CStr(newval)
Exit Function
End If
Next
err.Raise nexist, , "Variable Not initialized"
End Function

Function getval(nam)
For i = 0 To UBound(avar) - 1
arr1 = Split(avar(i), "<>")
If arr1(0) = CStr(nam) Then
getval = arr1(1)
Exit Function
End If
Next
err.Raise nexist, , "Variable not initialized"
End Function
