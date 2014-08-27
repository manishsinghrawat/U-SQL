Attribute VB_Name = "calculator"
Dim breaked As String
Dim arrpost() As String
Dim func(26) As String
Dim fcnt As String

Function setfunc()
func(0) = "sin"
func(1) = "cos"
func(2) = "tan"
func(3) = "cosec"
func(4) = "sec"
func(5) = "cot"
func(6) = "abs"
func(7) = "grt"
func(8) = "ln"
func(9) = "sgn"
fcnt = 9
End Function
Function getno(str As String)

End Function
Function isopr(str As String)

End Function
Function encode(str As String)
str = Replaceall(str, "<=", ":")
str = Replaceall(str, ">=", ";")
str = Replaceall(str, "!=", "'")

End Function
