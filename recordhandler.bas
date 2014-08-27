Attribute VB_Name = "recordhandler"
Function createtable(nam, record As Recordset)
Dim str As String
Dim rs As Recordset

str = "CREATE TABLE " + nam + " ("
For i = 0 To record.Fields.Count - 1
str = str + record.Fields(i).name
str = str + " " + getstr(record.Fields(i).Type, record.Fields(i).size) + " ,"
Next
str = Left(str, Len(str) - 1)
str = str + ");"
db.Execute (str)
db.TableDefs.Refresh
For i = 0 To db.TableDefs.Count - 1
If db.TableDefs(i).name = nam Then
Set rs = db.TableDefs(i).OpenRecordset(dbOpenDynaset)
Exit For
End If
Next

record.MoveFirst

While (Not record.EOF)
rs.AddNew

For i = 0 To record.Fields.Count - 1
rs(rs.Fields(i).name) = record(record.Fields(i).name)
Next
rs.Update
record.MoveNext
Wend

End Function

Function editit(record As Recordset, X, Y, plc)
record.MoveFirst
For i = 1 To X - 1
If record.EOF = True Then
err.Raise 512, , "Record too short"
Exit Function
End If

record.MoveNext
Next

record.Edit
record.Fields(Y - 1).Value = CStr(plc)
record.Update
End Function

Function getstr(num, size)
Select Case num
Case bigint
getstr = "bigint"
Exit Function
Case dbBinary
getstr = "binary"
Exit Function
Case dbBoolean
getstr = "boolean"
Exit Function
Case dbByte
getstr = "byte"
Exit Function
Case dbChar
getstr = "char"
Exit Function
Case dbCurrency
getstr = "currency"
Exit Function
Case dbDate
getstr = "date"
Exit Function
Case dbDecimal
getstr = "decimal (6,3)"
Exit Function
Case dbDouble
getstr = "double"
Exit Function
Case dbFloat
getstr = "float"
Exit Function
Case dbGUID
getstr = "guid"
Exit Function
Case dbInteger
getstr = "int"
Exit Function
Case dbLong
getstr = "long"
Exit Function
Case dbLongBinary
getstr = "longbinary"
Exit Function
Case dbNumeric
getstr = "numeric"
Exit Function
Case dbSingle
getstr = "single"
Exit Function
Case dbText
getstr = "char (" + CStr(size) + ")"
Exit Function
Case dbTime
getstr = "time"
Exit Function
Case dbVarBinary
getstr = "varbinary"
Exit Function
End Select

End Function
