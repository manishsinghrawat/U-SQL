VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form sql 
   Caption         =   "ASC SQL- Advanced Shell Compiler for SQL - Created By Manish"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   720
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   ScaleHeight     =   9180
   ScaleWidth      =   9900
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox ssql 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   1560
      Width           =   7695
   End
   Begin VB.ComboBox ql 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   6
      Text            =   "ssql"
      Top             =   1680
      Visible         =   0   'False
      Width           =   7695
   End
   Begin VB.TextBox file 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Text            =   "< File Name Here >"
      Top             =   960
      Width           =   7695
   End
   Begin VB.CommandButton opens 
      Caption         =   "Open"
      Height          =   495
      Left            =   7920
      TabIndex        =   4
      Top             =   960
      Width           =   1935
   End
   Begin VB.ListBox l 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   2775
   End
   Begin MSComctlLib.ListView lv 
      Height          =   6975
      Left            =   3000
      TabIndex        =   2
      Top             =   3840
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   12303
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton exec 
      Caption         =   "Execute"
      Height          =   495
      Left            =   7920
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ASC SQL Created By Manish"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   5
      Top             =   120
      Width           =   8775
   End
   Begin VB.Menu pop 
      Caption         =   "Popup"
      Begin VB.Menu rename 
         Caption         =   "Rename"
      End
      Begin VB.Menu delete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "sql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Dim str As String
Dim fil As String
Function execit(text As String)
On Error GoTo err
lv.ListItems.Clear
lv.ColumnHeaders.Clear

If InStr(1, text, "select", vbTextCompare) = 1 Then
Set rs = db.OpenRecordset(text, dbOpenDynaset)
showset
MsgBox " Query Executed ... " + vbNewLine + " Assuming Data Request Query ", vbInformation, "Information"
Else
db.Execute (text)
db.Close
otable
MsgBox " Request Completed ... " + vbNewLine + " Assuming " + vbNewLine + " 1) Data Defination Language (DDL)" + vbNewLine + " 2) Data Manipulation Language (DML) ", vbInformation
End If


Exit Function
err:
MsgBox " Error occured while executing SQL statement " + vbNewLine + " Error Description : " + err.Description, vbCritical, "Error Handler"
End Function

Private Sub exec_Click()

mem = ssql.text
While InStr(1, mem, vbNewLine) > 0
mem = Replace(mem, vbNewLine, " ", 1)
Wend


execit (mem)


End Sub
'Function addit(mem As String)
'ssql.AddItem ("")

'If ssql.ListCount <> 1 Then
'For i = ssql.ListCount - 1 To 1 Step -1
'ssql.List(i) = ssql.List(i - 1)
'Next
'End If
'ssql.List(0) = mem


'End Function

Function showset()
On Error Resume Next
For i = 0 To rs.Fields.Count - 1
lv.ColumnHeaders.Add , , rs.Fields(i).name
Next

rs.MoveFirst
While Not rs.EOF
j = j + 1
lv.ListItems.Add , , rs.Fields(0)
For i = 1 To rs.Fields.Count - 1
If rs.Fields(i) <> mems Then
If Not rs.Fields(i) Is Nothing Then
lv.ListItems(lv.ListItems.Count).SubItems(i) = rs.Fields(i)
End If
End If
Next
If j = 1000 Then

j = 0
End If
rs.MoveNext
Wend
rs.Close
End Function


Private Sub file_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
opens_click
End If
End Sub

Private Sub Form_Load()
top = GetSetting(App.Title, "settings", "top", top)
Left = GetSetting(App.Title, "settings", "left", Left)
Height = GetSetting(App.Title, "settings", "height", Height)
Width = GetSetting(App.Title, "settings", "width", Width)
WindowState = GetSetting(App.Title, "settings", "windowstate", WindowState)
l_Click
End Sub

Private Sub Form_Resize()
On Error Resume Next
lbl.Left = 0
lbl.Width = Width
lbl.top = 200
lbl.Height = 700

l.Left = 100
file.top = 1000
file.Left = 100

ssql.top = file.top + file.Height + 100
ssql.Left = 100

opens.Height = file.Height
exec.Height = ssql.Height

file.Width = Width - opens.Width - 400
ssql.Width = file.Width

opens.Left = file.Left + file.Width + 100
opens.top = file.top

exec.top = opens.top + opens.Height + 100
exec.Left = opens.Left

l.top = ssql.top + ssql.Height + 100
l.Width = 2 / 8 * Width

lv.top = exec.top + exec.Height + 100
lv.Left = l.Left + l.Width + 100
lv.Height = Height - lv.top - 1000
l.Height = lv.Height
lv.Width = Width - l.Width - 400


End Sub
Function otable()

On Error GoTo err
Dim fso As New FileSystemObject

If fso.FileExists(fil) = False Then
err.Raise 1
End If

Set db = OpenDatabase(fil)

l.Clear

For i = 0 To db.TableDefs.Count - 1
l.AddItem (db.TableDefs(i).name)
Next

Exit Function
err:
MsgBox " Error occured while Opening MDB Database " + vbNewLine + " Error Description : " + err.Description, vbCritical, "Error Handler"
End Function

Private Sub Form_Unload(Cancel As Integer)
If Not (WindowState = vbMinimized Or WindowState = vbMaximized) Then
SaveSetting App.Title, "settings", "top", top
SaveSetting App.Title, "settings", "left", Left
SaveSetting App.Title, "settings", "height", Height
SaveSetting App.Title, "settings", "width", Width
End If
If Not WindowState = vbMinimized Then
SaveSetting App.Title, "settings", "windowstate", WindowState
End If
End Sub

Private Sub l_Click()
If l.ListIndex = -1 Then
pop.Enabled = False
Else
pop.Enabled = True
End If

End Sub

Private Sub l_dblClick()
On Error GoTo err
lv.ListItems.Clear
lv.ColumnHeaders.Clear

If Not l.ListIndex = -1 Then
Set rs = db.OpenRecordset(l.List(l.ListIndex))
showset
End If
Exit Sub
err:
MsgBox " Unable to complete Data Request." + vbNewLine + " Error Description : " + err.Description, vbCritical, "Error Handler"

End Sub

Private Sub l_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If l.ListIndex <> -1 Then
If Button = vbRightButton Then
PopupMenu pop, , l.Left + X, l.top + Y
End If
End If
End Sub

Private Sub Rename_click()
On Error GoTo err:

mem = l.List(l.ListIndex)
nname = InputBox("Enter new name for table", "Rename", "")

For i = 0 To db.TableDefs.Count - 1
If db.TableDefs(i).name = mem Then
db.TableDefs(i).name = nname
End If
Next

db.Close
otable
Exit Sub
err:
MsgBox " Unable to Rename file " + vbNewLine + " Error Description : " + err.Description, vbCritical, "Error Handler"

End Sub
Private Sub delete_click()
On Error GoTo err
mem = l.List(l.ListIndex)
msg = MsgBox("Are you sure you want to delete table " + mem + " ?", vbQuestion + vbYesNo, "Confirm Delete")

If msg = vbYes Then
db.TableDefs.delete (mem)
db.Close
otable
End If

Exit Sub
err:
MsgBox " Unable to Delete file " + vbNewLine + " Error Description : " + err.Description, vbCritical, "Error Handler"

End Sub

Private Sub opens_click()
fil = file.text
otable
l_Click
End Sub

