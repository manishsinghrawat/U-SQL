VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmshow 
   Caption         =   "Table Display"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   9810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      Height          =   495
      Left            =   7320
      TabIndex        =   1
      Top             =   7080
      Width           =   2415
   End
   Begin MSComctlLib.ListView lv 
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   12091
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label lbl 
      Caption         =   "Table Editor : Can be used for viewing the tables or the recordsets based on queries"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   7200
      Width           =   7095
   End
End
Attribute VB_Name = "frmshow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function disrec(rs As Recordset)
lv.ListItems.Clear
lv.ColumnHeaders.Clear
For i = 0 To rs.Fields.Count - 1
lv.ColumnHeaders.Add , , rs.Fields(i).Name
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
sizeh
End Function


Private Sub Command1_Click()
Unload Me
End Sub

Function sizeh()
If Not lv.ColumnHeaders.Count = 0 Then
wid = (lv.Width - 300) / lv.ColumnHeaders.Count
For i = 1 To lv.ColumnHeaders.Count
lv.ColumnHeaders(i).Width = wid
Next
End If
End Function

Private Sub Form_Resize()
On Error Resume Next
lv.Width = Width - 500
lv.Height = Height - 1500
lbl.Top = lv.Top + lv.Height + 200
Command1.Top = lbl.Top
Command1.Left = lv.Left + lv.Width - Command1.Width
sizeh
End Sub
