VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmedit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recordset Editor"
   ClientHeight    =   11025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11025
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Reset Position"
      Height          =   495
      Left            =   8400
      TabIndex        =   10
      Top             =   9480
      Width           =   2175
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   12720
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   495
      Left            =   10680
      TabIndex        =   6
      Top             =   9480
      Width           =   2175
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   12960
      TabIndex        =   5
      Top             =   9480
      Width           =   2175
   End
   Begin MSComctlLib.ListView lv 
      Height          =   9255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   16325
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
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   6240
      TabIndex        =   3
      Top             =   9480
      Width           =   2055
   End
   Begin VB.CommandButton cmddel 
      Caption         =   "Delete"
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   9480
      Width           =   1935
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "Add"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   9480
      Width           =   1935
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "Edit"
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   9480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Editor for Recorsets generated using queries the changes made here will be reflected back in parent tables"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   10200
      Width           =   15015
   End
   Begin VB.Label lbl 
      Height          =   255
      Index           =   0
      Left            =   10680
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "frmedit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim record As Recordset
Function dis()
cmdedit.Enabled = False
cmdadd.Enabled = False
cmddel.Enabled = False
cmdquit.Enabled = False
cmdsave.Enabled = True
cmdcancel.Enabled = True
lv.Width = 10455
sizeh
End Function

Function ena()
cmdedit.Enabled = True
cmdadd.Enabled = True
cmddel.Enabled = True
cmdquit.Enabled = True
cmdsave.Enabled = False
cmdcancel.Enabled = False
lv.Width = 15075
sizeh
For i = txt.Count - 1 To 1 Step -1
Unload txt(i)
Unload lbl(i)
Next
txt(0).Visible = False
lbl(0).Visible = False

End Function

Function sizeh()
If Not lv.ColumnHeaders.Count = 0 Then
wid = (lv.Width - 300) / lv.ColumnHeaders.Count
For i = 1 To lv.ColumnHeaders.Count
lv.ColumnHeaders(i).Width = wid
Next
End If
End Function

Private Sub cmdadd_Click()
On Error GoTo err

record.MoveFirst
For i = 1 To lv.SelectedItem.Index - 1
record.MoveNext
Next
record.AddNew
dis

txt(0).Text = ""
lbl(0).Caption = record.Fields(0).Name
txt(0).Visible = True
lbl(0).Visible = True

For i = 1 To record.Fields.Count - 1
Load txt(i)
 Load lbl(i)
 txt(i).Left = txt(i - 1).Left
 txt(i).Width = txt(i - 1).Width
 txt(i).Height = txt(i - 1).Height
 txt(i).Top = txt(i - 1).Top + txt(i - 1).Height
 txt(i).Visible = True
 lbl(i).Left = lbl(i - 1).Left
 lbl(i).Width = lbl(i - 1).Width
 lbl(i).Height = lbl(i - 1).Height
 lbl(i).Top = txt(i).Top
 lbl(i).Visible = True
 
txt(i).Text = ""

lbl(i).Caption = record.Fields(i).Name
Next
err:
MsgBox "Error : Unable to Add new record " + vbNewLine + "Possible Reason(s) : " + err.Description, vbCritical, "Error Handler"
End Sub

Private Sub cmdcancel_Click()
ena
End Sub

Private Sub cmddel_Click()
On Error GoTo err
record.MoveFirst
For i = 1 To lv.SelectedItem.Index - 1
record.MoveNext
Next
msg = MsgBox("Are you sure you want to delete current record ?", vbQuestion + vbYesNo, "Confirm deletion")
If msg = vbYes Then
record.Delete
disrec record
End If

Exit Sub
err:
MsgBox "Error : Unable to delete the File " + vbNewLine + "Possible Reason(s) : " + err.Description, vbCritical, "Error Handler"
End Sub

Private Sub cmdedit_Click()
On Error GoTo err

record.MoveFirst
For i = 1 To lv.SelectedItem.Index - 1
record.MoveNext
Next
record.edit
dis

txt(0).Text = record(record.Fields(0).Name)
lbl(0).Caption = record.Fields(0).Name
txt(0).Visible = True
lbl(0).Visible = True

For i = 1 To record.Fields.Count - 1
Load txt(i)
 Load lbl(i)
 txt(i).Left = txt(i - 1).Left
 txt(i).Width = txt(i - 1).Width
 txt(i).Height = txt(i - 1).Height
 txt(i).Top = txt(i - 1).Top + txt(i - 1).Height
 txt(i).Visible = True
 lbl(i).Left = lbl(i - 1).Left
 lbl(i).Width = lbl(i - 1).Width
 lbl(i).Height = lbl(i - 1).Height
 lbl(i).Top = txt(i).Top
 lbl(i).Visible = True
 
txt(i).Text = ""
 If record.Fields(i) <> mems Then
If Not record.Fields(i) Is Nothing Then
 txt(i).Text = record(record.Fields(i).Name)
End If
End If

lbl(i).Caption = record.Fields(i).Name
Next
Exit Sub
err:
MsgBox "Error : Unable to Edit the File " + vbNewLine + "Possible Reason(s) : " + err.Description, vbCritical, "Error Handler"

End Sub

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
Set record = rs
sizeh
End Function

Private Sub cmdquit_Click()
Unload Me
End Sub

Private Sub Command2_Click()

End Sub

Private Sub cmdsave_Click()
On Error Resume Next


For i = 0 To record.Fields.Count - 1
record(record.Fields(i).Name) = txt(i).Text
Next

record.Update
disrec record
ena

Exit Sub
err:
MsgBox "Error : Unable to Complete Data Storage Request " + vbNewLine + "Error Description : " + err.Description, vbCritical, "error Handler"
End Sub

Private Sub Command1_Click()
On Error Resume Next
Me.Left = 0
Me.Top = 0
End Sub

Private Sub Form_Load()
ena
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = Not cmdquit.Enabled
End Sub

