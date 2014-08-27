VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form mcrorun 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Macro Run"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   9885
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Call Stand Alone SQL Query Manager"
      Height          =   495
      Left            =   6120
      TabIndex        =   10
      Top             =   1440
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      Height          =   375
      Left            =   7080
      TabIndex        =   9
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton opn 
      Caption         =   "Open"
      Height          =   375
      Left            =   8520
      TabIndex        =   8
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox file 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   6855
   End
   Begin VB.ListBox mc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      IntegralHeight  =   0   'False
      Left            =   3240
      TabIndex        =   4
      Top             =   2040
      Width           =   6495
   End
   Begin VB.CommandButton quit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   6600
      TabIndex        =   3
      Top             =   7320
      Width           =   3135
   End
   Begin VB.CommandButton rn 
      Caption         =   "Run"
      Height          =   615
      Left            =   3240
      TabIndex        =   2
      Top             =   7320
      Width           =   3255
   End
   Begin VB.ListBox ll 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      IntegralHeight  =   0   'False
      ItemData        =   "mcrorun.frx":0000
      Left            =   120
      List            =   "mcrorun.frx":0002
      TabIndex        =   1
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Macro Filename"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Macro Content"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Macro List"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   2895
   End
End
Attribute VB_Name = "mcrorun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error GoTo err
cdlg.CancelError = True
cdlg.DialogTitle = "Open file"
cdlg.Filter = "Macro Files(*.mcr)|*.mcr"
cdlg.ShowOpen
file.text = cdlg.FileName
opn_Click
err:
End Sub



Private Sub Command2_Click()
Load sql
sql.Show vbModal
End Sub

Private Sub ll_Click()
mc.Clear
arr1 = Split(arr(ll.ListIndex), "<>")

For i = 0 To UBound(arr1)
mc.AddItem (arr1(i))
Next

If mc.ListCount >= 1 Then
mc.ListIndex = 0
End If
End Sub

Private Sub opn_Click()

ll.Clear
mc.Clear
ReDim arr(0)
Dim fso As New FileSystemObject
Dim txt As TextStream
On Error Resume Next
Set txt = fso.OpenTextFile(file.text, ForReading, False)
If txt Is Nothing Then
MsgBox " An error was encountered opening file " + vbNewLine + " Error Description : File not found ", vbCritical, "Error Handler "
Exit Sub
End If

datbas = txt.ReadLine
While Not txt.AtEndOfStream
ll.AddItem (txt.ReadLine)
mem = txt.ReadLine
First = True
While (mem <> "break")
If First = True Then
arr(ll.ListCount - 1) = mem
First = False
Else
arr(ll.ListCount - 1) = arr(ll.ListCount - 1) + "<>" + mem
End If
mem = txt.ReadLine
Wend
ReDim Preserve arr(UBound(arr) + 1)
Wend

If ll.ListCount >= 1 Then
ll.ListIndex = 0
ll_Click
End If
Exit Sub
err:
ll.Clear
mc.Clear
MsgBox " An error was encountered opening file " + vbNewLine + " Error Description : " + err.Description, vbCritical, "Error Handler "

End Sub

Private Sub quit_Click()
Unload Me

End Sub

Private Sub rn_Click()
clearall
On Error GoTo lerr
 chkloop

On Error GoTo err
Set db = OpenDatabase(datbas)
clearall
While (atline < mc.ListCount)
runit (mc.List(atline))
atline = atline + 1
Wend


Exit Sub
lerr:
MsgBox "Error  : Compile Error in Loop Structure " + vbNewLine + "Error on Macro line : " + CStr(atline + 1) + vbNewLine + "Error Description : " + err.Description, vbCritical, "Error Handler"
mc.ListIndex = atline
Exit Sub

err:
MsgBox "Error  : Unable to successfully run macro " + vbNewLine + "Error on Macro line : " + CStr(atline + 1) + vbNewLine + "Error Description : " + err.Description, vbCritical, "Error Handler"
mc.ListIndex = atline
Exit Sub
End Sub

Public Function getline(num)
getline = mc.List(num)
End Function

Public Function gettot()
gettot = mc.ListCount
End Function
