VERSION 5.00
Begin VB.Form editdat 
   Caption         =   "Data Editor"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5520
      TabIndex        =   3
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "Save"
      Height          =   495
      Left            =   5520
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label lbl 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "editdat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function edit(record As Recordset, num As Integer)
record.MoveFirst
For i = 1 To num - 1
record.MoveNext
Next

txt(0).Text = record(record.Fields(0).Name)
lbl(0).Caption = record.Fields(0).Name
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
 
 
 If record.Fields(i) <> mems Then
If Not record.Fields(i) Is Nothing Then
 txt(i).Text = record(record.Fields(i).Name)
End If
End If



lbl(i).Caption = record.Fields(i).Name
Next

End Function

Private Sub tx1_Change(Index As Integer)

End Sub

