VERSION 5.00
Begin VB.Form frm_result 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Individual Account Report"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   3135
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   3135
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   885
      Width           =   735
   End
   Begin VB.TextBox txt_id 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1380
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Account No"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   960
   End
End
Attribute VB_Name = "frm_result"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command4_Click()
If Not txt_id.Text = "" Then
data.Command1 txt_id.Text
DataReport3.Show
txt_id.Text = ""
Me.Hide
Else
    txt_id.SetFocus
End If
End Sub

