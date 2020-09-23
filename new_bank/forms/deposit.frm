VERSION 5.00
Begin VB.Form deposit 
   BackColor       =   &H80000009&
   Caption         =   "Deposit"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10110
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "deposit.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9015
   ScaleWidth      =   10110
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Deposit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   7335
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   3120
         TabIndex        =   15
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   14
         Top             =   3120
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   3120
         TabIndex        =   12
         Top             =   2640
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         Height          =   330
         Left            =   3120
         TabIndex        =   10
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   3120
         TabIndex        =   9
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   345
         Left            =   3120
         TabIndex        =   0
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   1
         Top             =   4920
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   2
         Top             =   4920
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   360
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1080
         TabIndex        =   13
         Top             =   3360
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Final Balance"
         Height          =   210
         Left            =   1080
         TabIndex        =   11
         Top             =   2760
         Width           =   1320
      End
      Begin VB.Label Label4 
         Caption         =   "Actual Balance"
         Height          =   375
         Left            =   1080
         TabIndex        =   8
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Amount"
         Height          =   210
         Left            =   1080
         TabIndex        =   7
         Top             =   2280
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Account Number"
         Height          =   210
         Left            =   1080
         TabIndex        =   6
         Top             =   1200
         Width           =   1620
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Slip Number"
         Height          =   210
         Left            =   1080
         TabIndex        =   5
         Top             =   840
         Width           =   1215
      End
   End
End
Attribute VB_Name = "deposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Combo1_Click()
With RS_account
    .MoveFirst
  While Not .EOF
   If Combo1.List(Combo1.ListIndex) = .Fields(0) Then
    Text6.Text = .Fields(12)
   End If
    .MoveNext
  Wend
 End With
End Sub

Private Sub Command4_Click()
If Text3.Text = "" Then
MsgBox "Add Amount for Transaction", vbInformation
Exit Sub
Text3.SetFocus
End If
Dim actual As String

Min = 2000
actual = Val(Text2.Text)
dep = Val(Text3.Text)
final = actual + dep
Text4.Text = final


With RS_transaction

.AddNew
.Fields(0) = Text1.Text
.Fields(1) = Combo1.Text
.Fields(2) = "Deposit"
.Fields(3) = Text6.Text
.Fields(4) = Text3.Text
.Fields(5) = Text5.Text
.Fields(6) = Text4.Text
.Update
.MoveNext
With RS_account
.MoveFirst
.Update
.Fields(12) = Text4.Text
.MoveNext
End With


MsgBox "Your Account has being Credited by" & Text3.Text & ", Your final balance is " & Text4.Text, vbInformation, "Account Deposit"
 
End With

clearall
End Sub

Private Sub Command5_Click()
Me.Hide
clearall

End Sub

Private Sub Form_Load()
Call connect
With RS_account
    While Not .EOF
   Combo1.AddItem .Fields(0)
   .MoveNext
  Wend
 End With
 Text5.Text = Format(Date & Time)

 
End Sub
Public Sub clearall()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text6.Text = ""
Combo1.ListIndex = -1

End Sub


Private Sub Form_Unload(Cancel As Integer)
close_connect
End Sub
