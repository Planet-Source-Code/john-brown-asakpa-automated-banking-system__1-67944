VERSION 5.00
Begin VB.Form payment 
   BackColor       =   &H80000009&
   Caption         =   "Payment"
   ClientHeight    =   9900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11700
   Icon            =   "payment.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9900
   ScaleWidth      =   11700
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Payment"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   18
         Top             =   3600
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   3120
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3120
         TabIndex        =   14
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   4200
         Width           =   2055
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3120
         TabIndex        =   3
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
         Left            =   2040
         TabIndex        =   2
         Top             =   5280
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
         Left            =   4320
         TabIndex        =   1
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Rate in Percentage"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1080
         TabIndex        =   17
         Top             =   3600
         Width           =   1875
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Period of Loan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1080
         TabIndex        =   15
         Top             =   3240
         Width           =   1230
      End
      Begin VB.Label Label6 
         Caption         =   "Interest"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   12
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Actual Amount"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   10
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Date of Application"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1080
         TabIndex        =   8
         Top             =   2280
         Width           =   1875
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Customer Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1080
         TabIndex        =   6
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Loan Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1080
         TabIndex        =   5
         Top             =   1320
         Width           =   1320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Slip Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1080
         TabIndex        =   4
         Top             =   840
         Width           =   1215
      End
   End
End
Attribute VB_Name = "payment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SlipNoGen As Integer

Dim intr As Integer
Dim pay As String
Private Sub Combo1_Click()
With RS_loan
.MoveFirst
While Not .EOF
If Combo1.List(Combo1.ListIndex) = .Fields(0) Then
Text3.Text = .Fields(1)
Text4.Text = .Fields(4)
Text5.Text = .Fields(7)
Text2.Text = .Fields(9)
End If
.MoveNext
Wend

End With
End Sub

Private Sub Command4_Click()
Dim per As String
With RS_loan
.MoveFirst
'While Not .EOF
per = .Fields(8)
pay = Val(Text5.Text)
intr = pay * Val(Text7.Text) * per / 100
Text6.Text = intr
.MoveNext
'Wend
End With
If MsgBox("Are You Sure?", vbQuestion + vbYesNo, "AutoBank") = vbYes Then

  If check <> vbOK Then
  
   With RS_Payment
    .AddNew
    .Fields(0) = Text1.Text
    .Fields(1) = Combo1.Text
    .Fields(2) = Text3.Text
    .Fields(3) = Text4.Text
    .Fields(4) = Text5.Text
    .Fields(5) = Text7.Text
    .Fields(6) = Text6.Text
    .Fields(7) = Text2.Text
   .Update
   MsgBox "Your loan has been paid, you are to return payment on this date " & Text2.Text & "and your interest will be " & Text6.Text, vbOKOnly
   End With
    cleaall
  End If
End If

End Sub

Private Sub Command5_Click()
Unload Me
End Sub
Private Function check()
Dim temp As Integer
If Val(Text5.Text) < 10000 Then
temp = MsgBox("Amount Can't Be Less Than 100", vbCritical + vbOKOnly, "AutoBank")
End If

If Combo1.Text = "" Then
temp = MsgBox("Account Number Can't Be Empty", vbCritical + vbOKOnly, "AutoBank")
End If

check = temp
End Function

Private Sub DataCombo1_KeyPress(KeyAscii As Integer)
    If DataCombo1.Text = "" Then
    DataCombo1.SetFocus
    Exit Sub
    End If
Text1.SetFocus
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If Val(Text1.Text) < 100 Then
 MsgBox "!!! Amount Can't Be Less Than 100 !!!", vbCritical + vbOKOnly, "AutoBank"
 Text1.SetFocus
 SendKeys "{Home}+{End}"
 Exit Sub
End If
Text1.Text = Val(Text1.Text)
Command4.SetFocus
End If
End Sub


Private Sub Form_Load()
Call connect
Frame1.Enabled = True
With RS_loan
    While Not .EOF
    
    Combo1.AddItem .Fields(0)
    
   .MoveNext
  Wend
 End With
End Sub
Private Sub cleaall()
Text1.Text = ""
Combo1.ListIndex = -1
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
End Sub


