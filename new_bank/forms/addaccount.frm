VERSION 5.00
Begin VB.Form atm 
   BackColor       =   &H80000009&
   Caption         =   "ATM Application Form"
   ClientHeight    =   9075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9645
   Icon            =   "addaccount.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9075
   ScaleWidth      =   9645
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text14 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   11
      Top             =   3960
      Width           =   2655
   End
   Begin VB.TextBox Text13 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   10
      Top             =   3480
      Width           =   2655
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   9
      Top             =   3000
      Width           =   2655
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   8
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   7
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   6
      Top             =   1080
      Width           =   2655
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   5
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      TabIndex        =   3
      Top             =   7560
      Width           =   5295
      Begin VB.CommandButton cmdexit 
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
         Left            =   3720
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "Add"
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
         Left            =   360
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdsave 
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
         Left            =   2160
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label Label3 
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
      Left            =   6240
      TabIndex        =   19
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Card Code"
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
      Left            =   960
      TabIndex        =   18
      Top             =   3120
      Width           =   1020
   End
   Begin VB.Label lblphone 
      AutoSize        =   -1  'True
      Caption         =   "Account Name"
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
      Left            =   960
      TabIndex        =   17
      Top             =   2640
      Width           =   1410
   End
   Begin VB.Label lblmobile 
      AutoSize        =   -1  'True
      Caption         =   "Mobile number"
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
      Left            =   960
      TabIndex        =   16
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label lblcity 
      AutoSize        =   -1  'True
      Caption         =   "Account Number"
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
      Left            =   960
      TabIndex        =   15
      Top             =   2040
      Width           =   1620
   End
   Begin VB.Label lbladd 
      AutoSize        =   -1  'True
      Caption         =   "Customer Address"
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
      Left            =   960
      TabIndex        =   14
      Top             =   1560
      Width           =   1800
   End
   Begin VB.Label lblname 
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
      Left            =   960
      TabIndex        =   13
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblmail 
      AutoSize        =   -1  'True
      Caption         =   "Email"
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
      Left            =   960
      TabIndex        =   12
      Top             =   4080
      Width           =   540
   End
   Begin VB.Label Label1 
      Caption         =   "Customer Information"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   480
      TabIndex        =   4
      Top             =   600
      Width           =   8415
   End
End
Attribute VB_Name = "atm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command3_Click()

End Sub



Private Sub cmdexit_Click()

Me.Hide
End Sub



Private Sub Command1_Click()
On Error Resume Next
CommonDialog1.DialogTitle = "Backup"
CommonDialog1.Filter = "*.mdb"
CommonDialog1.FileName = "*.mdb"
CommonDialog1.ShowOpen
If Not Len(CommonDialog1.FileName) = 5 Then
    Text1.Text = CommonDialog1.FileName
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
CommonDialog1.DialogTitle = "Backup"
CommonDialog1.Filter = "*.mdb"
CommonDialog1.FileName = "*.mdb"
CommonDialog1.ShowOpen
If Not Len(CommonDialog1.FileName) = 5 Then
    Text2.Text = CommonDialog1.FileName
End If
End Sub

Private Sub cmdsave_Click()
If MsgBox("Are you sure?", vbYesNo) = vbYes Then

Generate
With RS_atm
'.MoveFirst
.AddNew
.Fields(0) = Text8.Text
.Fields(1) = Text9.Text
.Fields(2) = Text10.Text
.Fields(3) = Text11.Text
.Fields(4) = Text12.Text
.Fields(5) = Text13.Text
.Fields(6) = Text14.Text
.Fields(7) = Label3.Caption
.Update
MsgBox "Your ATM Card has being loaded, your ID is " & Text11.Text & " and your Access Code is " & Text12.Text, vbOKOnly
'.MoveNext
End With
End If
End Sub

Private Sub Form_Load()
Call connect

  
Label3.Caption = Date

End Sub


Private Sub Text10_lostfocus()
With RS_account
.MoveFirst
While Not .EOF
If Text10.Text = .Fields(0) Then
Text11.Text = .Fields(9)
'Generate

End If
.MoveNext
Wend
End With
End Sub
Private Sub Generate()
Dim id As String
Dim code As String
id = 10001
 Tmp = 3300
 With RS_atm
 Tmp = Tmp + RS_atm.RecordCount + 2
code = CStr(Tmp)
Text12.Text = code
End With
End Sub
