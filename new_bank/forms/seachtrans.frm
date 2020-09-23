VERSION 5.00
Begin VB.Form seachtrans 
   Caption         =   "Form1"
   ClientHeight    =   9750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12720
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9750
   ScaleWidth      =   12720
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3000
      TabIndex        =   10
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   9
      Top             =   840
      Width           =   855
   End
   Begin VB.ListBox lstGuestName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3900
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Frame fraDetails 
      BackColor       =   &H80000016&
      Caption         =   "Account Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   5295
      Left            =   2520
      TabIndex        =   1
      Top             =   2160
      Width           =   6495
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
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   3600
         Width           =   2175
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
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   3000
         Width           =   2175
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
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   2400
         Width           =   2175
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
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox Text2 
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
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1200
         Width           =   2175
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
         Height          =   375
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label lblGuestID 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         Caption         =   "Slip No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   360
         TabIndex        =   7
         Top             =   1200
         Width           =   780
      End
      Begin VB.Label lblState 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         Caption         =   "Trans Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   360
         TabIndex        =   6
         Top             =   3000
         Width           =   1440
      End
      Begin VB.Label lbladdress1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         Caption         =   "Trans Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   360
         TabIndex        =   5
         Top             =   1800
         Width           =   1860
      End
      Begin VB.Label lblCity 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         Caption         =   "Previous Balance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   360
         TabIndex        =   4
         Top             =   2400
         Width           =   1845
      End
      Begin VB.Label lbldate 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         Caption         =   "Account Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   1710
      End
      Begin VB.Label lbltime 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         Caption         =   "Final Balance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   360
         TabIndex        =   2
         Top             =   3600
         Width           =   1440
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Transaction Log"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   3240
      TabIndex        =   13
      Top             =   0
      Width           =   2115
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   11880
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter Account Number"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   360
      TabIndex        =   12
      Top             =   960
      Width           =   2205
   End
   Begin VB.Label lblcount 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   9120
      TabIndex        =   11
      Top             =   787
      Width           =   75
   End
End
Attribute VB_Name = "seachtrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrSql As String


Public Sub Findname()

  
    
     ' record variables
    Dim mark As Variant
    Dim count As Integer
    
   Call connect
              
    
    count = 0
    RS_Name.Find "acc_no LIKE '" & txtName.Text & "%'"
    Do While Not RS_transaction.EOF
        'continue if last find succeeded
       lstGuestName.AddItem RS_transaction!acc_no
        'count the last title found
       'count = count + 1
        ' note current position
       mark = RS_transaction.Bookmark
       RS_Name.Find "acc_name LIKE '" & txtName.Text & "%'", 1, adSearchForward, mark
        ' above code skips current record to avoid finding the same row repeatedly;
        ' last arg (bookmark) is redundant because Find searches from current position
    Loop
    If count = 0 Then
     MsgBox "No Match Found", vbOKOnly + vbInformation, "Information"
     txtName.SetFocus
    Else
     lblcount.Caption = "Total Matches found " & count
    End If
     ' clean up
    RS_transaction.Close
'    cnn.Close
    Set RS_transaction = Nothing
    Set cnn = Nothing

End Sub
Public Sub Findnumber()

  
    
     ' record variables
    Dim mark As Variant
    Dim count As Integer
    
   Call connect
              
    
    count = 0
    With RS_transaction
   
    .Find "acc_no LIKE '" & txtName.Text & "%'"
    Do While Not .EOF
        'continue if last find succeeded
       lstGuestName.AddItem RS_transaction!date1
        'count the last title found
       count = count + 1
        ' note current position
       mark = .Bookmark
       .Find "acc_no LIKE '" & txtName.Text & "%'", 1, adSearchForward, mark
        ' above code skips current record to avoid finding the same row repeatedly;
        ' last arg (bookmark) is redundant because Find searches from current position
      
    Loop
    If count = 0 Then
     MsgBox "No Match Found", vbOKOnly + vbInformation, "Information"
     txtName.SetFocus
    Else
     lblcount.Caption = "Total Number of Transaction made on your Account is " & count
    End If
     ' clean up
    RS_transaction.Close
    End With
'    cnn.Close
    Set RS_transaction = Nothing
    Set cnn = Nothing

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdGo_Click()

lstGuestName.clear
fraDetails.Visible = False

If txtName.Text = "" Then
 MsgBox "Enter the name", vbOKOnly + vbCritical, "Error"
 txtName.SetFocus
 Exit Sub
End If

Findnumber


End Sub

Private Sub Form_Load()
fraDetails.Visible = False

Call connect


lstGuestName.clear
End Sub

Private Sub lstGuestName_Click()
Call connect
With RS_transaction
 .MoveFirst
 While Not .EOF
  If lstGuestName.List(lstGuestName.ListIndex) = .Fields(5) Then
   Text1.Text = .Fields(1)
   Text2.Text = .Fields(0)
   Text3.Text = .Fields(2)
   Text4.Text = .Fields(3)
   Text5.Text = .Fields(4)
   Text6.Text = .Fields(6)
  ' Image1.Picture = LoadPicture(.Fields(13))

  End If
  .MoveNext
 Wend
   
  fraDetails.Visible = True
 
End With
End Sub

