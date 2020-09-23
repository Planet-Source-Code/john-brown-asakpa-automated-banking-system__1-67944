VERSION 5.00
Begin VB.Form Addloan 
   BackColor       =   &H80000009&
   ClientHeight    =   11640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14235
   Icon            =   "Addloans.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11640
   ScaleWidth      =   14235
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
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
      Left            =   4080
      TabIndex        =   32
      Top             =   9840
      Width           =   5415
      Begin VB.CommandButton Command2 
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
      Begin VB.CommandButton Command1 
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
      Begin VB.CommandButton Command3 
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
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Loan Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8775
      Left            =   2280
      TabIndex        =   23
      Top             =   960
      Width           =   9735
      Begin VB.TextBox Text9 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00FF8080&
         Caption         =   "Time  Period"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2055
         Left            =   120
         TabIndex        =   43
         Top             =   6360
         Width           =   2895
         Begin VB.OptionButton Option19 
            BackColor       =   &H00FF8080&
            Caption         =   "2 Years"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   1440
            Width           =   1455
         End
         Begin VB.OptionButton Option18 
            BackColor       =   &H00FF8080&
            Caption         =   "1 Year"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   1080
            Width           =   1575
         End
         Begin VB.OptionButton Option17 
            BackColor       =   &H00FF8080&
            Caption         =   "6 Months"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   720
            Width           =   1815
         End
         Begin VB.OptionButton Option16 
            BackColor       =   &H00FF8080&
            Caption         =   "90 Days"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.TextBox Text8 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox Text7 
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
         Left            =   2640
         TabIndex        =   40
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox Text3 
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
         Left            =   2640
         TabIndex        =   38
         Top             =   1440
         Width           =   3855
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
         Left            =   2640
         TabIndex        =   36
         Top             =   960
         Width           =   3015
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FF8080&
         Caption         =   "Comercial Segment"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2175
         Left            =   6480
         TabIndex        =   34
         Top             =   3120
         Width           =   2895
         Begin VB.Frame Frame7 
            BackColor       =   &H00FF8080&
            Height          =   1335
            Left            =   120
            TabIndex        =   35
            Top             =   720
            Width           =   2655
            Begin VB.OptionButton Option15 
               BackColor       =   &H00FF8080&
               Caption         =   "Above 1,000,000"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   960
               Width           =   2400
            End
            Begin VB.OptionButton Option14 
               BackColor       =   &H00FF8080&
               Caption         =   "Upto 1,000,000"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   120
               TabIndex        =   17
               Top             =   720
               Width           =   2400
            End
            Begin VB.OptionButton Option13 
               BackColor       =   &H00FF8080&
               Caption         =   "Upto 500,000"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   120
               TabIndex        =   16
               Top             =   480
               Width           =   2400
            End
            Begin VB.OptionButton Option12 
               BackColor       =   &H00FF8080&
               Caption         =   "Upto 250,000"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   120
               TabIndex        =   15
               Top             =   240
               Width           =   2400
            End
         End
         Begin VB.OptionButton Option11 
            BackColor       =   &H00FF8080&
            Caption         =   "Cash Credit"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Width           =   2535
         End
         Begin VB.OptionButton Option10 
            BackColor       =   &H00FF8080&
            Caption         =   "Term Loans"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FF8080&
         Caption         =   "Loan Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   120
         TabIndex        =   33
         Top             =   3480
         Width           =   2775
         Begin VB.OptionButton Option9 
            BackColor       =   &H00FF8080&
            Caption         =   "Comercial Loans"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   840
            Width           =   2295
         End
         Begin VB.OptionButton Option8 
            BackColor       =   &H00FF8080&
            Caption         =   "Personal Loans"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   480
            Width           =   2295
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FF8080&
         Caption         =   "Account Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2055
         Left            =   3240
         TabIndex        =   25
         Top             =   6360
         Width           =   3495
         Begin VB.TextBox Text6 
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
            Left            =   1080
            TabIndex        =   22
            Top             =   1560
            Width           =   2295
         End
         Begin VB.TextBox Text5 
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
            Left            =   1080
            TabIndex        =   21
            Top             =   1080
            Width           =   2295
         End
         Begin VB.TextBox Text4 
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
            Left            =   1080
            TabIndex        =   20
            Top             =   600
            Width           =   2295
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FF8080&
            Caption         =   "Joint"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Name 3"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   120
            TabIndex        =   28
            Top             =   1560
            Width           =   765
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Name 2"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   120
            TabIndex        =   27
            Top             =   1080
            Width           =   765
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Name 1"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   120
            TabIndex        =   26
            Top             =   600
            Width           =   765
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FF8080&
         Caption         =   "Personal Segment"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2775
         Left            =   3120
         TabIndex        =   24
         Top             =   3480
         Width           =   3135
         Begin VB.OptionButton Option7 
            BackColor       =   &H00FF8080&
            Caption         =   "Vahical Above 100000"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   2400
            Width           =   2775
         End
         Begin VB.OptionButton Option6 
            BackColor       =   &H00FF8080&
            Caption         =   "Vahical Upto 100000"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   2040
            Width           =   2775
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H00FF8080&
            Caption         =   "Vahical Upto 25000"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   1680
            Width           =   2775
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FF8080&
            Caption         =   "Personal Upto 100000"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   2655
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FF8080&
            Caption         =   "Personal Upto 50000"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   2535
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FF8080&
            Caption         =   "Housing Upto 200000"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   960
            Width           =   2655
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00FF8080&
            Caption         =   "Housing Above 200000"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   1320
            Width           =   2775
         End
      End
      Begin VB.TextBox Text1 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Date of Return"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   360
         TabIndex        =   49
         Top             =   3000
         Width           =   1440
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   360
         TabIndex        =   41
         Top             =   2520
         Width           =   1425
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Label8"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   5160
         TabIndex        =   39
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Phone Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   360
         TabIndex        =   37
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   360
         TabIndex        =   31
         Top             =   480
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   360
         TabIndex        =   30
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   360
         TabIndex        =   29
         Top             =   1440
         Width           =   795
      End
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "LOAN APPLICATION FORM"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   345
      Left            =   4440
      TabIndex        =   50
      Top             =   240
      Width           =   3765
   End
End
Attribute VB_Name = "Addloan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim LoanNoGen As Integer
Dim LoanSegment As Object
Dim LoanType As Object
Dim term As Integer

Private Sub Check1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If Check1.Value = 1 Then
  Text4.Enabled = True
  Text5.Enabled = True
  Text6.Enabled = True
  Label4.Enabled = True
  Label5.Enabled = True
  Label6.Enabled = True
  Text4.SetFocus
 ElseIf Check1.Value = 0 Then
  Text4.Enabled = False
  Text5.Enabled = False
  Text6.Enabled = False
  Label4.Enabled = False
  Label5.Enabled = False
  Label6.Enabled = False
  Command2.SetFocus
 End If
End If
End Sub

Private Sub Command1_Click()
With RS_loan
Dim id As String
Dim code As String
id = "LN"
temp = 1000

temp = temp + RS_loan.RecordCount + 1
code = id + "-" + CStr(temp)
Text1.Text = code
Frame1.Enabled = True
End With
Text2.SetFocus
End Sub

Private Sub Command2_Click()
Frame1.Enabled = False
If MsgBox("Are You Sure?", vbQuestion + vbYesNo, "AuotBank") = vbYes Then
    If check <> vbOK Then
     With RS_loan
     .AddNew
     
     .Fields(0) = Text1.Text
     .Fields(1) = Text2.Text
    .Fields(2) = Text3.Text
     .Fields(3) = Text7.Text
     .Fields(4) = Label8.Caption
     .Fields(5) = LoanSegment.Caption
     .Fields(6) = LoanType.Caption
     .Fields(7) = Text8.Text
      Select Case term
      Case 90
        .Fields(8) = 90
        maturitydate = Date + 90
        .Fields(9) = maturitydate
        
      Case 6
        .Fields(8) = 6
        maturitydate = Date + 180
        .Fields(9) = maturitydate
        
      Case 1
        .Fields(8) = 1
        maturitydate = Date + 365
        .Fields(9) = maturitydate
        
      Case 2
        .Fields(8) = 2
        maturitydate = Date + 365 + 365
        .Fields(9) = maturitydate
        
      End Select
     .Fields(10) = Text4.Text
     .Fields(11) = Text5.Text
     .Fields(12) = Text6.Text
     .Update
     MsgBox "Your Request has being sent for verification", vbInformation, "Loan Application Form"
          MsgBox "You shall be contacted as soon as the verification is complete", vbInformation, "Loan Application Form"
     cleaall
     Exit Sub
     End With
    End If
End If
Frame1.Enabled = False

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call Connect
Frame1.Enabled = False
Label8.Caption = Format(Date) & " " & (Time)
End Sub

Public Sub cleaall()
Text1.Text = ""
Text3.Text = ""
Text2.Text = ""
Text7.Text = ""
Frame2.Enabled = False
Frame6.Enabled = False
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text8.Text = ""
Text9.Text = ""
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Label4.Enabled = False
Label5.Enabled = False
Label6.Enabled = False
Frame1.Enabled = False
Command1.SetFocus
End Sub

Private Sub DataCombo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If DataCombo1.Text = "" Then
  DataCombo1.SetFocus
  Exit Sub
  End If
Text3.SetFocus
End If
End Sub


Private Sub option16_Click()
term = 90
maturitydate = Date + 90
Text9.Text = maturitydate
End Sub

Private Sub option17_Click()
term = 6
maturitydate = Date + 180
Text9.Text = maturitydate
End Sub

Private Sub option18_Click()
maturitydate = Date + 365
Text9.Text = maturitydate
term = 1
End Sub

Private Sub option19_Click()
maturitydate = Date + 365 + 365
Text9.Text = maturitydate
term = 2
End Sub

Private Sub Option1_Click()
Set LoanType = Option1
Check1.SetFocus
Text8.Text = 50000
End Sub

Private Sub Option10_Click()
If KeyAscii = 13 Then
Option12.SetFocus
End If
End Sub
Private Sub Option11_click()
If KeyAscii = 13 Then
Option12.SetFocus
End If
End Sub

Private Sub Option12_Click()
Set LoanType = Option12
Check1.SetFocus
Text8.Text = 250000
End Sub


Private Sub Option13_Click()
Set LoanType = Option12
Check1.SetFocus
Text8.Text = 500000
End Sub


Private Sub Option14_Click()
Set LoanType = Option12
Check1.SetFocus
Text8.Text = 1000000
End Sub


Private Sub Option15_Click()
Set LoanType = Option12
Check1.SetFocus
Text8.Text = 1500000
End Sub




Private Sub Option2_Click()
Check1.SetFocus
Set LoanType = Option2
Text8.Text = 100000
End Sub

Private Sub Option3_Click()
Check1.SetFocus
Set LoanType = Option3
Text8.Text = 200000
End Sub

Private Sub Option4_Click()
Check1.SetFocus
Set LoanType = Option4
Text8.Text = 3000000
End Sub

Private Sub Option5_Click()
Check1.SetFocus
Set LoanType = Option5
Text8.Text = 25000
End Sub

Private Sub Option6_Click()
Check1.SetFocus
Set LoanType = Option6
Text8.Text = 100000
End Sub

Private Sub Option7_Click()
Check1.SetFocus
Set LoanType = Option7
Text8.Text = 200000
End Sub

Private Sub Option8_Click()
Frame6.Enabled = False
Frame2.Enabled = True
Option1.SetFocus
Set LoanSegment = Option8
End Sub
Private Sub Option9_Click()
Frame2.Enabled = False
Frame6.Enabled = True
Option10.SetFocus
Set LoanSegment = Option9
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Val(Text3.Text) < 2000 Then
  MsgBox "!!! Amount Can't Be Less Then 2000 !!!", vbCritical + vbOKOnly, "AutoBank"
  Text3.SetFocus
  SendKeys "{Home}+{End}"
  Exit Sub
  End If
Text3.Text = Val(Text3.Text)
Option8.SetFocus
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text4.Text = UCase(Text4.Text)
Text5.SetFocus
End If
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text5.Text = UCase(Text5.Text)
Text6.SetFocus
End If
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text6.Text = UCase(Text6.Text)
Command2.SetFocus
End If
End Sub


Private Function check()
Dim temp As Integer
If Check1.Value = 1 _
And Text4.Text = "" _
And Text5.Text = "" _
And Text6.Text = "" Then
temp = MsgBox("!!!No Additional Name Found!!!", vbOKOnly + vbCritical, "AutoBank")
End If

check = temp
End Function
