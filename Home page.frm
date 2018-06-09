VERSION 5.00
Begin VB.Form Form2 
   Caption         =   " "
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "WELCOME!"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7935
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   14535
      Begin VB.CommandButton Command7 
         Caption         =   "Exit Application"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   10320
         TabIndex        =   7
         Top             =   3840
         Width           =   2295
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Log Out"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   10320
         TabIndex        =   6
         Top             =   2040
         Width           =   2295
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Purchase Medicines"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5640
         TabIndex        =   5
         Top             =   6120
         Width           =   2655
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Stock Management"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5640
         TabIndex        =   4
         Top             =   4800
         Width           =   2655
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Customer Management"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5640
         TabIndex        =   3
         Top             =   3480
         Width           =   2655
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Dealer Management"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5640
         TabIndex        =   2
         Top             =   2160
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Employee Management"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5640
         TabIndex        =   1
         Top             =   720
         Width           =   2535
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form3.Show
Unload Me
End Sub

Private Sub Command2_Click()
Form4.Show
Unload Me
End Sub

Private Sub Command3_Click()
Form5.Show
Unload Me
End Sub

Private Sub Command4_Click()
Form6.Show
Unload Me
End Sub

Private Sub Command5_Click()
Form7.Show
Unload Me
End Sub

Private Sub Command6_Click()
Form1.Show
Unload Me
MsgBox "Logged out!", vbInformation + vbOKOnly
End Sub

Private Sub Command7_Click()
Unload Me
End Sub
