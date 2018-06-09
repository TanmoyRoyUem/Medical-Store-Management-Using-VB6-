VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form4 
   Caption         =   "Dealer Management"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form4"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Dealer Management.frx":0000
      Height          =   4695
      Left            =   840
      TabIndex        =   18
      Top             =   2520
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   8281
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   22
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   7440
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1508
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Dealer Management.frx":0015
      OLEDBString     =   $"Dealer Management.frx":00B9
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from table1"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Add/Update Company Details"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8055
      Left            =   10200
      TabIndex        =   6
      Top             =   600
      Width           =   6735
      Begin VB.CommandButton Command5 
         Caption         =   "Add New Company"
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
         Left            =   840
         TabIndex        =   13
         Top             =   5280
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   12
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2880
         TabIndex        =   11
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2880
         TabIndex        =   10
         Top             =   2520
         Width           =   2895
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   9
         Top             =   3840
         Width           =   3135
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Update Details"
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
         Left            =   4080
         TabIndex        =   8
         Top             =   5280
         Width           =   1815
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Clear"
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
         Left            =   3000
         TabIndex        =   7
         Top             =   6720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Company Name"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   17
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Address "
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   16
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Contact person"
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
         Left            =   480
         TabIndex        =   15
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Phone Number"
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
         Left            =   480
         TabIndex        =   14
         Top             =   3840
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Back"
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
      Left            =   7080
      TabIndex        =   5
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete Company"
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
      Left            =   3960
      TabIndex        =   4
      Top             =   7800
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "View All Records"
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
      Left            =   960
      TabIndex        =   3
      Top             =   7800
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search by Company Name"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   5655
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4200
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Adodc1.RecordSource = "Select * from Table1 where Company_Name = '" + Text1.Text + "' "
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox "No Such Company Found!", vbCritical + vbOKCancel, "Error!"
End If
End Sub

Private Sub Command2_Click()
Adodc1.RecordSource = "Select * from Table1"
Adodc1.Refresh
End Sub

Private Sub Command3_Click()
answer = MsgBox("Are you sure you want to delete this company from the records?", vbQuestion + vbYesNo, "Warning!")
If answer = vbYes Then
Adodc1.Recordset.Delete
MsgBox "Record deleted successfully.", vbInformation + vbOKOnly, "Confirmation"
Adodc1.Recordset.Update
End If
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
End Sub

Private Sub Command4_Click()
Form2.Show
Unload Me
End Sub

Private Sub Command5_Click()
If Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Then
MsgBox "Field shouldn't be empty!", vbOKOnly + vbCritical, "Warning!"
Else
Adodc1.Recordset.AddNew
With Adodc1.Recordset
.Fields(0) = Text2.Text
.Fields(1) = Text3.Text
.Fields(2) = Text4.Text
.Fields(3) = Text5.Text
End With
MsgBox "New company added successfully!", vbOKOnly + vbInformation, "Add successful!"
Adodc1.Recordset.Update
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
End If
End Sub

Private Sub Command6_Click()
If Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Then
MsgBox "Field shouldn't be empty!", vbOKOnly + vbCritical, "Warning!"
Else
With Adodc1.Recordset
.Fields(0) = Text2.Text
.Fields(1) = Text3.Text
.Fields(2) = Text4.Text
.Fields(3) = Text5.Text
End With
MsgBox "Details updated successfully!", vbOKOnly + vbInformation, "Update successful!"
Adodc1.Recordset.Update
Text2.Text = Adodc1.Recordset.Fields(0)
Text3.Text = Adodc1.Recordset.Fields(1)
Text4.Text = Adodc1.Recordset.Fields(2)
Text5.Text = Adodc1.Recordset.Fields(3)
End If
End Sub

Private Sub Command7_Click()
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
End Sub

Private Sub DataGrid1_Click()
Text2.Text = Adodc1.Recordset.Fields(0)
Text3.Text = Adodc1.Recordset.Fields(1)
Text4.Text = Adodc1.Recordset.Fields(2)
Text5.Text = Adodc1.Recordset.Fields(3)
End Sub

