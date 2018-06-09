VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form5 
   Caption         =   "Customer Management"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form5"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   975
      Left            =   7320
      Top             =   840
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1720
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
      Connect         =   $"Customer management.frx":0000
      OLEDBString     =   $"Customer management.frx":00A6
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
      Left            =   6240
      TabIndex        =   20
      Top             =   7920
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete Customer Details"
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
      Left            =   3120
      TabIndex        =   16
      Top             =   7920
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
      Left            =   720
      TabIndex        =   15
      Top             =   7920
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Add/Update Customer Details"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8295
      Left            =   10200
      TabIndex        =   4
      Top             =   600
      Width           =   7335
      Begin VB.TextBox Text6 
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
         Left            =   3360
         TabIndex        =   19
         Top             =   4800
         Width           =   2655
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
         Height          =   615
         Left            =   3120
         TabIndex        =   11
         Top             =   7080
         Width           =   1455
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
         Height          =   615
         Left            =   4200
         TabIndex        =   10
         Top             =   6120
         Width           =   2295
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
         Left            =   3360
         TabIndex        =   9
         Top             =   3840
         Width           =   2775
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
         Left            =   3360
         TabIndex        =   8
         Top             =   2880
         Width           =   2775
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
         Left            =   3360
         TabIndex        =   7
         Top             =   1920
         Width           =   2775
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
         Left            =   3360
         TabIndex        =   6
         Top             =   840
         Width           =   2775
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Add New Customer"
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
         Left            =   960
         TabIndex        =   5
         Top             =   6120
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "Last Bought on :"
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
         Left            =   720
         TabIndex        =   18
         Top             =   4800
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Reference ID"
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
         Left            =   840
         TabIndex        =   17
         Top             =   1920
         Width           =   1575
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
         Height          =   615
         Left            =   840
         TabIndex        =   14
         Top             =   3960
         Width           =   2055
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
         Left            =   960
         TabIndex        =   13
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Customer Name"
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
         Left            =   840
         TabIndex        =   12
         Top             =   840
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search by Customer Name/Customer ID"
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
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   5655
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
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
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
         TabIndex        =   1
         Top             =   480
         Width           =   3615
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Customer management.frx":014C
      Height          =   4815
      Left            =   480
      TabIndex        =   3
      Top             =   2520
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   8493
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
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.RecordSource = "Select * from Table1 where Customer_Name = '" + Text1.Text + "' or Customer_ID='" + Text1.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox "No Such Customer Found!", vbCritical + vbOKCancel, "Error!"
End If
End Sub

Private Sub Command2_Click()
Adodc1.RecordSource = "Select * from Table1"
Adodc1.Refresh
End Sub

Private Sub Command3_Click()
answer = MsgBox("Are you sure you want to delete this customer from the records?", vbQuestion + vbYesNo, "Warning!")
If answer = vbYes Then
Adodc1.Recordset.Delete
MsgBox "Record deleted successfully.", vbInformation + vbOKOnly, "Confirmation"
Adodc1.Recordset.Update
End If
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
End Sub

Private Sub Command4_Click()
Form2.Show
Unload Me
End Sub

Private Sub Command5_Click()
If Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Then
MsgBox "Field shouldn't be empty!", vbOKOnly + vbCritical, "Warning!"
Else
Adodc1.Recordset.AddNew
With Adodc1.Recordset
.Fields(0) = Text2.Text
.Fields(1) = Text3.Text
.Fields(2) = Text4.Text
.Fields(3) = Text5.Text
.Fields(4) = Text6.Text
End With
MsgBox "New customer added successfully!", vbOKOnly + vbInformation, "Add successful!"
Adodc1.Recordset.Update
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
End If
End Sub

Private Sub Command6_Click()
If Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Then
MsgBox "Field shouldn't be empty!", vbOKOnly + vbCritical, "Warning!"
Else
With Adodc1.Recordset
.Fields(0) = Text2.Text
.Fields(1) = Text3.Text
.Fields(2) = Text4.Text
.Fields(3) = Text5.Text
.Fields(4) = Text6.Text
End With
MsgBox "Details updated successfully!", vbOKOnly + vbInformation, "Update successful!"
Adodc1.Recordset.Update
Text2.Text = Adodc1.Recordset.Fields(0)
Text3.Text = Adodc1.Recordset.Fields(1)
Text4.Text = Adodc1.Recordset.Fields(2)
Text5.Text = Adodc1.Recordset.Fields(3)
Text6.Text = Adodc1.Recordset.Fields(4)
End If
End Sub

Private Sub Command7_Click()
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
End Sub

Private Sub DataGrid1_Click()
Text2.Text = Adodc1.Recordset.Fields(0)
Text3.Text = Adodc1.Recordset.Fields(1)
Text4.Text = Adodc1.Recordset.Fields(2)
Text5.Text = Adodc1.Recordset.Fields(3)
Text6.Text = Adodc1.Recordset.Fields(4)
End Sub

