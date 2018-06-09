VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form7 
   Caption         =   "Selling medicines"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form7"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Back to Home"
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
      Left            =   11520
      TabIndex        =   10
      Top             =   5280
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show Total Price"
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
      Left            =   11280
      TabIndex        =   9
      Top             =   1800
      Width           =   2895
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
      Height          =   735
      Left            =   6120
      TabIndex        =   8
      Top             =   6000
      Width           =   2895
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   855
      Left            =   6960
      Top             =   7800
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      Connect         =   $"Selling Medicines.frx":0000
      OLEDBString     =   $"Selling Medicines.frx":00A6
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from table1"
      Caption         =   "Adodc2"
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
   Begin VB.ComboBox Combo2 
      DataSource      =   "Adodc2"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5880
      TabIndex        =   6
      Top             =   4560
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sell "
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
      Left            =   11400
      TabIndex        =   5
      Top             =   3240
      Width           =   2775
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
      Left            =   5880
      TabIndex        =   4
      Top             =   3000
      Width           =   3255
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   4200
      Top             =   7800
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      Connect         =   $"Selling Medicines.frx":014C
      OLEDBString     =   $"Selling Medicines.frx":01EE
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *  from table1"
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
   Begin VB.ComboBox Combo1 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      ItemData        =   "Selling Medicines.frx":0290
      Left            =   5880
      List            =   "Selling Medicines.frx":0292
      TabIndex        =   3
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label Label4 
      Caption         =   "Total price"
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
      Left            =   2280
      TabIndex        =   7
      Top             =   6000
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Customer name"
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
      Left            =   2280
      TabIndex        =   2
      Top             =   4320
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Quantity"
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
      Left            =   2280
      TabIndex        =   1
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Medicine name"
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
      Left            =   2280
      TabIndex        =   0
      Top             =   1680
      Width           =   2415
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Combo1.Text = "" Or Combo2.Text = "" Or Text1.Text = "" Then
MsgBox "Blank field detected!", vbOKOnly + vbCritical, "Warning!"
Else
Adodc1.RecordSource = "select * from table1 where Medicine_Name='" + Combo1.Text + "'"
Adodc1.Refresh
If Val(Text1.Text) > Val(Adodc1.Recordset.Fields("Quantity")) Then
MsgBox "Order quantity has exceeded stock quantity!", vbOKOnly + vbCritical, "Warning!"
Else
MsgBox "Order has been successfully placed.", vbOKOnly + vbInformation, "Order Confirmation"
Adodc1.Recordset.Fields("Quantity") = Val(Adodc1.Recordset.Fields("Quantity")) - Val(Text1.Text)
Adodc1.Recordset.Update
End If
End If
End Sub

Private Sub Command2_Click()
Adodc1.RecordSource = "select * from table1 where Medicine_Name='" + Combo1 + "'"
Adodc1.Refresh
Text2.Text = Val(Text1.Text) * Val(Adodc1.Recordset.Fields("Price"))
Label4.Visible = True
Text2.Visible = True
End Sub

Private Sub Command3_Click()
Form2.Show
Unload Me
End Sub

Private Sub Form_Load()
Label4.Visible = False
Text2.Visible = False
With Adodc1.Recordset
.MoveFirst
Do Until .EOF
Combo1.AddItem .Fields(0)
Adodc1.Recordset.MoveNext
Loop
End With
With Adodc2.Recordset
.MoveFirst
Do Until .EOF
Combo2.AddItem .Fields(0)
Adodc2.Recordset.MoveNext
Loop
End With

End Sub

Private Sub Text1_Click()
If Combo1.Text = "" Then
MsgBox "Enter the name of medicine first!", vbOKOnly + vbCritical, "Warning!"
End If
End Sub
