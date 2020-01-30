VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00800080&
   Caption         =   "Form2"
   ClientHeight    =   10335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22800
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   10335
   ScaleMode       =   0  'User
   ScaleWidth      =   44186.05
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Logindb 
      Height          =   330
      Left            =   480
      Top             =   6600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\bus project\Databases folder\Login.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\bus project\Databases folder\Login.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "logindb"
      Caption         =   "Logindb"
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
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   13320
      TabIndex        =   13
      Top             =   8520
      Width           =   1815
   End
   Begin VB.CommandButton lastbtn 
      Caption         =   "Last"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   16200
      TabIndex        =   12
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton nextbtn 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   16200
      TabIndex        =   11
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton prevbtn 
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   16200
      TabIndex        =   10
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton firstbtn 
      Caption         =   "First"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   16200
      TabIndex        =   9
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton addtn 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5880
      TabIndex        =   8
      Top             =   8520
      Width           =   1815
   End
   Begin VB.CommandButton updatetn 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   9240
      TabIndex        =   7
      Top             =   8520
      Width           =   2415
   End
   Begin VB.TextBox txtpword 
      DataField       =   "Password"
      DataSource      =   "Logindb"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      IMEMode         =   3  'DISABLE
      Left            =   10200
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   4440
      Width           =   4335
   End
   Begin VB.TextBox txtcpword 
      DataField       =   "Confirm Password"
      DataSource      =   "Logindb"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      IMEMode         =   3  'DISABLE
      Left            =   10200
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   6105
      Width           =   4455
   End
   Begin VB.TextBox txtuname 
      DataField       =   "Username"
      DataSource      =   "Logindb"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      IMEMode         =   3  'DISABLE
      Left            =   10200
      TabIndex        =   4
      Top             =   2640
      Width           =   4335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3480
      TabIndex        =   3
      Top             =   6120
      Width           =   5535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5760
      TabIndex        =   2
      Top             =   4440
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   5640
      TabIndex        =   1
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   10800
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Log in/Sign up"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8280
      TabIndex        =   0
      Top             =   600
      Width           =   4335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub addtn_Click()
Logindb.Recordset.AddNew
End Sub

Private Sub Command1_Click()
Me.Hide
End Sub

Private Sub deletercd_Click()
confirmation = MsgBox("Do you want to delete this record", vbYesNo + vbCritical, "Delete Record Confirmation")
If confirmation = vbYes Then
Logindb.Recordset.Delete
MsgBox "Record has been deleted successfully", vbInformation, "Message"
Else
MsgBox "Record Not Deleted!!!", vbInformation, "Message"
End If
End Sub

Private Sub firstbtn_Click()
Logindb.Recordset.MoveFirst
End Sub

Private Sub lastbtn_Click()
Logindb.Recordset.MoveLast
End Sub

Private Sub nextbtn_Click()
Logindb.Recordset.MoveNext
End Sub

Private Sub prevbtn_Click()
Logindb.Recordset.MovePrevious
End Sub

Private Sub updatetn_Click()
Logindb.Recordset.Fields("Username") = txtuname.Text
Logindb.Recordset.Fields("Password") = txtpword.Text
Logindb.Recordset.Fields("Confirm Password") = txtcpword.Text
Logindb.Recordset.Update
MsgBox "Record has been updated successfully"
End Sub
