VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF8080&
   Caption         =   "Departmental Stores"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   ScaleHeight     =   12375
   ScaleWidth      =   22800
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbuname 
      Height          =   315
      Left            =   7560
      TabIndex        =   12
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton cmddelete 
      BackColor       =   &H80000013&
      Caption         =   "Delete"
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
      Left            =   8040
      TabIndex        =   7
      ToolTipText     =   "select the uname of List"
      Top             =   7080
      Width           =   1215
   End
   Begin VB.TextBox txtcpword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   7560
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   5160
      Width           =   1935
   End
   Begin VB.TextBox txtpword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   7560
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox txtuname 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      MaxLength       =   8
      TabIndex        =   4
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton cmdadd 
      BackColor       =   &H80000013&
      Caption         =   "Add"
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
      Left            =   4920
      TabIndex        =   3
      ToolTipText     =   "adda new record"
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdmain 
      BackColor       =   &H80000013&
      Cancel          =   -1  'True
      Caption         =   "Main"
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
      Left            =   9600
      TabIndex        =   2
      ToolTipText     =   "back to main"
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H80000013&
      Caption         =   "Save"
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
      Left            =   6480
      TabIndex        =   1
      ToolTipText     =   "saves record"
      Top             =   7080
      Width           =   1215
   End
   Begin VB.ComboBox cmbt_user 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "login_details.frx":0000
      Left            =   7560
      List            =   "login_details.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "user"
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      Height          =   3015
      Left            =   5280
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   5055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LOGIN DETAILS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   6120
      TabIndex        =   13
      Top             =   2760
      Width           =   3225
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   11
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   " Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   10
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   9
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "Type of User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   8
      Top             =   5760
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim S As String

Private Sub cmbuname_click()
Set rs = New ADODB.Recordset
rs.Open "SELECT * FROM user_details", cn, adOpenDynamic, adLockOptimistic
While Not rs.EOF
If rs(0) = cmbuname.Text Then
txtpword.Text = rs(1)
txtcpword.Text = rs(2)
cmbt_user.Text = rs(3)
End If
rs.MoveNext
Wend

End Sub

Private Sub cmdadd_Click()
cmbuname.Visible = False
'rs.CancelUpdate
txtuname.Text = ""
txtpword.Text = ""
txtcpword.Text = ""
'cmbt_user.AddItem "user"
txtuname.SetFocus
'rs.AddNew
End Sub

Private Sub cmdmain_Click()
Unload Me

End Sub

Private Sub cmdsave_Click()
 If LTrim(RTrim(txtuname.Text)) = "" Then
 txtuname.Text = cmbuname.Text
 End If
 

    If LTrim(RTrim(txtuname.Text)) = "" Then
       MsgBox " PLZ ENTER USERNAME"
       txtuname.Text = ""
       txtuname.SetFocus
        GoTo endpara
    ElseIf LTrim(RTrim(txtpword.Text)) = "" Then
       MsgBox " PLZ ENTER  PASSWORD"
       txtpword.Text = ""
       txtpword.SetFocus
       GoTo endpara
    ElseIf LTrim(RTrim(txtcpword.Text)) = "" Then
       MsgBox " PLZ ENTER  CONFIRM PASSWORD"
       txtcpword.Text = ""
       txtcpword.SetFocus
       GoTo endpara
     ElseIf Not (txtpword.Text = txtcpword.Text) Then
       MsgBox " PASSWORD AND CONFIRM PASSWORD NOT MATCHING"
       GoTo endpara
     ElseIf LTrim(RTrim(cmbt_user.Text)) = "" Then
       MsgBox " PLZ ENTER type of USERID"
      GoTo endpara
    Else
Set rs = New ADODB.Recordset
rs.Open "SELECT * FROM user_details", cn, adOpenDynamic, adLockOptimistic

While Not rs.EOF
If txtuname.Text = rs(0) Then
MsgBox "name already exists"
cmdadd_Click
GoTo endpara
End If
rs.MoveNext
Wend
 

cn.Execute "INSERT INTO user_details VALUES('" & txtuname.Text & "','" & txtpword.Text & "','" & txtcpword.Text & "','" & cmbt_user.Text & "')"

Dim r As Integer

MsgBox "RECOrD IS SAVED"
End If
txtuname.Text = ""
txtpword.Text = ""
txtcpword.Text = ""

endpara:
  'txtcpword.SetFocus
Form_Load
End Sub

Private Sub Form_Load()
'Call open_cn
cmbt_user.Clear
 Set cn = New ADODB.Connection
cn.ConnectionString = "DSN=p;UID=scott;PWD=tiger;"
cn.Open

Set rs = New ADODB.Recordset
Set rs = New ADODB.Recordset
rs.Open "SELECT * FROM user_details", cn, adOpenDynamic, adLockOptimistic
cmbuname.Visible = True
cmbuname.Clear
While Not rs.EOF
cmbuname.AddItem rs(0)
rs.MoveNext
Wend

cmbt_user.AddItem "user"
cmbt_user.AddItem "admin"
cmbt_user.Visible = True


End Sub

Public Sub initgrid()
Dim S As String
S$ = "<sl_no|user name|type_user>"
Grid1.FormatString = S$
Grid1.ColWidth(0) = 800
Grid1.ColWidth(1) = 1200
Grid1.ColWidth(2) = 1200
Grid1.Rows = 2
End Sub



Private Sub Form_Activate()
'get_users
rs.AddNew
txtuname.Text = ""
txtpword.Text = ""
txtcpword.Text = ""
'txtcpword.SetFocus
End Sub


Private Sub cmddelete_Click()
  If LTrim(RTrim(cmbuname.Text)) = "" Then
       MsgBox " PLZ ENTER USERNAME"
       txtuname.Text = ""
       txtuname.SetFocus
        GoTo endpara
    ElseIf LTrim(RTrim(txtpword.Text)) = "" Then
       MsgBox " PLZ ENTER  PASSWORD"
       txtpword.Text = ""
       txtpword.SetFocus
       GoTo endpara
    ElseIf LTrim(RTrim(txtcpword.Text)) = "" Then
       MsgBox " PLZ ENTER  CONFIRM PASSWORD"
       txtcpword.Text = ""
       txtcpword.SetFocus
       GoTo endpara
     ElseIf Not (txtpword.Text = txtcpword.Text) Then
       MsgBox " PASSWORD AND CONFIRM PASSWORD NOT MATCHING"
       GoTo endpara
     ElseIf LTrim(RTrim(cmbt_user.Text)) = "" Then
       MsgBox " PLZ ENTER type of USERID"
      GoTo endpara
    Else

   Dim a As String
   a = MsgBox("do you want to delete the user", vbYesNo)
   If a = vbYes Then
     
      
        cn.Execute " DELETE FROM user_details WHERE uname='" & cmbuname.Text & "'"
       ' get_users
        MsgBox " USER DELETED"
        txtuname.Text = ""
        txtpword.Text = ""
        txtcpword.Text = ""
        'cmbt_user.Text = ""
        'txtcpword.SetFocus
       GoTo endpara
    
      
  
  Else
  txtuname.Text = ""
  txtpword.Text = ""
  txtcpword.Text = ""
 GoTo endpara
 
 End If
 End If
' get_users
endpara:
Form_Load
End Sub


