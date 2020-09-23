VERSION 5.00
Begin VB.Form frmLogon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Professional Address Book v2.0"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5745
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   5745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHelp 
      Caption         =   "?"
      Height          =   255
      Left            =   5400
      TabIndex        =   7
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Ex&it"
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1320
      Width           =   2895
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblName 
      Caption         =   "Login Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblMsg 
      Caption         =   "Please Enter your Login name and Password."
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "frmLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset
Public nUSER As String


Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdHelp_Click()
MsgBox "If you are Running the Address Book First Time the username and password is " & "< admin >", vbInformation, "Help"
End Sub

Private Sub cmdOK_Click()
If rs.State <> 1 Then
    rs.Open "Other", c, adOpenDynamic, adLockOptimistic, adCmdTable
End If
rs.MoveFirst
Do While Not rs.EOF
    If txtName.Text = rs.Fields("Username").Value And txtPassword.Text = rs.Fields("Password").Value Then
    nUSER = rs.Fields("Username")
    Me.Hide
    AddressMain.Show
        Exit Sub
    End If
rs.MoveNext
Loop
If rs.EOF Then
    MsgBox "Invalid Username or Password!", vbCritical, "Error"
End If
rs.Close
End Sub

Private Sub Form_Load()
    Call CenterForm(Me)
    Call Connect
End Sub
