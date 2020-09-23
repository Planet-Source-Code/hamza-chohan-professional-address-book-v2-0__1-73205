VERSION 5.00
Begin VB.Form frmChangePass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5685
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChangePass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   5685
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Can&cel"
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Frame fraMain 
      Height          =   1935
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5415
      Begin VB.TextBox txtNewPass2 
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
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtNewPass1 
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
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1080
         Width           =   2055
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
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   720
         Width           =   2055
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
         Left            =   2040
         TabIndex        =   0
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label lblNewPass2 
         Caption         =   "Confirm Password:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblNewPass1 
         Caption         =   "New Password:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblOldPass 
         Caption         =   "Old Password:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblName 
         Caption         =   "Login Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmChangePass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Dim rs As New ADODB.Recordset
If rs.State = 1 Then
rs.Close
End If
If rs.State <> 1 Then
    rs.Open "Other", c, adOpenDynamic, adLockOptimistic, adCmdTable
End If
rs.MoveFirst
Do While Not rs.EOF
    If rs.Fields("Username").Value = txtName.Text And rs.Fields("Password").Value = txtPassword.Text And txtNewPass1.Text = txtNewPass2.Text Then
    rs.Fields("Password") = txtNewPass1.Text
    rs.Update
    rs.Close
    MsgBox "Password Changed!", vbInformation
    Unload Me
    Exit Sub
End If
    rs.MoveNext
Loop
If rs.EOF Then
    MsgBox "Invalid Old Password!", vbCritical
End If
    rs.Close
End Sub

Private Sub Form_Load()
    Call CenterForm(Me)
End Sub
