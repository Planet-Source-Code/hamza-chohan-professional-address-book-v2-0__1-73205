VERSION 5.00
Begin VB.Form frmNewUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New User"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6975
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNewUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   6975
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Can&cel"
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Frame fraMain 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6735
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
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   4575
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
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label lblName 
         Caption         =   "Login Name:"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblPassword 
         Caption         =   "Password:"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmNewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Dim r As New ADODB.Recordset
If txtName.Text = "" Then
    MsgBox "Enter User Name.", vbCritical
Else
    If r.State <> 1 Then
        r.Open "Other", c, adOpenDynamic, adLockOptimistic, adCmdTable
    End If
r.MoveFirst
Do While Not r.EOF
If txtName.Text = r.Fields("Username") Then
    MsgBox "User Already Exists!", vbCritical
    r.Close
    Exit Sub
End If
r.MoveNext
Loop
r.AddNew
r.Fields("Username").Value = txtName.Text
r.Fields("Password").Value = txtPassword.Text
r.Update
r.Close
MsgBox "New User Created!", vbInformation
End If
End Sub

Private Sub Form_Load()
Call CenterForm(Me)
End Sub
