VERSION 5.00
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6195
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   6195
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtCriteria 
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
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   3735
   End
   Begin VB.Label lblName 
      Caption         =   "Enter Name"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSearch_Click()
If search = "Name" Then
    rSearch.Open "Select * from Addresses where Name='" & txtCriteria & "'", c, adOpenDynamic, adLockOptimistic
ElseIf search = "Address" Then
    rSearch.Open "Select * from Addresses where Address='" & txtCriteria & "'", c, adOpenDynamic, adLockOptimistic
ElseIf search = "City" Then
    rSearch.Open "Select * from Addresses where City='" & txtCriteria & "'", c, adOpenDynamic, adLockOptimistic
ElseIf search = "Country" Then
    rSearch.Open "Select * from Addresses where Country='" & txtCriteria & "'", c, adOpenDynamic, adLockOptimistic
ElseIf search = "Phone" Then
    rSearch.Open "Select * from Addresses where Phone='" & txtCriteria & "'", c, adOpenDynamic, adLockOptimistic
ElseIf search = "Email" Then
    rSearch.Open "Select * from Addresses where Email='" & txtCriteria & "'", c, adOpenDynamic, adLockOptimistic
End If
If rSearch.EOF Then
    MsgBox "No Record Found!", vbCritical
    rSearch.Close
Else
    frmSearchResult.Show
    Call frmSearchResult.Fill
    Unload Me
End If
End Sub

Private Sub Form_Load()
lblName.Caption = "Enter" & search
Call modMain.CenterForm(Me)
End Sub
