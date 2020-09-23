VERSION 5.00
Begin VB.Form frmSearchResult 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Result"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6405
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSearchResult.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   6405
   Begin VB.Frame fraAddress 
      Caption         =   "Address Book"
      Height          =   4215
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   6135
      Begin VB.TextBox txtID 
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
         Top             =   575
         Width           =   1215
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
         Left            =   1560
         TabIndex        =   1
         Top             =   960
         Width           =   3855
      End
      Begin VB.TextBox txtAddress 
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
         TabIndex        =   2
         Top             =   1320
         Width           =   3855
      End
      Begin VB.TextBox txtCity 
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
         TabIndex        =   3
         Top             =   1680
         Width           =   3855
      End
      Begin VB.TextBox txtCountry 
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
         TabIndex        =   4
         Top             =   2040
         Width           =   3855
      End
      Begin VB.TextBox txtPhone 
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
         TabIndex        =   5
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox txtEmail 
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
         TabIndex        =   6
         Top             =   2760
         Width           =   3855
      End
      Begin VB.Label lblID 
         Caption         =   "Record ID:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblName 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblAddress 
         Caption         =   "Address:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblCity 
         Caption         =   "City:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label lblCountry 
         Caption         =   "Country:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label lblPhone 
         Caption         =   "Phone:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label lblEmail 
         Caption         =   "Email:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   2760
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "<<"
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<"
      Height          =   495
      Left            =   960
      TabIndex        =   10
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   495
      Left            =   1680
      TabIndex        =   9
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">>"
      Height          =   495
      Left            =   2400
      TabIndex        =   8
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   5280
      TabIndex        =   7
      Top             =   4560
      Width           =   855
   End
End
Attribute VB_Name = "frmSearchResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdFirst_Click()
r.MoveFirst
Call Fill
End Sub

Private Sub cmdLast_Click()
r.MoveLast
Call Fill
End Sub

Private Sub cmdNext_Click()
r.MoveNext
If r.EOF Then
    r.MoveLast
End If
Call Fill
End Sub

Private Sub cmdPrevious_Click()
r.MovePrevious
If r.BOF Then
    r.MoveFirst
End If
Call Fill
End Sub

Sub Fill()
txtID.Text = rSearch("ID")
txtName.Text = rSearch("Name")
txtAddress.Text = rSearch("Address")
txtCity.Text = rSearch("City")
txtCountry.Text = rSearch("Country")
txtPhone.Text = rSearch("Phone")
txtEmail.Text = rSearch("Email")
End Sub



Private Sub Form_Load()
Call modMain.CenterForm(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
rSearch.Close
End Sub
