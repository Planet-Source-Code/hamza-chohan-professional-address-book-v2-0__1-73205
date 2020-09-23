VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Form"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6390
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   6390
   Begin VB.CommandButton cmdReport 
      Caption         =   "&View Report"
      Height          =   495
      Left            =   4560
      TabIndex        =   24
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   3600
      TabIndex        =   23
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   2760
      TabIndex        =   22
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   1920
      TabIndex        =   21
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   1080
      TabIndex        =   20
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   495
      Left            =   240
      TabIndex        =   19
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">>"
      Height          =   495
      Left            =   2400
      TabIndex        =   18
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   495
      Left            =   1680
      TabIndex        =   17
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<"
      Height          =   495
      Left            =   960
      TabIndex        =   16
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "<<"
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   4560
      Width           =   735
   End
   Begin VB.Frame fraAddress 
      Caption         =   "Address Book"
      Height          =   4215
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6135
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
      Begin VB.Label lblEmail 
         Caption         =   "Email:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label lblPhone 
         Caption         =   "Phone:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label lblCountry 
         Caption         =   "Country:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label lblCity 
         Caption         =   "City:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label lblAddress 
         Caption         =   "Address:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblName 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblID 
         Caption         =   "Record ID:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
Call Fill
Call AddMode
txtID.Text = ""
txtName.Text = ""
txtAddress.Text = ""
txtCity.Text = ""
txtCountry.Text = ""
txtPhone.Text = ""
txtEmail.Text = ""
txtID.SetFocus
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
c.Execute ("Delete from Addresses where ID = '" & txtID.Text & "';")
r.Requery
Call Fill
MsgBox "Record Deleted Sucessfully.", vbInformation
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

Private Sub cmdReport_Click()
rptAddresses.Show
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrHand
If txtID.Text = "" Then
    MsgBox "Enter ID First!", vbCritical, "Save"
ElseIf txtName.Text = "" Then
    MsgBox "Enter Name!", vbCritical, "Save"
ElseIf txtAddress.Text = "" Then
    MsgBox "Enter Address!", vbCritical, "Save"
Else
    Set rCheck = c.Execute("Select * from Addresses where ID='" & txtID.Text & "';")
    If rCheck.EOF Then
        c.Execute "Insert into Addresses(ID, Name, Address, City, Country, Phone, Email) values('" & txtID.Text & "','" & txtName.Text & "', '" & txtAddress.Text & "','" & txtCity.Text & "','" & txtCountry.Text & "','" & txtPhone.Text & "','" & txtEmail.Text & "');"
        MsgBox "Record Saved Successfully!", vbInformation, "Save"
        r.Requery
        Call Fill
        Call ViewMode
    Else
        MsgBox "Record already exists!"
    End If
End If
Exit Sub
ErrHand:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo ErrHand
If txtName.Text = "" Then
    MsgBox "Enter Name!", vbCritical, "Save"
    txtName.SetFocus
ElseIf txtAddress.Text = "" Then
    MsgBox "Enter Address!", vbCritical, "Save"
    txtAddress.SetFocus
Else
    c.Execute ("Update Addresses set Name='" & txtName.Text & "',Address='" & txtAddress.Text & "',City='" & txtCity.Text & "',Country='" & txtCountry.Text & "',Phone='" & txtPhone.Text & "',Email='" & txtEmail.Text & "' where ID='" & txtID.Text & "';")
    MsgBox "Record Updated Successfully!", vbInformation, "Save"
End If
Exit Sub
ErrHand:
    MsgBox Err.Description, vbCritical, "Error"
End Sub


Public Sub Fill()
txtID.Text = r.Fields("ID")
txtName.Text = r.Fields("Name")
txtAddress.Text = r.Fields("Address")
txtCity.Text = r.Fields("City")
txtCountry.Text = r.Fields("Country")
txtPhone.Text = r.Fields("Phone")
txtEmail.Text = r.Fields("Email")
End Sub

Private Sub Form_Load()
Call Fill
Call CenterForm(Me)
r.Requery
End Sub

Public Sub AddMode()
cmdAdd.Enabled = False
cmdSave.Enabled = True
cmdDelete.Enabled = False
cmdUpdate.Enabled = False
cmdClose.Enabled = False
End Sub

Public Sub ViewMode()
cmdAdd.Enabled = True
cmdSave.Enabled = False
cmdDelete.Enabled = True
cmdUpdate.Enabled = True
cmdClose.Enabled = True
End Sub


