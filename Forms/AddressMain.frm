VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.MDIForm AddressMain 
   BackColor       =   &H8000000C&
   Caption         =   "Professional Address Book v2.0"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   9690
   Icon            =   "AddressMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "AddressMain.frx":09EA
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   7680
      Width           =   9690
      _ExtentX        =   17092
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4410
            MinWidth        =   4410
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "6:14 PM"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            TextSave        =   "6/10/2010"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuAdd 
         Caption         =   "Add Address"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Address"
      End
      Begin VB.Menu mnuUpdate 
         Caption         =   "Update Address"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "Backup"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuByName 
         Caption         =   "Search by Name"
      End
      Begin VB.Menu mnubyAddress 
         Caption         =   "Search by Address"
      End
      Begin VB.Menu mnubyPhone 
         Caption         =   "Search by Phone"
      End
      Begin VB.Menu mnubyCity 
         Caption         =   "Search by City"
      End
      Begin VB.Menu mnubyCountry 
         Caption         =   "Search by Country"
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "&Settings"
      Begin VB.Menu mnuAddUser 
         Caption         =   "Add User"
      End
      Begin VB.Menu mnuChangePass 
         Caption         =   "Change Password"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuNeedHelp 
         Caption         =   "Help Topics"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAuthor 
         Caption         =   "About Author"
      End
   End
End
Attribute VB_Name = "AddressMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Load()
StatusBar.Panels(1).Text = "Welcome:  " & frmLogon.nUSER
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload frmAbout
    Unload frmBackup
    Unload frmChangePass
    Unload frmLogon
    Unload frmNewUser
    Unload frmRestore
    Unload frmSearch
    Unload frmSearchResult
    Unload rptAddresses
    End
End Sub

Private Sub mnuAdd_Click()
frmMain.Show
End Sub

Private Sub mnuAddUser_Click()
frmNewUser.Show
End Sub


Private Sub mnuAuthor_Click()
frmAbout.Show
End Sub

Private Sub mnuBackup_Click()
frmBackup.Show
End Sub

Private Sub mnubyAddress_Click()
modMain.search = "Address"
frmSearch.Show
End Sub

Private Sub mnubyCity_Click()
modMain.search = "City"
frmSearch.Show
End Sub

Private Sub mnubyCountry_Click()
modMain.search = "Country"
frmSearch.Show
End Sub

Private Sub mnuByName_Click()
modMain.search = "Name"
frmSearch.Show
End Sub

Private Sub mnubyPhone_Click()
modMain.search = "Phone"
frmSearch.Show
End Sub

Private Sub mnuChangePass_Click()
frmChangePass.Show
End Sub

Private Sub mnuDelete_Click()
frmMain.Show
End Sub

Private Sub mnuExit_Click()
If MsgBox("Sure To Exit?", vbYesNo + vbQuestion, "Exit") = vbYes Then
End
Else
Exit Sub
Cancel = True
End If
End Sub


Private Sub mnuNeedHelp_Click()
MsgBox "For any Help Query" & vbCrLf & vbCrLf & "Email: hamzajhang@yahoo.com" & vbCrLf & vbCrLf & "Phone#: +92-334-632-0905", vbInformation, "Help Topic"
End Sub

Private Sub mnuRestore_Click()
frmRestore.Show
End Sub

Private Sub mnuUpdate_Click()
frmMain.Show
End Sub
