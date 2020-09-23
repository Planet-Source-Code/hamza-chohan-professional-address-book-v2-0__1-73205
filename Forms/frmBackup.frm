VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBackup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup Data"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7800
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBackup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   7800
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
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin MSComDlg.CommonDialog ComDlg 
         Left            =   360
         Top             =   1080
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   375
         Left            =   6240
         TabIndex        =   5
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdBackup 
         Caption         =   "&Backup"
         Height          =   375
         Left            =   5040
         TabIndex        =   4
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "B&rowse"
         Height          =   375
         Left            =   6240
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtPath 
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
         Left            =   840
         TabIndex        =   2
         Top             =   480
         Width           =   5175
      End
      Begin VB.Label lblPath 
         Caption         =   "Path:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim Files As New FileSystemObject
    
Private Sub cmdBackup_Click()
On Error GoTo ErrEnd
    If Trim(txtPath.Text) = Empty Then
        MsgBox "Enter Valid Path", vbCritical, "Backup"
        Exit Sub
    End If
    If Files.FileExists(txtPath.Text) = True Then
        If txtPath.Text = App.Path & "\Database\AddressBook.mdb" Then
            MsgBox "Enter Valid Path", vbCritical, "Backup"
            Exit Sub
        End If
        Files.DeleteFile (txtPath.Text)
        Files.CopyFile App.Path & "\Database\AddressBook.mdb", txtPath.Text, True
    Else
        Files.CopyFile App.Path & "\Database\AddressBook.mdb", txtPath.Text, True
    End If
    MsgBox "Data Backup Successfully", vbInformation, "Backup"
    Exit Sub
ErrEnd:
    MsgBox "Backup unsucessfull due to " & Err.Description, vbCritical
End Sub

Private Sub cmdBrowse_Click()
    ComDlg.Filter = "Microsoft Access Files (*.mdb)|*.mdb"
    ComDlg.FilterIndex = 1
    ComDlg.ShowSave
    txtPath.Text = ComDlg.FileName
    Exit Sub
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call CenterForm(Me)
txtPath.Text = App.Path & "\Database\AddressBook.mdb"
End Sub
