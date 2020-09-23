VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About -->Professional Address Book v2.0<--"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7410
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   7410
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label 
      Caption         =   "Phone#: +92-334-632-0905"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   4335
   End
   Begin VB.Label Label 
      Caption         =   "email: hamzajhang@yahoo.com"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   4335
   End
   Begin VB.Label Label 
      Caption         =   "Created by: Hamza Saleem Chohan"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   4335
   End
   Begin VB.Label Label 
      Caption         =   "PROFESSIONAL ADDRESS BOOK VERSION 2.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6855
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call CenterForm(Me)
End Sub
