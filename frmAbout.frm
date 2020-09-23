VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00C00000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   5460
   ClientLeft      =   4620
   ClientTop       =   3765
   ClientWidth     =   5730
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3768.589
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      ClipControls    =   0   'False
      Height          =   540
      Left            =   135
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   6
      Top             =   210
      Width           =   540
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1950
      Left            =   45
      TabIndex        =   5
      Top             =   3330
      Width           =   5640
      Begin VB.Image Image1 
         Height          =   1575
         Left            =   -75
         Picture         =   "frmAbout.frx":030A
         Stretch         =   -1  'True
         Top             =   330
         Width           =   5700
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4260
      TabIndex        =   0
      Top             =   2655
      Width           =   1260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   70.429
      X2              =   5295.313
      Y1              =   1511.577
      Y2              =   1511.577
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":611A
      ForeColor       =   &H000000FF&
      Height          =   1155
      Left            =   1020
      TabIndex        =   1
      Top             =   945
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "Application Title"
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   1050
      TabIndex        =   3
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1521.93
      Y2              =   1521.93
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   1035
      TabIndex        =   4
      Top             =   600
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":622C
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   285
      TabIndex        =   2
      Top             =   2265
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
End Sub

