VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   3960
   ClientLeft      =   5475
   ClientTop       =   5250
   ClientWidth     =   2775
   DrawWidth       =   3
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   264
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   185
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H0000FFFF&
      Height          =   1005
      Left            =   1455
      Pattern         =   "*.bmp;*.dib;*.jpg;*.gif"
      TabIndex        =   8
      Top             =   2520
      Width           =   1320
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   1455
      TabIndex        =   7
      Top             =   2175
      Width           =   1365
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   630
      Left            =   2100
      ScaleHeight     =   570
      ScaleWidth      =   495
      TabIndex        =   6
      Top             =   420
      Width           =   555
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   3525
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Random Dots"
      Height          =   375
      Left            =   1455
      TabIndex        =   5
      Top             =   1830
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   405
      Left            =   1455
      TabIndex        =   4
      Top             =   3570
      Width           =   1365
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      DrawWidth       =   2
      Height          =   3735
      Left            =   0
      ScaleHeight     =   100
      ScaleMode       =   0  'User
      ScaleWidth      =   40
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1560
      Left            =   1470
      ScaleHeight     =   104
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   0
      Top             =   270
      Width           =   600
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C00000&
      Caption         =   "Color:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2130
      TabIndex        =   3
      Top             =   195
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "Pattern Draw"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Stereogram generator
' Created By Andy Nova*
' andy@highsupport.com
' http://www.highsupport.com
Option Explicit

Private Sub Command1_Click()
Form2.Hide
End Sub
Public Sub Command2_Click()
Dim n1 As Integer, n2 As Integer
'Picture1.Cls
For n1 = 0 To Picture1.ScaleHeight
For n2 = 0 To Picture1.ScaleWidth
If Int(Rnd * 2) = 1 Then Picture1.PSet (n2, n1), Picture3.BackColor
Next n2
Next n1
End Sub
Private Sub Command3_Click()
Picture1.Cls
Picture2.Cls
End Sub

Private Sub File1_Click()
'If Combo1.Text = "Pictures" Then Picture1.Picture = LoadPicture(""): Exit Sub
Picture1.Picture = LoadPicture(File1)
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.AutoRedraw = True
Picture1.PSet (X, Y), Picture3.BackColor
Picture2.PSet (X, Y), Picture3.BackColor
End Sub
Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Picture1.AutoRedraw = True
Picture1.PSet (X, Y), Picture3.BackColor
Picture2.PSet (X, Y), Picture3.BackColor
End Sub
Private Sub Picture3_Click()
CommonDialog1.ShowColor
Picture3.BackColor = CommonDialog1.Color
End Sub
