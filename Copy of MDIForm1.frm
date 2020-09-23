VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Strereogram"
   ClientHeight    =   3480
   ClientLeft      =   4905
   ClientTop       =   3720
   ClientWidth     =   4965
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu ma 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuStereogram 
      Caption         =   "&Stereogram"
      Begin VB.Menu mnuGenerate 
         Caption         =   "&Generate"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options"
         Shortcut        =   ^T
      End

   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub
Private Sub mnuExit_Click()
End
End Sub
Private Sub mnuCopy_Click()
copier MDIForm1.activeform.Picture1.Image
End Sub
Private Sub MDIForm_Load()
Form2.Command2_Click
mnuSave.Enabled = False
mnuCopy.Enabled = False
mnuGenerate.Enabled = False
End Sub
Private Sub mnuOptions_Click()
Form2.Show vbModal, Me
End Sub
Private Sub mnuOpen_Click()
opener Me
End Sub
Private Sub mnuSave_Click()
saver MDIForm1.activeform.Picture1.Image
End Sub
Private Sub mnuPaste_Click()
paster
End Sub
Private Sub mnuGenerate_Click()
MDIForm1.activeform.copyable = True
MDIForm1.mnuCopy = True
MDIForm1.mnuSave = True
generate MDIForm1.activeform.Picture1, Form2.Picture1
End Sub

Private Sub mnustop_Click()

End Sub
