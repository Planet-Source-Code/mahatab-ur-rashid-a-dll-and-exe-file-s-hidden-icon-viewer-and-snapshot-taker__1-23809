VERSION 5.00
Begin VB.Form frmCopyIcon 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   2745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3000
      Top             =   1440
   End
   Begin VB.PictureBox Pic1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "frmCopyIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
frmCopyIcon.Height = Pic1.Height
frmCopyIcon.Width = Pic1.Width
End Sub

Private Sub Timer1_Timer()
Static k As Integer

ExtractIconAndShow frmMain.File1.FileName, (IconNumber - 1), Pic1

If (k >= 2) Then
keybd_event VK_MENU, 0, 0, 0  ' Press Alt
keybd_event VK_SNAPSHOT, 0, 0, 0  ' Press PrintScreen
keybd_event VK_SNAPSHOT, 0, KEYEVENTF_KEYUP, 0  ' Release Alt
keybd_event VK_MENU, 0, KEYEVENTF_KEYUP, 0  ' Release PrintScreen
End If

If (k >= 3) Then
k = 0
Unload Me
End If

k = k + 1
End Sub
