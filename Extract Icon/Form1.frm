VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "[ DLL- EXE File's ] Icon Viewer"
   ClientHeight    =   6690
   ClientLeft      =   2535
   ClientTop       =   555
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   6870
   Begin VB.Frame Frame3 
      Caption         =   "[ Control Panel ]"
      Height          =   975
      Left            =   2760
      TabIndex        =   55
      Top             =   5400
      Width           =   3975
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   375
         Left            =   2640
         TabIndex        =   59
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy &Icon"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1440
         TabIndex        =   58
         ToolTipText     =   "Copy Icon to Clipboard"
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   57
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSComctlLib.ProgressBar Prog1 
      Height          =   150
      Left            =   4800
      TabIndex        =   54
      Top             =   6500
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   53
      Top             =   6435
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   " Extracted Icons "
      Height          =   5295
      Left            =   2760
      TabIndex        =   4
      Top             =   120
      Width           =   3975
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   47
         Left            =   3240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   52
         Top             =   4560
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   46
         Left            =   2640
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   51
         Top             =   4560
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   45
         Left            =   2040
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   50
         Top             =   4560
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   44
         Left            =   1440
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   49
         Top             =   4560
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   43
         Left            =   840
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   48
         Top             =   4560
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   42
         Left            =   240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   47
         Top             =   4560
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   41
         Left            =   3240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   46
         Top             =   3960
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   40
         Left            =   2640
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   45
         Top             =   3960
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   39
         Left            =   2040
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   44
         Top             =   3960
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   38
         Left            =   1440
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   43
         Top             =   3960
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   37
         Left            =   840
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   42
         Top             =   3960
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   36
         Left            =   240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   41
         Top             =   3960
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   35
         Left            =   3240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   40
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   34
         Left            =   2640
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   39
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   33
         Left            =   2040
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   38
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   32
         Left            =   1440
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   37
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   31
         Left            =   840
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   36
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   30
         Left            =   240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   35
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   29
         Left            =   3240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   34
         Top             =   2760
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   28
         Left            =   2640
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   33
         Top             =   2760
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   27
         Left            =   2040
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   32
         Top             =   2760
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   26
         Left            =   1440
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   31
         Top             =   2760
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   25
         Left            =   840
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   30
         Top             =   2760
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   24
         Left            =   240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   29
         Top             =   2760
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   23
         Left            =   3240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   28
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   22
         Left            =   2640
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   27
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   21
         Left            =   2040
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   26
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   20
         Left            =   1440
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   25
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   19
         Left            =   840
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   24
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   18
         Left            =   240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   23
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   12
         Left            =   240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   22
         Top             =   1560
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   13
         Left            =   840
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   21
         Top             =   1560
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   14
         Left            =   1440
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   20
         Top             =   1560
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   15
         Left            =   2040
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   19
         Top             =   1560
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   16
         Left            =   2640
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   18
         Top             =   1560
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   17
         Left            =   3240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   17
         Top             =   1560
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   10
         Left            =   2640
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   16
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   11
         Left            =   3240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   15
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   6
         Left            =   240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   14
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   7
         Left            =   840
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   13
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   8
         Left            =   1440
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   12
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   9
         Left            =   2040
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   11
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   5
         Left            =   3240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   10
         Top             =   360
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   0
         Left            =   240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   9
         Top             =   360
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   2
         Left            =   1440
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   8
         Top             =   360
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   1
         Left            =   840
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   7
         Top             =   360
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   4
         Left            =   2640
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   3
         Left            =   2040
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Select a File "
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.DirListBox Dir1 
         Height          =   1890
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   2295
      End
      Begin VB.FileListBox File1 
         Height          =   2235
         Left            =   120
         Pattern         =   "*.exe;*.dll"
         TabIndex        =   2
         Top             =   2880
         Width           =   2295
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2160
      Top             =   5880
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      Caption         =   "Status : None"
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      TabIndex        =   56
      Top             =   5520
      Width           =   2535
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim NextClicked As Integer 'Click counter of cmdNext

Private Sub Pic1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'Storing actual Index number of a Icon as ToolTip Text
If NextClicked <= 0 Then
Pic1(Index).ToolTipText = "Index number " & (Index + 1)
Else
Pic1(Index).ToolTipText = "Index number " & ((NextClicked * 46) + (Index + 1)) + NextClicked
End If
End Sub

Private Sub Timer1_Timer()

If FileName <> "" Then
NumberOfIcon = ExtractIconAndShow(FileName, -1, Pic1(I))

StatusBar1.SimpleText = "Total " & NumberOfIcon & " Icon(s) in " & FileName

If NumberOfIcon <> 0 Then
'Initializing progress bar
Prog1.Min = 0: Prog1.Max = NumberOfIcon
Else
Exit Sub
End If
End If

'Calling ExtractIconAndShow
ExtractIconAndShow FileName, j, Pic1(I)

' If all picture box is full
If (I >= 47) Then
lblInfo.Caption = "Status : Click Next button to view more!"
cmdNext.Enabled = True
Beep
Timer1.Enabled = False
Exit Sub
End If

I = I + 1       'Increasing picture boxes index
j = j + 1       'Increasing icon's index of file
Prog1.Value = j 'Increasing progress bar's value

'If all icon shown
If Prog1.Value = NumberOfIcon Then
Prog1.Value = 0
I = 0
temp = j
j = 0
Timer1.Enabled = False
Exit Sub
End If
End Sub

'Clear all Picture Box contents
Private Sub clrPicbox()
For I = 0 To 47
Pic1(I).Cls
Next I
End Sub

Private Sub cmdClose_Click()
End
End Sub

Private Sub cmdNext_Click()
clrPicbox
I = 0
lblInfo.Caption = ""
Timer1.Enabled = True
NextClicked = NextClicked + 1 ' Counting Mouse click on this button
cmdNext.Enabled = False
End Sub


'Copy a Icon to ClipBoard
Private Sub cmdCopy_Click()
Dim rc As Long
On Error GoTo errhnd

If File1.FileName = "" Then
MsgBox "You must select a file!", vbInformation, "No file selected!"
cmdCopy.Enabled = False
Exit Sub
End If

'Get a icon's Index number
IconNumber = InputBox("Icon Index:", "Enter a Icon index")

If IconNumber > NumberOfIcon Then Exit Sub 'Invalid Index Number

'Load the for which get ScreenShot
Load frmCopyIcon
frmCopyIcon.Show

errhnd:
If Err.Number <> 0 Then
MsgBox Err.Description, vbCritical, "Error!"
End If
End Sub


Private Sub Dir1_Change()
File1.Path = Dir1.Path
lblInfo.Caption = "Status : You are in " & Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
FileName = File1.FileName
clrPicbox
I = 0: j = 0
NextClicked = 0
cmdCopy.Enabled = True
Timer1.Enabled = True
End Sub


