VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H008080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Protected"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   4755
   FillColor       =   &H00C0E0FF&
   ForeColor       =   &H0000FF00&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4755
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List2 
      Height          =   1230
      ItemData        =   "Form1.frx":1708A
      Left            =   2400
      List            =   "Form1.frx":1709D
      TabIndex        =   7
      Top             =   3480
      Width           =   4095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Lock USB"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   360
      TabIndex        =   6
      Top             =   2520
      Width           =   3975
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2040
      Width           =   3975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Unlock USB"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "a"
      TabIndex        =   0
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "Defaults PassWord : NO"
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   675
      Left            =   120
      Picture         =   "Form1.frx":17148
      Stretch         =   -1  'True
      Top             =   120
      Width           =   720
   End
   Begin VB.Label L 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "Enter PassWord"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   3855
   End
   Begin VB.Menu a 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Const LB_DIR As Long = &H18D
Enum DirListConstant
    DDL_EXCLUSIVE = &H8000
    DDL_ARCHIVE = &H20
    DDL_DIRECTORY = &H10
    DDL_DRIVES = &H4000
    DDL_HIDDEN = &H2
    DDL_READONLY = &H1
    DDL_SYSTEM = &H4
End Enum
Dim APPPath As String
Function bFileExists(sFile As String) As Boolean
 On Error Resume Next
 bFileExists = ((GetAttr(sFile) And vbDirectory) = 0)
End Function

Private Sub a_Click()
MsgBox vbCrLf & vbCrLf & vbCrLf & vbCrLf & "This Program CopyRight By DinhQuangTrung90@yahoo.com" & vbCrLf & vbCrLf & vbCrLf & vbCrLf
End Sub

Private Sub Command1_Click()
If Text1.Text = ReadIniFile(APPPath & "Log.QuangTrung", "Protect", "Pass", "NO") Then
'Dang Nhap Thanh cong
MsgBox "OK", vbOKOnly, "OK"
a = Shell("explorer" & " " & APPPath, vbMaximizedFocus)
MoKhoaUSB


'Dang nhap thanh cong roi thi THOATTT
End
Else
'Dang Nhap That Bai
MsgBox "Password Correct !!!"
SendKeys "+{TAB}"
Text1.Text = ""
End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
CP.Show
End Sub


Private Sub Command4_Click()
KhoaUSB
MsgBox "Locked"
End Sub

Private Sub Form_Load()

APPPath = App.Path
If Right(APPPath, 1) <> "\" Then
APPPath = APPPath & "\"
End If



List1.Clear
Dim sFile As String
    sFile = Dir$(APPPath, vbSystem + vbHidden + vbDirectory)
    Do Until Len(sFile) = 0
        List1.AddItem sFile
        sFile = Dir$
        Loop
For x = 0 To List1.ListCount - 1

If List1.List(x) = "." Then
List1.RemoveItem (x)
End If

If List1.List(x) = ".." Then
List1.RemoveItem (x)
End If

If List1.List(x) = "autorun.inf" Then
List1.RemoveItem (x)
End If

If List1.List(x) = "Log.QuangTrung" Then
List1.RemoveItem (x)
End If


If List1.List(x) = "USBProtect.exe" Then
List1.RemoveItem (x)
End If

Next x
If bFileExists(APPPath & "autorun.inf") = False Then
On Error GoTo thoat
Dim i, FileName, FileNumber, afiLename
   ' Obtain Folder where this program's EXE file resides
   FileName = App.Path
   ' Make sure FileName ends with a backslash
   If Right(FileName, 1) <> "\" Then FileName = FileName & "\"
   FileName = FileName & "autorun.inf"  ' name output text file MyList.txt
   ' Obtain an available filenumber from the operating system
   FileNumber = FreeFile
   ' Open the FileName as an output file , using FileNumber as FileHandle
   Open FileName For Output As FileNumber
   ' Now iterate through each item of lstNames
   For i = 0 To List2.ListCount - 1
      ' Write the List item to file. Make sure you use symbol # in front of FileNumber
      Print #FileNumber, List2.List(i)
   Next
   Close FileNumber
thoat:
'SetAttr APPPath & "autorun.inf", vbHidden + vbReadOnly + vbSystem
   End If
End Sub




Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Command1_Click
SendKeys "{TAB}"
End If
End Sub




Private Sub KhoaUSB()
For h = 0 To List1.ListCount - 1
'SetAttr APPPath & List1.List(h), vbHidden + vbSystem + vbReadOnly
Next h
End Sub
Private Sub MoKhoaUSB()
For h = 0 To List1.ListCount - 1
'SetAttr APPPath & List1.List(h), vbNormal
Next h
End Sub
