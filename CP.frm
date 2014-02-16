VERSION 5.00
Begin VB.Form CP 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Change Password"
   ClientHeight    =   1785
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4650
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Cofirm Password"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "New Password"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Old Password"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "CP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim APPPath As String
Private Sub Command1_Click()
If Text1.Text = ReadIniFile(APPPath & "Log.QuangTrung", "Protect", "Pass", "NO") And Text2.Text = Text3.Text Then
MsgBox "OK"
WriteIniFile APPPath & "Log.QuangTrung", "Protect", "Pass", Text2.Text
SetAttr APPPath & "Log.QuangTrung", vbHidden + vbReadOnly + vbSystem
Unload Me
Else
MsgBox "Password Correct !"
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
APPPath = App.path
If Right(APPPath, 1) <> "\" Then
APPPath = APPPath & "\"
End If
End Sub
