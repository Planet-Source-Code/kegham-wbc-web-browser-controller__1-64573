VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Web browser controller"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6645
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   6375
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Remember my password on this computer"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox txtuser 
         Enabled         =   0   'False
         Height          =   285
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox txtpass 
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   360
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1320
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&OK"
         Default         =   -1  'True
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   325
         Left            =   360
         TabIndex        =   4
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label lblwait 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   330
         Left            =   1320
         TabIndex        =   10
         Top             =   2280
         Width           =   4815
      End
      Begin VB.Image Image1 
         Height          =   1455
         Left            =   4080
         Picture         =   "Form1.frx":2982
         Top             =   360
         Width           =   1965
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Email Address:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Width           =   1155
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Height          =   2535
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   6135
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   6000
      Top             =   120
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   4845
      Left            =   120
      TabIndex        =   11
      Top             =   3600
      Width           =   6360
      ExtentX         =   11218
      ExtentY         =   8546
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Web Browser Controller"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   3240
      Width           =   2850
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
lblwait.Caption = " Please wait a few moments..."
web.Document.All.Item("txtEmailAddress").Value = txtuser.Text
web.Document.All.Item("txtPassword").Value = txtpass.Text
web.Document.All.Item("chkRememberPassword").Value = Check1.Value
web.Document.All.Item("cmOk").Click
End Sub

Private Sub Form_Load()
web.Navigate "http://www.pscode.com/vb/authentication/DeleteCookies.asp?txtReturnURL=%2Fvb%2Fdefault%2Easp%3FlngWId%3D1%26txtForceRefresh%3D3720061819638320"
Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
web.Navigate "http://www.pscode.com/vb/authentication/Login.asp?txtReturnURL=%2Fvb%2Fdefault%2Easp%3FlngWId%3D1"
Timer1.Enabled = False
End Sub

Private Sub web_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
txtuser.Enabled = True
txtpass.Enabled = True
Check1.Enabled = True
Command1.Enabled = True
Exit Sub
End Sub
