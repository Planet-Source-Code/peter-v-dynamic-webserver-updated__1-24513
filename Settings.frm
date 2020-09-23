VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FTP upload Settings.."
   ClientHeight    =   1905
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4350
   Icon            =   "Settings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1320
      TabIndex        =   11
      Text            =   "Text5"
      ToolTipText     =   "give the path & file remote .."
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Text            =   "Text4"
      ToolTipText     =   "give path & file where the javascript is"
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   7
      Text            =   "Text3"
      ToolTipText     =   "give password"
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Text            =   "Text2"
      ToolTipText     =   "Give your username "
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Text            =   "Text1"
      ToolTipText     =   "Fill down the URL of free web-provider"
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Apply"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Remote File"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Local File"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Password"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Username"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "URL ftp server"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = modSettings.strUrl
Text2.Text = modSettings.strLogin
Text3.Text = modSettings.strPassword
Text4.Text = modSettings.strSource
Text5.Text = modSettings.strRemote
End Sub

Private Sub OKButton_Click()
modSettings.strLogin = Text2.Text
modSettings.strPassword = Text3.Text
modSettings.strUrl = Text1.Text
modSettings.strSource = Text4.Text
modSettings.strRemote = Text5.Text
'Saven data to inf file...

modINI.WriteINI modSettings.ThisDir & "DynaSetting.inf", "Settings", "FTP", modSettings.strUrl
modINI.WriteINI modSettings.ThisDir & "DynaSetting.inf", "Settings", "Username", modSettings.strLogin
modINI.WriteINI modSettings.ThisDir & "DynaSetting.inf", "Settings", "Password", modSettings.strPassword
modINI.WriteINI modSettings.ThisDir & "DynaSetting.inf", "Settings", "LocalFile", modSettings.strSource
modINI.WriteINI modSettings.ThisDir & "DynaSetting.inf", "Settings", "RemoteFile", modSettings.strRemote
Unload Me
End Sub

