VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6580F760-7819-11CF-B86C-444553540000}#1.0#0"; "EZFTP.OCX"
Begin VB.Form frmWebserver 
   Caption         =   "Home Web relinker"
   ClientHeight    =   660
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   3690
   Icon            =   "WebServer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   660
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin EZFTPLib.EZFTP FTP 
      Left            =   2160
      Top             =   0
      _Version        =   65536
      _ExtentX        =   800
      _ExtentY        =   800
      _StockProps     =   0
      LocalFile       =   ""
      RemoteFile      =   ""
      RemoteAddres    =   ""
      UserName        =   ""
      Password        =   ""
      Binary          =   0   'False
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   2640
      Picture         =   "WebServer.frx":030A
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   2
      Top             =   0
      Width           =   615
   End
   Begin MSWinsockLib.Winsock WS1 
      Left            =   1560
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label IP 
      Caption         =   "IP address: "
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Menu mnuPWS 
      Caption         =   "mnuPWS"
      Begin VB.Menu mnuIPaddr 
         Caption         =   "Your IP"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu Settings 
         Caption         =   "S&ettings"
      End
      Begin VB.Menu StartRelink 
         Caption         =   "&Start"
      End
      Begin VB.Menu StopRelink 
         Caption         =   "S&top"
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmWebserver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------
'Program made by Peter Verburgh @ 2001
'First read the Readme_Web included in the zip file..
'This application let it make a dynamic Webserver..
'If you have a free account by provider ex. A , and you got some free webspace...
'but the provider wouldn't let you use the full advantage of free ASP.. PHP...
'and you have a PWS , IIS ,....
'You can now make your own webserver ..
'How it works..
'-----------------
'you start this program .. then you can see a ico .. right down..,
'this application detects the ip adres..& then it  change some data in the index.html..
'local on your computer on drive.. x  (changable it by settings..)
'& then it would be uploaded to your free web- provider..
'Next step is, i have written some very- easy code in index.html (javascript) hat reads the
'ip adres & the script detects if your webserver is online.. if its true ,
'the user navigates to your free account & the person will be directed to your webserver.
'(standard port 80)... but you can change it in the javascript..'
'AND now you GOT THE FULL POWER... to do anythings...

'REMARK : its possible , that it wouldn't work if you have multiple IP , have installed on
'your pc... because WINSOCK.OCX use the default..
'you can handle it by binding...

'If the OCX file for ftp. isn't included in the file.. you can find it on my site (under construction)
'http://users.skynet.be/verburgh.peter

'Please Vote for me !!
'Tnx !!!!
'-------------------------------------------------------------------------------------------------------



Dim blnStart As Boolean
Dim FileSend As Boolean
Dim fs, f, Text
Dim fso1, a

Const ForReading = 1, ForWriting = 2, ForAppending = 3



Private Sub Form_Load()
Dim fso, msg
'-------------------------------------------------------- check for local IP... --------------------
' if you have multiple IP installed on your pc , that way doesn't work.. because ,
'winsock takes the default value....
'In the future i will look to change it by API..

Label1.Caption = WS1.LocalIP
mnuIPaddr.Caption = WS1.LocalIP
'---------------------------------------------------------
'Reading Inf file
modSettings.ThisDir = CurDir

modSettings.ThisDir = modSettings.ThisDir & "\"
'----------------------------------------------------------------
'Check if file exist... otherwise make the file...

   Set fso = CreateObject("Scripting.FileSystemObject")
   If (fso.FileExists((modSettings.ThisDir & "DynaSetting.inf"))) Then
      'File Exist...no problem.. error
   Else
      MsgBox "This is maybe the first time that you use this application , so you have to fill down the settings..", vbInformation
        Dialog.Show
       CreateIcon
       StopRelink.Visible = False
       Me.Hide
       GoTo End1:
   End If
'---------------------------------------------------------------------------------------
modSettings.strUrl = modINI.sGetINI(modSettings.ThisDir & "DynaSetting.inf", "Settings", "FTP", "?")
modSettings.strLogin = modINI.sGetINI(modSettings.ThisDir & "DynaSetting.inf", "Settings", "Username", "?")
modSettings.strPassword = modINI.sGetINI(modSettings.ThisDir & "DynaSetting.inf", "Settings", "Password", "?")
modSettings.strSource = modINI.sGetINI(modSettings.ThisDir & "DynaSetting.inf", "Settings", "LocalFile", "?")
modSettings.strRemote = modINI.sGetINI(modSettings.ThisDir & "DynaSetting.inf", "Settings", "RemoteFile", "?")
'---------------------------------------------------------

CreateIcon
StopRelink.Visible = False
Me.Hide
End1:
End Sub

Sub SendFiletoServer()
'This error handling => tnx to AutoBot
On Error GoTo SendFiletoServer_Err

FTP.UserName = modSettings.strLogin
FTP.Password = modSettings.strPassword
FTP.RemoteAddress = modSettings.strUrl
FTP.LocalFile = modSettings.strSource
FTP.RemoteFile = modSettings.strRemote
FTP.Connect
FTP.PutFile
'Wait until data is transferred..

'-------------------------
FTP.Disconnect
FileSend = False
Exit Sub

SendFiletoServer_Err:

    MsgBox "Unable To Connect To FTP Server"
    
    Resume Next

End Sub


Function HTMLData(str As String) As String
pos = InStr(1, Text, "var IPADR =", vbTextCompare)
'----------------- Check if the EXACT data   "var IPADR ="  exist... otherwise the file is not correct !!
If pos = 0 Then GoTo FileCorrupt
pos1 = InStr(pos, Text, "'", vbTextCompare)
pos2 = InStr(pos1 + 1, Text, "'", vbTextCompare)
Txt1 = Mid(Text, 1, pos1)
Txt2 = Mid(Text, pos2, Len(Text))
TotalTXT = Txt1 & str & Txt2
HTMLData = TotalTXT
Exit Function
FileCorrupt:
MsgBox "File is corrupt ! Use the index.html file included in the ZIP file !!!"
HTMLData = "ERROR"
End Function

Public Sub CreateIcon()
    Dim Tic As NOTIFYICONDATA
    Tic.cbSize = Len(Tic)
    Tic.hwnd = Picture1.hwnd
    Tic.uID = 1&
    Tic.uFlags = NIF_DOALL
    Tic.uCallbackMessage = WM_MOUSEMOVE
    Tic.hIcon = Picture1.Picture
    Tic.szTip = "WebRelink " & Chr$(0)
    erg = Shell_NotifyIcon(NIM_ADD, Tic)
End Sub

Public Sub DeleteIcon()
    Dim Tic As NOTIFYICONDATA
    Tic.cbSize = Len(Tic)
    Tic.hwnd = Picture1.hwnd
    Tic.uID = 1&
    erg = Shell_NotifyIcon(NIM_DELETE, Tic)
End Sub

Private Sub Form_Terminate()
DeleteIcon
End Sub

Private Sub Form_Unload(Cancel As Integer)
DeleteIcon
End Sub



Private Sub FTP_TransferProgress(ByVal BytesTransferred As Long, ByVal TotalBytes As Long)
If BytesTransferred = TotalBytes Then FileSend = True
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    X = X / Screen.TwipsPerPixelX


    Select Case X
        Case WM_LBUTTONDOWN
        
        Case WM_RBUTTONDOWN
        
        PopupMenu mnuPWS
        Case WM_MOUSEMOVE
        
        Case WM_LBUTTONDBLCLK
        
    End Select
End Sub

Private Sub Settings_Click()
Dialog.Show
End Sub

Private Sub StartRelink_Click()

'---------------------------- CHECK if THE LOCAL FILE IS CORRECT !!!!-------------------
'----------- check if File Exist..
  Set fso1 = CreateObject("Scripting.FileSystemObject")
   If (fso1.FileExists(modSettings.strSource)) Then
   'Okay .... file Exist...
   Else
   MsgBox "Source file " & modSettings.strSource & "  NOT found !!  Check & change the Settings !! ", vbCritical
   GoTo End2:
   End If
'----------------------------------
Dim IPADR As String
IPADR = WS1.LocalIP

Set f = fso1.OpenTextFile(modSettings.strSource, ForReading)
    Text = f.ReadAll
    f.Close

'Data is read in txt...
Text = HTMLData(IPADR)  'change te data - strings..
'Save data
If Text <> "ERROR" Then
   'Check if the Remotepath & file > len(0)
   If modSettings.strRemote = "" Then
   MsgBox "You have to fill down a path & file !!", vbCritical
   Exit Sub
   End If
Set a = fso1.CreateTextFile(modSettings.strSource, True)
    a.Write Text
    a.Close
StopRelink.Visible = True
StartRelink.Visible = False
SendFiletoServer
Else
MsgBox "Did you fixed the contents of the file ?" & vbCrLf & modSettings.strSource, vbInformation
End If
End2:
End Sub

Private Sub StopRelink_Click()
Dim IPADR As String
IPADR = "NONE"
'Data Read
Text = HTMLData(IPADR)
'Data changed
Set a = fso1.CreateTextFile(modSettings.strSource, True)
    a.Write Text
    a.Close
StopRelink.Visible = False
StartRelink.Visible = True
SendFiletoServer
End Sub
