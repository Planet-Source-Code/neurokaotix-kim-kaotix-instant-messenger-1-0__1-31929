VERSION 5.00
Begin VB.Form SignOn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sign On Kaotix Network"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3000
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SignOn.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   3000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdSignOn 
      Caption         =   "Sign On"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   1335
   End
   Begin VB.ComboBox cmbUsername 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "SignOn.frx":030A
      Left            =   120
      List            =   "SignOn.frx":030C
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      ToolTipText     =   "Click Here To Enter Your Screen Name"
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New user?"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   885
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Get a ScreenName!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1200
      MouseIcon       =   "SignOn.frx":030E
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   2280
      Width           =   1680
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   3480
      X2              =   -240
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   3480
      X2              =   -240
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line Line1 
      X1              =   3480
      X2              =   -240
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Forgot Password?"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1395
      MouseIcon       =   "SignOn.frx":0618
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   3240
      Width           =   1500
   End
   Begin VB.Image Green 
      Height          =   150
      Left            =   600
      Picture         =   "SignOn.frx":0922
      Top             =   4560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image Red 
      Height          =   150
      Left            =   360
      Picture         =   "SignOn.frx":099B
      Top             =   4560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   315
      TabIndex        =   6
      Top             =   2640
      Width           =   990
   End
   Begin VB.Image Image3 
      Height          =   150
      Left            =   120
      Top             =   2685
      Width           =   150
   End
   Begin VB.Image Image2 
      Height          =   150
      Left            =   120
      Top             =   1710
      Width           =   150
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KIM ScreenName:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   320
      TabIndex        =   5
      Top             =   1665
      Width           =   1710
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   0
      Picture         =   "SignOn.frx":0A14
      Top             =   30
      Width           =   3000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KIM ScreenName:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   330
      TabIndex        =   4
      Top             =   1680
      Width           =   1710
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   330
      TabIndex        =   7
      Top             =   2655
      Width           =   990
   End
   Begin VB.Line Line4 
      X1              =   3480
      X2              =   -240
      Y1              =   1545
      Y2              =   1545
   End
End
Attribute VB_Name = "SignOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbUsername_Change()
If cmbUsername.Text <> "" Then
Image2.Picture = Green.Picture
Else
Image2.Picture = Red.Picture
End If
End Sub

Private Sub cmdSignOn_Click()
    If Client.Winsock1.State <> sckClosed Then Client.Winsock1.Close
    Client.Winsock1.RemotePort = 1008
    'Client.Winsock1.RemoteHost = "216.77.225.246" 'put your IP here and comment out the one below
    Client.Winsock1.RemoteHost = "127.0.0.1"       'to allow people to connect to your IP
    Client.Winsock1.Connect
    
Do Until Client.Winsock1.State = sckConnected
    DoEvents: DoEvents: DoEvents: DoEvents
    If Client.Winsock1.State = sckError Then
        MsgBox "Could not connect to server! The server may be down or you may not be connected to the Internet. Check your connection and try again. If you still cannot connect wait until a later time when the server will be up."
        Exit Sub
    End If
Loop
    Client.Winsock1.SendData (".login" & " " & LCase(cmbUsername.Text) & " " & LCase(txtPassword.Text))
End Sub

Private Sub Form_Load()
Image2.Picture = Red.Picture
Image3.Picture = Red.Picture
End Sub

Private Sub Command1_Click()
FinalClose = True
Unload Me
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Not FinalClose Then
Me.WindowState = 1
Cancel = 1
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub


Private Sub txtPassword_Change()
If txtPassword.Text <> "" Then
Image3.Picture = Green.Picture
Else
Image3.Picture = Red.Picture
End If
End Sub
