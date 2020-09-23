VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7875
   Icon            =   "MainScreen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "MainScreen.frx":0442
   ScaleHeight     =   6660
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblTips 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reset your PC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   180
      Index           =   1
      Left            =   5040
      MouseIcon       =   "MainScreen.frx":AB80C
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   2160
      Width           =   945
   End
   Begin VB.Label lblTips 
      BackStyle       =   0  'Transparent
      Caption         =   "Send me a mail to give me your opinion"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   420
      Index           =   5
      Left            =   5040
      MouseIcon       =   "MainScreen.frx":AB95E
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label lblTips 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "About this program"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   180
      Index           =   4
      Left            =   2160
      MouseIcon       =   "MainScreen.frx":ABAB0
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   4380
      Width           =   1185
   End
   Begin VB.Label lblTips 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Close this programe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   180
      Index           =   3
      Left            =   5040
      MouseIcon       =   "MainScreen.frx":ABC02
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   3000
      Width           =   1260
   End
   Begin VB.Label lblTips 
      BackStyle       =   0  'Transparent
      Caption         =   "Show the sesion screen to  be changed"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   420
      Index           =   2
      Left            =   2160
      MouseIcon       =   "MainScreen.frx":ABD54
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   2925
      Width           =   2415
   End
   Begin VB.Label lblTips 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Only close your PC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   180
      Index           =   0
      Left            =   2160
      MouseIcon       =   "MainScreen.frx":ABEA6
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   2160
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ExitWindows Operations"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   495
      Left            =   1440
      TabIndex        =   8
      Top             =   600
      Width           =   4845
   End
   Begin VB.Label lblCommands 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit Windows"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   225
      Index           =   0
      Left            =   2160
      MouseIcon       =   "MainScreen.frx":ABFF8
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblCommands 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reset PC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   225
      Index           =   1
      Left            =   5040
      MouseIcon       =   "MainScreen.frx":AC14A
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1800
      Width           =   795
   End
   Begin VB.Label lblCommands 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change sesion"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   225
      Index           =   2
      Left            =   2160
      MouseIcon       =   "MainScreen.frx":AC29C
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2640
      Width           =   1275
   End
   Begin VB.Label lblCommands 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel && unload"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   225
      Index           =   3
      Left            =   5040
      MouseIcon       =   "MainScreen.frx":AC3EE
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2640
      Width           =   1365
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   1440
      MouseIcon       =   "MainScreen.frx":AC540
      MousePointer    =   99  'Custom
      Picture         =   "MainScreen.frx":AC692
      Top             =   1680
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   4320
      MouseIcon       =   "MainScreen.frx":ACAD4
      MousePointer    =   99  'Custom
      Picture         =   "MainScreen.frx":ACC26
      Top             =   1680
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   4320
      MouseIcon       =   "MainScreen.frx":AD068
      MousePointer    =   99  'Custom
      Picture         =   "MainScreen.frx":AD1BA
      Top             =   2520
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   1440
      MouseIcon       =   "MainScreen.frx":AD5FC
      MousePointer    =   99  'Custom
      Picture         =   "MainScreen.frx":AD74E
      Top             =   2520
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   1440
      MouseIcon       =   "MainScreen.frx":ADB90
      MousePointer    =   99  'Custom
      Picture         =   "MainScreen.frx":ADCE2
      Top             =   3960
      Width           =   480
   End
   Begin VB.Label lblCommands 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "About..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   225
      Index           =   4
      Left            =   2160
      MouseIcon       =   "MainScreen.frx":AE124
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   4080
      Width           =   600
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   4320
      MouseIcon       =   "MainScreen.frx":AE276
      MousePointer    =   99  'Custom
      Picture         =   "MainScreen.frx":AE3C8
      Top             =   3960
      Width           =   480
   End
   Begin VB.Label lblCommands 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Send me a mail"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   225
      Index           =   5
      Left            =   5040
      MouseIcon       =   "MainScreen.frx":AE80A
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   4080
      Width           =   1305
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      Caption         =   "Operations"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   1320
      TabIndex        =   1
      Top             =   1320
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      Caption         =   "Others"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   1320
      TabIndex        =   0
      Top             =   3600
      Width           =   465
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C000&
      Height          =   1935
      Left            =   1080
      Top             =   1440
      Width           =   5655
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C000&
      Height          =   1095
      Left            =   1080
      Top             =   3720
      Width           =   5655
   End
   Begin VB.Image imgRes 
      Height          =   600
      Index           =   0
      Left            =   1320
      Picture         =   "MainScreen.frx":AE95C
      Top             =   1620
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgRes 
      Height          =   600
      Index           =   1
      Left            =   4260
      Picture         =   "MainScreen.frx":B0172
      Top             =   1600
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgRes 
      Height          =   600
      Index           =   2
      Left            =   1345
      Picture         =   "MainScreen.frx":B1694
      Top             =   2430
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgRes 
      Height          =   600
      Index           =   3
      Left            =   4250
      Picture         =   "MainScreen.frx":B2996
      Top             =   2445
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgRes 
      Height          =   600
      Index           =   4
      Left            =   1345
      Picture         =   "MainScreen.frx":B3C98
      Top             =   3920
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgRes 
      Height          =   495
      Index           =   5
      Left            =   4230
      Picture         =   "MainScreen.frx":B51BA
      Top             =   3940
      Visible         =   0   'False
      Width           =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Trimmer As New cSelShapeForm

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



Private Sub Form_Load()
    Trimmer.TrimForm Me: Load Form2: Trimmer.TrimForm Form2
    'MsgBox "If you have Microsoft Windows XP, this program will not work correctly. Only will work Change Session Option." & Chr(13) & "a big shit no? :(", vbExclamation, App.Title
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Trimmer.GrabForm Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCommands(0).ForeColor = &H8000000B: lblCommands(1).ForeColor = &H8000000B
    lblCommands(2).ForeColor = &H8000000B: lblCommands(3).ForeColor = &H8000000B
    lblCommands(4).ForeColor = &H8000000B: lblCommands(5).ForeColor = &H8000000B
    With Me
        .lblTips(0).ForeColor = &H8000000B
        .lblTips(1).ForeColor = &H8000000B
        .lblTips(2).ForeColor = &H8000000B
        .lblTips(3).ForeColor = &H8000000B
        .lblTips(4).ForeColor = &H8000000B
        .lblTips(5).ForeColor = &H8000000B
        .imgRes(0).Visible = False
        .imgRes(1).Visible = False
        .imgRes(2).Visible = False
        .imgRes(3).Visible = False
        .imgRes(4).Visible = False
        .imgRes(5).Visible = False
    End With
End Sub

Private Sub Image1_Click()
    ClosePC
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgRes(0).Visible = True
    lblCommands(0).ForeColor = vbWhite
    lblTips(0).ForeColor = vbWhite
End Sub

Private Sub Image2_Click()
    ResetPC
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgRes(1).Visible = True
    lblCommands(1).ForeColor = vbWhite
    lblTips(1).ForeColor = vbWhite
End Sub

Private Sub Image3_Click()
    ChangeSesion
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgRes(2).Visible = True
    lblCommands(2).ForeColor = vbWhite
    lblTips(2).ForeColor = vbWhite
End Sub

Private Sub Image4_Click()
    End
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgRes(3).Visible = True
    lblCommands(3).ForeColor = vbWhite
    lblTips(3).ForeColor = vbWhite
End Sub

Private Sub Image5_Click()
    Trimmer.TrimForm Form2
    Form2.Show 1
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgRes(4).Visible = True
    lblCommands(4).ForeColor = vbWhite
    lblTips(4).ForeColor = vbWhite
End Sub

Private Sub Image6_Click()
    ShellExecute Me.hwnd, "open", "mailto:lambdero18@hotmail.com", vbNullString, "C:\", 5
End Sub

Private Sub Image6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgRes(5).Visible = True
    lblCommands(5).ForeColor = vbWhite
    lblTips(5).ForeColor = vbWhite
End Sub

Private Sub lblCommands_Click(Index As Integer)
    Select Case Index
        Case 0
            ClosePC
        Case 1
            ResetPC
        Case 2
            ChangeSesion
        Case 3
            End
        Case 4
            Trimmer.TrimForm Form2
            Form2.Show 1
        Case 5
            ShellExecute Me.hwnd, "open", "mailto:lambdero18@hotmail.com", vbNullString, "C:\", 5

    End Select
End Sub

Private Sub lblCommands_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i%
    For i% = 0 To 5
        If i% = Index Then
            
            lblCommands(i).ForeColor = &H8000000E
            lblTips(i).ForeColor = &H8000000E
            imgRes(i).Visible = True
        Else
            lblCommands(i).ForeColor = &H8000000B
            lblTips(i).ForeColor = &H8000000B
            imgRes(i).Visible = False
        End If
    Next
End Sub
