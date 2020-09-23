VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3960
   LinkTopic       =   "Form2"
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   3345
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblCommands 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clik any location to unload this screen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   210
      Index           =   2
      Left            =   720
      MouseIcon       =   "frmAbout.frx":2B8C2
      TabIndex        =   4
      Top             =   1920
      Width           =   2715
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(c) 2002"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   225
      Left            =   1560
      TabIndex        =   3
      Top             =   720
      Width           =   705
   End
   Begin VB.Label lblCommands 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Close, Reset or change sesion!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   210
      Index           =   1
      Left            =   840
      MouseIcon       =   "frmAbout.frx":2BA14
      TabIndex        =   2
      Top             =   1560
      Width           =   2250
   End
   Begin VB.Label lblCommands 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "By Pau Jansa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   210
      Index           =   0
      Left            =   1440
      MouseIcon       =   "frmAbout.frx":2BB66
      TabIndex        =   1
      Top             =   1200
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ExitWindows Operations"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   225
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   2190
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Trimmer As New cSelShapeForm

Private Sub Form_Click()
    
    Unload Me
End Sub

