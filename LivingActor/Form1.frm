VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Living Actor"
   ClientHeight    =   2520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2760
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2520
   ScaleWidth      =   2760
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Animations"
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2535
      Begin VB.CommandButton Command5 
         Caption         =   "Load List"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Play"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Big Size"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Speak"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LA As CantocheLivingActor
Private Actor As CantocheActor

Private Sub Command3_Click()
    Actor.SetScaleNow 2
End Sub

Private Sub Command4_Click()
    Actor.Play Combo1.Text
End Sub

Private Sub Command5_Click()
    Command4.Enabled = True
    Command5.Enabled = False
    Combo1.Enabled = True
    
'Load Animation List
    Dim i As Integer
    
    For i = 0 To Actor.AnimationCount - 1
        Combo1.AddItem Actor.animation(i)
    Next
'</>
End Sub

Private Sub Form_Load()
    Set LA = New CantocheLivingActor
    Set Actor = New CantocheActor
End Sub

Private Sub Command1_Click()

'Load Character
    'Actor Path
    LA.LoadActor App.Path & "\bob.liv"
    
    Set Actor = LA.Actor(0)
    
    'set TTS Engine (SAPI4)
    LA.SetTTSID "{CA141FD0-AC7F-11D1-97A3-006008273002}"
    
    Actor.Show
    
    Command1.Enabled = False
    Command2.Enabled = True
'</>

End Sub

Private Sub Command2_Click()
    Actor.Speak "Hello!, Good Morning. My name is Bob, No, No. Not Microsoft's Bob but Cantoche's Bob"
    Actor.Play "Wave"
    Actor.Speak "Please Vote If you Liked it."
End Sub
