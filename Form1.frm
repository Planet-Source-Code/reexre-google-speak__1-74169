VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.ocx"
Begin VB.Form Form1 
   Caption         =   "Google Speak"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8475
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   8475
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.ProgressBar PB 
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "Form1.frx":17AA
      Top             =   240
      Width           =   7935
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1200
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Speak it !"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   1
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":1836
      Top             =   2040
      Width           =   7935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   5400
      TabIndex        =   3
      Top             =   2160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Thanks to Leandro Ascierto (www.leandroascierto.com.ar)"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   7920
      Width           =   7935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    SpeekMoreThan100 Text1.Text, Right$(Combo1, 2)


End Sub

Private Sub Form_Activate()
GoogleSpeakLess100chars "Welcome to Google Speak.", "en"
End Sub

Private Sub Form_Initialize()
    InitCommonControls

End Sub

Private Sub Form_Load()
    Call FixThemeSupport(Me.Controls) 'add by theBatch 1.0

    '    Combo1.AddItem "Alemán: de"
    '    Combo1.AddItem "Danés: da"
    '    Combo1.AddItem "Español: es"
    '    Combo1.AddItem "Finlandia: fi"
    '    Combo1.AddItem "Francés: fr"
    '    Combo1.AddItem "Inglés: en"
    '    Combo1.AddItem "Italiano: it"
    '    Combo1.AddItem "Neerlandés: nl"
    '    Combo1.AddItem "Polaco: pl"
    '    Combo1.AddItem "Portugués: pt"
    '    Combo1.AddItem "Sueco: sv"

    Combo1.AddItem "German: de"
    Combo1.AddItem "Danish: da"
    Combo1.AddItem "Spanish: es"
    Combo1.AddItem "Finnish: fi"
    Combo1.AddItem "French: fr"
    Combo1.AddItem "English: en"
    Combo1.AddItem "Italian: it"
    Combo1.AddItem "Dutch: nl"
    Combo1.AddItem "Polish: pl"
    Combo1.AddItem "Portuguese: pt"
    Combo1.AddItem "Swedish: sv"

    Combo1.ListIndex = 2



End Sub


Private Sub Command2_Click()
    '    Debug.Print GoogleSpeak("Antes era sexo droga y rock and roll, ahora es paja mate y chamame", "es", True)
End Sub

Private Sub Form_Terminate()
GoogleSpeakLess100chars "good bye.", "en"
End Sub
