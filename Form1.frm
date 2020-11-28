VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "×¢²á»ú"
   ClientHeight    =   3435
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   6255
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "×¢²á"
      Height          =   615
      Left            =   1200
      TabIndex        =   4
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1500
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ÐòÁÐºÅ£º"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   1200
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "ÇëÊäÈë£º"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   5175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Label1.Visible = False
   Label2.Visible = True
   Text1.Visible = False
   For i = 1 To 9
         a = i Xor 2
         a = Right(a, 1)
        code = code & a
   Next
   Text2.Visible = True
   Command1.Visible = False
    Text2.Text = code
End Sub

Private Sub Form_Load()
 Label1.Visible = True
 Label2.Visible = False
 Text1.Visible = True
 Text2.Visible = False
 
End Sub
