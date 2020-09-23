VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "MorphTextBox Demo - Matthew R. Usner"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5535
   ScaleWidth      =   8370
   StartUpPosition =   2  'CenterScreen
   Begin prjTextBox.ucGradContainer ucGradContainer1 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   9340
      IconSize        =   0
      HeaderColor2    =   64
      HeaderColor1    =   8421631
      BackColor2      =   16761024
      BackColor1      =   4194368
      BorderColor     =   255
      CaptionColor    =   65535
      Caption         =   "MorphTextBox Demo - Matthew R. Usner"
      HeaderHeight    =   30
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderVisible   =   0   'False
      CurveTopLeft    =   60
      CurveBottomRight=   60
      Begin prjTextBox.MorphTextBox MorphTextBox4 
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   4680
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "Ilia HD (MoMoYa) is a lowlife code thief!"
         PasswordChar    =   "*"
         SelectOnFocus   =   -1  'True
      End
      Begin prjTextBox.MorphTextBox MorphTextBox3 
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   1920
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "Non TT: MS Serif, 10, Bold"
         FocusColor1     =   4210688
         Picture         =   "frmTest.frx":0000
         SelColor1       =   12582912
         SelColor2       =   16761024
         SelGradHeight   =   0
         CaretHeight     =   0
         SelTextColor    =   8454016
      End
      Begin prjTextBox.MorphTextBox MorphTextBox2 
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "Non TT: MS Sans Serif, 14, Bold"
         DefaultColor1   =   4210752
         DefaultColor2   =   16772895
         FocusColor1     =   8421376
         MaxLength       =   0
         SelGradHeight   =   0
         SelTextColor    =   8388736
      End
      Begin prjTextBox.MorphTextBox MorphTextBox1 
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "TrueType: Tahoma, 10, Bold"
         DefaultTextColor=   65535
         DefaultBorderColor=   65280
         FocusTextColor  =   65535
         FocusBorderColor=   65280
         MaxLength       =   0
         Picture         =   "frmTest.frx":2FA2
         SelColor1       =   64
         SelColor2       =   8421631
         Curvature       =   30
         SelGradHeight   =   0
         SelTextColor    =   14737632
      End
      Begin prjTextBox.MorphTextBox MorphTextBox7 
         Height          =   495
         Left            =   3840
         TabIndex        =   6
         Top             =   360
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "TrueType: Tahoma, 12, Regular"
         DefaultTextColor=   65535
         DefaultBorderColor=   65280
         FocusTextColor  =   65535
         FocusBorderColor=   65280
         MaxLength       =   0
         Picture         =   "frmTest.frx":12640
         SelColor1       =   64
         SelColor2       =   8421631
         Curvature       =   30
         SelTextColor    =   14737632
      End
      Begin prjTextBox.MorphTextBox MorphTextBox8 
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   1440
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "Non TrueType: MS Sans Serif, 14, Regular"
         DefaultColor1   =   4210752
         DefaultColor2   =   16772895
         FocusColor1     =   8421376
         MaxLength       =   0
         SelGradHeight   =   0
         SelTextColor    =   8388736
      End
      Begin prjTextBox.MorphTextBox MorphTextBox9 
         Height          =   375
         Left            =   4080
         TabIndex        =   8
         Top             =   1920
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "Non TT: MS Serif, 14, Regular"
         FocusColor1     =   4210688
         Picture         =   "frmTest.frx":21CDE
         SelColor1       =   12582912
         SelColor2       =   16761024
         CaretHeight     =   0
         SelTextColor    =   8454016
      End
      Begin prjTextBox.MorphTextBox MorphTextBox10 
         Height          =   495
         Left            =   360
         TabIndex        =   9
         Top             =   2400
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Handwriting"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "TT: Lucida Handwriting, 14, Italic"
         DefaultTextColor=   65535
         DefaultColor1   =   8388608
         DefaultColor2   =   16777088
         DefaultMiddleOut=   0   'False
         FocusColor1     =   128
         FocusColor2     =   8421631
         MaxLength       =   0
         SelGradHeight   =   0
         SelTextColor    =   8388736
      End
      Begin prjTextBox.MorphTextBox MorphTextBox11 
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   3000
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Handwriting"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "TT: Lucida Handwriting, 12, Bold/Italic"
         DefaultTextColor=   65535
         DefaultColor1   =   8388608
         DefaultColor2   =   16777088
         DefaultMiddleOut=   0   'False
         FocusColor1     =   128
         FocusColor2     =   8421631
         MaxLength       =   0
         SelGradHeight   =   0
         SelTextColor    =   8388736
      End
      Begin prjTextBox.MorphTextBox MorphTextBox12 
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   3600
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "Non TrueType: FixedSys, 9, Bold"
         DefaultColor1   =   0
         DefaultColor2   =   14737632
         FocusBorderColor=   65280
         MaxLength       =   0
         SelColor1       =   64
         SelColor2       =   8421631
         Curvature       =   20
         SelGradHeight   =   0
         SelTextColor    =   14737632
      End
      Begin prjTextBox.MorphTextBox MorphTextBox13 
         Height          =   495
         Left            =   360
         TabIndex        =   12
         Top             =   4090
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Matisse ITC"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "TrueType: Matisse ITC, 16, Italic"
         DefaultTextColor=   16776960
         DefaultColor1   =   64
         DefaultColor2   =   8421631
         FocusBorderColor=   65280
         MaxLength       =   0
         SelColor1       =   64
         SelColor2       =   8421631
         Curvature       =   20
         SelGradHeight   =   0
         SelTextColor    =   14737632
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "<< Selected text for top left MorphTextBox:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   3840
         TabIndex        =   14
         Top             =   120
         Width           =   3975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   13
         Top             =   120
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PasswordChar/SelectOnFocus Properties Set."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   3840
         TabIndex        =   5
         Top             =   4800
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'MorphTextBox1.Text = "Hello"
'  just a simple example how to change font through code.
'   Dim X As New StdFont
'   If MorphTextBox1.Font.Name = "Tahoma" Then
'      X.Name = "MS Sans Serif"
'      X.Bold = True
'     X.Size = 12
'      Set MorphTextBox1.Font = X
'   Else
'      X.Name = "Tahoma"
'      X.Bold = True
'      X.Size = 9
'      Set MorphTextBox1.Font = X
'   End If

End Sub

Private Sub MorphTextBox1_DblClick()
Label2.Caption = MorphTextBox1.SelText
End Sub


Private Sub MorphTextBox1_KeyDown(KeyCode As Integer, Shift As Integer)
Label2.Caption = MorphTextBox1.SelText

End Sub

Private Sub MorphTextBox1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Debug.Print Now, "M"
Label2.Caption = MorphTextBox1.SelText
Me.Refresh ' need this just for the demo; it'll cause some flicker.
End Sub
