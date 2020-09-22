VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H0000CCFF&
   BorderStyle     =   0  'None
   Caption         =   "Visual Color Converter"
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   3840
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2760
      ScaleHeight     =   585
      ScaleWidth      =   945
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   3480
      Width           =   1095
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2760
      ScaleHeight     =   585
      ScaleWidth      =   945
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   1095
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2760
      ScaleHeight     =   585
      ScaleWidth      =   945
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   600
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      MaxLength       =   3
      TabIndex        =   2
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   480
      MaxLength       =   3
      TabIndex        =   1
      Top             =   600
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2760
      ScaleHeight     =   585
      ScaleWidth      =   945
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      MaxLength       =   3
      TabIndex        =   0
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Command4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Convert"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ">>"
      Height          =   195
      Left            =   1250
      TabIndex        =   27
      Top             =   3480
      Width           =   270
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VBColor to RGB:"
      Height          =   195
      Left            =   120
      TabIndex        =   26
      Top             =   3240
      Width           =   1440
   End
   Begin VB.Label Command3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Convert"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ">>"
      Height          =   195
      Left            =   1250
      TabIndex        =   23
      Top             =   2520
      Width           =   270
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VBColor to HEX:"
      Height          =   195
      Left            =   120
      TabIndex        =   22
      Top             =   2280
      Width           =   1410
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "X"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3650
      TabIndex        =   21
      Top             =   15
      Width           =   120
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   0
      Y1              =   240
      Y2              =   4200
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   3840
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line1 
      X1              =   3825
      X2              =   3825
      Y1              =   255
      Y2              =   4200
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ">>"
      Height          =   195
      Left            =   1360
      TabIndex        =   19
      Top             =   1560
      Width           =   270
   End
   Begin VB.Label Command2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Convert"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hex="
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   1560
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HEX to VBColor:"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   1320
      Width           =   1410
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ">>"
      Height          =   195
      Left            =   1250
      TabIndex        =   15
      Top             =   600
      Width           =   270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RBG to VBColor:"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   360
      Width           =   1440
   End
   Begin VB.Label Title1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Visual Color Converter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   30
      TabIndex        =   13
      Top             =   15
      Width           =   1560
   End
   Begin VB.Label Header1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   3975
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Convert"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   2535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error GoTo hell
    Picture1.BackColor = RGB(Text1.Text, Text2.Text, Text3.Text)
    Text4.Text = Picture1.BackColor
    Exit Sub
hell:
    MsgBox Err.Description, vbCritical, Err.Number
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command1.BackColor = vbBlack
    Command1.ForeColor = vbWhite
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command2.BackColor = vbBlack
    Command2.ForeColor = vbWhite
End Sub

Private Sub Command3_Click()
On Error Resume Next
    Picture3.BackColor = Text7.Text
    Text8.Text = "&H" & Text7.Text
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command3.BackColor = vbBlack
    Command3.ForeColor = vbWhite
End Sub

Private Sub Command4_Click()
On Error Resume Next
Picture4.BackColor = Text9.Text
Dim lVBCol As Long
    lVBCol = CLng(Text9.Text)
    Text10.Text = CStr(lVBCol And &HFF&) & ", " & CStr((lVBCol \ &H100&) And &HFF&) & ", " & CStr(lVBCol \ &H10000)
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command4.BackColor = vbBlack
    Command4.ForeColor = vbWhite
End Sub

Private Sub Form_Load()
    'Me.BackColor = RGB(255, 204, 0)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command1.BackColor = vbWhite
    Command1.ForeColor = vbBlack
    Command2.BackColor = vbWhite
    Command2.ForeColor = vbBlack
    Command3.BackColor = vbWhite
    Command3.ForeColor = vbBlack
    Command4.BackColor = vbWhite
    Command4.ForeColor = vbBlack
    Label4.ForeColor = vbWhite
End Sub

Private Sub Form_Resize()
    Header1.Width = Me.Width
End Sub

Private Sub Command2_Click()
    Picture2.BackColor = HexConvert(Text5.Text)
    Text6.Text = Picture2.BackColor
End Sub

Function HexConvert(hexnum As String) As Long
On Error Resume Next
   HexConvert = "&H" & hexnum
End Function

Private Sub Header1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label4.ForeColor = vbWhite
End Sub

Private Sub Label4_Click()
    Unload Me
    End
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label4.ForeColor = &HCCFF&
End Sub

Private Sub Text1_GotFocus()
    Text1.BackColor = vbBlack
    Text1.ForeColor = vbWhite
    Text1.SelStart = 0
    Text1.SelLength = 3
End Sub

Private Sub Text1_LostFocus()
    Text1.BackColor = vbWhite
    Text1.ForeColor = vbBlack
End Sub

Private Sub Text2_GotFocus()
    Text2.BackColor = vbBlack
    Text2.ForeColor = vbWhite
    Text2.SelStart = 0
    Text2.SelLength = 3
End Sub

Private Sub Text2_LostFocus()
    Text2.BackColor = vbWhite
    Text2.ForeColor = vbBlack
End Sub

Private Sub Text3_GotFocus()
    Text3.BackColor = vbBlack
    Text3.ForeColor = vbWhite
    Text3.SelStart = 0
    Text3.SelLength = 3
End Sub

Private Sub Text3_LostFocus()
    Text3.BackColor = vbWhite
    Text3.ForeColor = vbBlack
End Sub

Private Sub Text4_GotFocus()
    Text4.BackColor = vbBlack
    Text4.ForeColor = vbWhite
    Text4.SelStart = 0
    Text4.SelLength = 10
End Sub

Private Sub Text4_LostFocus()
    Text4.BackColor = vbWhite
    Text4.ForeColor = vbBlack
End Sub

Private Sub Text5_GotFocus()
    Text5.BackColor = vbBlack
    Text5.ForeColor = vbWhite
    Text5.SelStart = 0
    Text5.SelLength = 6
End Sub

Private Sub Text5_LostFocus()
    Text5.BackColor = vbWhite
    Text5.ForeColor = vbBlack
End Sub

Private Sub Text6_GotFocus()
    Text6.BackColor = vbBlack
    Text6.ForeColor = vbWhite
    Text6.SelStart = 0
    Text6.SelLength = 10
End Sub

Private Sub Text6_LostFocus()
    Text6.BackColor = vbWhite
    Text6.ForeColor = vbBlack
End Sub

Private Sub Text7_GotFocus()
    Text7.BackColor = vbBlack
    Text7.ForeColor = vbWhite
    Text7.SelStart = 0
    Text7.SelLength = 10
End Sub

Private Sub Text7_LostFocus()
    Text7.BackColor = vbWhite
    Text7.ForeColor = vbBlack
End Sub

Private Sub Text8_GotFocus()
    Text8.BackColor = vbBlack
    Text8.ForeColor = vbWhite
    Text8.SelStart = 0
    Text8.SelLength = 10
End Sub

Private Sub Text8_LostFocus()
    Text8.BackColor = vbWhite
    Text8.ForeColor = vbBlack
End Sub

Private Sub Text9_GotFocus()
    Text9.BackColor = vbBlack
    Text9.ForeColor = vbWhite
    Text9.SelStart = 0
    Text9.SelLength = 10
End Sub

Private Sub Text9_LostFocus()
    Text9.BackColor = vbWhite
    Text9.ForeColor = vbBlack
End Sub

Private Sub Text10_GotFocus()
    Text10.BackColor = vbBlack
    Text10.ForeColor = vbWhite
    Text10.SelStart = 0
    Text10.SelLength = 11
End Sub

Private Sub Text10_LostFocus()
    Text10.BackColor = vbWhite
    Text10.ForeColor = vbBlack
End Sub
