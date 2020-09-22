VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Text box basics"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   8280
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text8 
      DragMode        =   1  'Automatic
      Height          =   975
      Left            =   2400
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatic
      TabIndex        =   15
      Text            =   "Form1.frx":0000
      Top             =   7200
      Width           =   1935
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00C0FFC0&
      Height          =   1575
      Left            =   4320
      MultiLine       =   -1  'True
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      TabIndex        =   13
      Top             =   4800
      Width           =   2175
   End
   Begin VB.TextBox Text6 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      OLEDragMode     =   1  'Automatic
      TabIndex        =   12
      Text            =   "Form1.frx":0040
      Top             =   4800
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Lock"
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Height          =   1575
      Left            =   3240
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   720
      Width           =   3255
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1200
      MaxLength       =   2
      TabIndex        =   7
      Text            =   "5"
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2160
      MaxLength       =   1
      TabIndex        =   5
      Text            =   "*"
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "password"
      Top             =   2040
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Mask with this character"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "123"
      Top             =   480
      Width           =   2655
   End
   Begin VB.Shape Shape5 
      Height          =   1935
      Left            =   0
      Top             =   4680
      Width           =   6615
   End
   Begin VB.Label Label6 
      Caption         =   ">>Go ahead and drag>>"
      Height          =   255
      Left            =   2400
      TabIndex        =   14
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Shape Shape4 
      Height          =   2775
      Left            =   3120
      Top             =   0
      Width           =   3495
   End
   Begin VB.Label Label5 
      Caption         =   "This text box is multi-line enabled and has scrollbars. You can also lock and unlock the box."
      Height          =   735
      Left            =   3240
      TabIndex        =   10
      Top             =   0
      Width           =   3255
   End
   Begin VB.Shape Shape3 
      Height          =   975
      Left            =   0
      Top             =   3120
      Width           =   6615
   End
   Begin VB.Shape Shape2 
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      Height          =   1575
      Left            =   0
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "Some people use this trick in options and settings panels."
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   4695
   End
   Begin VB.Label Label3 
      Caption         =   "This text box is          times cooler because it's properties are set to blend in with the label. "
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   6615
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "This text box can be used as a password field. Check the box to use a character mask of your choice."
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "This text box only allows numbers."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'These declarations are for the drag and drop functions.
Dim intXoffset As Integer
Dim intYoffset As Integer

Public Sub NumbersOnly(KeyAscii As Integer)
'Anything with an ascii value less than 0 or greater than
'9 will not be entered into the box.

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If

End Sub

Private Sub Check1_Click()
'Mask the password.

If Check1.Value = 1 Then
Text2.PasswordChar = Text3.Text
Else
Text2.PasswordChar = ""
End If


End Sub




Private Sub Command1_Click()
'Lets you do 2 things with 1 button and saves form space. :)

If Command1.Caption = "Lock" Then
Text5.Locked = True
Command1.Caption = "Unlock"
Else
Text5.Locked = False
Command1.Caption = "Lock"
End If

End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
'Moves the textbox.

Source.Move X - intXoffset, Y - intYoffset

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'Be sure to put this in the keypress event, not
'keydown or change.

NumbersOnly KeyAscii

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
'Remember this one? :)

NumbersOnly KeyAscii

End Sub

Private Sub Text6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Begin to drag data.

Text6.OLEDrag

End Sub

Private Sub Text6_OLECompleteDrag(Effect As Long)
'Make the text disappear.

Text6.Text = ""

End Sub

Private Sub Text6_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
'Set data to drag.

Data.SetData Text6.Text, vbCFText
AllowedEffects = vbDropEffectMove

End Sub

Private Sub Text7_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Grabs the text.

Text7.Text = Data.GetData(vbCFText)

End Sub

Private Sub Text8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

intoffset = X
intoffset = Y

End Sub
