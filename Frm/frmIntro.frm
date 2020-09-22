VERSION 5.00
Begin VB.Form frmIntro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FlatButton Test"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   Icon            =   "frmIntro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4830
   StartUpPosition =   2  'CenterScreen
   Begin FlatButtonTest.FlatButton cmdExit 
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Exit"
   End
   Begin VB.CheckBox chkFocusRectangle 
      Caption         =   "Use Focus Rectangle"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin FlatButtonTest.FlatButton cmdSetCaption 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   1320
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   503
      Caption         =   "Set"
   End
   Begin VB.TextBox txtCaption 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Text            =   "Caption"
      Top             =   1320
      Width           =   1815
   End
   Begin FlatButtonTest.FlatButton FlatButton1 
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Buton 1"
   End
   Begin VB.Label Label1 
      Caption         =   "Flat Button"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   4575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   120
      X2              =   4680
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   120
      X2              =   4680
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label lblCaption 
      Caption         =   "Settings:"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   2
      X1              =   120
      X2              =   4680
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   1
      X1              =   120
      X2              =   4680
      Y1              =   2400
      Y2              =   2400
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdSetCaption_Click()
  FlatButton1.Caption = txtCaption.Text
  
  If chkFocusRectangle.Value = 1 Then
  FlatButton1.FocusRect = True
  Else
  FlatButton1.FocusRect = False
  End If
End Sub
