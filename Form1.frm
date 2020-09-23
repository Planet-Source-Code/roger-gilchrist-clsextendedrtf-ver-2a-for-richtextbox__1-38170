VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Press ESC to continue"
   ClientHeight    =   810
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   810
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   1200
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   1931
      _Version        =   393217
      BackColor       =   -2147483633
      BorderStyle     =   0
      Enabled         =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Silly As New ClsRTFFontLooks

Private Sub Command1_Click()
Unload Form1
End Sub

Private Sub Form_Load()
RichTextBox1.Text = "Welcome! This demo program is  based on manipulating RTF code "
RichTextBox1.SelLength = Len(RichTextBox1.Text)
Silly.AssignControls RichTextBox1, ExtendedRTFDemo.CommonDialog1
'Silly.RippleEngine BaseLine, 10, False, 0, 1
'Silly.RainBow
Timer1.Enabled = True
End Sub


Private Sub Form_Unload(Cancel As Integer)
Timer1.Enabled = False
Unload Form1
End Sub

Private Sub Timer1_Timer()
Static flicker As Boolean
flicker = Not flicker
Silly.RippleEngine BaseLine, 8, flicker, 3, 3
Silly.SpectrumSector s3GreenCyan, flicker, flicker, False
'due to time considerations of the code manipulation this text length and
'Ripple settings are near the maximum you can use for this sort of animation
End Sub
