VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form ExtendedRTFDemo 
   Caption         =   "ClsExtendedRTF Demo"
   ClientHeight    =   8685
   ClientLeft      =   165
   ClientTop       =   -2685
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   11880
   Begin VB.Frame Frame1 
      Caption         =   "RichWordOver                   InstantTranlator"
      Height          =   855
      Left            =   8760
      TabIndex        =   10
      Top             =   5280
      Width           =   3375
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Label6"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "current word"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.CheckBox Check2 
      Caption         =   "HighLighted Text"
      Height          =   195
      Left            =   0
      TabIndex        =   9
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Hidden Text"
      Height          =   195
      Left            =   2280
      TabIndex        =   8
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reset TextRTF"
      Height          =   435
      Index           =   0
      Left            =   0
      TabIndex        =   7
      ToolTipText     =   "Read the document before playing with me"
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reset SelRTF"
      Height          =   435
      Index           =   1
      Left            =   4560
      TabIndex        =   5
      ToolTipText     =   "Read the document before playing with me"
      Top             =   5880
      Width           =   735
   End
   Begin RichTextLib.RichTextBox text1 
      Height          =   2055
      Left            =   0
      TabIndex        =   4
      Top             =   6480
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3625
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"extendedRTFDemo.frx":0000
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   4560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "extendedRTFDemo.frx":008B
      Top             =   7920
      Width           =   7335
   End
   Begin VB.TextBox Text2 
      Height          =   1095
      Left            =   4560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "extendedRTFDemo.frx":0091
      Top             =   6480
      Width           =   7455
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   8493
      _Version        =   393217
      HideSelection   =   0   'False
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      FileName        =   "C:\Program Files\Microsoft Visual Studio\VB98\QND Programs\ExtendedRTF\Extended RTF code for VB6.rtf"
      TextRTF         =   $"extendedRTFDemo.frx":0097
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "bold"
            Object.ToolTipText     =   "this is VB's native Bold"
            ImageKey        =   "bold"
            Style           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "italic"
            Object.ToolTipText     =   "this is VB's native Italic "
            ImageKey        =   "italic"
            Style           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "underline"
            Object.ToolTipText     =   "ClsExtendedRTF Underlines"
            ImageKey        =   "underline"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   9
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "standard"
                  Text            =   "standard"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "dot"
                  Text            =   "dot"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "dash"
                  Text            =   "dash"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "dashdot"
                  Text            =   "dashdot"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "dashdotdot"
                  Text            =   "dashdotdot"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "thick"
                  Text            =   "thick"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "wave"
                  Text            =   "wave"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "word"
                  Text            =   "word"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "double"
                  Text            =   "double"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "fcolor"
            Object.ToolTipText     =   "VB Selcolor and CDlg ShowColor"
            ImageKey        =   "fcolor"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   6
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "fblack"
                  Text            =   "Black"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "fblue"
                  Text            =   "blue"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "fred"
                  Text            =   "red"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "fyellow"
                  Text            =   "yellow"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "fselect"
                  Text            =   "Select..."
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "hihglight"
            Object.ToolTipText     =   "API highlight"
            ImageKey        =   "highlight"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   11
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "clear"
                  Text            =   "Clear"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "red"
                  Text            =   "Red"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "blue"
                  Text            =   "Blue"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "green"
                  Text            =   "Green"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "yellow"
                  Text            =   "Yellow"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "user"
                  Text            =   "Select Highlight..."
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "userauto"
                  Text            =   "Select Highlight AutoText..."
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "useruser"
                  Text            =   "Select HighLight and Text Colours..."
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "fontincrease"
            Object.ToolTipText     =   "Increase FontSize by one point ClsExtendedRTF "
            ImageKey        =   "fontincrease"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "fontdecrease"
            Object.ToolTipText     =   "Decrease FontSize by one point ClsExtendedRTF "
            ImageKey        =   "fontdecrease"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "restore"
            Description     =   "Restore"
            Object.ToolTipText     =   "Reload last saved version of document ClsExtendedRTF "
            ImageKey        =   "restore"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "activeupdate"
            Object.ToolTipText     =   "Activate RTF text boxes and other time wasters"
            ImageKey        =   "activeoff"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "manifest"
            Object.ToolTipText     =   "ClsManifestation Switch compiled program appearance WindosXP/Classic"
            ImageKey        =   "manifest"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "zoom"
            Style           =   4
            Object.Width           =   1000
         EndProperty
      EndProperty
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   15600
         Style           =   2  'Dropdown List
         TabIndex        =   16
         ToolTipText     =   "Zoom"
         Top             =   120
         Width           =   1095
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7920
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "extendedRTFDemo.frx":1ACA7
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "extendedRTFDemo.frx":1ADB9
            Key             =   "fontdecrease"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "extendedRTFDemo.frx":1B4B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "extendedRTFDemo.frx":1BBAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "extendedRTFDemo.frx":1BCBF
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "extendedRTFDemo.frx":1BDD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "extendedRTFDemo.frx":1BEE3
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "extendedRTFDemo.frx":1BFF5
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "extendedRTFDemo.frx":1C107
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "extendedRTFDemo.frx":1C219
            Key             =   "highlight"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "extendedRTFDemo.frx":1C66B
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "extendedRTFDemo.frx":1CD65
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "extendedRTFDemo.frx":1D45F
            Key             =   "restore"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "extendedRTFDemo.frx":1D9A1
            Key             =   "fcolor"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "extendedRTFDemo.frx":1DAB3
            Key             =   "fontincrease"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "extendedRTFDemo.frx":1E1AD
            Key             =   "activeoff"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "extendedRTFDemo.frx":1E5FF
            Key             =   "activeon"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "extendedRTFDemo.frx":1EA51
            Key             =   "manifest"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7200
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label4 
      Caption         =   "RichTextBox1.Text"
      Height          =   255
      Left            =   6360
      TabIndex        =   15
      Top             =   7680
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   255
      Left            =   3960
      TabIndex        =   14
      Top             =   5400
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "RichTextBox1.TextRTF"
      Height          =   255
      Left            =   840
      TabIndex        =   13
      Top             =   5970
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "RichTextBox1.SelRTF"
      Height          =   255
      Left            =   5640
      TabIndex        =   6
      Top             =   5970
      Width           =   2055
   End
   Begin VB.Menu Ffmnu 
      Caption         =   "&File"
      Begin VB.Menu fmnu 
         Caption         =   "&New"
         Index           =   0
      End
      Begin VB.Menu fmnu 
         Caption         =   "&Open"
         Index           =   1
      End
      Begin VB.Menu fmnu 
         Caption         =   "&Restore"
         Index           =   2
      End
      Begin VB.Menu fmnu 
         Caption         =   "&Save"
         Index           =   3
      End
      Begin VB.Menu fmnu 
         Caption         =   "S&ave As"
         Index           =   4
      End
      Begin VB.Menu fmnu 
         Caption         =   "E&xit"
         Index           =   5
      End
   End
   Begin VB.Menu MnuUL 
      Caption         =   "&UnderLines"
      Begin VB.Menu Ulsmnu 
         Caption         =   "Standard"
         Index           =   0
      End
      Begin VB.Menu Ulsmnu 
         Caption         =   "dotted"
         Index           =   1
      End
      Begin VB.Menu Ulsmnu 
         Caption         =   "dash"
         Index           =   2
      End
      Begin VB.Menu Ulsmnu 
         Caption         =   "DashDot"
         Index           =   3
      End
      Begin VB.Menu Ulsmnu 
         Caption         =   "DashDotDot"
         Index           =   4
      End
      Begin VB.Menu Ulsmnu 
         Caption         =   "Thick"
         Index           =   5
      End
      Begin VB.Menu Ulsmnu 
         Caption         =   "hair"
         Index           =   6
      End
      Begin VB.Menu Ulsmnu 
         Caption         =   "Wave"
         Index           =   7
      End
      Begin VB.Menu Ulsmnu 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu Ulsmnu 
         Caption         =   "Problematic"
         Index           =   9
         Begin VB.Menu probul 
            Caption         =   "word"
            Index           =   0
         End
         Begin VB.Menu probul 
            Caption         =   "double"
            Index           =   1
         End
      End
      Begin VB.Menu Ulsmnu 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu Ulsmnu 
         Caption         =   "Remove AnyUnderline"
         Index           =   11
      End
   End
   Begin VB.Menu hmnu 
      Caption         =   "&HighLighter"
      Begin VB.Menu HCmnu 
         Caption         =   "Red"
         Index           =   0
      End
      Begin VB.Menu HCmnu 
         Caption         =   "Blue"
         Index           =   1
      End
      Begin VB.Menu HCmnu 
         Caption         =   "Green"
         Index           =   2
      End
      Begin VB.Menu HCmnu 
         Caption         =   "Yellow"
         Index           =   3
      End
      Begin VB.Menu HCmnu 
         Caption         =   "Grey"
         Index           =   4
      End
      Begin VB.Menu HCmnu 
         Caption         =   "Mauve"
         Index           =   5
      End
      Begin VB.Menu HCmnu 
         Caption         =   "Pink"
         Index           =   6
      End
      Begin VB.Menu HCmnu 
         Caption         =   "Beige"
         Index           =   7
      End
      Begin VB.Menu HCmnu 
         Caption         =   "Marigold"
         Index           =   8
      End
      Begin VB.Menu HCmnu 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu HCmnu 
         Caption         =   "Select Highlight..."
         Index           =   10
      End
      Begin VB.Menu HCmnu 
         Caption         =   "Select Highlight Auto Text Colour..."
         Index           =   11
      End
      Begin VB.Menu HCmnu 
         Caption         =   "Select Highlight and Text Colour..."
         Index           =   12
      End
      Begin VB.Menu HCmnu 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu HCmnu 
         Caption         =   "Clear Highlight"
         Index           =   14
      End
   End
   Begin VB.Menu vtmuu 
      Caption         =   "&Visible"
      Begin VB.Menu vmnu 
         Caption         =   "Hide (Toggle)"
         Index           =   0
      End
      Begin VB.Menu vmnu 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu vmnu 
         Caption         =   "Show All Hidden"
         Index           =   2
      End
      Begin VB.Menu vmnu 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu vmnu 
         Caption         =   "Tag Selected"
         Index           =   4
      End
      Begin VB.Menu vmnu 
         Caption         =   "Hide Tagged"
         Index           =   5
      End
   End
   Begin VB.Menu MMmnu 
      Caption         =   "&Miscellaneous"
      Begin VB.Menu mmnu 
         Caption         =   "Single Quote"
         Index           =   0
      End
      Begin VB.Menu mmnu 
         Caption         =   "Double Quote"
         Index           =   1
      End
      Begin VB.Menu mmnu 
         Caption         =   "Bold <b> <\b>"
         Index           =   2
      End
      Begin VB.Menu mmnu 
         Caption         =   "LineNumbers"
         Index           =   3
      End
      Begin VB.Menu mmnu 
         Caption         =   "Show Para Marks"
         Index           =   5
      End
      Begin VB.Menu mmnu 
         Caption         =   "silly"
         Index           =   6
      End
   End
   Begin VB.Menu Formatmnu 
      Caption         =   "Font Looks"
      Begin VB.Menu flmnu 
         Caption         =   "Text Look Panel"
         Index           =   0
      End
      Begin VB.Menu flmnu 
         Caption         =   "Text Colour Panel"
         Index           =   1
      End
      Begin VB.Menu flmnu 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu flmnu 
         Caption         =   "Remove All Formatting"
         Index           =   3
      End
      Begin VB.Menu flmnu 
         Caption         =   "Remove UpDownSubSuper"
         Index           =   4
      End
      Begin VB.Menu flmnu 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu flmnu 
         Caption         =   "Remove All Colour Formats"
         Index           =   6
      End
      Begin VB.Menu flmnu 
         Caption         =   "Remove Text Colour Formats"
         Index           =   7
      End
      Begin VB.Menu flmnu 
         Caption         =   "Remove Back Colour Formats"
         Index           =   8
      End
      Begin VB.Menu flmnu 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu flmnu 
         Caption         =   "Remove Excess Spaces"
         Index           =   10
      End
   End
   Begin VB.Menu frmtmnu 
      Caption         =   "Format"
      Begin VB.Menu italic 
         Caption         =   "Italic"
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu jump 
      Caption         =   "&Jump"
      Begin VB.Menu jmnu 
         Caption         =   "0%"
         Index           =   0
      End
      Begin VB.Menu jmnu 
         Caption         =   "10%"
         Index           =   1
      End
      Begin VB.Menu jmnu 
         Caption         =   "20%"
         Index           =   2
      End
      Begin VB.Menu jmnu 
         Caption         =   "30%"
         Index           =   3
      End
      Begin VB.Menu jmnu 
         Caption         =   "40%"
         Index           =   4
      End
      Begin VB.Menu jmnu 
         Caption         =   "50%"
         Index           =   5
      End
      Begin VB.Menu jmnu 
         Caption         =   "60%"
         Index           =   6
      End
      Begin VB.Menu jmnu 
         Caption         =   "70%"
         Index           =   7
      End
      Begin VB.Menu jmnu 
         Caption         =   "80%"
         Index           =   8
      End
      Begin VB.Menu jmnu 
         Caption         =   "90%"
         Index           =   9
      End
      Begin VB.Menu jmnu 
         Caption         =   "100%"
         Index           =   10
      End
   End
   Begin VB.Menu hlpmnu 
      Caption         =   "&Help"
      Begin VB.Menu hlp 
         Caption         =   "ClsManifestation"
         Index           =   0
      End
      Begin VB.Menu hlp 
         Caption         =   "clsAPIHighlight"
         Index           =   1
         Begin VB.Menu hlp2 
            Caption         =   "Programmer"
            Index           =   0
         End
         Begin VB.Menu hlp2 
            Caption         =   "End-User"
            Index           =   1
         End
      End
      Begin VB.Menu hlp 
         Caption         =   "ClsAPIZoom"
         Index           =   2
         Begin VB.Menu hlp3 
            Caption         =   "Programmer"
            Index           =   0
         End
         Begin VB.Menu hlp3 
            Caption         =   "End-User"
            Index           =   1
         End
      End
      Begin VB.Menu hlp 
         Caption         =   "ClsRTFFontLooks"
         Index           =   3
      End
      Begin VB.Menu hlp 
         Caption         =   "CslExtendedRTF"
         Index           =   4
      End
   End
End
Attribute VB_Name = "ExtendedRTFDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public FileName As String

Public ActiveUpdate As Boolean

Private Sub Combo1_Click()

    RTBZoom.ZoomDo

End Sub

Private Sub Command1_Click(Index As Integer)

    Select Case Index
      Case 0
      
        RichTextBox1.TextRTF = Text1.Text
        RichTextBox1_SelChange
        
      Case 1
        RichTextBox1.SelRTF = Text2.Text
        RichTextBox1_SelChange
    End Select

End Sub



Private Sub DoAvtivate()

    RichTextBox1_SelChange ' force one last update before disabling
    ' (if you are enabling what's one more call:))
    'You could start a timer which makes occassional
    'updates from here if you feel the need
    ActiveUpdate = Not ActiveUpdate
    'reset tool button image
    Toolbar1.Buttons("activeupdate").Image = IIf(ActiveUpdate, "activeon", "activeoff")
    'reset tooltips
    Text1.ToolTipText = MyRTB.AddSlug(Text1.ToolTipText, "(", IIf(ActiveUpdate, "ON", "OFF"), ")")
    Text2.ToolTipText = MyRTB.AddSlug(Text2.ToolTipText, "(", IIf(ActiveUpdate, "ON", "OFF"), ")")
    Text3.ToolTipText = MyRTB.AddSlug(Text3.ToolTipText, "(", IIf(ActiveUpdate, "ON", "OFF"), ")")
    Command1(0).Enabled = ActiveUpdate
    Command1(1).Enabled = ActiveUpdate
'    ExtendedRTFDemo.Height = IIf(ActiveUpdate, 9420, 7260)

End Sub

Private Sub flmnu_Click(Index As Integer)

    Select Case Index
      Case 0
'The Panels can look after themselves and for demo purposes this is better
'but if you don't want them popping up unless there is a selection then
'uncomment the If..Else..End If structures below
'        If RichTextBox1.SelLength > 0 Then
            TextLookPanel.Show vbModal, Me
'          Else 'NOT RICHTEXTBOX1.SELLENGTH...
'            MsgBox "Select some text first", , "Text Look Selector"
'        End If
      Case 1
'        If RichTextBox1.SelLength > 0 Then
            ColourPanel.Show vbModal, Me
'          Else 'NOT RICHTEXTBOX1.SELLENGTH...
'            MsgBox "Select some text first", , "Colour Style Selector"
'        End If
'        'case 2
      Case 3
        RTBLooks.NoFormatting
        'case 4
      Case 4
        RTBLooks.FormatRemove
      Case 6
        RTBLooks.ColourRemoveAll

      Case 7
        RTBLooks.ColourRemove True
      Case 8
        RTBLooks.ColourRemove False
        ' Case 9
      Case 10
        RTBLooks.ExcessSpaceDelete
    End Select

End Sub

Private Sub fmnu_Click(Index As Integer)

    Select Case Index
      Case 0 'new
        MyRTB.FileNew
      Case 1 'Open
        MyRTB.FileOpen
      Case 2
        MyRTB.FileReLoad
      Case 3 'FileSave

        MyRTB.FileSave

      Case 4 'SaveAs
        MyRTB.FileSaveAs
      Case 5
        MyRTB.FileSafeSave
        End
    End Select

End Sub

Private Sub Form_Initialize()

    Manfst.Manifest

End Sub

Private Sub Form_Load()

  'RichTextBox1.Top = Toolbar1.Height + 600
Form1.Show vbModal, Me
    Me.Width = Screen.Width
    Me.Left = 0
     Me.Top = 0
    Manfst.ToolBarButton Toolbar1, "manifest", False, False, False, "manifest", "manifest"
    MyRTB.AssignControls RichTextBox1, CommonDialog1
    RTBZoom.AssignControls RichTextBox1, Combo1
    MyRTB.FileName = RichTextBox1.FileName
    RTBLooks.AssignControls RichTextBox1, CommonDialog1
    RTBHigh.AssignControls RichTextBox1, CommonDialog1
    Text1.Text = RichTextBox1.TextRTF
    Text2.Text = RichTextBox1.SelRTF
    Text3.Text = RichTextBox1.SelText
    RichTextBox1_SelChange
    Show
 PlaceControlOnToolBar Combo1, Toolbar1, "zoom", True
    
    DoAvtivate
    DoAvtivate ' second call disabels it
    Text1.ToolTipText = "This is the RTFcode for the document above. ActiveUpdate"
    Text2.ToolTipText = "This is the RTFcode for the Selection in the document above. ActiveUpdate"
    Text3.ToolTipText = "This is the selected text in the document above. ActiveUpdate"
    Text1.ToolTipText = MyRTB.AddSlug(Text1.ToolTipText, "(", IIf(ActiveUpdate, "ON", "OFF"), ")")
    Text2.ToolTipText = MyRTB.AddSlug(Text2.ToolTipText, "(", IIf(ActiveUpdate, "ON", "OFF"), ")")
    Text3.ToolTipText = MyRTB.AddSlug(Text3.ToolTipText, "(", IIf(ActiveUpdate, "ON", "OFF"), ")")

End Sub

Private Sub Form_Resize()
Dim Halfscreen As Long
    With ExtendedRTFDemo
        RichTextBox1.Width = .Width - 130
        Halfscreen = (.Width - 130) / 2
        Text1.Width = Halfscreen
        Command1(0).Left = Text1.Left
        Label2.Left = Text1.Left + Command1(0).Width + 10
        
        Text2.Left = Text1.Width + 100
        Text2.Width = Halfscreen - 130
        Command1(1).Left = Text2.Left
        Label3.Left = Text2.Left + Command1(1).Width + 10
        Label4.Left = Label3.Left
        Text3.Left = Text1.Width + 100
        Text3.Width = Halfscreen - 130
        
        Frame1.Left = .Width - 200 - Frame1.Width
    End With 'EXTENDEDRTFDEMO

End Sub

Private Sub Form_Unload(Cancel As Integer)

    MyRTB.FileSafeSave

End Sub

Private Sub HCmnu_Click(Index As Integer)

    Select Case Index
      Case 0
        RTBHigh.APIHighlightHard vbRed
      Case 1
        RTBHigh.APIHighlightHard vbBlue
      Case 2
        RTBHigh.APIHighlightHard vbGreen
      Case 3
        RTBHigh.APIHighlightHard vbYellow
      Case 4
        RTBHigh.APIHighlightHard RGB(175, 175, 175)  'grat
      Case 5
        RTBHigh.APIHighlightHard RGB(255, 200, 255)  'mauve
      Case 6
        RTBHigh.APIHighlightHard RGB(255, 200, 175)  'pink
      Case 7
        RTBHigh.APIHighlightHard RGB(255, 200, 100)   'beige
      Case 8
        RTBHigh.APIHighlightHard RGB(255, 255, 200)   'marigold
        ' RGB(200, 255, 200) 'pale green
        ' RGB(200, 255, 100) 'lime green
        ' RGB(255, 200, 100) 'beige
      Case 10
        RTBHigh.APIHighlightUser
      Case 11
        RTBHigh.APIHighlightUserAuto
      Case 12
        RTBHigh.APIHighlightUserUser
      Case 13
        RTBHigh.APIHighlightRemove
    End Select

End Sub

Private Sub hlp2_Click(Index As Integer)

    RTBHigh.About Index = 1

End Sub

Private Sub hlp3_Click(Index As Integer)

    RTBZoom.About Index = 1

End Sub

Private Sub hlp_Click(Index As Integer)

    Select Case Index
      Case 0
        Manfst.About
Case 3
RTBLooks.About
Case 4
MyRTB.About
    End Select

End Sub

Private Sub italic_Click()

    RichTextBox1.SelItalic = Not RichTextBox1.SelItalic

End Sub

Private Sub jmnu_Click(Index As Integer)

    MyRTB.DocPercent = Index * 10

End Sub

Private Sub mmnu_Click(Index As Integer)

  Dim PreserveSelStart As Long, PreserveSelLen As Long

    Select Case Index
      Case 0
        MyRTB.SelQuoteSng = Not MyRTB.SelQuoteSng
      Case 1

        MyRTB.SelQuoteDbl = Not MyRTB.SelQuoteDbl

      Case 2
        MyRTB.ActOnTag "<b>", "<\b>", "\b", True
      Case 3
        MyRTB.LineNumbers
      Case 4
      Case 5
        MyRTB.VisibleParagraphMarks = Not MyRTB.VisibleParagraphMarks
        Case 6
        Form1.Show vbModal, Me
    End Select

End Sub

Private Sub probul_Click(Index As Integer)

  'ulw  undelineWord not supported appears as single underline
  '   (only real difference is that you can remove pieces rather than whole underline)
  'uldb undelineDouble not supported appear as single underline
  'Although not fully supported in RichTextBox the RTF code is preserved
  'and if you open your doc in an big RTF Word Processor (not WordPad)
  'they show properly.

    Select Case Index
      Case 0
        MyRTB.SelUlWord = Not MyRTB.SelUlWord
      Case 1
        MyRTB.SelUlDouble = Not MyRTB.SelUlDouble
    End Select

End Sub

Private Sub RichTextBox1_Change()

    MyRTB.Dirty = True

End Sub

Private Sub RichTextBox1_KeyDown(KeyCode As Integer, Shift As Integer)

    MyRTB.KeyDown KeyCode, Shift

End Sub

Private Sub RichTextBox1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

    MyRTB.MouseMove Button, Shift, x, Y
    Label1.Caption = MyRTB.RichWordOverMod
    Label6.Caption = MyRTB.InstantTran(Label1.Caption)

End Sub

Private Sub RichTextBox1_SelChange()

    If MyRTB.Busy Then
        Exit Sub 'Busy will stop updates while class is very busy '>---> Bottom
    End If
    'these are not major time wasters but Busy will stop them while class is very busy
    ExtendedRTFDemo.Caption = RTBZoom.ZoomSlug(ExtendedRTFDemo.Caption, False)
    Label5.Caption = "Document Position: " & MyRTB.DocPercent & "%"
    ExtendedRTFDemo.Caption = MyRTB.AddSlug(ExtendedRTFDemo.Caption, "+Colours Total: ", MyRTB.ColoursUsed & " Selected: " & MyRTB.ColoursUsed(False), "+")
    ExtendedRTFDemo.Caption = RTBHigh.HighLightSlug(ExtendedRTFDemo.Caption, "^Highlight=", ColourRGB, "^")
    ExtendedRTFDemo.Caption = RTBHigh.HighLightSlug(ExtendedRTFDemo.Caption, "(HighlightMessage=", Description, ")")
    With Toolbar1
        .Buttons("bold").Value = IIf(RichTextBox1.SelBold, tbrPressed, tbrUnpressed)
        .Buttons("underline").Value = IIf(RichTextBox1.SelUnderline, tbrPressed, tbrUnpressed)
        .Buttons("italic").Value = IIf(RichTextBox1.SelItalic, tbrPressed, tbrUnpressed)
    End With 'TOOLBAR1

    If ActiveUpdate Then 'stop them if user selects to
        ' these are real speed killers so can be deactivated from Toolbar traffic light button
        Check2.Value = IIf(MyRTB.HasHighlitText, vbChecked, vbUnchecked)
        Check1.Value = IIf(RTBLooks.SelVisible, vbChecked, vbUnchecked)

        With RichTextBox1
            Text1.Text = .TextRTF
            Text2.Text = .SelRTF
            Text3.Text = .SelText
        End With 'RICHTEXTBOX1
        Text1.SelStart = Len(Text1.Text) / 100 * MyRTB.DocPercent
    End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
      Case "bold"
        RichTextBox1.SelBold = Not RichTextBox1.SelBold
      Case "italic"
        RichTextBox1.SelItalic = Not RichTextBox1.SelItalic
      Case "underline"
        MyRTB.SelRTFToggle MyRTB.CurrentUnderlineStyle
      Case "fontincrease"
        MyRTB.FontSizeStep True
      Case "fontdecrease"
        MyRTB.FontSizeStep False
      Case "restore"
        MyRTB.FileReLoad
      Case "activeupdate"
        DoAvtivate
      Case "manifest"
      If IsDebugMode Then
      MsgBox "Changing the manifest only works with compiled program"
      Else
        Manfst.Action
      End If
    End Select

End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

  Dim BckCol As Long

    Select Case ButtonMenu.Key
      Case "standard"
        RichTextBox1.SelUnderline = Not RichTextBox1.SelUnderline
      Case "dot"
        MyRTB.SelDot = Not MyRTB.SelDot
      Case "dash"
        MyRTB.SelDash = Not MyRTB.SelDash
      Case "dashd"
        MyRTB.SelDashd = Not MyRTB.SelDashd
      Case "dashdd"
        MyRTB.SelDashdd = Not MyRTB.SelDashdd
      Case "thick"
        MyRTB.SelThick = Not MyRTB.SelThick
      Case "wave"
        MyRTB.SelWave = Not (MyRTB.SelWave) ' True 'ExtendedRTFToggle "\ulwave"
      Case "word"
        MyRTB.SelUlWord = Not MyRTB.SelUlWord
      Case "double"
        MyRTB.SelUlDouble = Not MyRTB.SelUlDouble
      Case "clear"
        RTBHigh.APIHighlightRemove

      Case "red"
        RTBHigh.APIHighlightHard vbRed
      Case "blue"
        RTBHigh.APIHighlightHard vbBlue
      Case "green"
        RTBHigh.APIHighlightHard vbGreen
      Case "yellow"
        RTBHigh.APIHighlightHard vbYellow
      Case "select"
        RTBHigh.APIHighlightUser
      Case "fblack"
        MyRTB.SelColor = vbBlack
      Case "fblue"
        MyRTB.SelColor = vbBlue
      Case "fred"
        MyRTB.SelColor = vbRed
      Case "fyellow"
        MyRTB.SelColor = vbYellow
      Case "fselect"
        MyRTB.SelColor = MyRTB.ColourUser
      Case "useruser"
        RTBHigh.APIHighlightUserUser
      Case "userauto"
        RTBHigh.APIHighlightUserAuto
      Case "user"
        RTBHigh.APIHighlightUser
    End Select

End Sub

Private Sub Ulsmnu_Click(Index As Integer)

    Select Case Index
      Case 0
        MyRTB.SelUnderline = Not MyRTB.SelUnderline
      Case 1
        MyRTB.SelDot = Not MyRTB.SelDot
      Case 2
        MyRTB.SelDash = Not MyRTB.SelDash
      Case 3
        MyRTB.SelDashd = Not MyRTB.SelDashd
      Case 4
        MyRTB.SelDashdd = Not MyRTB.SelDashdd
      Case 5
        MyRTB.SelThick = Not MyRTB.SelThick
      Case 6
        MyRTB.SelHair = Not MyRTB.SelHair
      Case 7
        MyRTB.SelWave = Not MyRTB.SelWave
      Case 11
        RichTextBox1.SelUnderline = False
    End Select

End Sub

Private Sub vmnu_Click(Index As Integer)

    Select Case Index
      Case 0
        RTBLooks.SelVisible = Not RTBLooks.SelVisible
        '-
      Case 2
        MyRTB.HiddenTextShow
        '-
      Case 4
        MyRTB.ApplyTag "*I*", "*V*"
      Case 5
        MyRTB.ActOnTag "*I*", "*V*", "\v"

    End Select

End Sub

':) Ulli's VB Code Formatter V2.13.6 (22/08/2002 2:13:28 AM) 4 + 438 = 442 Lines
