VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form ColourPanel 
   Caption         =   "Colour Style Selector"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   8955
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      Caption         =   "Preserve Colour Selection"
      Height          =   375
      Left            =   5640
      TabIndex        =   15
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mode"
      Height          =   735
      Left            =   7800
      TabIndex        =   12
      Top             =   120
      Width           =   975
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   600
         TabIndex        =   14
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "Text1"
         BuddyDispid     =   196611
         OrigLeft        =   600
         OrigTop         =   240
         OrigRight       =   855
         OrigBottom      =   615
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
   End
   Begin RichTextLib.RichTextBox RTBClrDemo 
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   873
      _Version        =   393217
      TextRTF         =   $"ColorPanel.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Description"
      Height          =   2655
      Left            =   1920
      TabIndex        =   8
      Top             =   120
      Width           =   2655
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   2175
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.ListBox List2 
      Height          =   2400
      Left            =   6120
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   7680
      TabIndex        =   6
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Do It"
      Height          =   375
      Index           =   0
      Left            =   6480
      TabIndex        =   5
      Top             =   3960
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Back"
      Height          =   195
      Index           =   1
      Left            =   7800
      TabIndex        =   4
      Top             =   2145
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Text"
      Height          =   195
      Index           =   0
      Left            =   7800
      TabIndex        =   3
      Top             =   1830
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "InOut"
      Height          =   195
      Index           =   1
      Left            =   7800
      TabIndex        =   2
      Top             =   1275
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "LeftRight"
      Height          =   195
      Index           =   0
      Left            =   7800
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   $"ColorPanel.frx":00CE
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Width           =   5295
   End
End
Attribute VB_Name = "ColourPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Copyright 2002 Roger Gilchrist
'rojgilkrist@hotmail.com
'very new; not much comment
'you'll have to work it out
Private Description(10) As String
Private Enum List2Fillers
    blank
    Spectrum
    rainbows
    Material
    random
End Enum
Private Const NL As String = vbNewLine
Private Demo As New ClsRTFFontLooks
Private PreserveColour As Boolean

Private Sub ActivateTools(Inout As Boolean, LeftRight As Boolean, TxtBck As Boolean, Mde As Boolean, DescNumber As Integer, Optional LRCaption As Boolean = True)

    Check1(0).Visible = LeftRight
    Check1(0).Caption = IIf(LRCaption, "LeftRight", "LightDark")
    Check1(1).Visible = Inout
    List2.Visible = List2.ListCount > 0
    Label1.Caption = Description(DescNumber)
    Option1(0).Visible = TxtBck
    Option1(1).Visible = TxtBck
    Frame2.Visible = Mde
    Check2.Value = vbUnchecked 'turn off Preserve Colour Selection
    Check2.Visible = (InStr(List1.List(List1.ListIndex), "*") > 0) 'Hide if not needed

End Sub

Private Sub Check1_Click(Index As Integer)

    DemoShow

End Sub

Private Sub Check2_Click()

    PreserveColour = (Check2.Value = vbChecked)

End Sub

Private Sub Command1_Click(Index As Integer)

    If Index = 0 Then
        TakeAction False
    End If
    Unload ColourPanel

End Sub

Private Sub DemoShow()

    RTBClrDemo.SelStart = 0
    RTBClrDemo.SelLength = Len(RTBClrDemo.Text)
    If Left$(LCase$(List1.List(List1.ListIndex)), 6) = "highli" Then
        RTBClrDemo.Find "a strong contrasting", 1
    End If
    Demo.ColourRemoveAll
    TakeAction True

End Sub

Private Sub Form_Load()

    Demo.AssignControls RTBClrDemo, ExtendedRTFDemo.CommonDialog1
    'The name of this   V_________V needs to match that being used by you RichTextBox
    Command1(0).Enabled = RTBLooks.IsSelection
    Command1(0).Caption = IIf(Command1(0).Enabled, "Do It", "No Selection")
    'List1 is Sorted=True as the rest of the form actually reads the lcase string value of this list
    'You can add either Lcase,Ucase or ProperCase strings here but make sure you use lcase everywhere else
    'Dont forget to add a description
    'Add a 1 * or 2 ** if the option includes 1 or 2 colour selection(s)
    With List1
        .Clear
        .AddItem "Blend **"
        .AddItem "Candy"
        .AddItem "Dither *"
        .AddItem "Materials"
        .AddItem "Rainbow"
        .AddItem "Random"
        .AddItem "Spectrum"
        .AddItem "HighlightUser *"
        .AddItem "HighlightUserAuto *"
        .AddItem "HighlightUserUser **"

    End With 'LIST1

    Description(0) = "Select a colour style" & NL & _
                "If there is 1 * you will be asked to select one colour" & NL & _
                "If there are 2 ** you will be asked to select two colours"
    Description(1) = "Candy" & NL & "" & NL & _
                "1. Select a Spectrum Setion." & NL & _
                "2. Select Text or Back."
    Description(2) = "Blend" & NL & _
                "Two colours are blended into each other." & NL & _
                "1. Check/Uncheck Inout." & NL & _
                "2. Select Text or Back."
    Description(3) = "Rainbow" & NL & _
                "A full width rainbow spectrum is created." & NL & _
                "1. Select a Direction." & NL & _
                "2. Select Text or Back."
    Description(4) = "Spectrum" & NL & _
                "One of 6 rainbow colour ranges is created." & NL & _
                "1. Select a Spectrum Setion." & NL & _
                "2. Check/UnCheck InOut" & NL & _
                "3. Check/UnCheck LeftRight" & NL & _
                "4. Select Text or Back."
    Description(5) = "Materials" & NL & _
                "Smoothly changing material colour spread is created." & NL & _
                "1. Select a Material" & NL & _
                "2. Check/UnCheck InOut" & NL & _
                "3. Check/UnCheck LeftRight" & NL & _
                "4. Select Text or Back."
    Description(6) = "Random" & NL & _
                "Each Character gets its own colour" & NL & _
                "1. Select a Colour range" & NL & _
                "2. Check/UnCheck InOut" & NL & _
                "3. Select Text or Back." & NL & NL & _
                "Remember this is really random. The sample is NOT exactly what you will get in the main document."
    Description(7) = "Dither" & NL & _
                "A selected colour is dithered from dark to light" & NL & _
                "1. Check/UnCheck LightDark" & NL & _
                "2. Select Text or Back."
    Description(8) = "RTFHighlightUser" & NL & _
                "Back Colour is selected by user." & NL & _
                "1. Select Back colour from ColorDialog" & NL & NL & _
                "Note there are also HighlghtHard versions of this routine if you want to hard code a colour." & NL & _
                "The Preserve Colour Selection Checkbox uses the Hard version to save time"
    Description(9) = "RTFHighlightUserAuto" & NL & _
                "Back Colour is selected by user, Text colour by Program." & NL & _
                "1. Select Back Colour from ColorDialog" & NL _
                & "2. Class creates a Contrasting Text Colour" & NL & NL & _
                "Note there are also HighlghtHard versions of this routine if you want to hard code a colour." & NL & _
                "The Preserve Colour Selection Checkbox uses the Hard version to save time"
    Description(10) = "RTFHighlightUserUser" & NL & _
                "Text and Back colour are selected by user." & NL & _
                "1. Select Back Colour from ColorDialog" & NL & _
                "2. Select Text Colour from ColorDialog" & NL & NL & _
                "Note there are also HighlghtHard versions of this routine if you want to hard code a colour." & NL & _
                "The Preserve Colour Selection Checkbox uses the Hard version to save time"

    ActivateTools False, False, False, False, 0 'turn everything off at first

End Sub

Private Sub List1_Click()

  'note use lcase name in Case "whatever" when adding new ones
  'work out which tools should be active
  'and whether List2 needs to show any thing

    Select Case LCase$(List1.List(List1.ListIndex))
      Case "candy"
        List2Filler Spectrum
        ActivateTools False, False, True, True, 1
        UpDown1.Max = 3
      Case "rainbow"
        List2Filler rainbows
        ActivateTools False, False, True, False, 3
      Case "spectrum"
        List2Filler Spectrum
        ActivateTools True, True, True, False, 4
      Case "blend **"
        List2Filler blank
        ActivateTools True, False, True, False, 2
      Case "materials"
        List2Filler Material
        ActivateTools True, True, True, False, 5
      Case "random"
        List2Filler random
        ActivateTools False, False, True, False, 6
      Case "dither *"
        List2Filler blank
        ActivateTools True, True, True, False, 7, False
      Case "highlightuser *"
        List2Filler blank
        ActivateTools False, False, False, False, 8
      Case "highlightuserauto *"
        List2Filler blank
        ActivateTools False, False, False, False, 9
      Case "highlightuseruser **"
        List2Filler blank
        ActivateTools False, False, False, False, 10

      Case Else
        List2Filler blank
        ActivateTools False, False, False, False, 0
    End Select

    DemoShow
    'turn on Preserve Colour Selection if necessary
    'value has to be set after DemoShow or you don't get initial colour choice option
    Check2.Value = IIf(InStr(LCase$(List1.List(List1.ListIndex)), "*"), vbChecked, vbUnchecked)

End Sub

Private Sub List2_Click()

  'List2 is deliberately left Sorted=False so that the panel can use list postion to read selections

    DemoShow

End Sub

Private Sub List2Filler(Mode As List2Fillers)

  'List2 is deliberately left Sorted=False so that the panel can use list postion to read selections
  'For neatness try to set them in Alpha order

    With List2
        .Clear
        Select Case Mode
          Case Spectrum
            .AddItem "s1RedYellow"
            .AddItem "s2YellowGreen"
            .AddItem "s3GreenCyan"
            .AddItem "s4CyanBlue"
            .AddItem "s5BlueMagenta"
            .AddItem "s6MagentaRed"
            .ListIndex = 0
          Case blank 'do nothing

          Case rainbows
            .AddItem "RedMagenta"
            .AddItem "MagentaRed"

          Case Material
            .AddItem "Blue Steel"
            .AddItem "Diamond"
            .AddItem "Flesh1"
            .AddItem "Ice"
            .AddItem "Gold"
            .AddItem "Lead"
            .AddItem "Milkchocolate"
            .AddItem "Oldgold"
            .AddItem "Pineboard"
            .AddItem "Rubber"
            .AddItem "Silver"
            .AddItem "Slate"
            .AddItem "Yellowrose"
            .AddItem "Test"
          Case random
            .AddItem "All colours"
            .AddItem "Rainbow"
            .AddItem "s1RedYellow"
            .AddItem "s2YellowGreen"
            .AddItem "s3GreenCyan"
            .AddItem "s4CyanBlue"
            .AddItem "s5BlueMagenta"
            .AddItem "s6MagentaRed"
            .AddItem "grey"
        End Select
    End With 'LIST2

End Sub

Private Sub Option1_Click(Index As Integer)

    DemoShow

End Sub

Private Sub TakeAction(DemoDoc As Boolean)

  'PreserveColour prevents the colour dialog from firing every time you change things
  'for those tools which need user to select colours

  Dim Target As Variant
  Dim TextBack As Boolean
  Dim InOuter As Boolean
  Dim LeftRighter As Boolean

    TextBack = (Option1(0).Value = True)
    LeftRighter = (Check1(0).Value = vbChecked)
    InOuter = (Check1(1).Value = vbChecked)
    Demo.ColourRemoveAll

    If DemoDoc Then
        Set Target = Demo
      Else 'DEMODOC = FALSE
        Set Target = RTBLooks
    End If
    Select Case LCase$(List1.List(List1.ListIndex))
      Case "blend **"

        If DemoDoc Then
            If PreserveColour Then
                Target.BlenderHard Demo.LastBlendStartColour, Demo.LastBlendEndColour, InOuter, TextBack
              Else 'PRESERVECOLOUR = FALSE
                Target.Blender InOuter, TextBack
            End If
          Else 'DEMODOC = FALSE
            Target.BlenderHard Demo.LastBlendStartColour, Demo.LastBlendEndColour, InOuter, TextBack
        End If
      Case "candy"
        Target.Candy List2.ListIndex, Val(Text1.Text), TextBack
      Case "dither *"
        If DemoDoc Then
            If PreserveColour Then
                Target.DitherHard Demo.LastDitherColor, LeftRighter, InOuter, TextBack
              Else 'PRESERVECOLOUR = FALSE
                Target.Dither LeftRighter, InOuter, TextBack
            End If
          Else 'DEMODOC = FALSE

            Target.DitherHard Demo.LastDitherColor, LeftRighter, InOuter, TextBack
        End If
      Case "materials"
        Target.Materials LCase$(List2.List(List2.ListIndex)), InOuter, LeftRighter, TextBack
      Case "rainbow"
        Target.RainBow List2.List(List2.ListIndex) = "RedMagenta", TextBack
      Case "random"
        Target.RandomColour List2.ListIndex, TextBack
      Case "spectrum"
        Target.SpectrumSector List2.ListIndex, InOuter, LeftRighter, TextBack
      Case "highlightuser *"
        If DemoDoc Then
            If PreserveColour Then
                Target.RTFHighlightUser
              Else 'PRESERVECOLOUR = FALSE
                Target.RTFHighlightHard Demo.LastHighlightColour
            End If
          Else 'DEMODOC = FALSE
            Target.RTFHighlightHard Demo.LastHighlightColour
        End If
      Case "highlightuserauto *"
        If DemoDoc Then
            If PreserveColour Then
                Target.RTFHighlightHardauto Demo.LastHighlightColour
              Else 'PRESERVECOLOUR = FALSE
                Target.RTFHighlightUserAuto
            End If

          Else 'DEMODOC = FALSE
            Target.RTFHighlightHardauto Demo.LastHighlightColour
        End If
      Case "highlightuseruser **"
        If DemoDoc Then
            If PreserveColour Then

                Target.RTFHighlightHardHard Demo.LastHighlightColour, Demo.LastHighlightForeColour
              Else 'PRESERVECOLOUR = FALSE
                Target.RTFHighlightUserUser
            End If
          Else 'DEMODOC = FALSE
            Target.RTFHighlightHardHard Demo.LastHighlightColour, Demo.LastHighlightForeColour
        End If
    End Select

End Sub

':) Ulli's VB Code Formatter V2.13.6 (22/08/2002 11:30:06 AM) 16 + 337 = 353 Lines
