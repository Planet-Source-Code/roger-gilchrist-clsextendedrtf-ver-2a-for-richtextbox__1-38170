VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsRTFFontLooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'ClsRTFFontColours
'copyright 2002 Roger Gilchrist
'This class modifies the TextRTF string to create various colour effects on text in RichTextboxes
'All colour routines contain an Optional Boolean switch ForeBack which determines whether to text
'(True) or the background(False) is coloured.
'I have created most of the colour sets myself but the following sources served as inspiration
' and constructed the math behind the routines (I have rewritten them into spearate routines)
'------------IMPORTANT INFORMATION-------------------------------
'there is a problem with supporting highlighting (which is used for back colouring in this class)
'This requires the Riched20.dll (version 3)
'This also requires Riched32.dll (5.00.2008.1)
'and probably the Richx32.ocx control
'VB6 SP5 (what I use) has no difficult with this
'thanks to Thomas Görtler (maniac@ colorarts.de) who let me know of the problem then solved it himself
'and Sergio Perciballi (oigres@postmaster.co.uk) whose code contained the detailed version stuff above

'I borrowed and modified the main colour routines from then following people's code
'Thanks to all
'Routine:      | Source
'-------------------------------------------------------------------------------------
'Blend:-        Dan Redding (bwsoft@revealed.net)
'RainbowColor:- DrawRainBow © oigres P Email: oigres@postmaster.co.uk
'Materials:-    John Colman (john_colman@hotmail.com)in metalCBprj.OCX
'
'If you construct an interesting variant or expand the colour set of Materials
'send me details & code and I'll probably add it to the next edition of the code
'and add you to the list above

'Search this routine for the following line
'*---PROGRAMMER MODIFICATION POINT---*
'it marks points at which you might edit or delete
'If you modify other routines you will need to be
'very careful they are usually vital to the class

'I'm a bit compulsive about modularity so each module has its own
'Private calls to to the API. This means if you copy just one module
'you don't have to keep coming back to find that missing Public variable
'in another module. Once you have inculde a module in your program you can
'remove the Private Declare statements if you have a Public one in your project

'copyright 2002 Roger Gilchrist
'rojagilkrist@hotmail.com

Option Explicit
Private m_DefFormat As String
Private m_CDlg As CommonDialog
Private m_busy As Boolean
Private m_RTB As RichTextBox
Private m_End_OF_RTF As Long
Private Const DQ As String = """"      ' DB=DoubleQuote NL=NewLine
Private Const NL As String = vbNewLine ' I use these for long MsgBoxes layout.
Private m_Start() As Long
Private m_Len() As Long
Private m_LastHighlightColour As Long
Private m_LastHighlightForeColour As Long
Private m_LastBlendStartColour As Long
Private m_LastBlendEndColour As Long
Private m_LastDitherColor As Long
Private Const SafeJoinSkip As String = "~!~!~!*~!"
Private Enum LIFO
    Push = True
    pop = False
End Enum
'#If False Then'Enforce Case For Enums (does not compile but fools IDE)
'Dim Push
'Dim Pop
'#End If'Barry Garvin VBPJ 101 Tech Tips 11 March 2001 p1

Private Enum RTFCodeLocs
    NotPresent
    EmbedOr1st
    LastWithBlank
    LastInSelection
    LastBeforeText
    EndOfString
End Enum
'#If False Then
'Dim NotPresent
'Dim EmbedOr1st
'Dim LastWithBlank
'Dim LastInSelection
'Dim LastBeforeText
'dim EndOfString
'#End If

Public Enum Spectrum
    s1RedYellow = 0
    s2YellowGreen
    s3GreenCyan
    s4CyanBlue
    s5BlueMagenta
    s6MagentaRed
End Enum
'#If False Then
'Dim s1RedYellow
'Dim s2YellowGreen
'Dim s3GreenCyan
'Dim s4CyanBlue
'Dim s5BlueMagenta
'Dim s6Magentared
'#End If
'
Private Enum SpectrumMode
    LeftRight = 0
    RightLeft = 1
    InOutLeftRight
    InOutRightLeft
End Enum
'#If False Then 'Enforce Case For Enums (does not compile but fools IDE)
'Dim LeftRight
'Dim InOutLeftRight
'Dim RightLeft
'Dim InOutRightLeft
'#End if  'Barry Garvin VBPJ 101 Tech Tips 11 March 2001 p1

Private Enum RTFPartsType
    blank
    RTFOnly
    TextOnly
    Mixed
End Enum
'#If False Then
'Dim Blank
'Dim RTFOnly
'Dim TextOnly
'Dim Mixed
'#End If
Private m_rainbowPhase As Spectrum
Private m_rainbowStep As Integer
Private Declare Function IsCharAlphaNumeric Lib "user32" Alias "IsCharAlphaNumericA" (ByVal cChar As Byte) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Public Sub About()
Dim Msg As String
Msg = "Not settled enough for a hard coded description yet. You'll have to read the code itself. This class is behind the Panels on the Font Looks menu."
MsgBox Msg, , "ClsRTFFontLooks"
End Sub
Public Sub AssignControls(R As RichTextBox, C As CommonDialog)

  'Copyright 2002 Roger Gilchrist
  'Place the Call to this in Form_Load or Sub Main
  'RTBLooks.AssignControls RichTextBox1, CommonDialog1

    m_busy = False
    Set m_CDlg = C
    Set m_RTB = R

End Sub

Private Function Blend(RGB1 As Long, RGB2 As Long, Percent As Single) As Long

  'Dan Redding (bwsoft@revealed.net)
  'http://home.revealed.net/bwsoft
  'developed this code which I have modified to use my ColourLong2RGB
  '-----------------------------------------------------------
  'These are his original notes for this routine
  '----------------------------------------------------------
  'This one doesn't really use the HSL routines, just the
  'RGB Component routines.  I threw it in as a bonus ;)
  'Takes two colors and blends them according to a
  'percentage given as a Single
  'For example, .3 will return a color 30% of the way
  'between the first color and the second.
  '.5, or 50%, will be an even blend (halfway)
  'Can create some nice effects inside a For loop
  '-----------------------------------------------------------
  'NOTE the Percent has to be expressed as a value between 0 and 1
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if not needed

  Dim R As Integer, R1 As Integer, R2 As Integer
  Dim G As Integer, G1 As Integer, G2 As Integer
  Dim B As Integer, B1 As Integer, B2 As Integer

    If Percent >= 1 Then
        Blend = RGB2
        Exit Function '>---> Bottom
      ElseIf Percent <= 0 Then 'NOT PERCENT...
        Blend = RGB1
        Exit Function '>---> Bottom
    End If

    ColourLong2RGB RGB1, R1, G1, B1
    ColourLong2RGB RGB2, R2, G2, B2

    R = ((R2 * Percent) + (R1 * (1 - Percent)))
    G = ((G2 * Percent) + (G1 * (1 - Percent)))
    B = ((B2 * Percent) + (B1 * (1 - Percent)))

    Blend = RGB(R, G, B)

End Function

Public Sub Blender(Optional Inout As Boolean = False, Optional ForeBack As Boolean = True)

  'Copyright 2002 Roger Gilchrist
  'This Sub creates an array of Blended colours for the ColourApplicator
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if not needed
  'use this as a template for your own colour routines

  Dim i As Long
  Dim SPos As Long, EPos As Long 'Range of Text and hence BlendArray
  Dim BlendArray() As Long 'Store colours for passing to ColourApplicator
  Dim perc As Single ' this allows the maximum level of detail in the blend array
  Dim CStart As Long, CEnd As Long 'Start and End Colours
  Dim PerClen As Long, PerCur As Long 'calculate percentage of string to drive Blend routine

    If GetStartEnd(SPos, EPos) = False Then 'Use at start of any modifications you build;
        Exit Sub                            'sets start and end and Exits if no Selection '>---> Bottom
    End If
    CStart = ColourUser     'Get the colours
    CEnd = ColourUser
    m_LastBlendStartColour = CStart
    m_LastBlendEndColour = CEnd
    PerClen = EPos          'save range of blend target text
    'EPos = SPos + EPos - 1  'get upper boundry of selection
    'for ease of adding colour to text the colour array is initialised to
    'the character position rather than a standard zero based array
    ReDim BlendArray(SPos To EPos) As Long ' initialise array

    If Inout Then
        'This creates a blend from Cstart to Cend and back to Cstart
        For i = SPos To EPos
            PerCur = PerCur + 1 ' current member of PerClen range
            'blend needs percent expressed decimally; /50 increases the speed of change
            perc = Percent(PerCur, PerClen) / 50
            BlendArray(i) = Blend(CStart, CEnd, perc)
            BlendArray(EPos - PerCur + 1) = BlendArray(i)
            If PerCur > PerClen \ 2 Then
                Exit For '>---> Next
            End If
        Next i
      Else 'INOUT = FALSE
        'This creates a blend from Cstart to Cend
        For i = SPos To EPos
            PerCur = PerCur + 1
            'blend needs percent expressed decimally hence the /100
            BlendArray(i) = Blend(CStart, CEnd, Percent(PerCur, PerClen) / 100)
        Next i
    End If

    ColourApplicator BlendArray, ForeBack

End Sub

Public Sub BlenderHard(CStart As Long, CEnd As Long, Optional Inout As Boolean = False, Optional ForeBack As Boolean = True)

  'Copyright 2002 Roger Gilchrist
  'This Sub creates an array of Blended colours for the ColourApplicator
  'Designed to provide a hard coded version of Blender
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if not needed
  'use this as a template for your own colour routines

  Dim i As Long
  Dim SPos As Long, EPos As Long 'Range of Text and hence BlendArray
  Dim BlendArray() As Long 'Store colours for passing to ColourApplicator
  Dim perc As Single ' this allows the maximum level of detail in the blend array
  Dim PerClen As Long, PerCur As Long 'calculate percentage of string to drive Blend routine

    If GetStartEnd(SPos, EPos) = False Then 'Use at start of any modifications you build;
        Exit Sub                            'sets start and end and Exits if no Selection '>---> Bottom
    End If
    PerClen = EPos          'save range of blend target text
    'EPos = SPos + EPos - 1  'get upper boundry of selection
    'for ease of adding colour to text the colour array is initialised to
    'the character position rather than a standard zero based array
    ReDim BlendArray(SPos To EPos) As Long ' initialise array

    If Inout Then
        'This creates a blend from Cstart to Cend and back to Cstart
        For i = SPos To EPos
            PerCur = PerCur + 1 ' current member of PerClen range
            'blend needs percent expressed decimally; /50 increases the speed of change
            perc = Percent(PerCur, PerClen) / 50
            BlendArray(i) = Blend(CStart, CEnd, perc)
            BlendArray(EPos - PerCur + 1) = BlendArray(i)
            If PerCur > PerClen \ 2 Then
                Exit For '>---> Next
            End If
        Next i
      Else 'INOUT = FALSE
        'This creates a blend from Cstart to Cend
        For i = SPos To EPos
            PerCur = PerCur + 1
            'blend needs percent expressed decimally hence the /100
            BlendArray(i) = Blend(CStart, CEnd, Percent(PerCur, PerClen) / 100)
        Next i
    End If

    ColourApplicator BlendArray, ForeBack

End Sub

Public Property Get Busy() As Variant

  'Read only use to stop updates in external program if class is working hard

    Busy = m_busy

End Property

Public Sub Candy(Range As Spectrum, Optional Spread As Integer = 0, Optional ForeBack As Boolean = True)

  'Copyright 2002 Roger Gilchrist
  ' variant of rainbow discovered by accident while developing Rainbow
  'just too pretty to dump so I modifed it as a separate routine
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if not needed, play with it to develop other colour sets

  Dim i As Long
  Dim SPos As Long, EPos As Long
  Dim RbowMod As Long, Rstep As Long, CurChar As Long
  Dim RainBowArray() As Long

    If GetStartEnd(SPos, EPos) = False Then 'Use at start of any modifications you build;
        Exit Sub                            'sets start and end and Exits if no Selection '>---> Bottom
    End If
    RbowMod = (EPos - SPos) \ 5
    ReDim RainBowArray(SPos To EPos) As Long
    For i = SPos To EPos
        CurChar = CurChar + 1
        Select Case Spread
          Case 0
            Rstep = Rstep + RbowMod
          Case 1
            Rstep = 255 * (((CurChar + RbowMod) * EPos) / 100)
          Case 2
            Rstep = 255 * ((i * EPos) / 100)
          Case 3
            Rstep = 255 * CurChar / RbowMod
        End Select
        RainBowArray(i) = RainbowColor(CInt(Range), Rstep Mod 255)
    Next i

    ColourApplicator RainBowArray, ForeBack

End Sub

Private Function ColorRTFCode(col As Long) As String

  'Copyright 2002 Roger Gilchrist
  'creates an RTF colour code member
  'support for ColourManipulation routine

  Dim R As Integer, G As Integer, B As Integer

    ColourLong2RGB col, R, G, B

    ColorRTFCode = "\red" & R & "\green" & G & "\blue" & B & ";"

End Function

Private Sub ColourApplicator(ClrArray, ForeBack As Boolean, Optional RndPos As Boolean = False)

  'Copyright 2002 Roger Gilchrist
  ' this is the core of the various coloured text routines

  Dim RTFHead As String 'hold the part of the selection which is not changed(unless it contains a colortbl)
  Dim clrTbl As String 'Extract colortbl from RTFHead if it exists or fake it otherwise
  Dim OrigclrTbl As String 'store the ColorTbl so that it can be Replaced
  Dim Parts As Variant 'hold and manipulate the part of the selection which can be changed
  Dim curcolr As Long  'keep track of position in ClrArray
  Dim i As Long        'For Next through the Selection.

    PositionStore Push '
    clrTbl = GetColorTable()    '
    OrigclrTbl = clrTbl         'store for later Replace call
    If Len(clrTbl) = 0 Then ' if No colorTbl create a new one with auto first value ";"
        clrTbl = "{\colortbl{;"
      Else 'else strip the end off the existing one'NOT LEN(CLRTBL)...
        clrTbl = Left$(clrTbl, Len(clrTbl) - 3)
    End If
    SeparateRTFString RTFHead, Parts
    curcolr = LBound(ClrArray) 'get first colour from array
    For i = LBound(Parts) To UBound(Parts)  ')
        If RndPos Then
            ForeBack = IIf(Rnd > 0.5, True, False)
        End If
        ColourManipulate Parts(i), clrTbl, ClrArray, curcolr, ForeBack
    Next i
    clrTbl = clrTbl & "}  " 'finish colortbl
    If InStr(RTFHead, "{\colortbl") Then 'Replace Original colortbl with new one
        RTFHead = Replace(RTFHead, OrigclrTbl, clrTbl)
      Else ' if there was no original colortbl add it and another "}" to end it'NOT INSTR(RTFHEAD,...
        RTFHead = Left$(RTFHead, Len(RTFHead) - 1) & clrTbl & "}"
    End If
    WorkingStringWrite Trim$(RTFHead) & Join(Parts)
    PositionStore pop

End Sub

Private Sub ColourLong2RGB(ByVal col As Long, Optional ByRef Red As Integer, Optional ByRef Green As Integer, Optional ByRef Blue As Integer)

  'don't know where I got this but there are a million variations around

    Red = col Mod 256
    Green = ((col And &HFF00FF00) / 256)
    Blue = (col And &HFF0000) / (65536)

End Sub

Private Sub ColourManipulate(Part, clrTbl$, ClrArray, curcolr As Long, ForeBack As Boolean)

  'Copyright 2002 Roger Gilchrist
  'does the actual work of inserting colour codes into colortbl and rtf code

  Dim ClrCount As Long 'get current number of colors in colortbl
  Dim i As Long ' loop driver' probably doesn't need to be a long but its safer
  Dim tmpP As String 'bulid new string to replace Part

    ClrCount = CountOccurances(clrTbl, "\red") 'get length of ColorTbl
    Select Case RTFPartType(CStr(Part))
      Case blank 'blanks have to absorb an array member
        tmpP$ = tmpP$ & IIf(ForeBack, "\cf", "\highlight") & ClrCount + 1 & " "
        clrTbl$ = clrTbl$ & ColorRTFCode(CLng(ClrArray(curcolr)))
        curcolr = curcolr + 1
      Case RTFOnly
        tmpP = Part
      Case Mixed, TextOnly
        'add a colour to each letter and advance colour counter for text letter but ignore RTF code
        For i = 1 To Len(Part)
            If InStr("\}", Mid$(Part, i, 1)) = 0 Then
                tmpP$ = tmpP$ & IIf(ForeBack, "\cf", "\highlight") & ClrCount + i & " " & Mid$(Part, i, 1)
                clrTbl$ = clrTbl$ & ColorRTFCode(CLng(ClrArray(curcolr)))
                curcolr = curcolr + 1
              Else 'NOT INSTR("\}",...
                'becuase of the way Part is generated once you reach an RTF code token there is not more text
                tmpP$ = tmpP$ & Mid$(Part, i)
                Exit For '>---> Next
            End If
        Next i
    End Select
    Part = tmpP

End Sub

Public Function ColourNamed(ColStr$) As Long

  'Copyright 2002 Roger Gilchrist
  '*---PROGRAMMER MODIFICATION POINT---*
  'If you want to use specific colours for highlighting in you program
  'give them names and set them here.
  'the name can of course be anything you like
  'Descriptive: "spellingEerror" or "SpellingErrorBack" and "SpellingErrorFore"
  'Colour:"beige"' if you do this make it a believable one especially if end-users will see it
  'Anything that makes sense to you:"fred","Wilma","98o9876r65","!*&#@$"
  '--------------IMPORTANT-----------------------
  'Remember this routine is designed to be case insensitive so use lower case when you add your colours
  '----------------------------------------------

    Select Case LCase$(ColStr$) ' this means that if you mistype capitals the routine doesn't care
      Case "beige"
        ColourNamed = RGB(255, 200, 100)
      Case "mauve"
        ColourNamed = RGB(255, 200, 255)
      Case "grey", "gray" ' so both English and American spellers get the same thing!
        ColourNamed = RGB(175, 175, 175)
      Case "pink"
        ColourNamed = RGB(255, 200, 175)
      Case "beige"
        ColourNamed = RGB(255, 200, 100)
      Case "marigold"
        ColourNamed = RGB(255, 255, 200)
      Case "palegreen"
        ColourNamed = RGB(200, 255, 200)
      Case "limegreen"
        ColourNamed = RGB(200, 255, 100)
      Case "random", "don't care", "dumb luck"
        ColourNamed = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
      Case Else ' for programmers only; end-user's should never see this
        MsgBox "Specified ColourNamed " & DQ & ColStr & DQ & " does not exist or is misspelled.", , "ColourNamed Error"
        End     ' and this is why
    End Select

End Function

Public Sub ColourRemove(Optional ForeBack As Boolean)

  'Copyright 2002 Roger Gilchrist
  'This strips all occurrances of "\cf#" or "\highlight#"
  'by getting a maximim possible # and counting down
  'Counting up misfires as "\cf1" would convert "\cf10" to "0"
  'leaving a visible character "0" in the text

  Dim RTFHead As String 'hold the part of the selection which is not changed the colortbl is changed but the RTF engine does this, so you don't have to treat it programmatically
  Dim Parts As Variant 'hold and manipulate the part of the selection which can be changed
  Dim i As Long ', tmp As String
  Dim TMp As String ' convert Parts to string and back for ease of manipulation by the SafeJoinTest
  Dim Target As String ' hold the type of colour to remove
  Dim NumberPart As Long 'current number value for test
  Dim InitNumberParts As Long 'Maximum possible number value because SelRTf is renumbered as a subset of TextRTF

    PositionStore Push '
    SeparateRTFString RTFHead, Parts
    Target$ = IIf(ForeBack, "\cf", "\highlight")
    InitNumberParts = CountOccurances(Join(Parts), Target) * 4 'allows for back and fore colour contributing to colortbl
    If InitNumberParts = 0 Then
        Exit Sub '>---> Bottom
    End If

    For i = LBound(Parts) To UBound(Parts)
        NumberPart = InitNumberParts
        TMp = Parts(i)
        Do While InStr(TMp, Target$)
            TMp$ = Replace(TMp$, Target & NumberPart, Target & "0")
            NumberPart = NumberPart - 1
            If NumberPart < 0 Then
                Exit Do '>---> Loop
            End If
        Loop
        Parts(i) = SafeJoinTest(Parts(i), TMp)
    Next i
    WorkingStringWrite Trim$(RTFHead) & SafeJoin(Parts)
    PositionStore pop

End Sub

Public Sub ColourRemoveAll()

  'This is just a wrapper for two calls to ColourRemove
  'to hit bot fore and back colours

    ColourRemove True
    ColourRemove False

End Sub

Public Sub ColourTemplate(Optional Inout As Boolean = False, Optional LeftRight As Boolean = True, Optional ForeBack As Boolean = True)

  'POSSIBLE INPUTS
  'ForeBack  = Text or Background colour  ALWAYS USE It is not usually used in your colour modification but tells ColourApplication where to work
  'LeftRight = if colour array is ordered use this to reverse the order
  'LightDark = same as LeftRight but more descriptive if the order is basically intensity rather than colour
  'Inout     = use with any LeftRight or LightDark set to create a L-R-L or L-D-L pattern rather than the standard L-R or L-D pattern

  Dim i As Long ' used to step through Selection
  Dim SPos As Long, EPos As Long 'Range of Text and ColourArray
  Dim SelLen As Long 'distance between SPos and EPos use to manipulate colour
  Dim ColourArray() As Long 'Store colours for passing to ColourApplicator

    'ALL TEMPLATES NEED To Get Selection Start and End
    If GetStartEnd(SPos, EPos) = False Then 'Use at start of any modifications you build;
        Exit Sub                            'sets start and end and Exits if no Selection '>---> Bottom
    End If
    '
    ReDim ColourArray(SPos To EPos) As Long ' initialise ColourArray
    'Use (Spos To Epos) style as that is what ColourApplicator expects
    'Initialize any values your ColourStyle requires
    SelLen = EPos - SPos

    For i = SPos To EPos

        'ColourArray(i)= get a colour

    Next i

    If LeftRight = False Then ' or LightDark invert spectrum
        ColourArray = InvertArray(ColourArray)
    End If

    If Inout Then 'Inout spectrum 'Always do this after inversion
        ColourArray = InOutArray(ColourArray)
    End If

    ColourApplicator ColourArray, ForeBack

End Sub

Public Function ColourUser() As Long

  'creator(?) Roger Gilchrist
  'don't know where I got this, may have worked it out myself,
  'have been using it for ages when I need user's color input
  'This  has been left Public so that the programmers can access
  'to get colours elsewhere in their programs
  'left Public so you can use it elsewhere in your program
  'change to Private if you don't want it outside this class

    On Error GoTo error_cancel
    With m_CDlg
        '.DialogTitle = 'NO! It would be nice to set it to
        ' "Highlight Colour|Font Colour"
        ' so of course it is not supported
        'you'll just have to make it clear in documentation
        'USE THESE IF YOU WISH TO ENHANCE/LIMIT UserColour
        '        .Flags
        '               cdlCCFullOpen           &H2 Entire dialog box is displayed, including the Define Custom Colors section.
        '               cdlCCHelpButton         &H8 Causes the dialog box to display a Help button.
        '               cdlCCPreventFullOpen    &H4 Disables the Define Custom Colors command button and prevents the user from defining custom colors.
        '               cdlCCRGBInit            &H1 Sets the initial color value for the dialog box.
        '.Flags = cdlCCRGBInit
        .CancelError = True
        .ShowColor
        ColourUser = .Color
    End With 'CDLG1'M_CDLG
error_cancel:
    'This error trap just falls through returning VbBlack, 0, as user colour

End Function

Private Function CountOccurances(x, Item$) As Long

  ' not mine but don't remember where I got it
  'this is a cheap and nasty way to count sub strings

    If Len(x) Then ' remove this test if you would like a return of -1 for zero len strings
        CountOccurances = UBound(Split(x, Item))
    End If

End Function

Private Sub DeleteRTFCodeX(Part$, CS$)

  'Copyright 2002 Roger Gilchrist
  'slightly different version of previous routine
  ' don't have time/inclination to work out which is better
  'both are used in different places

  Dim Parts2 As Variant, i As Integer, tpart As String, bt As String

    Parts2 = Split(Part$, "\")
    For i = LBound(Parts2) To UBound(Parts2)
        bt$ = Parts2(i)
        ' guards against short matching start of long
        If Left$("\" & bt$ & " ", Len(CS$) + 1) = CS$ & " " Then
            bt$ = Replace(bt$, Mid$(CS$, 2), "")
            Parts2(i) = SafeJoinTest(Parts2(i), bt$)
        End If
    Next i
    Part$ = SafeJoin(Parts2, "\")

End Sub

Public Sub Dither(Optional LightDark As Boolean = True, Optional Inout As Boolean = True, Optional ForeBack As Boolean = True)

  Dim i As Long
  Dim R As Integer, G As Integer, B As Integer
  Dim SPos As Long, EPos As Long 'Range of Text and hence BlendArray
  Dim Dither() As Long
  Dim ditherSection As Integer
  Dim DitherStep As Long

    If GetStartEnd(SPos, EPos) = False Then 'Use at start of any modifications you build;
        Exit Sub                            'sets start and end and Exits if no Selection '>---> Bottom
    End If
    ReDim Dither(SPos To EPos) As Long ' initialise array

    DitherStep = (EPos - SPos) / 255
    DitherStep = IIf(DitherStep = 0, 1, DitherStep)
    m_LastDitherColor = ColourUser
    ColourLong2RGB m_LastDitherColor, R, G, B

    For i = SPos To EPos
        If i Mod DitherStep = 0 Then
            ditherSection = ditherSection + 1
        End If
        Dither(i) = RGB(IIf(R - ditherSection > 0, R - ditherSection, 0), IIf(G - ditherSection > 0, G - ditherSection, 0), IIf(B - ditherSection > 0, B - ditherSection, 0))
    Next i
    If LightDark = False Then
        Dither = InvertArray(Dither)
    End If
    If Inout Then
        Dither = InOutArray(Dither)
    End If
    ColourApplicator Dither, ForeBack

End Sub

Public Sub DitherHard(DColr As Long, Optional LightDark As Boolean = True, Optional Inout As Boolean = True, Optional ForeBack As Boolean = True)

  Dim i As Long
  Dim R As Integer, G As Integer, B As Integer
  Dim SPos As Long, EPos As Long 'Range of Text and hence BlendArray
  Dim Dither() As Long
  Dim ditherSection As Integer
  Dim DitherStep As Long

    If GetStartEnd(SPos, EPos) = False Then 'Use at start of any modifications you build;
        Exit Sub                            'sets start and end and Exits if no Selection '>---> Bottom
    End If
    ReDim Dither(SPos To EPos) As Long ' initialise array

    DitherStep = (EPos - SPos) / 255
    DitherStep = IIf(DitherStep = 0, 1, DitherStep)
    ColourLong2RGB DColr, R, G, B

    For i = SPos To EPos
        If i Mod DitherStep = 0 Then
            ditherSection = ditherSection + 1
        End If
        Dither(i) = RGB(IIf(R - ditherSection > 0, R - ditherSection, 0), IIf(G - ditherSection > 0, G - ditherSection, 0), IIf(B - ditherSection > 0, B - ditherSection, 0))
    Next i
    If LightDark = False Then
        Dither = InvertArray(Dither)
    End If
    If Inout Then
        Dither = InOutArray(Dither)
    End If
    ColourApplicator Dither, ForeBack

End Sub

Public Function ExcessSpaceDelete() As Boolean

  'Copyright 2002 Roger Gilchrist
  'Remove extra spacing
  '*---PROGRAMMER MODIFICATION POINT---*
  'if you don't want it delete

  Dim RTFHead As String     'Initial section of RTF probably never useful
  Dim Parts As Variant      'array for searching
  Dim i As Integer          'For variable to manipulate Parts

    With m_RTB
        PositionStore Push
        SeparateRTFString RTFHead, Parts
        'Mark excess spacesfor deletion
        For i = LBound(Parts) + 1 To UBound(Parts)
            If Parts(i) = "" Then
                If RTFPartType(CStr(Parts(i - 1))) = TextOnly Then
                    Parts(i) = SafeJoinSkip
                End If
            End If
        Next i
        WorkingStringWrite RTFHead & SafeJoin(Parts, " ")
        PositionStore pop
    End With 'M_RTB

End Function

Private Function ExtractFontSize(str$) As Long

  'Copyright 2002 Roger Gilchrist

  Dim fsLoc As Long, RTFend As Long

    fsLoc = InStr(str$, "\fs")       '.SelRTF has a standard format in which \fsN comes last
    If fsLoc Then
        ExtractFontSize = Mid$(str$, fsLoc + 3)
      Else 'FSLOC = FALSE
        PositionStore Push
        With m_RTB
            .SelStart = 1
            .SelLength = 1
            ExtractFontSize = .SelFontSize
        End With 'M_RTB
        PositionStore pop
    End If

End Function

Private Sub FontApplicator(FontArray)

  'Copyright 2002 Roger Gilchrist
  ' this is the core of the various font look routines
  '*---PROGRAMMER MODIFICATION POINT---*
  'work in progress
  'Does not handle fonttbl yet

  'fontArray format RTFfontcode & "***"
  'FontApplicator replaces "***" with an individual character

  'this is just a crude redit of ColourApplicator at present

  Dim RTFHead As String 'hold the part of the selection which is not changed(unless it contains a colortbl)
  Dim FntTbl As String 'Extract fonttbl from RTFHead if it exists or fake it otherwise
  Dim OrigFntTbl As String 'store the fonttbl so that it can be Replaced
  Dim Parts As Variant 'hold and manipulate the part of the selection which can be changed
  Dim CurFont As Long  'keep track of position in FontArray
  Dim i As Long        'For Next through the Selection.

    PositionStore Push '
    '    FntTbl = GetFontTable()    '
    '    OrigFntTbl = FntTbl         'store for later Replace call
    '    If Len(FntTbl) = 0 Then ' if No colorTbl create a new one with auto first value ";"
    '        FntTbl = "{\fonttbl{;"
    '      Else 'else strip the end off the existing one'NOT LEN(FntTbl)...
    '        FntTbl = Left$(FntTbl, Len(FntTbl) - 3)
    '    End If
    SeparateRTFString RTFHead, Parts
    CurFont = LBound(FontArray) 'get first colour from array
    For i = LBound(Parts) To UBound(Parts)  ')

        FontManipulate Parts(i), FntTbl, FontArray, CurFont
    Next i
    '''    FntTbl = FntTbl & "}  " 'finish colortbl
    '''    If InStr(RTFHead, "{\fonttbl") Then 'Replace Original colortbl with new one
    '''        RTFHead = Replace(RTFHead, OrigFntTbl, FntTbl)
    '''      Else ' if there was no original fonttbl add it and another "}" to end it'NOT INSTR(RTFHEAD,...
    '''        RTFHead = Left$(RTFHead, Len(RTFHead) - 1) & FntTbl & "}"
    '''    End If
    WorkingStringWrite Trim$(RTFHead) & Join(Parts)
    PositionStore pop

End Sub

Private Sub FontManipulate(Part, FntTbl$, FntArray, CurFont As Long)

  'Copyright 2002 Roger Gilchrist
  'does the actual work of inserting colour codes into colortbl and rtf code
  '*---PROGRAMMER MODIFICATION POINT---*
  'work in progress
  'Does not handle fonttbl yet

  Dim FntCount As Long 'get current number of colors in colortbl
  Dim i As Long ' loop driver' probably doesn't need to be a long but its safer
  Dim tmpP As String 'bulid new string to replace Part

    'FntCount = CountOccurances(clrTbl, "\fcharset") 'get length of fontTbl
    Select Case RTFPartType(CStr(Part))
      Case blank 'blanks have to absorb an array member
        tmpP$ = Replace(FntArray(CurFont), "***", " ") 'tmpP$ & "\fn" & FntCount + 1 & " "
        'FntTbl$ = FntTbl$ & ColorRTFCode(CurFont, FntArray(CurFont))
        CurFont = CurFont + 1
      Case RTFOnly
        tmpP = Part
      Case Mixed, TextOnly
        'add a colour to each letter and advance colour counter for text letter but ignore RTF code
        For i = 1 To Len(Part)
            If InStr("\}", Mid$(Part, i, 1)) = 0 Then
                tmpP$ = tmpP$ & Replace(FntArray(CurFont), "***", Mid$(Part, i, 1))
                'FntTbl$ = FntTbl$ & FontRTFCode(, curFontFntArray(CurFont))
                CurFont = CurFont + 1
              Else 'NOT INSTR("\}",...
                'becuase of the way Part is generated once you reach an RTF code token there is not more text
                tmpP$ = tmpP$ & Mid$(Part, i)
                Exit For '>---> Next
            End If
        Next i
    End Select
    Part = tmpP

End Sub

Private Function FontRTFCode(CurFont As Long, FntName) As String

  '*---PROGRAMMER MODIFICATION POINT---*
  'work in porgress do not use
  'Copyright 2002 Roger Gilchrist
  'creates an RTF colour code member
  'support for ColourManipulation routine
  'NEEDS WORK TO FILL FONT PROPERTIES PROPERLYVVVVVVVVVV

    FontRTFCode = "\f" & CurFont & "\froman\fprq2\fcharset0 " & FntName & ";"

End Function

Private Function FontSize1stCharacter() As Long

  Dim RTFHead As String, Parts As Variant

    SeparateRTFString RTFHead, Parts
    FontSize1stCharacter = ExtractFontSize(CStr(Parts(0)))

End Function

Public Sub FormatRemove()

  'Copyright 2002 Roger Gilchrist
  'This strips all occurrances of "\cf#" or "\highlight#"
  'by getting a maximim possible # and counting down
  'Counting up misfires as "\cf1" would convert "\cf10" to "0"
  'leaving a visible character "0" in the text

  Dim RTFHead As String 'hold the part of the selection which is not changed the colortbl is changed but the RTF engine does this, so you don't have to treat it programmatically
  Dim Parts As Variant 'hold and manipulate the part of the selection which can be changed
  Dim i As Long ', tmp As String
  Dim TMp As String ' convert Parts to string and back for ease of manipulation by the SafeJoinTest
  Dim Target As String ' hold the type of colour to remove
  Dim NumberPart As Long 'current number value for test
  Dim InitNumberParts As Long 'Maximum possible number value because SelRTf is renumbered as a subset of TextRTF
  Dim Targets As Variant
  Dim tnum As Integer

    Targets = Array("\up", "\dn", "\sub", "\super")
    PositionStore Push '
    Do While tnum < UBound(Targets)

        SeparateRTFString RTFHead, Parts
        Target$ = Targets(tnum)
        InitNumberParts = CountOccurances(Join(Parts), Target) * 4 'allows for back and fore colour contributing to colortbl
        If InitNumberParts = 0 Then
            Exit Sub '>---> Bottom
        End If

        For i = LBound(Parts) To UBound(Parts)
            NumberPart = InitNumberParts
            TMp = Parts(i)
            Do While InStr(TMp, Target$)
                TMp$ = Replace(TMp$, Target & NumberPart, Target & "0")
                TMp$ = Replace(TMp$, Target & "0", "")
                NumberPart = NumberPart - 1
                If NumberPart < 0 Then
                    Exit Do '>---> Loop
                End If
            Loop
            Parts(i) = SafeJoinTest(Parts(i), TMp)
        Next i
        WorkingStringWrite Trim$(RTFHead) & SafeJoin(Parts)
        tnum = tnum + 1
    Loop
    PositionStore pop

End Sub

Private Function GetColorTable(Optional ForceDoc As Boolean = False) As String

  'Copyright 2002 Roger Gilchrist
  'Specialised version of GetFormatString
  'used by ColoursUsed and ColourApplicator
  'returns the colortbl from either the selection or whole document(ForcedDoc=True)

  Dim ClrTblStart As Long ' Cutpoints for various pieces
  Dim ClrTblEnd As Long
  Dim RTFTail As String

    RTFTail$ = Left$(WorkingStringRead(ForceDoc), m_End_OF_RTF)
    ClrTblStart = InStr(RTFTail$, "{\colortbl")      'gets end of fonttbl
    If ClrTblStart = 0 Then                           'blank documents use default color table which is not present
        GetColorTable = ""
        Exit Function '>---> Bottom
      Else 'NOT CLRTBLSTART...
        'find certain end of colortbl note a multiple color table has ";}}"
        ClrTblEnd = InStr(ClrTblStart, RTFTail$, ";}")
        'so distance to next lump of code varies so get certain start at next "\"
        ClrTblEnd = InStr(ClrTblEnd, RTFTail, "\") - ClrTblStart
        GetColorTable = Mid$(RTFTail$, ClrTblStart, ClrTblEnd)
    End If

End Function

Private Function GetFontTable(Optional ForceDoc As Boolean = False) As String

  'Copyright 2002 Roger Gilchrist
  'Specialised version of GetFormatString
  'used by ColoursUsed and ColourApplicator
  'returns the colortbl from either the selection or whole document(ForcedDoc=True)

  Dim FntTblStart As Long ' Cutpoints for various pieces
  Dim FntTblEnd As Long
  Dim RTFTail As String

    RTFTail$ = Left$(WorkingStringRead(ForceDoc), m_End_OF_RTF)
    FntTblStart = InStr(RTFTail$, "{\fonttbl")      'gets end of fonttbl
    If FntTblStart = 0 Then                           'blank documents use default color table which is not present
        GetFontTable = ""
        Exit Function '>---> Bottom
      Else 'NOT CLRTBLSTART...'NOT FNTTBLSTART...
        'find certain end of colortbl note a multiple color table has ";}}"
        FntTblEnd = InStr(FntTblStart, RTFTail$, ";}")
        'so distance to next lump of code varies so get certain start at next "\"
        FntTblEnd = InStr(FntTblEnd, RTFTail, "\") - FntTblStart
        GetFontTable = Mid$(RTFTail$, FntTblStart, FntTblEnd)
    End If

End Function

Private Function GetGetValue(RTFCode$) As Variant

  'Copyright 2002 Roger Gilchrist
  'support for the Get SelXXXX properties
  '*---PROGRAMMER MODIFICATION POINT---*
  'this is provisional until I find/work out how to make it behave properly

    With m_RTB
        If .SelLength Then
            GetGetValue = Not (RTFCodeLoc(.SelRTF, RTFCode$) <> NotPresent) ' Code is present so remove
          Else '.SELLENGTH = FALSE
            GetGetValue = HasRTFCode(RTFCode$)
        End If
    End With 'M_RTB

End Function

Private Sub GetRGB(ByVal col As Long, Optional ByRef Red As Integer, Optional ByRef Green As Integer, Optional ByRef Blue As Integer)

  'don't know where I got this but there are a million variations around

    Red = col Mod 256
    Green = ((col And &HFF00FF00) / 256)
    Blue = (col And &HFF0000) / (65536)

End Sub

Private Function GetStartEnd(SPos As Long, EPos As Long) As Boolean

  'Support for Colour routines
  'call before anything else
  'PROGRAMMER WARNING

    With m_RTB
        SPos = .SelStart
        EPos = .SelLength
        EPos = SPos + EPos - 1
    End With 'M_RTB
    GetStartEnd = EPos > SPos
    'this lets this routine serve as a do nothing signal if there is no selection

End Function

Private Function HasRTFCode(RTFCode$) As Boolean

  'Copyright 2002 Roger Gilchrist
  'Use this if the RTF tag you are interested in does not
  'have a numerical component OR numeric component is known
  'WARNING WARNING WARNING WARNING WARNING WARNING WARNING WARNING WARNING
  'WARNING WARNING *---PROGRAMMER MODIFICATION POINT---*  WARNING WARNING
  'WARNING WARNING WARNING WARNING WARNING WARNING WARNING WARNING WARNING
  'THIS IS REALLY slow don't use it or anything that calls it in real world work
  'I 'm hoping to find a fast way to do this using API but only partial success so far

    With m_RTB
        HasRTFCode = InStr(.TextRTF, RTFCode$ & "\")
        If HasRTFCode = False Then
            HasRTFCode = InStr(.TextRTF, RTFCode$ & " ")
        End If
    End With 'M_RTB
    '                     End if RTF Code group                    embedded in RTF Code group

End Function

Private Function InOutArray(sArray) As Variant

  Dim TmpA() As Long
  Dim i As Long
  Dim LocalCounter As Long
  Dim TopOfArray As Long

    TopOfArray = LBound(sArray) + UBound(sArray)
    ReDim TmpA(LBound(sArray) To UBound(sArray))
    LocalCounter = LBound(sArray)

    For i = LBound(sArray) To UBound(sArray) Step 2 ' step through array at double speed
        TmpA(LocalCounter) = sArray(i)              'Add to bottom and top of TmpA
        TmpA(TopOfArray - LocalCounter) = sArray(i)
        LocalCounter = LocalCounter + 1
    Next i
    InOutArray = TmpA

End Function

Private Function InvertArray(sArray) As Variant

  'support code for Dither
  'use anywhere you want to reverse a colour array

  Dim TmpA() As Long
  Dim i As Long, LocalCounter As Long

    ReDim TmpA(LBound(sArray) To UBound(sArray))
    LocalCounter = UBound(sArray)
    For i = LBound(sArray) To UBound(sArray)
        TmpA(LocalCounter) = sArray(i)
        LocalCounter = LocalCounter - 1
    Next i
    InvertArray = TmpA

End Function

Private Function InvertColor(col As Long) As Long

  Dim R As Integer, G As Integer, B As Integer, GREYTEST As Long

    GetRGB col, R, G, B
    'the inversion is mine
    ' the different test should keep you from getting unusable colours
    ' if RGB = 128,128,128 then the usual invert will just return to itself
    'only if R,G & B are close in value (after all (0,0,255) Dark Blue is fine but would trigger the GREYTEST
    If Abs(R - G) < 10 And Abs(G - B) < 10 And Abs(R - B) < 10 Then
        'if RGB are close they are usually some sort of grey
        GREYTEST = (R + G + B) / 3
        'and the average is near dark gray
        If GREYTEST >= 120 And GREYTEST <= 135 Then
            R = R + 128 ' shift the colour half the RGB spectrum away
            B = B + 128
            G = G + 128
        End If
    End If
    InvertColor = RGB(Abs(R - 255), Abs(G - 255), Abs(B - 255))
    If InvertColor = col Then ' just a bit of paranoia
        InvertColor = RGB(Abs(R - 128), Abs(G - 128), Abs(B - 128))
    End If

End Function

Public Property Get LastBlendEndColour() As Long

  'These properties are mostly for the TextLookPanel's Preserve Selection Colour System

    LastBlendEndColour = m_LastBlendEndColour

End Property

Public Property Get LastBlendStartColour() As Long

  'These properties are mostly for the TextLookPanel's Preserve Selection Colour System

    LastBlendStartColour = m_LastBlendStartColour

End Property

Public Property Get LastDitherColor() As Long

  'These properties are mostly for the TextLookPanel's Preserve Selection Colour System

    LastDitherColor = m_LastDitherColor

End Property

Public Property Get LastHighlightColour() As Long

  'These properties are mostly for the TextLookPanel's Preserve Selection Colour System
  'and only apply to the RTFHighlight methods

    LastHighlightColour = m_LastHighlightColour

End Property

Public Property Get LastHighlightForeColour() As Long

  'These properties are mostly for the TextLookPanel's Preserve Selection Colour System
  'and only apply to the RTFHighlight methods

    LastHighlightForeColour = m_LastHighlightForeColour

End Property

Private Sub ManipulateRTFCode(Part$, CS$, Mode As Boolean)

    If Mode = False Then
        Select Case RTFPartType(CStr(Part))
          Case RTFOnly, Mixed ' only these need be treated
            DeleteRTFCodeX Part$, CS$
        End Select
      Else 'NOT MODE...
        Select Case RTFPartType(CStr(Part))
          Case RTFOnly
            Part = Part & CS$ ' add to end of Pure RTF it will override any counter or rival code
          Case Else
            Part = CS$ & " " & Part ' add to start of mixed or textonly so current char is affected.
        End Select
    End If

End Sub

Private Sub MaterialGradientFill(CArray() As Long, Min As Integer, Max As Integer, R As Integer, G As Integer, B As Integer, Optional Multiplier As Double = 1)

  'based on code originally written by
  'John Colman (john_colman@hotmail.com) in metalCBprj.OCX

  Dim i As Long
  Dim C As Integer
  Dim st As Double
  Dim counter As Long

    On Error Resume Next
    MaterialsSafeValues Min, Max, R, G, B, Multiplier
        Max = Max - Min
        st = (UBound(CArray) - LBound(CArray)) / 3.142 * Abs(Multiplier) * 2
        ' MaterialSafeValues, the * 2 and Counter are my only modification of the routine
        '*2 as I only need the 1st half of the material spectrum
        ' the InOutArray call takes care of full spectrums
        ' the Counter because LBound and UBound are absolute not relative values
        'and this routine needs relative to start and end math
        '   If st = 0 Then st = 1E-25'DEBUG LINE NOT NEEDED
        If Multiplier > 0 Then
            For i = LBound(CArray) To UBound(CArray)
                C = Abs(Max * Sin(counter / st)) + Min
                counter = counter + 1
                CArray(i) = RGB(R + C, G + C, B + C)
            Next i
          Else 'NOT MULTIPLIER...
            For i = LBound(CArray) To UBound(CArray)
                C = Abs(Max * Cos(counter / st)) + Min
                counter = counter + 1
                CArray(i) = RGB(R + C, G + C, B + C)
            Next i
        End If
    On Error GoTo 0

End Sub

Public Sub Materials(Clr As String, Optional Inout As Boolean = False, Optional LightDark As Boolean = True, Optional ForeBack As Boolean = True)

  'based on code originally written by
  'John Colman (john_colman@hotmail.com)in metalCBprj.OCX
  'I liked the subtle colours so i modified it from buttons to text
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if not needed

  'John Colman worked out these colours and used them in his project
  'if you create nice new ones let me know and I'll add them to future versions

  'Name         min,    max,    r , g , b ,     multiplier
  '-----------|------|-------|----|----|------|---------------
  'blue steel   0,      255,    0,  33, 99,      1.3
  'diamond      0,      255,    100,100,255,     1.3
  'gold         0,      255,    70, 30, 0,       1.3
  'ice          100,    255,    0,  40, 60,      3
  'lead         0,      100,    33, 33, 33,      1.5
  'rubber       0,      100,    100,33, 33,      1.5
  'silver       0,      255,    33, 33, 33,      1.3
  'slate        0,      100,    33, 33, 100,     1.3

  'these are some I made
  'Name             min,    max,    r , g , b ,     multiplier
  '-----------|----------|-------|----|----|------|-----------
  'yellowrose ,     100,    255,   54,  24, 0,     1.3
  'flesh1           100,    255,   44,  4,  0,     1.3
  'milk chocolate   0,     100,    95,  55, 45,    1.5
  'dark chocolate   0,      65,    95,  55, 45,    1.5
  'pineboard        100,   165,    100, 85, 35,    1.5
  'oldgold          100,    165,   180, 85, 35,    1.3

  'feel free to add to them just send me this line if you want
  'me to add to the distrubuted code

  'Name             min,    max,    r , g , b ,     multiplier
  '-----------|----------|-------|----|----|------|-----------

  'I may modify the name so it does not clash with other offerings
  'There is a 'test' name provided and already programmed into the ColourPanel
  'so you can experiment more quickly
  ' if you look at the next routine you will see that I have started working on
  'a routine which allows you to develop new materials on the run but it is not
  'in use yet
  
  Dim i As Long, SPos As Long, EPos As Long
  Dim tmpC() As Long

    If GetStartEnd(SPos, EPos) = False Then 'Use at start of any modifications you build;
        Exit Sub                            'sets start and end and Exits if no Selection '>---> Bottom
    End If
    'EPos = SPos + EPos - 1

    ReDim tmpC(SPos To EPos) As Long
    Select Case Clr
      Case "blue steel"
        MaterialGradientFill tmpC, 0, 255, 0, 33, 99, 1.3
      Case "diamond"
        MaterialGradientFill tmpC, 0, 255, 100, 100, 255, 1.3
      Case "gold"
        MaterialGradientFill tmpC, 0, 255, 70, 30, 0, 1.3
      Case "ice"
        MaterialGradientFill tmpC, 100, 255, 0, 40, 60, 3
      Case "lead"
        MaterialGradientFill tmpC, 0, 100, 33, 33, 33, 1.5
      Case "rubber"
        MaterialGradientFill tmpC, 0, 100, 100, 33, 33, 1.5
      Case "silver"
        MaterialGradientFill tmpC, 0, 255, 33, 33, 33, 1.3
      Case "slate"
        MaterialGradientFill tmpC, 0, 100, 33, 33, 100, 1.3
      Case "flesh1"
        MaterialGradientFill tmpC, 100, 255, 44, 4, 0, 1.3
      Case "yellowrose"
        MaterialGradientFill tmpC, 100, 255, 54, 24, 0, 1.3
      Case "milkchocolate"
        MaterialGradientFill tmpC, 0, 100, 95, 55, 45, 1.5
      Case "pineboard"
        MaterialGradientFill tmpC, 100, 165, 100, 85, 35, 1.5
      Case "oldgold"
        MaterialGradientFill tmpC, 100, 165, 180, 85, 35, 1.3 '1.5
      Case "test" ' this is here for you to experiment with
        '               min,max,  r , g , b , multiplier
        MaterialGradientFill tmpC, 100, 165, 180, 85, 35, 10 '1.3 '1.5

    End Select

    If LightDark = False Then ' invert spectrum
        tmpC = InvertArray(tmpC)
    End If

    If Inout Then
        tmpC = InOutArray(tmpC)
    End If

    ColourApplicator tmpC, ForeBack

End Sub

Public Sub MaterialsTest(MaterialName$, Inout As Boolean, LightDark As Boolean, ForeBack As Boolean, Min As Integer, Max As Integer, R As Integer, G As Integer, B As Integer, Multiplier As Double)

  Dim i As Long, SPos As Long, EPos As Long
  Dim tmpC() As Long

    If GetStartEnd(SPos, EPos) = False Then 'Use at start of any modifications you build;
        Exit Sub                            'sets start and end and Exits if no Selection '>---> Bottom
    End If
'MaterialName$ supplied so you can save it some where NOT IMPLIMENETED YET
    ReDim tmpC(SPos To EPos) As Long
    KeepInBounds 0, Min, Max  'saftey tests
    MaterialGradientFill tmpC, Min, Max, R, G, B, Multiplier
    If LightDark = False Then ' invert spectrum
        tmpC = InvertArray(tmpC)
    End If
    If Inout Then
        tmpC = InOutArray(tmpC)
    End If
    ColourApplicator tmpC, ForeBack

End Sub
Sub MaterialsSafeValues(Min As Integer, Max As Integer, R As Integer, G As Integer, B As Integer, Multiplier As Double)
'Keep all values sent to MaterialsGradientFill in bounds
    KeepInBounds Min, Max, 255
    KeepInBounds 0, R, 255
    KeepInBounds 0, G, 255
    KeepInBounds 0, B, 255
    'Multiplier can be anything you like -ive, +ive BUT NOT 0
    If Multiplier = 0 Then
    Multiplier = 0.00001 'this is a safe minimum value you could try smaller ones
    'there is no upper level but the effect is less prominant as you go up not very interesting
    End If

End Sub
Private Sub KeepInBounds(Lo, Value, Hi)

  'copyright 1988 Roger Gilchrist
  'Modified from slower all Variant version
  'Enforce minimum and maximum limits to Value
  'You could make this a function but it looks messier in code: Val = KeepinBounds(lo, Val, hi)

    Value = IIf(Value < Lo, Lo, Value)
    Value = IIf(Value > Hi, Hi, Value)

End Sub

Public Sub NoFormatting()

  'Copyright 2002 Roger Gilchrist
  'this strips all formatting in three slightly different ways
  '1.A Selection: Selection is reset to that of the character immediately before selection
  '               (if that is a whitespace the actual format may be visible on last character but could have been aplied just to the whitespace)
  '2.Whole Doc Selection: would reset to first charcter format (Probably not a good idea, I tried it on the Demo document, very ugly)
  '                       but has been programmed to deselect and use the No selection approach
  '                       If you want to duplicate this ugliness: Select the whole document, used Shift+LeftArrow to de-select last character, then call NoFormatting
  '3.No selection : Whole document is reset to RichTextBox default (You can set this from the IDE or accept VB defaults)

  '*---PROGRAMMER MODIFICATION POINT---*
  'if you would prefer the Ugly version of WholeDoc Selection(see above)
  'delete the 'If .SelLength = Len(.Text) Then' 'End If' structure

  Dim RTFHead As String     'Initial section of RTF probably never touched
  Dim Parts As Variant      'array for searching
  Dim WholeDocSwitch As Boolean

    With m_RTB
        PositionStore Push
        If .SelLength >= Len(.Text) Then ' deselect if it s whole doc to avoid ugliness
            .SelLength = 0
        End If
        If .SelLength Then
            .SelRTF = .SelText
          Else '.SELLENGTH = FALSE
            .Locked = True
            .SelStart = 1
            .SelLength = Len(.Text)
            .SelRTF = .SelText ' set whole text to first character
            .Locked = False
            SeparateRTFString RTFHead, Parts
            Parts(0) = m_DefFormat ' enforce def format
            WorkingStringWrite RTFHead & SafeJoin(Parts, " ")
        End If
        PositionStore pop
    End With 'M_RTB

End Sub

Private Function Percent(a, B) As Single

  'calculate percentage given two numbers

  Dim C As Single

    '    If a > B Then a = B
    If B = 0 Then
        Percent = 0
        Exit Function '>---> Bottom
    End If
    C = a / B * 100
    Percent = IIf(C > 100, 100, C)

End Function

Private Sub PositionStore(SaveTSetF As LIFO)

  'Simple PushPop store of SelStart and SelLength allows routines
  'which manipulate through .seXXXXXX to restore SelStart and SelLength
  'arrays allow nesting of calls as long as you always call push and pop in calling routines

  '1. MOST IMPORTANT it sets the m_busy variable which stops a lot of updating while involved
  'in manipulating .SelXXX stuff

  '2. Set Mouse to Hourglass if m_busy = True
  '*---PROGRAMMER MODIFICATION POINT---*
  'modify/disable if you have your own mouse cursor routines to use

  '*---PROGRAMMER MODIFICATION POINT---*
  'If you edit this make sure you leave the m_busy stuff even if you delete everything else
  'The rest of this routine is only really needed if a routine
  'changes .SelStart or .SelLength to carry out its task
  'None of my colour routines actually do this but it is useful if you are trying to develop
  'something from scratch before going to TextRTF manipulation.

  Dim insertPos As Long

    'keep this bit no matter what
    With m_RTB
        If SaveTSetF Then
            m_busy = True
            .MousePointer = rtfHourglass
          Else 'SAVETSETF = FALSE
            .MousePointer = rtfDefault
            m_busy = False
        End If
    End With 'M_RTB

    On Error GoTo 0
    With m_RTB
        If SaveTSetF Then
            On Error Resume Next
                If IsArray(m_Start) Then 'initialize
                    insertPos = 0
                  Else 'ISNULL(M_START) = FALSE'ISARRAY(M_START) = FALSE
                    insertPos = UBound(m_Start)
                End If
                ReDim Preserve m_Start(insertPos + 1) As Long
                ReDim Preserve m_Len(insertPos + 1) As Long
                m_Start(UBound(m_Start)) = .SelStart
                m_Len(UBound(m_Len)) = .SelLength
              Else 'SAVETSETF = FALSE

                .SelStart = m_Start(UBound(m_Start))
                .SelLength = m_Len(UBound(m_Len))
                If UBound(m_Start) > 0 Then
                    ReDim Preserve m_Start(UBound(m_Start) - 1) As Long
                    ReDim Preserve m_Len(UBound(m_Len) - 1) As Long
                  Else 'NOT UBOUND(M_START)...
                    ReDim Preserve m_Start(0) As Long
                    ReDim Preserve m_Len(0) As Long
                End If
            End If
        End With 'M_RTB
    On Error GoTo 0

End Sub

Public Sub RainBow(Optional RedMagenta As Boolean = True, Optional ForeBack As Boolean = True)

  'Copyright 2002 Roger Gilchrist
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if not needed

  Dim i As Long, SPos As Long, EPos As Long
  Dim RbowMod As Long, Rsection As Integer, Rstep As Long
  Dim Cycler As Long, RainBowArray() As Long, drv As Long
  Dim TestA As Integer, TestB As Integer

    If GetStartEnd(SPos, EPos) = False Then 'Use at start of any modifications you build;
        Exit Sub                            'sets start and end and Exits if no Selection '>---> Bottom
    End If
    RbowMod = (EPos - SPos) \ 6            'this is used to change spectrum sections
    If RbowMod < 1 Then
        RbowMod = 1
    End If
    ReDim RainBowArray(SPos To EPos) As Long ' initialise array
    TestB = 0
    For i = SPos To EPos
        Rstep = (255 * Cycler / RbowMod)
        TestA = Int(Percent(i - SPos, EPos - SPos) / 16.6666666666667)
        If TestA <> TestB Then
            TestB = TestB + 1
            Rstep = 0
            Cycler = 0
        End If
        Cycler = Cycler + 1 '    count through the spectrum section
        RainBowArray(i) = RainbowColor(TestA, Rstep)
    Next i
    If RedMagenta = False Then ' invert spectrum
        RainBowArray = InvertArray(RainBowArray)
    End If
    ColourApplicator RainBowArray, ForeBack

End Sub

Private Function RainbowColor(SpectrumMember As Integer, Cval As Long) As Long

  'Based on
  'DrawRainBow © oigres P
  'Email: oigres@postmaster.co.uk

    Select Case SpectrumMember
      Case s1RedYellow   '0
        RainbowColor = RGB(255, Cval, 0)
      Case s2YellowGreen '1
        RainbowColor = RGB(255 - Cval, 255, 0)
      Case s3GreenCyan   '2
        RainbowColor = RGB(0, 255, Cval)
      Case s4CyanBlue    '3
        RainbowColor = RGB(0, 255 - Cval, 255)
      Case s5BlueMagenta '4
        RainbowColor = RGB(Cval, 0, 255)
      Case s6MagentaRed  '5
        RainbowColor = RGB(255, 0, 255 - Cval)
    End Select
    'Next

End Function

Public Sub RainBowIO(Optional RedMagenta As Boolean = True, Optional Inout As Boolean = False, Optional ForeBack As Boolean = True)

  'Copyright 2002 Roger Gilchrist
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if not needed

  Dim i As Long, SPos As Long, EPos As Long
  Dim RbowMod As Long, Rsection As Integer, Rstep As Long
  Dim Cycler As Long, RainBowArray() As Long, drv As Long
  Dim TestA As Integer, TestB As Integer

    If GetStartEnd(SPos, EPos) = False Then 'Use at start of any modifications you build;
        Exit Sub                            'sets start and end and Exits if no Selection '>---> Bottom
    End If
    RbowMod = (EPos - SPos) \ 5            'this is used to change spectrum sections
    '    If RbowMod < 1 Then
    '        RbowMod = 1
    '    End If
    ReDim RainBowArray(SPos To EPos) As Long ' initialise array
    TestB = 0
    For i = SPos To EPos
        Rstep = 255 * Cycler / RbowMod
        TestA = Int(Percent(i - SPos, EPos - SPos) / 16.6666666666667)
        If TestA <> TestB Then
            TestB = TestB + 1
            Rstep = 0
            Cycler = 0
        End If
        Cycler = Cycler + 1 '    count through the spectrum section
        RainBowArray(i) = RainbowColor(TestA, Rstep)
    Next i
    If RedMagenta = False Then ' invert spectrum
        RainBowArray = InvertArray(RainBowArray)
    End If
    If Inout Then ' invert spectrum
        RainBowArray = InOutArray(RainBowArray)
    End If

    ColourApplicator RainBowArray, ForeBack

End Sub

Public Sub RandomColour(Optional ColorSet As Integer = 0, Optional ForeBack As Boolean = True)

  'Copyright 2002 Roger Gilchrist
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if not needed
  'ColorSet 0 to 7 select which source to use for getting random colors

  Dim i As Long, SPos As Long, EPos As Long
  Dim RndForArray() As Long
  Dim Gray As Integer

    If GetStartEnd(SPos, EPos) = False Then 'Use at start of any modifications you build;
        Exit Sub                            'sets start and end and Exits if no Selection '>---> Bottom
    End If

    'EPos = SPos + EPos - 1

    ReDim RndForArray(SPos To EPos) As Long
    For i = SPos To EPos
        Select Case ColorSet
          Case 0
            RndForArray(i) = ColourNamed("random")
          Case 1
            RndForArray(i) = RainbowColor(Int(Rnd * 5), Int(Rnd * 255))
          Case 2 To 7
            RndForArray(i) = RainbowColor(ColorSet - 2, Int(Rnd * 255))
          Case 8
            Gray = Int(Rnd * 255)
            RndForArray(i) = RGB(Gray, Gray, Gray)
        End Select
        'if you don't want to use ColourNamed any where else then delete it and use the next line instead
        'RndForArray(i) = ColourNamed = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
    Next i
    ColourApplicator RndForArray, ForeBack

End Sub

Private Sub RandomColour2(ColorSet As Integer, Optional RndForeBack As Boolean = False)

  'Copyright 2002 Roger Gilchrist
  'applies random auto colour to selected text
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if not needed

  Dim RLen As Long, i As Long
  Dim SPos As Long, EPos As Long
  Dim RndHi() As Long, Grzy As Integer

    If GetStartEnd(SPos, EPos) = False Then 'Use at start of any modifications you build;
        Exit Sub                            'sets start and end and Exits if no Selection '>---> Bottom
    End If
    ReDim RndHi(SPos To EPos) As Long
    For i = SPos To EPos
        If Rnd > 0.5 Then
            Select Case ColorSet
              Case 0
                RndHi(i) = ColourNamed("random")
              Case 1
                RndHi(i) = RainbowColor(Int(Rnd * 5), Int(Rnd * 255))
              Case 2 To 7
                RndHi(i) = RainbowColor(ColorSet - 2, Int(Rnd * 255))
              Case 8
                Grzy = Int(Rnd * 255)
                RndHi(i) = RGB(Gray, Gray, Gray)
            End Select
          Else 'NOT RND...
            If Not RndForeBack Then
                RndHi(i) = IIf(Rnd > 0.5, vbBlack, vbWhite) 'standard colouring
              Else 'NOT NOT...
                RndHi(i) = ColourNamed("random") ' may as well be random any way
            End If
        End If
    Next i
    ColourApplicator RndHi, True, RndForeBack

End Sub

Public Sub Ransom(Optional Colorise As Integer = 0)

  'Copyright 2002 Roger Gilchrist
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if you dont want this
  'fontArray format RTFfontcode&"***"
  'FontApplicator replaces "***" with an individual character

  Dim i As Integer
  Dim t As String, s As String, Lp As String, Rp As String
  Dim FntArray() As String
  Dim SPos As Long, EPos As Long
  Dim InitialFontSize As Long

    If GetStartEnd(SPos, EPos) = False Then 'Use at start of any modifications you build;
        Exit Sub                            'sets start and end and Exits if no Selection '>---> Bottom
    End If
    ReDim FntArray(SPos To EPos) As String
    InitialFontSize = FontSize1stCharacter
    For i = SPos To EPos
        FntArray(i) = RansomRTF(InitialFontSize)

    Next i
    FontApplicator FntArray

    Select Case Colorise
      Case 0 ' no colour
      Case 1

        RandomColour Int(Rnd * 7)
      Case 2

        RandomColour Int(Rnd * 7), False
      Case 3

        RandomColour2 Int(Rnd * 7), True
    End Select

End Sub

Private Sub RansomFormat(Rstart$, rend$, freq As Long, frmt$)

  'Copyright 2002 Roger Gilchrist
  'support for RansomRTF
  '*---PROGRAMMER MODIFICATION POINT---*

    Rstart$ = Rstart$ & IIf(Rnd > freq, frmt, "")
    If InStr(Rstart$, frmt) Then
        rend$ = rend$ & frmt & "0"
    End If

End Sub

Private Function RansomRTF(InitialFontSize As Long) As String

  'Copyright 2002 Roger Gilchrist
  'generate random RTF formatting
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if unwanted
  'change freq value for more or less of specific effects

  Dim Rstart As String, rend As String

    Rstart$ = ""
    rend$ = ""
    '        remember \up-# convertes to \dn# 50% chance of happening
    Rstart$ = Rstart$ & "\up" & CStr((5 - Int(Rnd * 10 + 1)))
    rend$ = rend$ & "\up0"
    'Font generally gets bigger but has 20% chance of being smaller
    Rstart$ = Rstart$ & "\fs" & (InitialFontSize + 10 - Int(Rnd * 25 + 1))
    RansomFormat Rstart$, rend$, 0.5, "\b"
    RansomFormat Rstart$, rend$, 0.6, "\sub"
    RansomFormat Rstart$, rend$, 0.5, "\super"
    RansomFormat Rstart$, rend$, 0.5, "\i"
    RansomFormat Rstart$, rend$, 0.5, "\caps"
    RansomFormat Rstart$, rend$, 0.6, "\strike"
    RansomUnderliner Rstart$, rend$, 0.5
    Rstart$ = Rstart$ & IIf(Len(Rstart$) > 0, " ", "")
    rend$ = rend$ & " " 'IIf(Len(Rend$) > 0, " ", "")
    RansomRTF = Rstart$ & "***" & rend$

End Function

Private Sub RansomUnderliner(Rstart$, rend$, freq As Long)

  'Copyright 2002 Roger Gilchrist
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if unwanted
  'support for RansomRTF

    If Rnd > freq Then
        Rstart = Rstart & Choose(Int(Rnd * 7) + 1, "\ul", "\uld", "\uldash", "\uldashd", "\uldashdd", "\ulth", "\ulhair", "\ulwave")
        rend = rend & "\ulnone"
    End If

End Sub

Private Sub RippleBase(Part, CurrentValue As Integer, Amplitude As Integer, WaveLength As Integer, UPDOWN As Boolean)

  'Copyright 2002 Roger Gilchrist
  'apply Base ripple
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if unwanted

  Dim i As Long, TMp As String, TmpA As String, tmpB As String

    Select Case RTFPartType(CStr(Part))
      Case blank
        RippleBaseCount CurrentValue, Amplitude, UPDOWN, WaveLength
        Part = SafeJoinSkip
        ' do nothing
      Case RTFOnly
        For i = 1 To CountOccurances(Part, "\tab")
            RippleBaseCount CurrentValue, Amplitude, UPDOWN, WaveLength
        Next i
        ' do nothing
      Case TextOnly
        For i = 1 To Len(Part)
            RippleBaseCount CurrentValue, Amplitude, UPDOWN, WaveLength
            TMp$ = TMp$ & "\up" & CurrentValue & " " & Mid$(Part, i, 1)
        Next i
        Part = TMp$
      Case Mixed
        If InStr(Part, "\") Then
            TmpA = Left$(Part, InStr(Part, "\") - 1)
            tmpB = Mid$(Part, InStr(Part, "\"))

          ElseIf InStr(Part, "}") Then 'NOT INSTR(PART,...
            TmpA = Left$(Part, InStr(Part, "}") - 1)
            tmpB = Mid$(Part, InStr(Part, "}"))
        End If

        For i = 1 To Len(TmpA)
            RippleBaseCount CurrentValue, Amplitude, UPDOWN, WaveLength
            TMp$ = TMp$ & "\up" & CurrentValue & " " & Mid$(TmpA, i, 1)
        Next i
        Part = TMp$ & tmpB

    End Select

End Sub

Private Sub RippleBaseCount(CurLevel As Integer, Limit As Integer, UPDOWN As Boolean, WaveLength As Integer)

  'Copyright 2002 Roger Gilchrist
  'change values for BaseRipple
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if unwanted

    If Abs(CurLevel) = Abs(Limit) Then
        UPDOWN = Not UPDOWN
    End If
    CurLevel = CurLevel + IIf(UPDOWN, -WaveLength, WaveLength)

    CurLevel = IIf(CurLevel < -Limit, -Limit, CurLevel)
    CurLevel = IIf(CurLevel > Limit, Limit, CurLevel)

End Sub

Public Sub RippleEngine(Mode As RippleStyle, Optional Amplitude As Integer = 5, Optional UPDOWN As Boolean = True, Optional InitialNum As Integer = 0, Optional WaveLength As Integer = 1)

  'Copyright 2002 Roger Gilchrist

  '*---PROGRAMMER MODIFICATION POINT---*
  'this is the new and used version of
  'delete if unwanted
  ' apply ripple text; allows you to develop and apply other ripple effects
  ' by adding to the select case structure

  'fontArray format RTFfontcode & "***"
  'FontApplicator replaces "***" with an individual character

  Dim i As Integer          'For variable to manipulate Parts
  Dim LimitLo As Integer, LimitHi As Integer
  Dim FntArray() As String
  Dim SPos As Long, EPos As Long

    If GetStartEnd(SPos, EPos) = False Then 'Use at start of any modifications you build;
        Exit Sub                            'sets start and end and Exits if no Selection '>---> Bottom
    End If
    ReDim FntArray(SPos To EPos) As String

    For i = SPos To EPos
        Select Case Mode
          Case BaseLine
            RippleBaseCount InitialNum, Amplitude, UPDOWN, WaveLength
            FntArray(i) = "\up" & InitialNum & "***"

          Case THeight
            If i = SPos Then 'if first one set initialNum and LimitHi and LimitLo
                InitialNum = FontSize1stCharacter + InitialNum
                LimitLo = InitialNum - Amplitude * 2
                If LimitLo < 8 Then
                    LimitHi = InitialNum + Amplitude * 2 + Abs(8 - LimitLo)
                    LimitLo = 8 ' prevent lowe text becoming to low
                  Else 'NOT LIMITLO...
                    LimitHi = InitialNum + Amplitude * 2
                End If
            End If
            RippleTHeightCount InitialNum, LimitLo, LimitHi, UPDOWN, WaveLength
            FntArray(i) = "\fs" & InitialNum & "***"
        End Select
    Next i
    FontApplicator FntArray

End Sub

Private Sub RippleTHeight(Part, CurLevel As Integer, LimitLo As Integer, LimitHi As Integer, UPDOWN As Boolean, WaveLength As Integer)

  'Copyright 2002 Roger Gilchrist
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if unwanted
  ' apply THeight ripple

  Dim i As Long, TMp As String, TmpA As String, tmpB As String

    Select Case RTFPartType(CStr(Part))
      Case blank

        RippleTHeightCount CurLevel, LimitLo, LimitHi, UPDOWN, WaveLength
        Part = SafeJoinSkip
      Case RTFOnly

        For i = 1 To CountOccurances(Part, "\tab")
            RippleTHeightCount CurLevel, LimitLo, LimitHi, UPDOWN, WaveLength
        Next i
      Case TextOnly

        For i = 1 To Len(Part)
            RippleTHeightCount CurLevel, LimitLo, LimitHi, UPDOWN, WaveLength
            TMp$ = TMp$ & "\fs" & CurLevel & " " & Mid$(Part, i, 1)
        Next i
        Part = TMp$
      Case Mixed
        If InStr(Part, "\") Then
            TmpA = Left$(Part, InStr(Part, "\") - 1)
            tmpB = Mid$(Part, InStr(Part, "\"))
          ElseIf InStr(Part, "}") Then 'NOT INSTR(PART,...
            TmpA = Left$(Part, InStr(Part, "}") - 1)
            tmpB = Mid$(Part, InStr(Part, "}"))
          Else 'NOT INSTR(PART,...
            Exit Sub '>---> Bottom
        End If
        For i = 1 To Len(TmpA)
            RippleTHeightCount CurLevel, LimitLo, LimitHi, UPDOWN, WaveLength
            TMp$ = TMp$ & "\fs" & CurLevel & " " & Mid$(TmpA, i, 1)
        Next i
        Part = TMp$ & tmpB
    End Select

End Sub

Private Sub RippleTHeightCount(CurLevel As Integer, LimitLo As Integer, LimitHi As Integer, UPDOWN As Boolean, WaveLength As Integer)

  'Copyright 2002 Roger Gilchrist
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if unwanted
  ' change values for THeight

    If CurLevel <= LimitLo Or CurLevel >= LimitHi Then
        UPDOWN = Not UPDOWN
    End If
    CurLevel = CurLevel + IIf(UPDOWN, -(2 + WaveLength), 2 + WaveLength)

End Sub

Private Function RTFCodeLoc(TMPRTF$, RTFCode$) As RTFCodeLocs

  'Copyright 2002 Roger Gilchrist
  'WHERE THE RTFCODE OCCURS WITHIN A SelRTF string
  'A single space at the end of RTFCode is the END OF RTFCode Token
  'A double space is an RTFToken followed by a blank text character
  'This routine allows you to judge which format the code code is using
  ' and act accordingly

    If InStr(TMPRTF$, RTFCode$ & "\") Then
        RTFCodeLoc = EmbedOr1st
      ElseIf InStr(TMPRTF$, RTFCode$ & " ") Then 'NOT INSTR(TMPRTF$,...
        If InStr(TMPRTF$, RTFCode$ & "  " & "}") Then
            RTFCodeLoc = LastWithBlank
          ElseIf InStr(TMPRTF$, RTFCode$ & " " & "}") Then 'NOT INSTR(TMPRTF$,...
            RTFCodeLoc = LastInSelection
          Else 'NOT INSTR(TMPRTF$,...
            RTFCodeLoc = LastBeforeText
        End If
      ElseIf Right$(TMPRTF$, Len(RTFCode$)) = RTFCode$ Then 'NOT INSTR(TMPRTF$,... 'NOT INSTR(TMPRTF$,...
        RTFCodeLoc = EndOfString
      Else 'NOT RIGHT$(TMPRTF$,...
        RTFCodeLoc = NotPresent
    End If

End Function

Private Sub RTFHighlight(colr As Long)

  'This is an RTF manipulation version of APIHighLight

  Dim i As Long
  Dim SPos As Long, EPos As Long 'Range of Text and hence BlendArray
  Dim HiLightArray() As Long 'Store colours for passing to ColourApplicator

    If GetStartEnd(SPos, EPos) = False Then 'Use at start of any modifications you build;
        Exit Sub                            'sets start and end and Exits if no Selection '>---> Bottom
    End If
    m_LastHighlightColour = colr
    ReDim HiLightArray(SPos To EPos)
    For i = SPos To EPos
        HiLightArray(i) = colr
    Next i
    ColourApplicator HiLightArray, False

End Sub

Private Sub RTFHighlight2(BColr As Long, FColr As Long)

  'This is an RTF manipulation version of APIHighLight2
  'creator Roger Gilchrist
  'private access for all 2 colour modes

    RTFHighlight BColr
    m_RTB.SelColor = FColr
    m_LastHighlightForeColour = FColr

End Sub

Public Sub RTFHighlightHard(BColr As Long)

  'This is an RTF manipulation version of APIHighLightHard
  'creator Roger Gilchrist
  'Hard means hardcoded colour selection

    RTFHighlight BColr

End Sub

Public Sub RTFHighlightHardHard(BColr As Long, FColr As Long)

  'This is an RTF manipulation version of APIHighLightHardHArd
  'creator Roger Gilchrist
  'Hard means hardcoded colour selection

    RTFHighlight2 BColr, FColr

End Sub

Public Sub RTFHighlightRemove()

  'This is an RTF manipulation version of APIHighLightRemove

    RTFHighlight TranslateSysColor(vbWindowBackground)

End Sub

Public Sub RTFHighlightUser()

  'This is an RTF manipulation version of APIHighLightUser
  'creator Roger Gilchrist Modification of Sergio Perciballi's work  (see APIHihglight)
  'User sets highlight colour Program selects suitable(?) font colour

    RTFHighlight ColourUser

End Sub

Public Sub RTFHighlightUserAuto()

  'This is an RTF manipulation version of APIHighLightUserAuto
  'creator Roger Gilchrist Modification of Sergio Perciballi's work  (see APIHihglight)
  'User sets highlight colour Program selects suitable(?) font colour

  Dim UColr As Long

    UColr = ColourUser
    RTFHighlight2 UColr, InvertColor(UColr)

End Sub

Public Sub RTFHighlightUserUser()

  'This is an RTF manipulation version of APIHighLightUserUser
  'creator Roger Gilchrist Modification of Sergio Perciballi's work  (see APIHihglight)
  'User sets highlight colour Program selects suitable(?) font colour

    RTFHighlight2 ColourUser, ColourUser

End Sub

Private Function RTFPartType(s$) As RTFPartsType

  'Copyright 2002 Roger Gilchrist
  'work out what type of string you're dealing with

    If Len(s$) = 0 Then
        RTFPartType = blank
      ElseIf InStr("\}{" & vbTab & vbNewLine, Left$(s$, 1)) Then 'NOT LEN(S$)...
        RTFPartType = RTFOnly
      ElseIf InStr(s$, "\") = 0 Then 'NOT INSTR("\}{"...
        If InStr(s$, "}") = 0 Then
            RTFPartType = TextOnly

          Else 'NOT INSTR(S$,...
            RTFPartType = Mixed
        End If

      Else 'NOT INSTR(S$,...
        If Left$(s$, 4) = "uc1\" Then
            RTFPartType = RTFOnly     'It is the first part of the MixedBit string in SelRTFToggle
          Else 'NOT LEFT$(S$,...
            RTFPartType = Mixed      '
        End If
    End If

End Function

Private Function SafeJoin(Bits, Optional Del$ = " ") As String

  'Copyright 2002 Roger Gilchrist
  '*PURPOSE: Replace Join with the ability to recognise whether a blank
  'was in original or deleted by editing of member
  'and not insert a del$ character for deleted member
  'SEE ALSO SafeJoinTest

  Dim i As Long

    For i = LBound(Bits) To UBound(Bits)
        If Bits(i) <> SafeJoinSkip Then ' if marked not to reattach then don't add
            SafeJoin = SafeJoin & Bits(i) & Del$
        End If
    Next i
    Do While Right$(SafeJoin, Len(Del$)) = Del$ '' make sure last char<>del$
        SafeJoin = Left$(SafeJoin, Len(SafeJoin) - Len(Del$))
    Loop

End Function

Private Function SafeJoinTest(p, t$) As Variant

  'Copyright 2002 Roger Gilchrist
  '*PURPOSE: Deal with 3 possible events for a split array member
  '1. no edit               accept
  '2. changed content       accept
  '3. deleted               reset as SafeJoinSkip
  'if 3 then skips rather than adding a delimiter otherwise just
  'SEE ALSO SafeJoin

    SafeJoinTest = IIf(t$ <> "" Or p = t$, t$, SafeJoinSkip)

End Function

Public Property Get SelCaps() As Variant

  'Copyright 2002 Roger Gilchrist

    SelCaps = GetGetValue("\caps")

End Property

Public Property Let SelCaps(ByVal vNewValue As Variant)

  'Copyright 2002 Roger Gilchrist

    SelRTFToggle ("\caps")

End Property

Public Property Get SelDown() As Variant

  'Copyright 2002 Roger Gilchrist

    SelDown = GetGetValue("\dn")

End Property

Public Property Let SelDown(ByVal vNewValue As Variant)

  'Copyright 2002 Roger Gilchrist

    If Val(vNewValue) = 0 Then
        vNewValue = "0"
    End If
    SelRTFToggle ("\dn" & vNewValue)

End Property

Private Function SelRTFToggle(code$) As Boolean

  'Copyright 2002 Roger Gilchrist

  Dim RTFHead As String     'Initial section of RTF probably never useful
  Dim Parts As Variant      'Split MixedBit for application of code
  Dim FntTblStart As Long, ClrTblStart As Long, RTF_End As Long ' Cutpoints for various pieces
  Dim i As Integer          'For variable to manipulate Parts
  Dim TMp As String          'tmp for testing whether "" is original or edited

    With m_RTB
        PositionStore Push
        SelRTFToggle = Not (RTFCodeLoc(WorkingStringRead, code$) <> NotPresent) ' Code is present so remove
        SeparateRTFString RTFHead, Parts
        For i = LBound(Parts) To UBound(Parts)
            TMp$ = Parts(i)
            ManipulateRTFCode TMp$, code$, SelRTFToggle
            Parts(i) = SafeJoinTest(Parts(i), TMp$)
        Next i
        WorkingStringWrite RTFHead & SafeJoin(Parts, " ")
        PositionStore pop
        '  End If
    End With 'M_RTB

End Function

Public Property Get SelSub() As Variant

  'Copyright 2002 Roger Gilchrist

    SelSub = GetGetValue("\sub")

End Property

Public Property Let SelSub(ByVal vNewValue As Variant)

  'Copyright 2002 Roger Gilchrist

    SelRTFToggle ("\sub")

End Property

Public Property Get SelSuper() As Variant

  'Copyright 2002 Roger Gilchrist

    SelSuper = GetGetValue("\super")

End Property

Public Property Let SelSuper(ByVal vNewValue As Variant)

  'Copyright 2002 Roger Gilchrist

    SelRTFToggle ("\super")

End Property

Public Property Get SelUp() As Variant

  'Copyright 2002 Roger Gilchrist

    SelUp = GetGetValue("\up")

End Property

Public Property Let SelUp(ByVal vNewValue As Variant)

  'Copyright 2002 Roger Gilchrist

    If Val(vNewValue) = 0 Then
        vNewValue = "0"
    End If
    SelRTFToggle ("\up" & vNewValue)

End Property

Public Property Get SelVisible() As Variant

  'Copyright 2002 Roger Gilchrist

    SelVisible = GetGetValue("\v")

End Property

Public Property Let SelVisible(ByVal vNewValue As Variant)

  'Copyright 2002 Roger Gilchrist

    SelRTFToggle ("\v")

End Property

Private Sub SeparateRTFString(RTFHead As String, Parts)

  'Copyright 2002 Roger Gilchrist
  'separates the RTFCode into 1.RTFHead =the unchanging head of the string
  '                           2. Parts  = an array of the pieces that can be changed

  Dim TMPRTF As String, Cpoint As Long

    TMPRTF = WorkingStringRead
    Cpoint = InStr(TMPRTF, "\uc") - 1 ' but there might be "\viewkind# or might not be
    Do Until InStr("\ " & vbNewLine, Mid$(TMPRTF, Cpoint, 1))
        Cpoint = Cpoint - 1
    Loop
    If Mid$(TMPRTF, Cpoint, 1) = "\" Then
        Cpoint = Cpoint - 1
    End If
    RTFHead = Left$(TMPRTF, Cpoint) 'InStr(TMPRTF, "\uc") - 1)
    Parts = Split(Mid$(TMPRTF, Cpoint + 1)) 'InStr(TMPRTF, "\uc")))

End Sub

'Public Sub RainBow(Optional RedMagenta As Boolean = True, Optional ForeBack As Boolean = True)
'
'  'Copyright 2002 Roger Gilchrist
'  '*---PROGRAMMER MODIFICATION POINT---*
'  'delete if not needed
'
'  Dim i As Long, SPos As Long, EPos As Long
'  Dim RbowMod As Long, Rsection As Integer, Rstep As Long
'  Dim Cycler As Long, RainBowArray() As Long, drv As Long
'
'    If GetStartEnd(SPos, EPos) = False Then 'Use at start of any modifications you build;
'        Exit Sub                            'sets start and end and Exits if no Selection
'    End If
'    RbowMod = (EPos - SPos) \ 6            'this is used to change spectrum sections
'    If RbowMod < 1 Then
'        RbowMod = 1
'    End If
'    ReDim RainBowArray(SPos To EPos) As Long ' initialise array
'    For i = SPos To EPos
'        Rstep = (255 * Cycler / RbowMod)
'
'        If Cycler Mod RbowMod = 0 And Cycler <> 0 Then
'            Cycler = 0
'            Rsection = Rsection + 1
'            Rstep = 0
'        End If
'
'        Cycler = Cycler + 1 '    count through the spectrum section
'        RainBowArray(i) = RainbowColor(IIf(Rsection > 5, 5, Rsection), Rstep)
'    Next i
'If RedMagenta = False Then ' invert spectrum
'    RainBowArray = InvertArray(RainBowArray)
'End If
'    ColourApplicator RainBowArray, ForeBack
'
'End Sub

'
'Private Sub RainBowSubSet(Range As Spectrum, Optional InOut As Boolean = False, Optional LeftRight As Boolean = True, Optional ForeBack As Boolean = True)
'
'  'Copyright 2002 Roger Gilchrist
'  '*---PROGRAMMER MODIFICATION POINT---*
'  'delete if not needed
'
'  Dim i As Long
'  Dim SPos As Long, EPos As Long
'  Dim RbowMod As Long, Rstep As Long
'  Dim Cycler As Long, Rev As Boolean, ClrArray() As Long, RainBowArrayFlip() As Long, drv As Long
'  Dim DIST As Long
'
'    If GetStartEnd(SPos, EPos) = False Then 'Use at start of any modifications you build;
'        Exit Sub                            'sets start and end and Exits if no Selection
'    End If
'
'    RbowMod = IIf(InOut, (EPos - SPos) / 2, EPos - SPos) 'this is used to change spectrum sections
'    If RbowMod < 1 Then
'        RbowMod = 1
'    End If
'    ReDim ClrArray(SPos To EPos) As Long
'    ReDim RainBowArrayFlip(SPos To EPos) As Long
'    For i = SPos To EPos
'        Rstep = 255 * Cycler / RbowMod
'        If Rstep = 255 Then
'            Rev = True
'        End If
'        Cycler = Cycler + IIf(Rev, -1, 1)
'        drv = i
'        If LeftRight = False Then ' invert spectrum
'            drv = SPos + EPos - i
'        End If
'
'        ClrArray(drv) = RainbowColor(CInt(Range), Rstep)
'
'    Next i
'
'    If InOut And LeftRight = False Then ' offsets the array to allow style 4
'        DIST = EPos - SPos
'        For i = LBound(ClrArray) To LBound(ClrArray) + DIST \ 2
'            RainBowArrayFlip(i) = ClrArray(i + (DIST - 1) / 2)
'        Next i
'        For i = LBound(ClrArray) + DIST \ 2 To UBound(ClrArray)
'            RainBowArrayFlip(i) = RainbowColor(i - DIST \ 2, 0)
'            RainBowArrayFlip(i) = RainbowColor(i - DIST \ 2, 1)
'        Next i
'        ClrArray = RainBowArrayFlip
'    End If
'
'
'If InOut Then
'    ClrArray = InOutArray(ClrArray)
'
'End If
'
'If LeftRight = False Then ' invert spectrum
'    ClrArray = InvertArray(ClrArray)
'End If
'    ColourApplicator ClrArray, ForeBack
'
'End Sub
Public Sub SpectrumSector(Range As Spectrum, Optional Inout As Boolean = False, Optional LeftRight As Boolean = True, Optional ForeBack As Boolean = True)

  'Copyright 2002 Roger Gilchrist
  '*---PROGRAMMER MODIFICATION POINT---*
  'delete if not needed

  Dim i As Long 'For Next value
  Dim SPos As Long, EPos As Long 'StartPos and EndPos
  Dim SelLen As Long, Rstep As Long '
  Dim Cycler As Long, ClrArray() As Long '

    If GetStartEnd(SPos, EPos) = False Then 'Use at start of any modifications you build;
        Exit Sub                            'sets StartPos and EndPos and Exits if no Selection '>---> Bottom
    End If

    ReDim ClrArray(SPos To EPos) As Long
    SelLen = EPos - SPos

    For i = SPos To EPos
        Rstep = 255 * Cycler / SelLen ' get percentage of 255
        Cycler = Cycler + 1
        ClrArray(i) = RainbowColor(CInt(Range), Rstep)
    Next i

    If LeftRight = False Then ' invert spectrum
        ClrArray = InvertArray(ClrArray)
    End If

    If Inout Then
        ClrArray = InOutArray(ClrArray)
    End If

    ColourApplicator ClrArray, ForeBack

End Sub

Private Function TranslateSysColor(ByVal lColor As Long) As Long

    lColor = lColor And (Not &H80000000)
    'SysColor2RGB = GetSysColor(lColor)
    TranslateSysColor = GetSysColor(lColor)

End Function

'Public Sub SpectrumSector(col As Spectrum, Mode As Integer, Optional ForeBack As Boolean)
'
'  'Copyright 2002 Roger Gilchrist
'  'public interface for RainBowSubset
'  '*---PROGRAMMER MODIFICATION POINT---*
'  'delete if unwanted
'
'
'    Select Case Mode
'      Case LeftRight
'        RainBowSubSet col, False, True, ForeBack
'      Case RightLeft
'        RainBowSubSet col, False, False, ForeBack
'      Case InOutLeftRight
'        RainBowSubSet col, True, True, ForeBack
'      Case InOutRightLeft
'        RainBowSubSet col, True, False, ForeBack
'    End Select
'
'End Sub

Private Function WorkingStringRead(Optional ForceDoc As Boolean = False) As String

  'Copyright 2002 Roger Gilchrist
  'Get a string for routines to work on
  'and generate Private m_End_OF_RTF to report end of RTF/Start of Text

  Dim fsLoc As Long, GotTrueEnd As Boolean

    With m_RTB
        If .SelLength = 0 Or ForceDoc Then
            WorkingStringRead = .TextRTF
          Else '.SELLENGTH = FALSE'NOT .SELLENGTH...
            WorkingStringRead = .SelRTF
        End If
    End With 'M_RTB

    fsLoc = InStr(1, WorkingStringRead, "\fs")          'find RTF size code and move test point
    If fsLoc = 0 Then                                   'not found = there is no selected Text
        m_End_OF_RTF = 0                                'so exit
        Exit Function '>---> Bottom
    End If
    Do
        If IsNumeric(Mid$(WorkingStringRead, fsLoc + 4, 1)) Then   ' check is size not "fswiss" or other RTF code
            m_End_OF_RTF = InStr(fsLoc, WorkingStringRead, " ")    ' find space delimiting size code
            GotTrueEnd = True
        End If
        fsLoc = InStr(fsLoc + 1, WorkingStringRead, "\fs")         ' find next \fs code and move test point
    Loop Until GotTrueEnd Or fsLoc = 0
    If GotTrueEnd = False And fsLoc = 0 Then                       ' Safety should never hit
        m_End_OF_RTF = 0
    End If

End Function

Private Sub WorkingStringWrite(WStr$)

  'Copyright 2002 Roger Gilchrist
  'put a string from a routine back into the selection or whole control

    With m_RTB
        If .SelLength Then
            If .SelRTF <> WStr Then ' avoid flicker by not updating if no change has occurred
                .SelRTF = WStr
            End If
          Else '.SELLENGTH = FALSE
            If .TextRTF <> WStr Then
                .TextRTF = WStr
            End If
        End If
    End With 'M_RTB

End Sub

':) Ulli's VB Code Formatter V2.13.6 (22/08/2002 11:43:18 AM) 131 + 2280 = 2411 Lines

Public Property Get IsSelection() As Boolean
'this property is only used byt he Colour and Text Look panels to disable the Do It button if there is no selection
IsSelection = m_RTB.SelLength > 0
End Property

