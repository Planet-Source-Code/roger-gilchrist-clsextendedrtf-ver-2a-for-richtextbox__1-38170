<div align="center">

## ClsExtendedRTF Ver 2a for RichTextBox


</div>

### Description

ClsExtendedRTF Version 2a

OOPS SMALL UPGRADE FORGOT ABOUT LOW RESOLUTION SCREENS sort of fixed that and there's a small animation demo added as an apology.

This code manipulates the underlying Rich Text format string in

RichTextBoxes in several ways.

Highlight text, add interesting, strange and weird formats to RTF text.

Resize mixed sizes of fonts with one call.

Add a Zooming ComboBox to any RichTextBox.

CommonDialog is encapsulated to give File New / Open / Save / SaveAs

/ SafeSave and Reload functionality. ShowColor is also encapsulated.

This is a more stable version of this code. Insertion point is preserved

after changes to text. More of the work is done at string manipulation level.

New details added see * below for version 2's new elements.

Minimum Requirments

Riched20.dll (version 3)

Riched32.dll (5.00.2008.1)

probably the Richx32.ocx control

Highlight now comes in two flavours RTF and API there are advantages and disadvatages to each. Menu and Toolbar use API; Text Colour Panel (in Font Looks menu) uses RTF

Incorperated my ClsManifestation (see help) It is behind the spanner button on the tool bar only useful in WindowsXP to choose between XP and Classic look for compiled demo. Does not work in IDE.

Font Looks menu contains to Forms which can be added to other programs to give maximum freedom of using the Font format and colour options.

Added new font colour choices.

Improved Demo performance.

Corrected many minor bugs(extra spaces, spaces being deleted)

Added Zoom ability

Broke up the original ClsExtendedRTF into a couple of separate classes.

Rewrote, simplified and often renamed routines for speed and comprehension

Please comment, vote and send me interesting formats you develop and I'll included them

with acknowlegments. Feel free to take just the parts you want (leave copyrights

etc) this class is still 'Under Construction' if you like check back for

updates.

Demo Program is also a test bed for experimenting with RTF.

'ExtendedRTF Code for VB6.rtf'is both the help file and the demonstation file

and was created interactivly with the code.

NOTE uses VB6 only routines but if you have work-rounds for them should work

with VB5, VB4. Needs more recent versions of RichText control.

ExtendedRTF provides:

1. Highlight (NOT the same as select) *more ways of accessing this facility.

2.Other Underlines wave,dot,dash, dashdot, dashdotdot, hairline, and thick.

(word and double are partially supported; you can set them but they appear as

single in RichTextBox, in Word they look fine)

3. Hidden text/images

4. Weird font formats. Ripple, HeightRipple, Ransom

5.Font colour schemes *Rainbow, *Spectrum and *Materials.

6.subscript, superscript,up and down RTF codes. RichTextBox recognises '\up' and

'\dn' with as superscript but not '\sub' and '\super' ???? <P>

7. Get current insertion point as percentage of total document

8. Remove excess spaces in selection or whole document.

9. File handling: routines to simplify standard New/Open/Save/SaveAs.

*Reload routine reload last saved document without opening CommonDialog

*Document loaded in IDE is automatically reloaded from disk.

SafeSave routine place this in your exit program and file procedures

and you'll never lose your edits again.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2002-08-23 11:14:20
**By**             |[Roger Gilchrist](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/roger-gilchrist.md)
**Level**          |Advanced
**User Rating**    |4.7 (61 globes from 13 users)
**Compatibility**  |VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[ClsExtende1214028222002\.zip](https://github.com/Planet-Source-Code/roger-gilchrist-clsextendedrtf-ver-2a-for-richtextbox__1-38170/archive/master.zip)








