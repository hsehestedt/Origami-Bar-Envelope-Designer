' OBED - Origami Bar Envelope Designer
' by Hannes Sehestedt

' ********************************
' * Version history follows code *
' ********************************

Rem $DYNAMIC
$ExeIcon:'envelope.ico'
$VersionInfo:FILEVERSION#=4,0,5,16
$VersionInfo:LegalCopyright=2021 by Hannes Sehestedt
$VersionInfo:ProductName=Origami Bar Envelope Designer
Option _Explicit
$Console:Only
_Dest _Console: _Source _Console

Width 70, 44
Width 70, 44, 70, 44
Option Base 1

_ConsoleTitle "Origami Bar Envelope Designer 4.0.5.16 by Hannes Sehestedt"

Start:

Clear

' Define variables to be used in this program.

Dim ChangeInSize As Single
Dim dark_block As String
Dim Description As String
Dim EW_Reduction As Single
Dim light_block As String
Dim Main_Choice As Integer
Dim paper_height_integer As Single
Dim paper_height_fraction As Single
Dim paper_width_integer As Single
Dim paper_width_fraction As Single
Dim paper_width_frac As String, paper_height_frac As String
Dim PrintSpec As String
Dim ProgramStartDir As String
Dim Resolution As Single
Dim temp As Single
Dim TweakChoice As String
Dim TweakedPaperHeight As Single
Dim TweakedPaperWidth As Single
Dim TweakMode As Integer
Dim value As Single
Dim x As Integer
Dim y As Integer

' The following variables are for the dimensions of the envelope and paper used to create the envelope
' Some values will be supplied by the user while others will be calculated.

Dim line$(42)
Dim envelope_width As Single, envelope_height As Single, paper_width As Single, paper_height As Single
Dim upper_front As Single, left_front As Single, bottom_front As Single, right_front As Single
Dim upper_front_frac As String, left_front_frac As String, bottom_front_frac As String, right_front_frac As String
Dim upper_bar As Single, left_bar As Single, bottom_bar As Single, right_bar As Single
Dim upper_bar_frac As String, left_bar_frac As String, bottom_bar_frac As String, right_bar_frac As String
Dim face_height As Single, face_width As Single, bar_height As Single, bar_width As Single
Dim face_height_frac As String, face_width_frac As String, bar_height_frac As String, bar_width_frac As String

' Globally available (shared) variables

Dim Shared LowerValue As Single
Dim Shared ReducedDenominator As Single
Dim Shared UpperValue As Single
Dim Shared YN As String

' End of variable declarations

' Save the location where the program was run from

ProgramStartDir$ = _CWD$

Color 15
Cls
Print "Welcome to the Origami Bar Envelope Designer (OBED)"
Print "(c) 2021 by Hannes Sehestedt"
Print
Print
Print "Which would you like to do?"
Print
Color 14, 2: Print "1)";: Color 15: Print " Determine the size of the page needed to create an envelope of a"
Print "specific size."
Print
Color 14, 2: Print "2)";: Color 15: Print " Determine the size of the envelope that would be created from a"
Print "specific paper size."
Print
Color 14, 2: Print "3)";: Color 15: Print " Display program help."
Print
Color 14, 2: Print "4)";: Color 15: Print " Quit the program."
Print
Print "Make your choice by number ";: Color 14, 2: Print "(1 - 4)";: Color 15: Input ": ", Main_Choice
If Main_Choice = 1 Then GoTo calc_page_size
If Main_Choice = 2 Then GoTo calc_envelope_size
If Main_Choice = 3 Then GoTo show_help
If Main_Choice = 4 Then System
Cls
Print
Color 14, 4: Print "That was not a valid selection.";: Color 15: Print " You need to enter a 1, 2, or 3 to make"
Print "your selection."
Print
Print "Press any key to continue."
Print
_Delay .2
_KeyClear
Sleep
GoTo Start

calc_page_size:

' This procedure will calculate the paper size needed to create an envelope of the size that the user specifies.

Cls
Print "  ";: Color 0, 14: Print "******************************************************************"
Color 15: Print "  ";: Color 0, 14: Print "* Calculate Paper Size Needed to Create a Specific Envelope Size *"
Color 15: Print "  ";: Color 0, 14: Print "******************************************************************"
Color 15
Print
Print "We will need to know what size envelope you wish to create."
Print

' Set an initial value
envelope_width = 0

Do
    Input "What is the width of the envelope"; envelope_width
Loop While envelope_width = 0

' Set an initial value
envelope_height = 0

Do
    Input "What is the height of the envelope"; envelope_height
Loop While envelope_height = 0

' Ask if Tweak Mode should be enabled. Give the user an option to get an explanation of this feature.
' If user elects to use this feature, then we need to ask to what nearest fractional value
' (1/8, 1/16, 1/32, etc.) we should calculate the page size.

AskAboutTweakMode:

Do
    Cls
    Print "Do you want to want enable "; Chr$(34); "Tweak Mode"; Chr$(34); "?"
    Print
    Print "NOTE: Enter the word "; Chr$(34); "HELP"; Chr$(34); " if you want to see help describing what"
    Print "this option does."
    Print
    Print "Enable "; Chr$(34); "Tweak Mode"; Chr$(34);: Input TweakChoice$

    If UCase$(TweakChoice$) = "HELP" Then
        Cls
        Print "When the program calculates the page size, it will calculate and"
        Print "display the precise size of the page. However, you may prefer to"
        Print "have page sizes that fall on the nearest 1/8, 1/16, 1/32, etc. in"
        Print "order to make measurements easier. If you enable Tweak Mode, then"
        Print "the program will adjust values so that the page size falls precisely"
        Print "on the nearest such increment. You will be asked what resolution you"
        Print "want to use (eighths, sixteenths, etc.). Note that the end result"
        Print "may be an envelope that is tweaked to a just slightly larger size"
        Print "than you had originally specified. It will never be smaller."
        Print
        Print "NOTE: This feature will typically be more useful to people using"
        Print "imperial measurements rather than metric. This is because imperial"
        Print "measurements are typically done in increments of 1/2, 1/4, 1/8, etc."
        Print "rather than the tenths used by metric units of measurement."
        Pause
        GoTo AskAboutTweakMode
    End If
    YesOrNo TweakChoice$
Loop While YN$ = "X"

Select Case YN$
    Case "Y"
        TweakMode = 1
    Case "N"
        TweakMode = 0
End Select

Resolution = 0 ' Set initial value

Do
    Cls
    Print "Enter the resolution for fractional calculations. As an example,"
    Print "enter 8 for 8ths of an inch, 16 for 16ths, etc."
    Print
    Print "Note that if you also enabled Tweak Mode, the page size will also be"
    Print "adjusted to the nearest 8th, 16th, etc."
    Print
    Print "If you are using metic units of measurement, such as cm, it is"
    Print "suggested that you enter a value of 10."
    Print
    Input "Enter resolution: ", Resolution
Loop While Resolution = 0

' Convert resolution to a decimal value. As an example, 1/8 of an inch is by definition 1 divided by 8 or 0.125.
' As a result, we need to divide 1 by the number provided by the user.

BeginCalculations:

LowerValue = Resolution
Resolution = 1 / Resolution

' Calculate the paper size needed to make an envelope of the desired size.

paper_height = envelope_height / .3125
paper_width = envelope_width + (.25 * paper_height)

' Break out the integer and fractional portions of the paper size into seperate variables.

paper_height_integer = Int(paper_height)
paper_height_fraction = paper_height - paper_height_integer

paper_width_integer = Int(paper_width)
paper_width_fraction = paper_width - paper_width_integer

If TweakMode = 0 Then GoTo EndTweakModeCalculations

' If Tweak Mode is enabled, then determine if the paper height is a multiple of the resolution
' selected by the user and than the tweak the size of the page if it is not a multiple of the
' desired resolution.

If Int(paper_height / Resolution) <> (paper_height / Resolution) Then
    temp = paper_height_fraction / Resolution
    If (temp - Int(temp)) > 0 Then
        temp = Int(temp) + 1
        temp = temp * Resolution
        TweakedPaperHeight = paper_height_integer + temp
        ChangeInSize = TweakedPaperHeight - paper_height

        ' Reset the paper height to the new tweaked size

        paper_height = TweakedPaperHeight

        ' The width of the envelope is reduced by .25 x the page length increase. This envelope width reduction amount
        ' will be stored in EW_Reduction

        EW_Reduction = ChangeInSize * .25
        paper_width = paper_width + EW_Reduction
    End If
End If

If Int(paper_width / Resolution) <> (paper_width / Resolution) Then
    temp = paper_width_fraction / Resolution
    If (temp - Int(temp)) > 0 Then
        temp = Int(temp) + 1
        temp = temp * Resolution
        TweakedPaperWidth = paper_width_integer + temp

        ' Reset the paper width to the new tweaked paper width

        paper_width = TweakedPaperWidth
    End If
End If

EndTweakModeCalculations:

If envelope_height > (1.25 * envelope_width) Then
    Cls
    Print
    Color 14, 4: Print "WARNING!": Color 15
    Print
    Print "These are invalid values. The envelope height cannot exceed"
    Print "1.25 times the envelope width."
    Print
    Print "Press any key to continue."
    _Delay .2
    _KeyClear
    Sleep
    GoTo Start
End If

If envelope_width >= (1.2 * envelope_height) Then GoTo envelope_size_valid

invalid1:

Cls
Print
Color 14, 4: Print "WARNING!": Color 15
Print
Print "Normally, the width of the envelope must be at least 1.2 times the"
Print "envelope height. This configuration will still work but it will"
Print "require the left and right side of the envelope flaps to overlap."
Print
Print "Press any key to continue."
_Delay .2
_KeyClear
Sleep

envelope_size_valid:

GoSub Calculate_Printable_Areas

GoTo ask_for_more

calc_envelope_size:

' This procedure will calculate the size of the envelope created by a paper size specified by the user.
Cls
Print "   ";: Color 0, 14: Print "****************************************************************"
Color 15: Print "   ";: Color 0, 14: Print "* Calculate Envelope Size Resulting From a Specific Paper Size *"
Color 15: Print "   ";: Color 0, 14: Print "****************************************************************"
Color 15
Print
Print "We will need to know the size of the paper that you will be using."
Print
' Set initial values
paper_width = 0
paper_height = 0

Do
    Input "What is the width of the paper that you are using"; paper_width
Loop While paper_width = 0

Do
    Input "What is the height of the paper that you are using"; paper_height
Loop While paper_height = 0

If paper_height > (2 * paper_width) Then
    Cls
    Print
    Color 14, 4: Print "WARNING!": Color 15
    Print
    Print "These are invalid values. The page length must not exceed"
    Print "2 times the page width."
    Print
    Print "Press any key to continue."
    _Delay .2
    _KeyClear
    Sleep
    GoTo Start
End If

If paper_width >= (paper_height * .625) Then GoTo paper_size_valid

invalid2:

Cls
Print
Color 14, 4: Print "WARNING!": Color 15
Print
Print "Normally, the width of the paper must be at least .625 times the"
Print "height. This configuration will still work but it will require"
Print "the left and right side of the envelope flaps to overlap."
Print
Print "Press any key to continue."
_Delay .2
_KeyClear
Sleep

paper_size_valid:

Resolution = 0 ' Set initial value

Do
    Cls
    Print "Enter the resolution for fractional calculations. As an example, if"
    Print "you want to calculate to the nearest eighth of an inch, enter 8,"
    Print "for the nearest sixteenth of an inch, enter 16, etc."
    Print
    Input "Enter resolution: ", Resolution
Loop While Resolution = 0

envelope_width = paper_width - (.25 * paper_height)
envelope_height = .3125 * paper_height

' Now that we have calculated the envelope height and width, the calculations can continue in
' exactly the same way as for calculating the paper size so we are now going to jump into
' that routine.

GoTo BeginCalculations

ask_for_more:

' Save a printable copy of the output to a file

Open "OBED_Results.txt" For Output As #1
For x = 1 To 42
    For y = 1 To 70
        Print #1, Chr$(Screen(x, y));
    Next y
    Print #1, ""
Next x
Close #1

'Locate 44, 1: Input "Press any key to continue", temp
Locate 43, 1: Pause

Cls
Print "A printable copy of the output has been saved as:"
Print
Color 14, 2: Print ProgramStartDir$; "\OBED_Results.txt": Color 15
Print
Print "When printing this file I would suggest formatting the text using"
Print "the font called "; Chr$(34); "Terminal"; Chr$(34); " with a font style of "; Chr$(34); "Regular"; Chr$(34); ". This should"
Print "provide the best looking results."

Pause
GoTo Start

Calculate_Printable_Areas:

' Begin by getting a description for the envelope that would be especially useful
' for printed output to help you remember what that envelope design was for.

Description$ = "" 'Set initial value
Cls
Print "Enter a description for this envelope (53 character max) or press"
Print "<ENTER> for none."
Print
Input "Description: ", Description$
If Len(Description$) > 53 Then
    Cls
    Print "The description is limited to 53 characters. Please try again."
    Pause
    GoTo Calculate_Printable_Areas
End If

upper_front = (paper_height * .1875)
ReduceFraction (CalcNumerator!(upper_front)), (LowerValue)
'upper_front_frac$ = LTrim$(Str$(Int(upper_front))) + " " + LTrim$(Str$(UpperValue)) + "/" + LTrim$(Str$(ReducedDenominator))
upper_front_frac$ = MakeFraction$(upper_front)

bottom_front = (paper_height * .5)
ReduceFraction (CalcNumerator!(bottom_front)), (LowerValue)
bottom_front_frac$ = MakeFraction$(bottom_front)

left_front = (paper_height * .125)
ReduceFraction (CalcNumerator!(left_front)), (LowerValue)
left_front_frac$ = MakeFraction$(left_front)

right_front = (paper_width - (paper_height * .125))
ReduceFraction (CalcNumerator!(right_front)), (LowerValue)
right_front_frac$ = MakeFraction$(right_front)

upper_bar = (.875 * paper_height)
ReduceFraction (CalcNumerator!(upper_bar)), (LowerValue)
upper_bar_frac$ = MakeFraction$(upper_bar)

bottom_bar = (.9375 * paper_height)
ReduceFraction (CalcNumerator!(bottom_bar)), (LowerValue)
bottom_bar_frac$ = MakeFraction$(bottom_bar)

left_bar = (.25 * paper_height)
ReduceFraction (CalcNumerator!(left_bar)), (LowerValue)
left_bar_frac$ = MakeFraction$(left_bar)

right_bar = (paper_width - (.25 * paper_height))
ReduceFraction (CalcNumerator!(right_bar)), (LowerValue)
right_bar_frac$ = MakeFraction$(right_bar)

Cls
light_block$ = Chr$(176)
dark_block$ = " "

' dark_block$ = CHR$(219)
' vertical bar is chr$(179)
' horizontal bar is chr$(196)

' Lines 1 to 6: Solid
' Lines 6 to 20: Front face
' Lines 21 to 34: Solid
' Lines 35 to 38: Bar
' Lines 39 to 41: Solid

For x = 1 To 6
    line$(x) = "1111111111111111111111111111111111111111111111111111111111111111111111"
Next x

For x = 6 To 20
    line$(x) = "1111111111000000000000000000000000000000000000000000000000001111111111"
Next x

For x = 21 To 34
    line$(x) = "1111111111111111111111111111111111111111111111111111111111111111111111"
Next x

For x = 35 To 38
    line$(x) = "1111111111111111111100000000000000000000000000000011111111111111111111"
Next x

For x = 39 To 42
    line$(x) = "1111111111111111111111111111111111111111111111111111111111111111111111"
Next x

For x = 1 To 42
    For y = 1 To Len(line$(x))
        If Mid$(line$(x), y, 1) = "1" Then
            Color 0, 14
            Print light_block$;
            Color 15
        End If
        If Mid$(line$(x), y, 1) = "0" Then Print dark_block$;
    Next y
    Print
Next x

Print

' Print locations of envelope front

Locate 13, 16: Print "This is the front Face of the Envelope"
Locate 6, 12: Print "< Upper Left"
Locate 7, 11
value = upper_front: GoSub Format
Print Using PrintSpec$; upper_front;: Print " from top ";
value = left_front: GoSub Format
Print Using PrintSpec$; left_front;: Print "  from left"
Locate 8, 11
Print " "; upper_front_frac$; " from top "; left_front_frac$; " from left"
Locate 20, 47: Print "Lower Right >"
Locate 18, 24
value = bottom_front: GoSub Format
Print Using PrintSpec$; bottom_front;: Print " from top ";
value = right_front: GoSub Format
Print Using PrintSpec$; right_front;: Print "  from left"
Locate 19, 24
Print " "; bottom_front_frac$; " from top "; right_front_frac$; " from left"

' Print locations of envelope bar

Locate 36, 25: Print "This is the bar on the"
Locate 37, 26: Print "back of the envelope"
Locate 35, 8: Print "Upper Left >"
Locate 33, 4
value = upper_bar: GoSub Format
Print Using PrintSpec$; upper_bar;: Print " ("; upper_bar_frac$; ") from top ";
value = left_bar: GoSub Format
Print Using PrintSpec$; left_bar;: Print " ("; left_bar_frac$; ") from left"
Locate 38, 52
Print "< Lower Right"
Locate 40, 14
value = bottom_bar: GoSub Format
Print Using PrintSpec$; bottom_bar;: Print " ("; bottom_bar_frac$; ") from top ";
value = right_bar: GoSub Format
Print Using PrintSpec$; right_bar;: Print " ("; right_bar_frac$; ") from left"

' Print overall summary in middle of page

value = (bottom_front - upper_front): GoSub Format
face_height = value
ReduceFraction (CalcNumerator!(face_height)), (LowerValue)
face_height_frac$ = MakeFraction$(face_height)
Locate 25, 4: Print "Front face is ";
Print Using PrintSpec$; face_height;: Print " ("; face_height_frac$; ") high by ";
value = (right_front - left_front): GoSub Format
face_width = value
ReduceFraction (CalcNumerator!(face_width)), (LowerValue)
face_width_frac$ = MakeFraction$(face_width)
Print Using PrintSpec$; face_width;: Print " ("; face_width_frac$; ") wide"
Locate 26, 4
Print "Envelope bar is ";
value = (bottom_bar - upper_bar): GoSub Format
bar_height = value
ReduceFraction (CalcNumerator!(bar_height)), (LowerValue)
bar_height_frac$ = MakeFraction$(bar_height)
Print Using PrintSpec$; bar_height;: Print " ("; bar_height_frac$; ") high by ";
value = (right_bar - left_bar): GoSub Format
bar_width = value
ReduceFraction (CalcNumerator!(bar_width)), (LowerValue)
bar_width_frac$ = MakeFraction$(bar_width)
Print Using PrintSpec$; bar_width;: Print " ("; bar_width_frac$; ") wide"

If Description$ <> "" Then
    Locate 29, Int(((70 - (Len(Description$) + 17)) / 2))
    Print "> Description: "; Description$; " <"
End If

' NOTE: With tweak mode enabled, we will adjust the page size to fall precisely on an incrememt of the size
' specified by the user, such as 8ths, 16ths,32nds, etc. If a value comes out to an even integer value then
' we don't want to display these fractional portions. For example we want to display "8 inches" rather than
' "8 0/16 inches". The printed code below includes handling of this situation.

temp = paper_width - Int(paper_width)
If temp <> 0 Then
    UpperValue = CInt(temp / Resolution)
Else
    UpperValue = 0
End If

Locate 3, 7
value = paper_width: GoSub Format
ReduceFraction (CalcNumerator!(paper_width)), (LowerValue)
paper_width_frac$ = MakeFraction$(paper_width)
Print "Page size is";: Print Using PrintSpec$; paper_width;: Print " ("; paper_width_frac$; ") wide by";
value = paper_height: GoSub Format
ReduceFraction (CalcNumerator!(paper_height)), (LowerValue)
paper_height_frac$ = MakeFraction$(paper_height)
Print Using PrintSpec$; paper_height;: Print " ("; paper_height_frac$; ") high"
Locate 4, 7:

If Main_Choice = 1 Then
    Select Case TweakMode
        Case 0
            Print "Tweak Mode was DISABLED. ";
        Case 1
            Print "Teak Mode was ENABLED. ";
    End Select
End If

Print "Calculations are to nearest 1/"; LTrim$(Str$(LowerValue)); ".";
Return

Format:

' When printing values on the graphic, we want to ensure that the numbers are a consistent number of characters.
' Here we set

If value < 10 Then
    PrintSpec$ = "##.###"
ElseIf value < 100 Then
    PrintSpec$ = "###.###"
ElseIf value < 1000 Then
    PrintSpec$ = "####.###"
End If

Return



show_help:

Cls
Print "Help - Page 1 of 2"
Print "=================="
Print
Print "This program will allow you to create custom sized origami bar"
Print "envelopes. When you run the program, you will be provided two main"
Print "options:"
Print
Print "1) Determine the size of the page needed to create an envelope of the"
Print "   size that you specify."
Print
Print "2) Detrmine what the envelope size will be if you start with a known"
Print "   page size."
Print
Print "After making your selection, the program will gather the needed"
Print "information from you and will then display a representation of the"
Print "page showing the location of all elements. The information displayed"
Print "will be:"
Print
Print "> The width and height of the page needed to create the envelope."
Print "> The overall envelope size (same as the size of the front face)."
Print "> Distance from top of page to top of envelope front face."
Print "> Distance from left edge of page to left edge of the envelope face."
Print "> Distance from top of page to bottom of envelope front face."
Print "> Distance from left edge of page to right edge of the envelope face."
Print "> Distance from top of page to top of envelope bar."
Print "> Distance from left edge of page to left edge of the envelope bar."
Print "> Distance from top of page to bottom of envelope bar."
Print "> Distance from left edge of page to right edge of the envelope bar."
Print "> A description of the envelope design, if you have provided one."
Print
Print "With this information you will be able to precisely layout the"
Print "location of the front face and bar on the back of the envelope. This"
Print "will allow you to add text or graphics to these either manually or"
Print "from a printer prior to folding the envelope."
Print
Print "When the program gathers the needed information from you, it will ask"
Print "for the resolution with which to display fractional values. For"
Print "example, if you want values to be displayed to the nearest 16th, you"
Print "would specify 16 as the resolution. For 32nds, enter 32, etc. If you"
Print "are using metric units of measurements such as mm or cm, then you will"
Print "probably want to specify 10 as your resolution."
Pause
Cls
Print "Help - Page 2 of 2"
Print "=================="
Print
Print "All output will be provided in both decimal format and as a fraction."
Print "The decimal values will be precise numbers while the fractional"
Print "values will be rounded to the nearest fraction matching the"
Print "resolution you specified. Decimal results are most useful for laying"
Print "out a page on the computer in a program such as Microsoft Publisher"
Print "so that you can add text and graphics to the front face and bar. The"
Print "fractional values are best suited for manual layout such as when"
Print "using a ruler to layout all the elements."
Print
Print "Only applicable to determining the page size: When you select this"
Print "option, you will be asked if you want to enable "; Chr$(34); "Tweak Mode"; Chr$(34); ". With"
Print "Tweak Mode enabled, the page size will be adjusted so that it falls"
Print "on an exact increment of the resolution you specified. For example,"
Print "if the program determines the page width to be 8.497 and you have"
Print "chosen to enable Tweak Mode with a resolution of 1/32 increments,"
Print "then the program may tweak the page size to 8.500 or 8 16/32, which"
Print "is equivalent to 8 1/2. The program would display this as 8 1/2 since"
Print "it will perform automatic fraction reduction."
Print
Print "Note that the process of tweaking will ever so slightly change the"
Print "size of the resulting envelope. The program is designed to make"
Print "certain that the change to the page size will only result in"
Print "increases to the envelope size, never a decrease."
Print
Print "Printing tips: When the program is run, it will automatically create"
Print "a copy of the output as a plain text file named "; Chr$(34); "OBED_Results.txt"; Chr$(34); " in"
Print "the same folder from which the program was run. When printing this"
Print "file I would suggest formatting the text using the font called"
Print Chr$(34); "Terminal"; Chr$(34); " with a font style of "; Chr$(34); "Regular"; Chr$(34); ". This should provide the"
Print "best looking results."

Pause
GoTo Start


' Sub Procedures


Sub YesOrNo (YesNo$)

    ' This routine checks whether a user responded with a valid "yes" or "no" response. The routine will return a capital "Y" in YN$
    ' if the user response was a valid "yes" response, a capital "N" if it was a valid "no" response, or an "X" if not a valid response.
    ' Valid responses are the words "yes" or "no" or the letters "y" or "n" in any case (upper, lower, or mixed). Anything else is invalid.

    Select Case UCase$(YesNo$)
        Case "Y", "YES"
            YN$ = "Y"
        Case "N", "NO"
            YN$ = "N"
        Case Else
            YN$ = "X"
    End Select

End Sub


Sub Pause

    ' Displays one blank line and then the message "Press any key to contine"

    Print
    Shell "pause"
End Sub


Sub ReduceFraction (Numerator, Denominator)

    ' This routine will reduce a fraction. For example, a value of 10/16 would be reduced to 5/8. To use this routine, pass the numerator and
    ' denominator to it. The reduced numerator will be returned in the variable UpperValue and the reduced denominator will be returned in the
    ' variable ReducedDenominator.
    ' Example: ReduceFraction (val1),(val2)

    Dim NumeratorResult, DenominatorResult As Single
    Dim TopNumToTest As Integer
    Dim x As Integer

    StartReduction:

    If Numerator = 1 Then GoTo End_ReduceFraction

    TopNumToTest = Int(Numerator / 2)
    If TopNumToTest = 1 Then TopNumToTest = 2

    For x = 2 To TopNumToTest
        NumeratorResult = Numerator / x
        DenominatorResult = Denominator / x
        If (NumeratorResult = Int(Numerator / x)) And (DenominatorResult = Int(Denominator / x)) Then
            Numerator = NumeratorResult
            Denominator = DenominatorResult
            GoTo StartReduction
        End If
    Next x

    End_ReduceFraction:

    ' Set the final values to be returned from this routne

    UpperValue = Numerator
    ReducedDenominator = Denominator

End Sub


' Functions


Function CalcNumerator! (Value!)
    CalcNumerator! = CInt((Value - Int(Value)) * LowerValue)
End Function


Function MakeFraction$ (Value!)

    If UpperValue = 0 Then
        MakeFraction$ = LTrim$(Str$(Int(Value)))
    ElseIf UpperValue = ReducedDenominator Then
        MakeFraction$ = LTrim$(Str$(Int(Value) + 1))
    Else
        MakeFraction$ = LTrim$(Str$(Int(Value))) + " " + LTrim$(Str$(UpperValue)) + "/" + LTrim$(Str$(ReducedDenominator))
    End If

    If Left$(MakeFraction$, 1) = "0" Then
        MakeFraction$ = Right$(MakeFraction$, (Len(MakeFraction$) - 2))
    End If

End Function



' *******************
' * Version History *
' *******************

' 2.0.0.1 - Initial release

' 2.0.0.2 - Fix a problem with sleep states by adding a delay a clearing key buffer before running sleep commands

' 2.1.0.3 - Make a change to how program handles an out-of-spec envelope. Rather than stopping execution, we will now print
'           a warning, but will continue on to perform the calculations since it is still possible to fold the envelope.
' 2.1.0.5 - Added a console title, added $Versioninfo tages, recompiled on QB64 version 1.4

' 3.0.0.6 - Sep 25, 2020: Added "Tweak Mode". This mode only applies when calculating the page size needed to create an envelope
'           of a specific size. When performing this operation, you could end up with a calculated page size that has dimensions
'           that would be very difficult to measure. By enabling Tweak Mode you will be asked what fractional measurements you
'           would like to use. For example, maybe you want page sizes specified to the nearest 1/8, 1/16, 1/32, etc. inches.
'           The program will then tweak the page size by the smallest amount possible to still give you an envelope at least
'           as wide and high as you have specified while adjusting page sizes to the nearest 1/8, 1/16, etc.

' 3.0.1.7 - Sep 26, 2020: Tweaked the Tweak Mode :-). Since the idea of Tweak Mode is to provide the page size in an
'           increment of a specific fractional value (eigths, sixteenths, etc. of an inch), I have added a line at the top of
'           the page to show the page in fractional format. Directly below that we also show the page size in decimal format.

' 3.0.2.8 - Oct 16, 2020: There was an issue with spacing of some of the values displayed to the screen. For example, a
'           numerical value might be displayed immediately after some text with no space, or there might be 2 spaces rather
'           than just 1 space. This has been corrected. In addition, when tweak mode was enabled and values are displayed
'           in fractional values such as "8 1/16 inches", if the calculated value was a full integer value we were still
'           displaying the fractions. For example, "8 0/16 inches". Checking for this situation has been added so that this
'           would now be displayed as "8 inches".

' 3.0.3.9 - Dec 2, 2020: Fractional values would not be reduced. As an example, the program might return a value of 10/16
'           rather than 5/8. This has been corrected so that fractions will always be reduced. NOTE: We only display
'           fractional values when tweaking mode is selected by the user. As a result, only tweaking mode was affected by
'           this issue.

' 3.1.0.10 - Mar 25, 2021: Discovered a bug when tweaking mode is enabled. The program was returning incorrect fractional
'            values. This has been corrected.

' 4.0.0.11 - Mar 28, 2021: Added major new functionality to the program: Added a fractional display of all calculated values.
'            For example, if a calculation comes out to 4.25 the output will display 4.25 as well as 4 1/4. Calculations
'            will be made to the nearest 8th, 16th, 32nd, etc. as selected by the user. The program will now also automatically
'            save a copy of the output to a printable file.
'
'            Finally, with this new major release, we introduce a new name to the program: OBED (Origami Bar Envelope Designer).
'            The name of the executable also reflects this change.

' 4.0.1.12 - Mar 29, 2021: In the section of code that creates the fractional output strings, if the integer portion of the result
'            was a zero (0), the fraction was being displayed as in this example: 0 5/16. This has been change so that the leading
'            zero is now dropped to provide a cleaner looking output.

' 4.0.2.13 - No functional changes, just some rewording of a few messages and a little cleanup of the text in the program help.

' 4.0.3.14 - Apr 19, 2021: Add some text both in the main program and the help section to note that the best results for printing
'            will be obtained by formatting the output with the "Terminal" font in the font style called "Regular".

' 4.0.4.15 - May 26, 2021: After the program is run, we would inform the user that a printable copy of the output was saved to the
'            same location from where the program was run. This has been modified to display the full path to make this easier for
'            the user to locate.

' 4.0.5.16 - Sep 9, 2021: Removed a few unnecessary lines from the code. These lines woulkd display a message but we never paused
'            the program so the output was unreadable because the very next operation would clear the screen and display another
'            message. These unnecessary lines were caught by pure accident. They have now been removed from the program.

