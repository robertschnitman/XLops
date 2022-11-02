Attribute VB_Name = "FunText"
' ============================================================================
' REPO: XLops
' MODULE: FunText.bas
' DESCRIPTION: Functions for parsing text.
' REQUIRED: Micrsoft VBScript Regular Expressions 5.5 (in VBA, Tools > References) for RX*() functions.
' https://stackoverflow.com/questions/22542834/how-to-use-regular-expressions-regex-in-microsoft-excel-both-in-cell-and-loops
'
' LIST OF FUNCTIONS:
' FINDREPLACE()
' FINDREMOVE()
' FINDBEFORE()
' FINDAFTER()
' FINDBETWEEN()
' FIRSTNAME()
' LASTNAME()
' TEXTLIKE()
' TEXTSTRIPWS()
' TEXTINSERT()
' TEXTREVERSE()/TEXTREV()
' TEXTCOMPARE()
' TEXTJOINR()
' TRIML()
' TRIMR()
' TRIMLR()
' TRIMC()/CTRIM()
' RXLIKE()
' RXREPLACE()
' RXREMOVE()
' RXGET()
' RXGETALL()
' ============================================================================

' ============================================================================
' Function: FINDREPLACE()
' Description: In a cell, replace a string with another.

Function FINDREPLACE(cell As String, StringFind As String, StringNew As String)
    
    ' VBA's Replace() is NOT like Excel's REPLACE()!!! It is simpler.
    FINDREPLACE = Replace(cell, StringFind, StringNew)

End Function
' ============================================================================

' ============================================================================
' Function: FINDREMOVE()
' Description: In a cell, remove specified character(s).

Function FINDREMOVE(cell As String, findchar As String)

    FINDREMOVE = FINDREPLACE(cell, findchar, "")

End Function
' ============================================================================

' ============================================================================
' Function: FINDBEFORE()
' Description: In a cell, return the text before the first specified character(s).

Function FINDBEFORE(cell As String, findchar As String)

    ' VBA's Instr() is like Excel's Find().
    CharPos = InStr(cell, findchar)
    
    ' If char cannot be found, throw an error.
    If CharPos = 0 Then
        
        FINDBEFORE = CVErr(xlErrNA)
        
    ' Otherwise, get everything before the specified character.
    Else
    
        FINDBEFORE = Left(cell, CharPos - 1)
        
    End If

End Function
' ============================================================================

' ============================================================================
' Function: FINDAFTER()
' Description: In a cell, return the text after the first specified character(s).

Function FINDAFTER(cell As String, findchar As String)
    
    ' VBA's Instr() is like Excel's Find().
    CharPos = InStr(cell, findchar)
    
    ' If char cannot be found, throw an error.
    If CharPos = 0 Then

        FINDAFTER = CVErr(xlErrNA)
        
    ' Otherwise, get everything after char.
    Else
    
        FINDAFTER = Mid(cell, CharPos + Len(char)) ' We add Len(char) in case char has multiple characters (e.g. "Robert ").
        
    End If

End Function
' ============================================================================

' ============================================================================
' Function: FINDBETWEEN()
' Description: In a cell, return the text BETWEEN specified characters.

Function FINDBETWEEN(cell As String, CharStart As String, CharEnd As String)

    ' Where does CharStart start?
    NumStart = InStr(cell, CharStart)
        
    ' Where does CharEnd start?
    NumEnd = InStr(cell, CharEnd)

    ' Throw an error if Excel cannot find the specified characters.
    If NumStart = 0 Or NumEnd = 0 Then
    
        ' https://www.exceltip.com/custom-functions/return-error-values-from-user-defined-functions-using-vba-in-microsoft-excel.html
        FINDBETWEEN = CVErr(xlErrNA) ' #N/A error
        
        Exit Function
        
    Else

        ' To get the text inbetween CharStart and CharEnd, we need to get the positions of when CharStart ends and when CharEnd begins.
        PosStart = NumStart + Len(CharStart)
        PosEnd = NumEnd - PosStart
        
        FINDBETWEEN = Mid(cell, PosStart, PosEnd)
        
    End If

End Function

' ============================================================================
' Function: FIRSTNAME()
' Description: Get the first name (and middle name if applicable).

Function FIRSTNAME(cell As String, Optional NameOrder As Integer = 1)
    ' NOTES:
    '   1. NameOrder options
    '       1 = First Name Last Name
    '       2 = Last Name, First Name
    '   2. Reverse-order case assumes that there is a comma.
    '   3. Be careful of compound last names (e.g. Del Mul, Van Helsing, etc.)
    
    ' Remove extraneous spaces (left and right sides).
    Dim cell2 As String
    cell2 = Trim(cell) ' Have to name this cell2 because LASTNAME() also uses the "cell" argument and it will "remember" the code in FIRSTNAME().
    
    ' Regular Order
    If NameOrder = 1 Then
    
        'Remove suffixes
        If InStr(cell2, ",") Then
           
            Dim suffix As String
            suffix = FINDAFTER(cell2, ",")
            
            cell2 = FINDREMOVE(cell, suffix)
                
        ElseIf InStr(cell, " Jr") Then
            
            cell2 = FINDBEFORE(cell2, " Jr")
            
        ElseIf InStr(cell, " I") Then
        
            cell2 = FINDBEFORE(cell2, " I")
                
        End If
    
        ' To get the number of spaces, get the length of the whole cell and subtract the cell without spaces from it.
        ' This is so that we know whether to get the middle name as well.
        LenCell = Len(cell2)
        LenCellNoSpaces = Len(FINDREMOVE(cell2, " "))
        
        LenSpaces = LenCell - LenCellNoSpaces
        
        ' In the simple case (e.g. Robert Schnitman), get the text before the space.
        If LenSpaces < 2 Then
        
            FIRSTNAME = Trim(FINDBEFORE(cell2, " "))
        
        ' In the complex case (e.g. Robert Gary Schnitman), get the first and middle names separately and before concatenating them together.
        Else
            
            ' Have to use DIM to avoid VBA throwing a compile error.
            Dim first As String
            Dim MiddleLast As String
            Dim middle As String
            Dim last As String
            
            ' First name is before the first space.
            first = FINDBEFORE(cell2, " ") ' Robert
            
            ' Middle and last names are AFTER the first space.
            MiddleLast = FINDAFTER(cell2, " ") ' Gary Schnitman, Jr.
            
            ' Middle name is before the space in MiddleLast
            middle = FINDBEFORE(MiddleLast, " ") ' Gary
            
            'Last name is after the space after middle name,
            last = FINDAFTER(MiddleLast, " ")
            
            ' Output should be the concatenation of first and middle names.
            Dim fm As String
            
            fm = first & " " & middle ' Robert Gary
            
            FIRSTNAME = Trim(fm)
            
        End If
        
    ' Reverse order--ASSUMES THAT THERE IS A COMMA.
    ElseIf NameOrder = 2 Then
        
        Dim out As String
        out = Trim(FINDAFTER(LASTNAME(cell2), ","))
        
        If InStr(out, "Jr ") Or InStr(out, "JR ") Or InStr(out, "I ") Or InStr(out, "i ") Then
        
            FIRSTNAME = Trim(FINDAFTER(out, " "))
            
        Else
        
            FIRSTNAME = Trim(out)
            
        End If
    
    ' Error if NameOrder is not 1 or 2.
    Else
    
        FIRSTNAME = CVErr(xlErrValue)
        
    End If
    

End Function

' ============================================================================
' Function: LASTNAME()
' Description: Get the last name of a person.

Function LASTNAME(cell As String, Optional NameOrder As Integer = 1)
    '   1. NameOrder options
    '       1 = First Name Last Name
    '       2 = Last Name, First Name
    '   2. Be careful of compound last names (e.g. Del Mul, Van Helsing, etc.)

    ' Regular order
    If NameOrder = 1 Then
    
        ' Get the first name so that we know the part of the string that's the last name.
        Dim first As String
        first = FIRSTNAME(cell)
        
        ' Anything after the first name is the last name.
        last = FINDAFTER(cell, first)
        
        LASTNAME = Trim(last)
        
    ' Reverse order
    ElseIf NameOrder = 2 Then
    
        ' Comma situations
        If InStr(cell, ",") Then
    
            ' Get the first name so that we know the part of the string that's the last name.
            Dim first2 As String
            first2 = FIRSTNAME(cell, 2)
            
            ' Remove anything that's a part of the first name.
            Dim last2 As String
            last2 = Trim(FINDREMOVE(cell, first2))
            
            ' Additional comma left at the end behind needs to be removed.
            LASTNAME = Left(last2, Len(last2) - 1)
            
        ' Non-comma situations
        Else
            
            ' Get the first name so that we know the part of the string that's the last name.
            Dim first3 As String
            first3 = FIRSTNAME(cell, 2)
            
            ' Remove anything that's a part of the first name.
            LASTNAME = Trim(FINDREMOVE(cell, first3))
        
        End If
    
    ' Throw a value error if the NameOrder value is not 1 or 2.
    Else
    
        LASTNAME = CVErr(xlErrValue)
        
    End If

End Function
' ============================================================================

' ============================================================================
' Function: TEXTLIKE()
' Description: Determine whether a string meets at least one given pattern.

Function TEXTLIKE(cell As String, ParamArray patterns() As Variant)
    ' ParamArray allows us to give TEXTLIKE() the ability to have multiple inputs without naming them (https://docs.microsoft.com/en-us/office/vba/language/concepts/getting-started/understanding-parameter-arrays).
    
    ' Source of table below: https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/language-features/operators-and-expressions/how-to-match-a-string-against-a-pattern
    ' Characters in pattern   Matches in string
    ' ---------------------   -----------------
    ' ?                       Any single character.
    ' *                       Zero or more characters.
    ' #                       Any single digit (0-9).
    ' [ charlist ]            Any single character in charlist.
    ' [ !charlist ]           Any single character not in charlist.
    
    ' e.g TEXTLIKE("Robert Schnitman", "Robert*") ' prints TRUE.
    ' e.g TEXTLIKE("Robert Schnitman", "Craig*", "Robert*") ' prints TRUE.
    
    ' For each given pattern, see if the given string matches any of the specified patterns.
    For Each patt In patterns
    
        ' Does the string match the given pattern?
        detect = cell Like patt
        
        ' If the string matches a specified pattern, exit the loop and use the value in detect;
        '   otherwise, resume the loop until the end.
        ' If the last value is FALSE, then the detect variable will return FALSE.
        If detect = True Then
            
            Exit For
            
        End If
        
    Next
        
    ' The output of the function should be a Boolean value (TRUE/FALSE).
    TEXTLIKE = detect
    
End Function
' ============================================================================

' ============================================================================
' Function: TEXTSTRIPWS()
' Description: Remove all spaces.

Function TEXTSTRIPWS(cell As String)
    
    TEXTSTRIPWS = FINDREMOVE(cell, " ")

End Function
' ============================================================================

' ============================================================================
' Function: TEXTINSERT()
' Description: Insert a character at a specified position

Function TEXTINSERT(cell As String, findchar As String, position As Integer)

    ' The left side of the string should be everything up to just before the specified position.
    sideA = Left(cell, position - 1)
    
    ' The right side should be the concatenation of the specified character to insert AND whatever isn't captured by sideA
    sideB = findchar + Mid(cell, position)
    
    ' Output should concatenate left and right sides.
    TEXTINSERT = sideA + sideB

End Function
' ============================================================================

' ============================================================================
' Function: TEXTREVERSE(), TEXTREV()
' Description: Reverse the order of a string.

Function TEXTREVERSE(cell As String)

    TEXTREVERSE = StrReverse(cell) ' e.g. TEXTREVERSE("ABCD") = "DCBA"

End Function

' TEXTREVERSE() Synonym
Function TEXTREV(cell As String)

	TEXTREV = TEXTREVERSE(cell)
	
End Function
' ============================================================================

' ============================================================================
' Function: TEXTCOMPARE()
' Description: Compare two strings. Based on VBA's StrComp().

Function TEXTCOMPARE(string1, string2, Optional CompareType As Long = 1, Optional value As Boolean = False)

    ' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/strcomp-function
    ' CompareType = 1 --> Textual Comparison (ABCD = abcd) -- case insensitivity
    ' CompareType = 0 --> Binary Comparison (ABCD > abcd)
    
    ' RESULTS IF value = FALSE
    ' -1 --> string1 < string2
    '  0 --> string1 = string2
    '  1 --> string1 > string2
    
    ' By default, StrComp() outputs an integer.
    comp = StrComp(string1, string2, CompareType)
    
    If value = False Then
    
        output = comp
        
    ' If we want the "translated" value of what the integer means, then output the equivalent string.
    ElseIf value = True Then
    
        Select Case comp
        
            Case -1
                
                output = "<" ' string1 & " < " & string2
                
            Case 0
            
                output = "=" ' string1 & " = " & string2
                
            Case 1
            
                output = ">" ' string1 & " > " & string2
                
        End Select
        
    End If
    
    ' Output the desired value.
    TEXTCOMPARE = output


End Function
' ============================================================================

' ============================================================================
' Function: TEXTJOINR()
' Description: Join a range of strings into a single string, separated by an optional delimiter.

Function TEXTJOINR(StringRange As Range, Optional delimiter As String)

    TEXTJOINR = Application.WorksheetFunction.TextJoin(delimiter, True, StringRange)

End Function
' ============================================================================

' ============================================================================
' Function: TRIML()
' Description: Trim leading spaces.

Function TRIML(cell As String)

    TRIML = LTrim(cell)
    
End Function
' ============================================================================

' ============================================================================
' Function: TRIMR()
' Description: Trim trailing spaces.

Function TRIMR(cell As String)

    TRIMR = RTrim(cell)
    
End Function
' ============================================================================

' ============================================================================
' Function: TRIMLR()
' Description: Remove leading and trailing spaces.

Function TRIMLR(cell As String)

    TRIMLR = LTrim(RTrim(cell))

End Function
' ============================================================================

' ============================================================================
' Function: TRIMC/CTRIM()
' Description: Trim and cleanup excessive whitespace.

Function CTRIM(cell As String)

	With Application.WorksheetFunction
	
	output = .Clean(.Trim(cell))
	
	End With

	CTRIM = output
	
End Function

' CTRIM() Synonym
Function TRIMC(cell As String)

	TRIMC = CTRIM(cell)
	
End Function
' ============================================================================

' ============================================================================
' Function: RXLIKE()
' Description: Test whether a regular expression pattern has been met.

Function RXLIKE(cell As String, pattern As String, Optional IgnoreCase As Boolean = False)

    ' Make sure you have the regular expressions feature by going to Tools > References in VBA ("Microsoft VBScript Regular Expressions 5.5").
    ' https://stackoverflow.com/questions/22542834/how-to-use-regular-expressions-regex-in-microsoft-excel-both-in-cell-and-loops
    
    ' Setup
    Dim regex As New RegExp
    
    With regex
    
        .Global = True
        .MultiLine = True
        .IgnoreCase = IgnoreCase
        .pattern = pattern
        
    End With
    
    ' Outputs a Boolean value (TRUE/FALSE)
    RXLIKE = regex.Test(cell)

End Function
' ============================================================================

' ============================================================================
' Function: RXREPLACE()
' Description: Replace a string based on a regular expression pattern.

Function RXREPLACE(StringFind As String, StringPattern As String, StringNew As String, Optional IgnoreCase As Boolean = False)

    ' Make sure you have the regular expressions feature by going to Tools > References in VBA ("Microsoft VBScript Regular Expressions 5.5").
    ' https://stackoverflow.com/questions/22542834/how-to-use-regular-expressions-regex-in-microsoft-excel-both-in-cell-and-loops
    
    ' Setup
    Dim regex As New RegExp
    
    With regex
    
        .Global = True
        .MultiLine = True
        .IgnoreCase = IgnoreCase
        .pattern = StringPattern
        
    End With
    
    ' If the string matches the given pattern, replace it with the new string; otherwise, throw an error.
    If regex.Test(StringFind) = True Then
    
        output = regex.Replace(StringFind, StringNew)
        
    Else
    
        output = CVErr(xlErrNA)
        
    End If
    
    RXREPLACE = output

End Function
' ============================================================================

' ============================================================================
' Function: RXREMOVE()
' Description: Remove a string based on a regular expression pattern.

Function RXREMOVE(StringFind As String, StringPattern As String, Optional IgnoreCase As Boolean = False)

	RXREMOVE = RXREPLACE(StringFind, StringPattern, "", IgnoreCase)

End Function
' ============================================================================

' ============================================================================
' Function: RXGET()
' Description: Extract the first text that meets a regular expression pattern.

Function RXGET(cell As String, pattern As String, Optional IgnoreCase As Boolean = False)

    ' Make sure you have the regular expressions feature by going to Tools > References in VBA ("Microsoft VBScript Regular Expressions 5.5").
    ' https://stackoverflow.com/questions/22542834/how-to-use-regular-expressions-regex-in-microsoft-excel-both-in-cell-and-loops
    
    ' Setup
    Dim regex As New RegExp
    
    With regex
    
        .Global = True
        .MultiLine = True
        .IgnoreCase = IgnoreCase
        .pattern = pattern
        
    End With
    
    ' If the string matches the pattern, produce the first match; otherwise, throw an error.
    If regex.Test(cell) = True Then
    
       Set matches = regex.Execute(cell)
       
       output = matches.Item(0)
        
    Else
    
        output = CVErr(xlErrNA)
        
    End If
    
    RXGET = output

End Function
' ============================================================================

' ============================================================================
' Function: RXGETALL()
' Description: Extract ALL text that meet a regular expression pattern.

Function RXGETALL(cell As String, pattern As String, Optional sep As String = ",", Optional IgnoreCase As Boolean = False)

    ' Make sure you have the regular expressions feature by going to Tools > References in VBA ("Microsoft VBScript Regular Expressions 5.5").
    ' https://stackoverflow.com/questions/22542834/how-to-use-regular-expressions-regex-in-microsoft-excel-both-in-cell-and-loops
    
    ' Setup
    Dim regex As New RegExp
    
    With regex
    
        .Global = True
        .MultiLine = True
        .IgnoreCase = IgnoreCase
        .pattern = pattern
        
    End With
    
    ' Get all matches.
    Set matches = regex.Execute(cell)
    
    ' https://stackoverflow.com/questions/8146485/returning-a-regex-match-in-vba-excel

    ' Join all matches in a single string, separated by a delimiter ("sep").
    For i = 0 To matches.Count - 1
    
        output = output & sep & matches.Item(i)
        
    Next
    
    ' The concatenation loop above always puts the separator in the first position of the string, so we need to remove it.
    If Len(output) <> 0 Then
    
        output = Right(output, Len(output) - Len(sep))
        
    End If
    
    RXGETALL = output

End Function
' ============================================================================