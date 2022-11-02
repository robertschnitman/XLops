Attribute VB_Name = "FunDates"
' ============================================================================
' REPO: XLops
' MODULE: FunDates.bas
' DESCRIPTION: Functions for parsing dates.
'
' LIST OF FUNCTIONS:
' WEEKDAYN()
' YMD()
' MDY()
' DMY()
' ============================================================================

' ============================================================================
' Function: WEEKDAYN()
' Description: Outputs the name of the weekday for a given date.

Function WEEKDAYN(d As Date)
    
    WEEKDAYN = WEEKDAYNAME(d)

End Function
' ============================================================================

' ============================================================================
' Function: YMD()
' Description: Formats a date value into the ISO standard format ("yyyy-mm-dd").

Function YMD(d As Date, Optional sep As String = "-")

    YMD = Format(d, "yyyy" & sep & "mm" & sep & "dd")

End Function
' ============================================================================

' ============================================================================
' Function: MDY()
' Description: Formats a date value into the month-day-year order ("mm/dd/yyyy").

Function MDY(d As Date, Optional sep As String = "/")

    MDY = Format(d, "mm" & sep & "dd" & sep & "yyyy")

End Function
' ============================================================================

' ============================================================================
' Function: DMY()
' Description: Formats a date value into the day-month-year order ("dd/mm/yyyy").

Function DMY(d As Date, Optional sep As String = "/")

    DMY = Format(d, "dd" & sep & "mm" & sep & "yyyy")

End Function
' ============================================================================