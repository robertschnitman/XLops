Attribute VB_Name = "FunLogic"
' ============================================================================
' REPO: XLops
' MODULE: FunLogic.bas
' DESCRIPTION: Boolean functions for common scenarios.
'
' LIST OF FUNCTIONS:
' ISLEN0()
' IS0()
' IFBLANK()
' IF0()
' SKIPBLANK()
' DOIF()
' ISMAC()
' ============================================================================

' ============================================================================
' Function: ISLEN0()
' Description: Tests whether a string is of zero length.
Function ISLEN0(cell As String)
    
    ISLEN0 = (Len(cell) = 0)

End Function
' ============================================================================

' ============================================================================
' Function: IS0()
' Description: Tests whether a cell is zero or a zero-length string.
Function IS0(cell As String)
    
    IS0 = (Len(cell) = 0 or CInt(cell) = 0)

End Function
' ============================================================================

' ============================================================================
' Function: IFBLANK()
' Description: Similar to IF(), but performs an action depending on whether a cell is blank or not.

Function IFBLANK(cell As String, ValTrue, ValElse)

    If ISLEN0(cell) = True Then
    
        output = ValTrue
        
    Else
    
        output = ValElse
        
    End If
    
    IFBLANK = output

End Function

' ============================================================================
' Function: IF0()
' Description: Perform an action depending on whether a cell is zero or length zero.

Function IF0(cell As String, ValTrue)

    If ISLEN0(cell) = True or CInt(cell) = 0 Then
    
        output = ValTrue
        
    Else
    
        output = cell
        
    End If
    
    IF0 = output

End Function
' ============================================================================

' ============================================================================
' Function: SKIPBLANK()
' Description: Perform an action if a cell is non-blank; otherwise, output blank.

Function SKIPBLANK(cell As String, ValDefine)

    If ISLEN0(cell) = True Then
    
        output = ""
        
    Else
    
        output = ValDefine
        
    End If
    
    SKIPBLANK = output

End Function
' ============================================================================

' ============================================================================
' Function: DOIF()
' Description: Perform an action only if a condition is met; otherwise, output blank.

Function DOIF(condition As Boolean, ValTrue)

    If condition = True Then
    
        output = ValTrue
        
    Else
    
        output = ""
        
    End If
    
    DOIF = output

End Function
' ============================================================================

' ============================================================================
' Function: ISMAC()
' Description: Test whether the user's computer is a Mac.

Function ISMAC()

    If Application.OperatingSystem Like "Mac*" Then
    
        output = True
        
    Else
    
        output = False
        
    End If
    
    ISMAC = output

End Function
' ============================================================================