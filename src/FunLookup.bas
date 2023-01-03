Attribute VB_Name = "FunLookup"
' ============================================================================
' REPO: XLops
' MODULE: FunLookup.bas
' DESCRIPTION: Simplified lookup functions.
'
' LIST OF FUNCTIONS:
' FLOOKUP()
' SLOOKUP()
' SMATCH()
' INDEX0()
' INDEXH()
' INDEXH0()
' INDEXM()/INDEXMATCH()
' INDEX0()
' INDEXM0()/INDEXMATCH0()
' ============================================================================

' ============================================================================
' Function: FLOOKUP()
' DESCRIPTION: "Flexible Lookup", a simpler Index-Match formula.
Function FLOOKUP(IDLookup As Variant, _
                 DataRange As Variant, _
                 NamesPattern As Variant, _
                 NamesRange As Variant, _
                 Optional IDApproxMatch As Variant = False, _
                 Optional ColMatchType As Variant = 2)
				 
    ' VLOOKUP's arguments require Variants (except last, which is to be boolean)
    ' https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetfunction.vlookup				 
    With Application.WorksheetFunction
    
        Dim ColIndex, Output As Variant
    
        ColIndex = .XMatch(NamesPattern, NamesRange, ColMatchType)
        
        Output = .VLookup(IDLookup, DataRange, ColIndex, IDApproxMatch)
    
    End With
    
    FLOOKUP = Output

End Function
' ============================================================================

' ============================================================================
' Function: SLOOKUP()
' DESCRIPTION: Simplified VLOOKUP() with desired column specified as a string.
Function SLOOKUP(IDLookup As Variant, _
				DataRange As Variant, _
				FieldName As Variant, _
				Optional IDApproxMatch As Variant = False, _
				Optional ColMatchType As Variant = 2)

    With Application.WorksheetFunction
        
        ' VLOOKUP requires column number, which can be found with the XMatch function and the Rows property
        ColIndex = .XMatch(FieldName, DataRange.Rows(1), ColMatchType)
    
        Output = .VLookup(IDLookup, DataRange, ColIndex, IDApproxMatch)
    
    End With
    
    SLOOKUP = Output

End Function
' ============================================================================

' ============================================================================
' FUNCTION: SMATCH()
' DESCRIPTION: Simplified XMATCH(..., ..., 2)
Function SMATCH(pattern As String, RangeRef As Range)

    ' https://support.microsoft.com/en-us/office/xmatch-function-d966da31-7a6b-4a13-a1c6-5a33ed6a0312
    ' 2 = Wildcard match
    SMATCH = Application.WorksheetFunction.XMatch(pattern, RangeRef, 2)
    
End Function
' ============================================================================

' ============================================================================
' FUNCTION: INDEX0()
' DESCRIPTION: Output another value if INDEX() result is zero or a zero-length string.
Function INDEX0(DataRange, RowNum As Integer, ColNum As Integer, Optional ElseValue As String = "")

        ' Convert range to array for Index() to work.
        Dim ArrRef As Variant
        ArrRef = DataRange

        ' Initial Output
        output = Application.WorksheetFunction.Index(ArrRef, RowNum, ColNum)
        
        ' Convert Output based on ElseValue
        If Application.WorksheetFunction.IsText(output) Then
        
            If Len(output) = 0 Then
            
                output = ElseValue
                    
            End If
            
        Else
        
            If CInt(output) = 0 Then
            
                output = ElseValue
                
            End If
        
        End If
        
        ' Final Output
        INDEX0 = output

End Function
' ============================================================================

' ============================================================================
' FUNCTION: INDEXH()
' DESCRIPTION: Get Header value in range.
Function INDEXH(DataRef As Range, DataHeader As Range, pattern As String)

    ' The Index() function requires an Array as an input; however, we need to be able to select a range of data.
    ' So, we convert the Ranges into an Array for Index() to work.
    Dim ArrRef, ArrHeader As Variant
    ArrRef = DataRef
    ArrHeader = DataHeader

    ' Regular INDEX calculation with XMATCH
    xm = SMATCH(pattern, ArrHeader)
        
    Output = Application.WorksheetFunction.Index(ArrRef, 1, xm)
    
    INDEXH = Output
    
 
End Function
' ============================================================================

' ============================================================================
' FUNCTION: INDEXH0()
' DESCRIPTION: Get Header value in range. If zero or length zero, Output another value.
Function INDEXH0(DataRef As Range, _
				DataHeader As Range, _
				pattern As String, _
				Optional ElseValue As String = "")

    ' Initial Output
    Output = INDEXH(DataRef, DataHeader, pattern)
        
    ' Convert initial value based on ElseValue
    If Len(Output) = 0 Or CInt(Output) = 0 Then
	
		Output = ElseValue                
    
	End If
        
    ' Final result
    INDEXH0 = Output    
 
End Function
' ============================================================================

' ============================================================================
' FUNCTION: INDEXM()/INDEXMATCH()
' DESCRIPTION: Simplified index-matching [INDEX(..., MATCH(...), MATCH(...))].
Function INDEXMATCH(DataRange, _
					LookupVal1 As String, _
					MatchRange1 As Range, _
					LookupVal2 As String, _
					MatchRange2 As Range)

    Dim ArrRef As Variant
    ArrRef = DataRange

    With Application.WorksheetFunction
        
		RowNum = SMATCH(LookupVal1, MatchRange1)
			
		ColNum = SMATCH(LookupVal2, MatchRange2)
				
		Output = .Index(ArrRef, RowNum, ColNum)
        
    End With
        
    INDEXMATCH = Output

End Function

' INDEXMATCH() Synonym
Function INDEXM(DataRange, _
				LookupVal1 As String, _
				MatchRange1 As Range, _
				LookupVal2 As String, _
				MatchRange2 As Range)
        
	INDEXM = INDEXMATCH(DataRange, LookupVal1, MatchRange1, LookupVal2, MatchRange2)

End Function
' ============================================================================

' ============================================================================
' FUNCTION: INDEXM0()/INDEXMATCH0()
' DESCRIPTION: Output another value if INDEXM() result is zero or a zero-length string.
Function INDEXM0(DataRange, _
				LookupVal1 As String, _
				MatchRange1 As Range, _
				LookupVal2 As String, _
				MatchRange2 As Range, _
				Optional ElseValue As String = "")
        
	' Initial Output
    Output = INDEXM(DataRange, LookupVal1, MatchRange1, LookupVal2, MatchRange2)
        
    ' Convert Output based on ElseValue
    If Len(Output) = 0 Or CInt(Output) = 0 Then

        Output = ElseValue

	End If
        
	' Final Output
    INDEX0 = Output

End Function

'INDEXM0() Synonym
Function INDEXMATCH0(DataRange, _
				LookupVal1 As String, _
				MatchRange1 As Range, _
				LookupVal2 As String, _
				MatchRange2 As Range, _
				Optional ElseValue As String = "")

	INDEXMATCH0 = INDEXM0(DataRange, LookupVal1, MatchRange1, LookupVal2, MatchRange2, ElseValue)

End Function
' ============================================================================