Attribute VB_Name = "ReadCSV"

Public Function ParseCSVLine(strSource As String, Optional strDelimiter As String = ",", Optional strQuote As String = """") As Variant()
    'Function parses 1 line of text (from a CSV file), returning a variant array
    'containing the column values stored in the CSV.
    'Optionally may pass the delimiter that separates fields (default is comma),
    'and/or the text string quote character (default is ").
    'Note: Strings in the CSV do not have to be contained within quotes as long
    '      as the delimiter character is not part of the string.
    
    'All fields (including strings in quotes, and numbers) will be stored as
    'variant.  Fields will range from array(0) to array(UBound(array)).
    'To assign explicit values to your var's, use the following method:
    '   A = CInt(array(0)) 'Integer
    '   B = CStr(array(1)) 'String
    '  etc.
        
    Dim intTest As Integer
    Dim intCount As Integer, intEnd As Integer
    Dim parseText As String, chunk As String
    Dim varHold() As Variant
    
    'initialize for new array
    intCount = 0: intEnd = 1
    parseText = strSource
    ReDim varHold(0)

    'process fields until no more delimiters found
    Do While intEnd > 0
        
        If Len(parseText) > 0 And Left(LTrim(parseText), 1) = strQuote Then
            '----------------------------
            'Process quoted fields here!
            '----------------------------
            parseText = LTrim(parseText)
            If Len(parseText) > 1 Then
                'Find ending quote
                intEnd = InStr(2, parseText, strQuote)
                If intEnd = 0 Then intEnd = Len(parseText) + 1 '<-last field
                'Extract field value
                If intEnd = 2 Then chunk = "" Else chunk = Mid(parseText, 2, intEnd - 2)
                If intEnd < Len(parseText) Then
                    'Find next delimiter
                    intEnd = InStr(intEnd + 1, parseText, strDelimiter)
                Else
                    'If no delimiter, then end parsing
                    intEnd = 0
                End If
            Else
                'if opening quote is last character, set last field blank & end parsing
                chunk = "": intEnd = 0
            End If
        Else
            '------------------------------
            'Process non-quoted fields here!
            '------------------------------
            If Len(parseText) > 0 Then
                'Find next delimiter
                intEnd = InStr(1, parseText, strDelimiter)
                If intEnd = 0 Then intEnd = Len(parseText) + 1
                'Extract field value
                If intEnd = 1 Then chunk = "" Else chunk = Left(parseText, intEnd - 1)
                If intEnd > Len(parseText) Then intEnd = 0  'detect end of string
            Else
                'If last field is blank, set it and end parsing
                chunk = "": intEnd = 0
            End If
        End If
        
        'Remove current field from parsing string
        If intEnd = Len(parseText) Or intEnd = 0 Then
            parseText = ""
        Else
            parseText = Right(parseText, Len(parseText) - intEnd)
        End If
            
        'increase the array and store new field
        If intCount > 0 Then ReDim Preserve varHold(UBound(varHold) + 1)
        varHold(UBound(varHold)) = CVar(chunk)
        
        intCount = intCount + 1 'increment record count
    Loop
    
    'Assign temp array to function value:
    ParseCSVLine = varHold
    
    'for debugging to the immediate window
    'For intTest = LBound(varHold) To UBound(varHold)
    '    Debug.Print "#" & intTest & ": " & varHold(intTest)
    'Next
End Function

