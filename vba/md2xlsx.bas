Function Md2AdocXlsx(ByVal codeLine As String) As String

    Dim myMatches As SubMatches
    Dim r As Boolean
    Dim outString As String
    Dim splitter As String
    splitter = "|"
    
    r = RegMatches("(^#+)\s(.+)", codeLine, myMatches)
    If (r) Then
        outString = myMatches(0) & splitter & splitter & myMatches(1)
    End If

    r = RegMatches("(^\*)\s(.+)", codeLine, myMatches)
    If (r) Then
        outString = myMatches(0) & splitter & splitter & myMatches(1)
    End If
    
    r = RegMatches("(^\s+\*)\s(.+)", codeLine, myMatches)
    If (r) Then
        outString = "**" & splitter & splitter & myMatches(1)
    End If
    
    r = RegMatches("(`+)", codeLine, myMatches)
    If (r) Then
        outString = "code"
    End If

    If outString = "" Then
        outString = splitter & splitter & codeLine
    End If
    
    Md2AdocXlsx = outString
End Function

Function RegMatches(myPattern As String, myString As String, myMatches As SubMatches) As Boolean
    RegMatches = False

    'Create objects.
    Dim objRegExp As RegExp
    Dim objMatch As Match
    Dim colMatches As MatchCollection
    Dim RetStr As String

    ' Create a regular expression object.
    Set objRegExp = New RegExp

    'Set the pattern by using the Pattern property.
    objRegExp.Pattern = myPattern

    ' Set Case Insensitivity.
    objRegExp.IgnoreCase = True

    'Set global applicability.
    objRegExp.Global = True
    
    
    Dim matchval As String
    Dim replaceVal As String
    
    Dim retVal As String
    retVal = myString
    
    Dim myMatch As String

    'Test whether the String can be compared.
    If (objRegExp.Test(myString) = True) Then
        RegMatches = True
        Set colMatches = objRegExp.Execute(myString)   ' Execute search.]
        
        For Each objMatch In colMatches   ' Iterate Matches collection.
            Set myMatches = objMatch.SubMatches
        Next
    End If
End Function