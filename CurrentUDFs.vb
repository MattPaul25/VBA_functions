Function InRange(cell As String, myRange As range, Optional strict As Boolean = False) As Boolean
''checks wheather item exist in range - strict as true is an and criteria | all items must equal cell
If strict Then
    For Each citem In myRange
            If citem <> cell Then
                    InRange = False
                    Exit For
            End If
        InRange = True
    Next citem
Else
    For Each citem In myRange
            If citem = cell Then
                    InRange = True
                    Exit For
            End If
        InRange = False
    Next citem
End If
End Function
Function SearchInstance(cell As String, searchString As String, inst As Integer) As Integer
'can search for second or third item in string
instCount = 0
    For i = 1 To Len(cell)
            If Mid(cell, i, Len(searchString)) = searchString Then
                instCount = instCount + 1
                    If instCount = inst Then
                        SearchInstance = i
                        Exit For
                    End If
             End If
    Next i
End Function
Function ConcatRange(ByVal rnge As range, ByVal sep As String) As String
'Concats range and separates each item with a designated separator
r1 = ""
    For Each cell In rnge
            If cell <> "" Then
                r = cell & sep & " "
                r1 = r1 & r
             End If
    Next cell
 r1 = Left(r1, Len(r1) - Len(sep))
ConcatRange = r1
End Function
 Function RegExGet(aString, myExpression) As Variant
'function used to create array of regular expression results
'Requires pooint at vbscript_regexp library
    Dim regEx As New VBScript_RegExp_55.RegExp
     Dim newArray() As String
     Dim cnt As Integer
    regEx.Pattern = myExpression
    regEx.IgnoreCase = False
    regEx.Global = True
    s = ""
    Set matches = regEx.Execute(aString)
    x = matches.Count
    ReDim newArray(x - 1) As String
    cnt = 0
        For Each Match In matches
            newArray(cnt) = Match.Value
            cnt = cnt + 1
        Next
        RegExGet = newArray()
End Function
Function RegExPosition(aString, myExpression) As Variant
'function used to create array of regex positions in string
'Requires pooint at vbscript_regexp library
    Dim regEx As New VBScript_RegExp_55.RegExp
     Dim newArray() As Integer
     Dim cnt As Integer
    regEx.Pattern = myExpression
    regEx.IgnoreCase = False 'True to ignore case
    regEx.Global = True 'True matches all occurances, False matches the first occurance
    s = ""
    Set matches = regEx.Execute(aString)
    x = matches.Count
    ReDim newArray(x - 1) As Integer
    cnt = 0
        For Each Match In matches
            newArray(cnt) = Match.FirstIndex + 1
            cnt = cnt + 1
        Next
        RegExPosition = newArray()
End Function
Function SplitUp(MyString As String, Optional MySep As String = "", Optional vertical As Boolean = False) As Variant
'doesnt use vba split function - but does the same thing except in array format
Dim newArray() As String
Dim j As Integer
Dim i As Integer
If MySep = "" Then
x = Len(MyString)
Else
 x = ((Len(MyString) - Len(Replace(MyString, MySep, ""))) / Len(MySep)) + 1
 End If
 ReDim newArray(x - 1) As String
 For i = 1 To x
     If Len(MyString) - Len(Replace(MyString, MySep, "")) = 0 And MySep <> "" Then
        newArray(i - 1) = MyString
        Exit For
    End If
    For j = 1 To Len(MyString)
        If MySep = "" Then
            newArray(i - 1) = Left(MyString, 1)
            MyString = Mid(MyString, 2, Len(MyString) - 1)
            Exit For
        End If
        If Mid(MyString, j, Len(MySep)) = MySep Then
            newArray(i - 1) = Mid(MyString, 1, j - 1)
            MyString = Mid(MyString, j + Len(MySep), Len(MyString))
            Exit For
        End If
    Next j
 Next i
 If vertical Then
 SplitUp = Application.Transpose(newArray())
 Else
 SplitUp = newArray()
 End If
End Function
 Function ConcatIf(ConcatRange As Variant, criteriaRange As Variant, criteria As String, MySep As String) As String
'concats items in a range based on criteria | works much like a sum if
Dim currentString As String
currentString = ""
    For i = 1 To ConcatRange.Count
            If criteriaRange(i) = criteria Then
            currentString = currentString & ConcatRange(i) & MySep
            End If
    Next i
currentString = Left(currentString, Len(currentString) - Len(MySep))
ConcatIf = currentString
End Function
Function ConcatUnique(ConcatRange As Variant, MySep As String) As String
'concats unique items | fails when items are small parts of larger previously concatenated items - will fix
Dim x As Integer
Dim currentString As String
currentString = ""
x = 1
Do While x <= ConcatRange.Count
    If InStr(1, currentString, ConcatRange(x)) = 0 Then
        currentString = currentString & ConcatRange(x) & MySep
    End If
    x = x + 1
Loop
currentString = Left(currentString, Len(currentString) - Len(MySep))
ConcatUnique = currentString
End Function

 Function CountString(FullString As String, PartialString As String) As Integer
 'counts the amount of times one string occurs within another string | good for instance perameters
 Dim cnt As Integer
 cnt = 0
 For i = 1 To Len(FullString)
    If Mid(FullString, i, Len(PartialString)) = PartialString Then
    cnt = cnt + 1
    End If
Next i
CountString = cnt
 End Function
 
Function ReptSep(item, Optional num = 2, Optional sep = "")
'if you need to repeat a function this can do it without rewriting the memory intensive function
x = 1
rItem = item
Do While x < num
    rItem = rItem & sep & item
    x = x + 1
Loop
ReptSep = rItem
End Function
Function countWord(cell)
'counts the words of a cell to get a word count
cell = Application.WorksheetFunction.Trim(cell)
Count = 1
For i = 1 To Len(cell)
If Mid(cell, i, 1) = " " Then
Count = Count + 1
End If
Next i
countWord = Count
End Function
Function ConvertFromDec(number, ToBase)
'convers number in dec to another base |not perfected for numbers over 10
NewNum = ""
If number = 0 Then
    ConvertFromDec = "0"
    Exit Function
End If
Do While number >= 1
    Remainder = number Mod ToBase
    number = (number - Remainder) / ToBase
    If Len(Remainder & "") > 1 Then
        NewNum = NewNum & " and " & Remainder
    Else
        NewNum = Remainder & NewNum
    End If
Loop
ConvertFromDec = NewNum
End Function
Function ConvertToDec(conValue, base)
'converts non base 10 values to base 10
For i = 0 To Len(conValue) - 1
    x = x + (Application.WorksheetFunction.Power(base, i) * Mid(conValue, Len(conValue) - i, 1))
Next i
ConvertToDec = x
End Function
Function ConvertBase(Value, FromBase, ToBase)
'converts any base to any base | not great over base 10...
If Value = 0 Then
    ConvertBase = "0"
    Exit Function
End If
For i = 0 To Len(Value) - 1
    x = x + (Application.WorksheetFunction.Power(FromBase, i) * Mid(Value, Len(Value) - i, 1))
Next i
Value = x
NewNum = ""
Do While Value >= 1
    Remainder = Value Mod ToBase
    Value = (Value - Remainder) / ToBase
    NewNum = Remainder & NewNum
Loop
    ConvertBase = NewNum
End Function
Function UniqueArr(aRange)
'outputs an array of unique items from an array with duplicates
cnt = 1
Dim newArr() As Variant
aRange = SortArray(aRange)
rMax = UBound(aRange)
For i = 2 To rMax
    If aRange(i) <> aRange(i - 1) Then cnt = cnt + 1
Next i
ReDim newArr(cnt - 1) As Variant
newArr(0) = aRange(1)
For i = 2 To rMax
    If aRange(i) <> aRange(i - 1) Then
        j = j + 1
        newArr(j) = aRange(i)
    End If
Next i
UniqueArr = newArr
End Function
Function CountUniques(aRange)
 'counts the unique items in list
cnt = 1
aRange = SortArray(aRange)
rMax = UBound(aRange)
For i = 2 To rMax
    If aRange(i) <> aRange(i - 1) Then cnt = cnt + 1
Next i
CountUniques = cnt
End Function
Function BubbleSort(aRange)
 'basic bubble sort - should make a more efficient sort algorithm - this is processor intensive
aRange = Excel.WorksheetFunction.Transpose(aRange)
rMin = LBound(aRange)
rMax = UBound(aRange)
  For i = rMin To rMax - 1
        For j = i + 1 To rMax
            If aRange(i) > aRange(j) Then
              strTemp = aRange(i)
              aRange(i) = aRange(j)
              aRange(j) = strTemp
            End If
        Next j
  Next i
SortArray = aRange
End Function
