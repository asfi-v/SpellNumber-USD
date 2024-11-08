# SpellNumber-USD
This repository contains a VBA (Visual Basic for Applications) script that converts numerical values into their corresponding English words. The script is designed to handle large numbers and includes functionality to append “Only” at the end of the converted text, making it suitable for financial documents and other applications where precise wording is required.

This function is designed to be used in Excel VBA projects but can be adapted for other VBA environments.

## Features
- Converts numbers to words, including large values up to trillions.
- Handles both whole numbers and decimal values.
- Appends "Only" at the end of the converted text for clarity.
- Easy to integrate into Excel and other VBA-supported applications.
- Handles Large Numbers: Converts numbers up to the trillions.
- Decimal Support: Accurately spells out the cents portion for decimal values.

## Usage
1. Copy the VBA code from the repository.
2. Paste it into your VBA editor in Excel or another application.
3. Call the `SpellNumber` function with the number you want to convert.

### VBA Code
```vba
Option Explicit
'Main Function
Function SpellNumber(ByVal MyNumber)
        Dim Dollars, Cents, Temp
        Dim DecimalPlace, Count
        ReDim Place(9) As String
        Place(2) = " Thousand "
        Place(3) = " Million "
        Place(4) = " Billion "
        Place(5) = " Trillion "
 
        MyNumber = Trim(Str(MyNumber))
        DecimalPlace = InStr(MyNumber, ".")
        If DecimalPlace > 0 Then
                Cents = GetTens(Left(Mid(MyNumber, DecimalPlace + 1) & _
                                    "00", 2))
                MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
        End If
        Count = 1
        Do While MyNumber <> ""
                Temp = GetHundreds(Right(MyNumber, 3))
                If Temp <> "" Then Dollars = Temp & Place(Count) & Dollars
                If Len(MyNumber) > 3 Then
                        MyNumber = Left(MyNumber, Len(MyNumber) - 3)
                Else
                        MyNumber = ""
                End If
                Count = Count + 1
        Loop
        If Dollars <> "" Then
            Select Case Dollars
                    Case "One"
                            Dollars = "One Dollar"
                     Case Else
                            Dollars = Dollars & " Dollars"
            End Select
        End If
        If Cents <> "" Then
            Select Case Cents
                    Case "One"
                            Cents = "One Cent"
                    Case Else
                            Cents = Cents & " Cents"
            End Select
        End If
        If Dollars = "" Then
            SpellNumber = Cents & " Only"
        ElseIf Cents = "" Then
            SpellNumber = Dollars & " Only"
        Else
            SpellNumber = Dollars & " and " & Cents & " Only"
        End If
End Function
 
Function GetHundreds(ByVal MyNumber)
        Dim Result As String
        If Val(MyNumber) = 0 Then Exit Function
        MyNumber = Right("000" & MyNumber, 3)
        ' Convert the hundreds place.
        If Mid(MyNumber, 1, 1) <> "0" Then
                Result = GetDigit(Mid(MyNumber, 1, 1)) & " Hundred "
        End If
        ' Convert the tens and ones place.
        If Mid(MyNumber, 2, 1) <> "0" Then
                Result = Result & GetTens(Mid(MyNumber, 2))
        Else
                Result = Result & GetDigit(Mid(MyNumber, 3))
        End If
        GetHundreds = Result
End Function
 
Function GetTens(TensText)
        Dim Result As String
        Result = "" ' Null out the temporary function value.
        If Val(Left(TensText, 1)) = 1 Then   ' If value between 10-19…
                Select Case Val(TensText)
                        Case 10: Result = "Ten"
                        Case 11: Result = "Eleven"
                        Case 12: Result = "Twelve"
                        Case 13: Result = "Thirteen"
                        Case 14: Result = "Fourteen"
                        Case 15: Result = "Fifteen"
                        Case 16: Result = "Sixteen"
                        Case 17: Result = "Seventeen"
                        Case 18: Result = "Eighteen"
                        Case 19: Result = "Nineteen"
                        Case Else
                End Select
        Else ' If value between 20-99…
                Select Case Val(Left(TensText, 1))
                        Case 2: Result = "Twenty "
                        Case 3: Result = "Thirty "
                        Case 4: Result = "Forty "
                        Case 5: Result = "Fifty "
                        Case 6: Result = "Sixty "
                        Case 7: Result = "Seventy "
                        Case 8: Result = "Eighty "
                        Case 9: Result = "Ninety "
                        Case Else
                End Select
                Result = Result & GetDigit _
                        (Right(TensText, 1))  ' Retrieve ones place.
        End If
        GetTens = Result
End Function
 
Function GetDigit(Digit)
        Select Case Val(Digit)
                Case 1: GetDigit = "One"
                Case 2: GetDigit = "Two"
                Case 3: GetDigit = "Three"
                Case 4: GetDigit = "Four"
                Case 5: GetDigit = "Five"
                Case 6: GetDigit = "Six"
                Case 7: GetDigit = "Seven"
                Case 8: GetDigit = "Eight"
                Case 9: GetDigit = "Nine"
                Case Else: GetDigit = ""
        End Select
End Function
```

## Example
1. Type the formula =SpellNumber(A1) into the cell where you want to display a written number, where A1 is the cell containing the number you want to convert. You can also manually type the value like =SpellNumber(1234.56)
2. Press Enter to confirm the formula.
   
   ![image](https://github.com/user-attachments/assets/8d027f05-8261-4249-b8bc-48fda60b526d)
   
5. Save your SpellNumber function workbook.
6. _Excel cannot save a workbook with macro functions in the standard macro-free workbook format (.xlsx). If you click File > Save. A VB project dialog box opens. Click No._
   


