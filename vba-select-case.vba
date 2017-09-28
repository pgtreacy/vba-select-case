Option Explicit

' ************************************************************
' https://www.myonlinetraininghub.com/vba-select-case
' ************************************************************
' Some simple examples of using SELECT CASE ELSE
'
' To run the code, click into a sub with your mouse,
' then press F8 to step through each sub, line by line,
' and see the code working
'
' Make sure your Immediate window is visible : Press CTRL + G
' ************************************************************



Sub select_case_number()

    Dim MyNum As Long
    
    MyNum = 6
    
    Select Case MyNum

        Case 1
            Debug.Print "One"

        Case 2
            Debug.Print "Two"

        Case 3
            Debug.Print "Three"

        Case 4
            Debug.Print "Four"

        Case 5
            Debug.Print "Five"

        Case Else
            Debug.Print "Greater than Five"

    End Select

End Sub

Sub select_case_numbers()

    Dim MyNum As Long
    
    MyNum = 2
    
    Select Case MyNum

        Case 1, 3, 5
            Debug.Print "Odd"

        Case 2, 4
            Debug.Print "Even"

    End Select

End Sub

Sub select_case_string()

    Dim MyStr As String
    
    MyStr = "Three"
    
    Select Case MyStr

        Case "One"
            Debug.Print 1

        Case "Two"
            Debug.Print 2

        Case "Three"
            Debug.Print 3
        
        Case "Four"
            Debug.Print 4
        
        Case "Five"
            Debug.Print 5
    
    End Select

End Sub

Sub select_case_strings()

    Dim MyStr As String
    
    MyStr = "Three"
    
    Select Case MyStr

        Case "One", "Three", "Five"
            Debug.Print "Odd"

        Case "Two", "Four"
            Debug.Print "Even"
    
    End Select

End Sub

Sub select_case_is()

    Dim MyNum As Long
    
    MyNum = 75
    
    Select Case MyNum

        Case Is > 100
            Debug.Print "Greater than 100"

        Case Is > 75
            Debug.Print "Greater than 75"

        Case Is > 50
            Debug.Print "Greater than 50"
            
        Case Is > 25
            Debug.Print "Greater than 25"
        
        Case Is > 0
            Debug.Print "Greater than 0"
            
        Case Else
            Debug.Print "Less than or equal to 0"
    
    End Select

End Sub


Sub select_case_to_num()

    Dim MyNum As Long
    
    MyNum = 50
    
    Select Case MyNum

        Case 51 To 100 ' Must be low number to high number
            Debug.Print "> 50 and <= 100"

        Case 0 To 50
            Debug.Print ">= 0 and <= 50"
            
    End Select

End Sub

Sub select_case_to_char()

    Dim MyStr As String
    
    MyStr = "f"
    
    Select Case MyStr

        Case "a" To "m"
            Debug.Print "a to m"

        Case "n" To "z"
            Debug.Print "n to z"
            
    End Select

End Sub



Sub select_case_mixed_test()

    Dim MyNum As Long
    
    MyNum = -9
    
    Select Case MyNum

        Case Is > 100
            Debug.Print "Greater than 100"

        Case 50 To 100
            Debug.Print ">= 50 and <= 100"

        Case 1 To 49
            Debug.Print ">= 1 and <= 49"
                  
        Case 0
            Debug.Print "Zero"
                        
        Case Else
            Debug.Print "Less than 0"
    
    End Select

End Sub

Sub select_case_variant()

    Dim MyVar As Variant
    
    MyVar = "One"
    
    Select Case MyVar

        Case "One"
            Debug.Print "One"
        
        Case 1
            Debug.Print "1"

    End Select

End Sub



