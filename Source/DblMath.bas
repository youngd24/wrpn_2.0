Attribute VB_Name = "DblMath"
' This is an 128 bit add-on to the BigInt math libarary specifically for the
' wrpn project.  It demonstrates the extensibility of the library.

Private Const MAX_BYTE = 15

Type Big128Int
    n(0 To MAX_BYTE) As Byte
End Type

Option Explicit
Public Sub DblMult(x As BigInt, y As BigInt)
    ' multiply
    Dim i As Integer
    Dim x128 As Big128Int
    Dim y128 As Big128Int
    Dim ans128 As Big128Int
    
    ' convert 64 bit integers to 128 bit integers
    For i = 0 To 7
        x128.n(i) = x.n(i)
        y128.n(i) = y.n(i)
    Next
    
    ' do the math
    ans128 = Big128Mult(x128, y128)
    
    ' break it up into "word_size" chunks
    For i = 0 To 7
        y.n(i) = ans128.n(i) And word_mask.n(i)
    Next
    For i = 1 To word_size
        ans128 = Big128Right(ans128)
    Next
    For i = 0 To 7
        x.n(i) = ans128.n(i) And word_mask.n(i)
    Next
End Sub
Public Function DblDiv(x As BigInt, y As BigInt, z As BigInt) As BigInt
    ' divide
    Dim i As Integer
    Dim x128 As Big128Int
    Dim y128 As Big128Int
    Dim temp128 As Big128Int
    Dim ans128 As Big128Int
    Dim ans As BigInt
    
    'convert two 64 bit integers into one 128 bit integer
    For i = 0 To 7
        x128.n(i) = y.n(i)
        temp128.n(i) = z.n(i)
    Next
    For i = 1 To word_size
        x128 = Big128Left(x128)
    Next
    x128 = Big128Add(x128, temp128)
    ' x is the divisor
    For i = 0 To 7
        y128.n(i) = x.n(i)
    Next
    
    ' do the math
    ans128 = Big128Div(x128, y128)
    
    ' break it up
    For i = 0 To 7
        ans.n(i) = ans128.n(i) And word_mask.n(i)
    Next
    
    DblDiv = ans
End Function
Public Function DblMod(x As BigInt, y As BigInt, z As BigInt) As BigInt
    ' modulus (remainder after division)
    Dim i As Integer
    Dim x128 As Big128Int
    Dim y128 As Big128Int
    Dim ans128 As Big128Int
    Dim temp128 As Big128Int
    Dim ans As BigInt
    
    'convert two 64 bit integers into one 128 bit integer
    For i = 0 To 7
        x128.n(i) = y.n(i)
        temp128.n(i) = z.n(i)
    Next
    For i = 1 To word_size
        x128 = Big128Left(x128)
    Next
    x128 = Big128Add(x128, temp128)
    ' x is the divisor
    For i = 0 To 7
        y128.n(i) = x.n(i)
    Next
    
    ' do the math
    ans128 = Big128Mod(x128, y128)
    
    ' break it up
    For i = 0 To 7
        ans.n(i) = ans128.n(i) And word_mask.n(i)
    Next
    
    DblMod = ans
End Function
Public Function Big128Left(x As Big128Int) As Big128Int
    ' left shift
    Dim ans128 As Big128Int
    Dim i As Integer
    Dim temp As Integer
    Dim carry As Integer
    
    ' shift left is the same as "multiply by 2"
    For i = 0 To MAX_BYTE
        temp = (x.n(i) * 2) + carry
        ans128.n(i) = temp Mod 256
        carry = temp \ 256
    Next
    
    Big128Left = ans128
End Function
Public Function Big128Right(x As Big128Int) As Big128Int
    ' right shift
    Dim ans128 As Big128Int
    Dim i As Integer
    Dim carry As Integer
    
    ' shift right is the same as "divide by 2"
    For i = 0 To MAX_BYTE
        If i < MAX_BYTE Then
            If x.n(i + 1) And 1 Then
                carry = 128
            Else
                carry = 0
            End If
        End If
        ans128.n(i) = (x.n(i) \ 2) Or carry
    Next
    
    ' only the right most bit counts
    If x.n(0) And 1 Then
        carry_bit = True
    Else
        carry_bit = False
    End If
    
    Big128Right = ans128
End Function
Public Function Big128IsZero(x As Big128Int) As Boolean
    ' test if zero
    Dim i As Integer
    
    For i = 0 To MAX_BYTE
        If x.n(i) <> 0 Then
            Big128IsZero = False
            Exit Function
        End If
    Next
    
    Big128IsZero = True
End Function
Public Function Big128IsLess(x As Big128Int, y As Big128Int) As Boolean
    ' test if less than
    Dim i As Integer
        
    For i = MAX_BYTE To 0 Step -1
        If x.n(i) > y.n(i) Then
            Big128IsLess = False
            Exit Function
        End If
        If x.n(i) < y.n(i) Then
            Big128IsLess = True
            Exit Function
        End If
    Next
    ' is equal
    Big128IsLess = False
End Function
Public Function Big128IsGreater(x As Big128Int, y As Big128Int) As Boolean
    ' test if greater than
    Dim i As Integer
        
    For i = MAX_BYTE To 0 Step -1
        If x.n(i) > y.n(i) Then
            Big128IsGreater = True
            Exit Function
        End If
        If x.n(i) < y.n(i) Then
            Big128IsGreater = False
            Exit Function
        End If
    Next
    ' is equal
    Big128IsGreater = False
End Function
Public Function Big128Add(x As Big128Int, y As Big128Int) As Big128Int
    ' addition
    Dim ans As Big128Int
    Dim i As Integer
    Dim carry As Integer
    Dim temp As Integer
    
    ' Normal stuff... add two bytes, keep the "carry" for the next byte
    For i = 0 To MAX_BYTE
        temp = CInt(x.n(i)) + CInt(y.n(i)) + carry
        ans.n(i) = temp Mod 256
        carry = temp \ 256
    Next
    
    carry_bit = CBool(carry)

    Big128Add = ans
End Function
Public Function Big128Sub(x As Big128Int, y As Big128Int) As Big128Int
    Dim ans As Big128Int
    Dim temp As Integer
    Dim i As Integer
    Dim carry As Integer

    ' Normal stuff... subtract two bytes, generate a "borrow" if needed
    For i = 0 To MAX_BYTE
        temp = CInt(x.n(i)) - y.n(i) - carry
        ' do we need a "borrow"?
        If temp < 0 Then
            temp = temp + 256
            carry = 1
        Else
            carry = 0
        End If
        ans.n(i) = temp
    Next
    
    carry_bit = CBool(carry)

    Big128Sub = ans
End Function
Public Function Big128Mult(x As Big128Int, y As Big128Int) As Big128Int
    ' multiply
    Dim ans As Big128Int
    Dim i As Integer
    Dim j As Integer
    Dim temp As Long
    Dim z As Big128Int
    
    ' let's handle some special cases
    If Big128IsZero(x) Or Big128IsZero(y) Then
        Big128Mult = ans
        Exit Function
    End If

    ' This is just like you were taught is school...  multiply each digit
    ' in the first number with every digit in the second.  Keep track of
    ' what "column" to write down the intermediate values, then add the
    ' intermediate values together.
    overflow = False
    For i = 0 To NumDigits(x)
        For j = 0 To NumDigits(y)
            temp = CLng(x.n(i)) * y.n(j)
            ' anything to do?
            If temp <> 0 Then
                If (i + j) <= MAX_BYTE Then
                    ' This is a sneaky trick to create intermediate values
                    ' without using shifts.  Normally you'd multiply the
                    ' temp value times the value for that "decimal place".
                    z = PutAt(temp, (i + j) * 8)
                    ans = Big128Add(ans, z)
                Else
                    overflow = True
                End If
            End If
        Next
    Next

    Big128Mult = ans
End Function
Public Function Big128Div(x As Big128Int, y As Big128Int) As Big128Int
    ' divide
    Dim i As Integer
    Dim ans_digit As Integer
    Dim nx As Integer
    Dim ny As Integer
    Dim nc As Integer
    Dim chunk As Big128Int
    Dim guess As Big128Int
    Dim z As Big128Int
    Dim ans As Big128Int
    Dim tempx As Big128Int

    ' sanity check
    If Big128IsZero(y) Then
        overflow = True
        Big128Div = ans
        Exit Function
    End If
    
    ' special case
    If Big128IsLess(x, y) Then
        Big128Div = ans
        Exit Function
    End If
    
    ' This is almost like you were taught in school... For each digit in
    ' the answer, guess the number of times the divisor will go into a
    ' "chunk" of the dividend.  We know the size of the chunk will have
    ' the same number of digits as x or be one digit larger.  Subtract
    ' the intermediate value and repeat.
    
    tempx = x
    ny = NumDigits(y)
    Do
        ' if no more divisors... we're done
        If Big128IsLess(tempx, y) Then
            Exit Do
        End If
        
        nx = NumDigits(tempx)
        ' build a "chunk" that has the same number of digits as y
        Big128Clear chunk
        For i = 0 To ny
            chunk.n(i) = tempx.n(i + nx - ny)
        Next
        nc = ny

        ' if it is too small, build another chunk with one more digit
        If Big128IsLess(chunk, y) Then
            For i = 0 To ny + 1
                chunk.n(i) = tempx.n(i + nx - ny - 1)
            Next
            nc = ny + 1
        End If
    
        ' To keep from testing all 255 possibilites for each "answer digit",
        ' we use a simplified binary tree search alogrithm.
        ans_digit = 128
        Big128Clear guess
        For i = 6 To 0 Step -1
            guess.n(0) = ans_digit
            z = Big128Mult(y, guess)
            If Big128IsGreater(z, chunk) Then
                ans_digit = ans_digit - (2 ^ i)
            Else
                ans_digit = ans_digit + (2 ^ i)
            End If
        Next
        guess.n(0) = ans_digit
        z = Big128Mult(y, guess)
        If Big128IsGreater(z, chunk) Then
            ans_digit = ans_digit - 1
        End If
        
        ' place answer digit in proper "column"
        z = PutAt(CLng(ans_digit), (nx - nc) * 8)
        
        ' Add it to the answer.
        ans = Big128Add(ans, z)
        
        ' Generate the intermediate value by shifting the "last guess"
        ' multipler to it's proper place (which undoes the effect of the
        ' shift that occured when we developed the chunk).
        Big128Clear guess
        guess.n(nx - nc) = ans_digit
        z = Big128Mult(y, guess)

        ' Now subtract the intermediate value and continue...
        tempx = Big128Sub(tempx, z)
    Loop While nx - nc > 0
    
    If Big128IsZero(tempx) Then
        carry_bit = False
    Else
        carry_bit = True
    End If
    
    Big128Div = ans
End Function
Public Function Big128Mod(x As Big128Int, y As Big128Int) As Big128Int
    ' modulus (remainder after division)
    Dim i As Integer
    Dim ans_digit As Integer
    Dim nx As Integer
    Dim ny As Integer
    Dim nc As Integer
    Dim chunk As Big128Int
    Dim guess As Big128Int
    Dim z As Big128Int
    Dim ans As Big128Int
    Dim tempx As Big128Int
    
    ' sanity check
    If Big128IsZero(y) Then
        overflow = True
        Big128Mod = ans
        Exit Function
    End If
    
    ' For the modulus operator, it's OK to have a x < y
    If Big128IsLess(x, y) Then
        Big128Mod = x
        Exit Function
    End If
    
    ' This is almost like you were taught in school... For each digit in
    ' the answer, guess the number of times the divisor will go into a
    ' "chunk" of the dividend.  We know the size of the chunk will have
    ' the same number of digits as z or be one digit larger.  Subtract
    ' the intermediate value and repeat.
    
    tempx = x
    ny = NumDigits(y)
    Do
        ' if no more divisors... we're done
        If Big128IsLess(tempx, y) Then
            Exit Do
        End If
        
        nx = NumDigits(tempx)
        ' build a "chunk" that has the same number of digits as y
        Big128Clear chunk
        For i = 0 To ny
            chunk.n(i) = tempx.n(i + nx - ny)
        Next
        nc = ny

        ' if it is too small, build another chunk with one more digit
        If Big128IsLess(chunk, y) Then
            For i = 0 To ny + 1
                chunk.n(i) = tempx.n(i + nx - ny - 1)
            Next
            nc = ny + 1
        End If
    
        ' To keep from testing all 255 possibilites for each "answer digit",
        ' we use a simplified binary tree search alogrithm.
        ans_digit = 128
        Big128Clear guess
        For i = 6 To 0 Step -1
            guess.n(0) = ans_digit
            z = Big128Mult(y, guess)
            If Big128IsGreater(z, chunk) Then
                ans_digit = ans_digit - (2 ^ i)
            Else
                ans_digit = ans_digit + (2 ^ i)
            End If
        Next
        guess.n(0) = ans_digit
        z = Big128Mult(y, guess)
        If Big128IsGreater(z, chunk) Then
            ans_digit = ans_digit - 1
        End If
        
        ' Generate the intermediate value by shifting the "last guess"
        ' multipler to it's proper place (which undoes the effect of the
        ' shift that occured when we developed the chunk).
        Big128Clear guess
        guess.n(nx - nc) = ans_digit
        z = Big128Mult(y, guess)

        ' Now subtract the intermediate value and continue...
        tempx = Big128Sub(tempx, z)
    Loop While nx - nc > 0
    
    ' no need for mask... it can't be Big128ger than x
    Big128Mod = tempx
End Function
Public Sub Big128Clear(ByRef x As Big128Int)
    ' Zero out an Big128Int.  Notice the use of ByRef to keep my habitual C++
    ' coding style of using parenthesis from causing problems.
    Dim i As Integer
    
    For i = 0 To MAX_BYTE
        x.n(i) = 0
    Next
End Sub
Private Function PutAt(i As Long, bits As Integer) As Big128Int
    ' Put the integer value (express as a long to prevent sign extension)
    ' at the specified starting bit location within a Big128Int.  This prevents
    ' a lot of bit shifting or multiplication, when you can place the value
    ' where you want it.
    Dim ans As Big128Int
    Dim n_bit As Integer
    Dim n_byte As Integer
    Dim temp As Long
    
    n_byte = bits \ 8
    n_bit = bits Mod 8
    temp = i * (2 ^ n_bit)
    
    ans.n(n_byte) = temp Mod 256
    If temp > 255 Then
        If n_byte < MAX_BYTE Then
            ans.n(n_byte + 1) = temp \ 256
        Else
            overflow = True
        End If
    End If
    
    PutAt = ans
End Function
Private Function NumDigits(x As Big128Int) As Integer
    ' returns the location of the most significant digit (byte) in an Big128Int
    Dim i As Integer
    
    For i = MAX_BYTE To 0 Step -1
        If x.n(i) <> 0 Then
            Exit For
        End If
    Next
    
    NumDigits = i
End Function
