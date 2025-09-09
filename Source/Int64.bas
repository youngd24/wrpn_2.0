Attribute VB_Name = "BigInt64"
' BigInt - A big integer math libarary for Visual Basic.
'
' Uses an array of bytes (since that is the only unsigned type in VB).  It
' is designed to be easily extended to other word lengths. The working size
' of the word (which in case runs from 1 to 64 bits) is defined in the
' word_size variable.
'
' carry_bit is set or cleared on the following operations:
'   BigLeft
'   BigRight
'   BigAdd
'   BigSub
'   BigDiv
'
' overflow is set or cleared on the following operations:
'   BigAdd
'   BigSub
'   BigMult
'   BigDiv
'
' lop (loss of precision) is set or cleared on the following operations:
'   BigFromDbl
'   BigToDbl
'   BigToLng
'   BigToInt

Option Explicit

Private Const MAX_BYTE = 7

Type BigInt
    n(0 To MAX_BYTE) As Byte
End Type

Private Const DBL_MAX = 140737488355327#
Private Const DBL_MIN = -140737488355328#
Private Const DBL_SIGNMASK = 140737488355328#
Private Const LONG_MAX = 2147483647
Private Const LONG_SIGNMASK = 2147483648#
Private Const INT_MAX = 32767
Private Const INT_SIGNMASK = 32768

Public carry_bit As Boolean
Public overflow As Boolean
Public lop As Boolean

Public Function BigFromDbl(d As Double) As BigInt
    ' convert a double to BigInt
    Dim ans As BigInt
    Dim i As Integer
    Dim temp As Double
    Dim temp1 As Double
    Dim temp2 As Double
    
    ' The size of the mantissa portion of a VB IEEE double is 52 bits.
    ' However, the standard allows for normalization (which we can't
    ' control) and a sign bit.  So, the number of unsigned bits we can
    ' use is 50.  But, for ease or programming, I've limited the size
    ' of a double to 48 bits (an even multiple of 8).
    
    lop = False
    If d < DBL_MIN Or d > DBL_MAX Then
        ' under no circumstances is a Loss Of Precision (LOP) from Double
        ' allowed, so we just return zero.
        lop = True
        BigFromDbl = ans
        Exit Function
    End If
    
    If d < 0 Then
        ' make it "semi-unsigned" (without causing sign extension)
        temp = d + DBL_SIGNMASK
        ' fill in leading bytes
        For i = 6 To MAX_BYTE
            ans.n(i) = 255
        Next
    Else
        temp = d
    End If
    
    For i = 0 To 5
        ' too big for Integer divide and Mod operator, so this won't work:
        ' ans.n(i) = (temp \ (256 ^ i)) Mod 256
        temp1 = Fix(temp / (256 ^ i))
        If temp1 > 0 Then
            temp2 = Fix(temp1 / 256)
            ans.n(i) = temp1 - (temp2 * 256)
        End If
    Next
    
    ' add the sign extension back
    If d < 0 Then
        ans.n(5) = ans.n(5) Or 128
    End If
    
    BigFromDbl = BigAnd(ans, word_mask)
End Function
Public Function BigFromLng(l As Long) As BigInt
    ' convert a long to BigInt
    Dim ans As BigInt
    Dim temp As Double
    Dim i As Integer
    
    If l < 0 Then
        ' make it "semi-unsigned" (without causing sign extension)
        temp = l + LONG_SIGNMASK
        ' fill in high bits
        For i = 4 To MAX_BYTE
            ans.n(i) = 255
        Next
    Else
        temp = l
    End If
    
    For i = 0 To 3
        ans.n(i) = (temp \ (256 ^ i)) Mod 256
    Next
    
    ' add the sign extension back
    If l < 0 Then
        ans.n(3) = ans.n(3) Or 128
    End If

    BigFromLng = BigAnd(ans, word_mask)
End Function
Public Function BigFromInt(j As Integer) As BigInt
    ' convert an integer to BigInt
    Dim ans As BigInt
    Dim i As Integer
    Dim temp As Integer
    
    If j < 0 Then
        ' make it "semi-unsigned" (without causing sign extension)
        temp = j + INT_SIGNMASK
        ' fill in high bits
        For i = 2 To MAX_BYTE
            ans.n(i) = 255
        Next
    Else
        temp = j
    End If
    
    For i = 0 To 1
        ans.n(i) = (temp \ (256 ^ i)) Mod 256
    Next

    ' add the sign extension back
    If j < 0 Then
        ans.n(1) = ans.n(1) Or 128
    End If
    
    BigFromInt = BigAnd(ans, word_mask)
End Function
Public Function BigAnd(x As BigInt, y As BigInt) As BigInt
    ' logical AND
    Dim ans As BigInt
    Dim i As Integer
    
    For i = 0 To MAX_BYTE
        ans.n(i) = x.n(i) And y.n(i)
    Next
    
    BigAnd = ans
End Function
Public Function BigOr(x As BigInt, y As BigInt) As BigInt
    ' logical OR
    Dim ans As BigInt
    Dim i As Integer
    
    For i = 0 To MAX_BYTE
        ans.n(i) = x.n(i) Or y.n(i)
    Next
    
    BigOr = ans
End Function
Public Function BigNot(x As BigInt) As BigInt
    ' logical NOT
    Dim ans As BigInt
    Dim i As Integer
    
    For i = 0 To MAX_BYTE
        ans.n(i) = 255 - x.n(i)
    Next
    
    BigNot = BigAnd(ans, word_mask)
End Function
Public Function BigXor(x As BigInt, y As BigInt) As BigInt
    ' logical Exclusive OR
    Dim ans As BigInt
    Dim i As Integer
    
    For i = 0 To MAX_BYTE
        ans.n(i) = x.n(i) Xor y.n(i)
    Next
    
    BigXor = ans
End Function
Public Function BigLeft(x As BigInt) As BigInt
    ' left shift
    Dim ans As BigInt
    Dim i As Integer
    Dim temp As Integer
    Dim carry As Integer
    
    ' shift left is the same as "multiply by 2"
    For i = 0 To MAX_BYTE
        temp = (x.n(i) * 2) + carry
        ans.n(i) = temp Mod 256
        carry = temp \ 256
    Next
    
    ' only the left most bit counts
    If BigIsZero(BigAnd(x, sign_mask)) Then
        carry_bit = False
    Else
        carry_bit = True
    End If
    
    BigLeft = BigAnd(ans, word_mask)
End Function
Public Function BigRight(x As BigInt) As BigInt
    ' right shift
    Dim ans As BigInt
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
        ans.n(i) = (x.n(i) \ 2) Or carry
    Next
    
    ' only the right most bit counts
    If x.n(0) And 1 Then
        carry_bit = True
    Else
        carry_bit = False
    End If
    
    BigRight = ans
End Function
Public Function BigIsZero(x As BigInt) As Boolean
    ' test if zero
    Dim i As Integer
    
    For i = 0 To MAX_BYTE
        If x.n(i) <> 0 Then
            BigIsZero = False
            Exit Function
        End If
    Next
    
    BigIsZero = True
End Function
Public Function BigIsEqual(x As BigInt, y As BigInt) As Boolean
    ' test if equal
    Dim i As Integer
    
    For i = 0 To MAX_BYTE
        If x.n(i) <> y.n(i) Then
            BigIsEqual = False
            Exit Function
        End If
    Next
    
    BigIsEqual = True
End Function
Public Function BigIsLess(x As BigInt, y As BigInt) As Boolean
    ' test if less than
    Dim i As Integer
        
    For i = MAX_BYTE To 0 Step -1
        If x.n(i) > y.n(i) Then
            BigIsLess = False
            Exit Function
        End If
        If x.n(i) < y.n(i) Then
            BigIsLess = True
            Exit Function
        End If
    Next
    ' is equal
    BigIsLess = False
End Function
Public Function BigIsGreater(x As BigInt, y As BigInt) As Boolean
    ' test if greater than
    Dim i As Integer
        
    For i = MAX_BYTE To 0 Step -1
        If x.n(i) > y.n(i) Then
            BigIsGreater = True
            Exit Function
        End If
        If x.n(i) < y.n(i) Then
            BigIsGreater = False
            Exit Function
        End If
    Next
    ' is equal
    BigIsGreater = False
End Function
Public Function BigAdd(x As BigInt, y As BigInt, Optional skip_mask As Boolean = False) As BigInt
    ' addition
    Dim ans As BigInt
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
    ' did it fit?
    If BigIsEqual(ans, BigAnd(ans, word_mask)) = False Then
        overflow = True
    Else
        overflow = False
    End If
    
    If skip_mask Then
        BigAdd = ans
    Else
        BigAdd = BigAnd(ans, word_mask)
    End If
End Function
Public Function BigSub(x As BigInt, y As BigInt, Optional skip_mask As Boolean = False) As BigInt
    Dim ans As BigInt
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
    ' did it fit?
    If BigIsEqual(ans, BigAnd(ans, word_mask)) = False Then
        overflow = True
    Else
        overflow = False
    End If
    
    If skip_mask Then
        BigSub = ans
    Else
        BigSub = BigAnd(ans, word_mask)
    End If
End Function
Public Function BigMult(x As BigInt, y As BigInt, Optional skip_mask As Boolean = False) As BigInt
    ' multiply
    Dim ans As BigInt
    Dim i As Integer
    Dim j As Integer
    Dim temp As Long
    Dim z As BigInt
    
    ' let's handle some special cases
    If BigIsZero(x) Or BigIsZero(y) Then
        BigMult = ans
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
                    ans = BigAdd(ans, z, True)
                Else
                    overflow = True
                End If
            End If
        Next
    Next
    
    If BigIsEqual(ans, BigAnd(ans, word_mask)) = False Then
        overflow = True
    End If

    If skip_mask Then
        BigMult = ans
    Else
        BigMult = BigAnd(ans, word_mask)
    End If
End Function
Public Function BigDiv(x As BigInt, y As BigInt) As BigInt
    ' divide
    Dim i As Integer
    Dim ans_digit As Integer
    Dim nx As Integer
    Dim ny As Integer
    Dim nc As Integer
    Dim chunk As BigInt
    Dim guess As BigInt
    Dim z As BigInt
    Dim ans As BigInt
    Dim tempx As BigInt

    ' sanity check
    If BigIsZero(y) Then
        overflow = True
        BigDiv = ans
        Exit Function
    End If
    
    ' special case
    If BigIsLess(x, y) Then
        BigDiv = ans
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
        If BigIsLess(tempx, y) Then
            Exit Do
        End If
        
        nx = NumDigits(tempx)
        ' build a "chunk" that has the same number of digits as y
        BigClear chunk
        For i = 0 To ny
            chunk.n(i) = tempx.n(i + nx - ny)
        Next
        nc = ny

        ' if it is too small, build another chunk with one more digit
        If BigIsLess(chunk, y) Then
            For i = 0 To ny + 1
                chunk.n(i) = tempx.n(i + nx - ny - 1)
            Next
            nc = ny + 1
        End If
    
        ' To keep from testing all 255 possibilites for each "answer digit",
        ' we use a simplified binary tree search alogrithm.
        ans_digit = 128
        BigClear guess
        For i = 6 To 0 Step -1
            guess.n(0) = ans_digit
            z = BigMult(y, guess, True)
            If BigIsGreater(z, chunk) Then
                ans_digit = ans_digit - (2 ^ i)
            Else
                ans_digit = ans_digit + (2 ^ i)
            End If
        Next
        guess.n(0) = ans_digit
        z = BigMult(y, guess, True)
        If BigIsGreater(z, chunk) Then
            ans_digit = ans_digit - 1
        End If
        
        ' place answer digit in proper "column"
        z = PutAt(CLng(ans_digit), (nx - nc) * 8)
        
        ' Add it to the answer.
        ans = BigAdd(ans, z, True)
        
        ' Generate the intermediate value by shifting the "last guess"
        ' multipler to it's proper place (which undoes the effect of the
        ' shift that occured when we developed the chunk).
        BigClear guess
        guess.n(nx - nc) = ans_digit
        z = BigMult(y, guess, True)

        ' Now subtract the intermediate value and continue...
        tempx = BigSub(tempx, z, True)
    Loop While nx - nc > 0
    
    If BigIsZero(tempx) Then
        carry_bit = False
    Else
        carry_bit = True
    End If
    
    BigDiv = BigAnd(ans, word_mask)
End Function
Public Function BigMod(x As BigInt, y As BigInt) As BigInt
    ' modulus (remainder after division)
    Dim i As Integer
    Dim ans_digit As Integer
    Dim nx As Integer
    Dim ny As Integer
    Dim nc As Integer
    Dim chunk As BigInt
    Dim guess As BigInt
    Dim z As BigInt
    Dim ans As BigInt
    Dim tempx As BigInt
    
    ' sanity check
    If BigIsZero(y) Then
        overflow = True
        BigMod = ans
        Exit Function
    End If
    
    ' For the modulus operator, it's OK to have a x < y
    If BigIsLess(x, y) Then
        BigMod = x
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
        If BigIsLess(tempx, y) Then
            Exit Do
        End If
        
        nx = NumDigits(tempx)
        ' build a "chunk" that has the same number of digits as y
        BigClear chunk
        For i = 0 To ny
            chunk.n(i) = tempx.n(i + nx - ny)
        Next
        nc = ny

        ' if it is too small, build another chunk with one more digit
        If BigIsLess(chunk, y) Then
            For i = 0 To ny + 1
                chunk.n(i) = tempx.n(i + nx - ny - 1)
            Next
            nc = ny + 1
        End If
    
        ' To keep from testing all 255 possibilites for each "answer digit",
        ' we use a simplified binary tree search alogrithm.
        ans_digit = 128
        BigClear guess
        For i = 6 To 0 Step -1
            guess.n(0) = ans_digit
            z = BigMult(y, guess, True)
            If BigIsGreater(z, chunk) Then
                ans_digit = ans_digit - (2 ^ i)
            Else
                ans_digit = ans_digit + (2 ^ i)
            End If
        Next
        guess.n(0) = ans_digit
        z = BigMult(y, guess, True)
        If BigIsGreater(z, chunk) Then
            ans_digit = ans_digit - 1
        End If
        
        ' Generate the intermediate value by shifting the "last guess"
        ' multipler to it's proper place (which undoes the effect of the
        ' shift that occured when we developed the chunk).
        BigClear guess
        guess.n(nx - nc) = ans_digit
        z = BigMult(y, guess, True)

        ' Now subtract the intermediate value and continue...
        tempx = BigSub(tempx, z, True)
    Loop While nx - nc > 0
    
    ' no need for mask... it can't be bigger than x
    BigMod = tempx
End Function
Public Function BigSqrt(x As BigInt) As BigInt
    ' square root using Newton's method
    Dim tempx As BigInt
    Dim r As BigInt
    Dim one As BigInt
    
    If BigIsZero(x) Then
        overflow = True
        BigSqrt = tempx
        Exit Function
    End If
    
    tempx = x
    r.n(0) = 1
    Do While BigIsGreater(tempx, r)
        ' remainder
        r = BigDiv(x, tempx)
        ' using right shift instead of division by 2
        tempx = BigRight(BigAdd(tempx, r, True))
    Loop
    
    ' carry on remainder (not a perfect square)
    r = BigSub(x, BigMult(tempx, tempx, True), True)
    If BigIsZero(r) Then
        carry_bit = False
    Else
        carry_bit = True
    End If
    
    BigSqrt = tempx
End Function
Public Function BigToDbl(x As BigInt) As Double
    ' convert a BigInt to a Double
    Dim ans As Double
    Dim i As Integer
    
    ' we can't use anything beyond 6 bits
    lop = False
    For i = 6 To MAX_BYTE
        If x.n(i) <> 0 Then
            lop = True
        End If
    Next

    ' This is essentially a base 256 numbering system.
    For i = 0 To 5
        ans = ans + (x.n(i) * (256 ^ i))
    Next

    ' is it negative?
    If ans > DBL_MAX Then
        ans = ans - (DBL_MAX * 2#) - 2
    End If
    
    BigToDbl = ans
End Function
Public Function BigToLng(x As BigInt) As Long
    ' convert a BigInt to a Long
    Dim ans As Double
    Dim i As Integer

    ' we can't use anything beyond 4 bits
    lop = False
    For i = 4 To MAX_BYTE
        If x.n(i) <> 0 Then
            lop = True
        End If
    Next

    ' This is essentially a base 256 numbering system.
    For i = 0 To 3
        ans = ans + (x.n(i) * (256 ^ i))
    Next
    
    ' is it negative?
    If ans > LONG_MAX Then
        ans = ans - (LONG_MAX * 2#) - 2
    End If
    
    BigToLng = CLng(ans)
End Function
Public Function BigToInt(x As BigInt) As Integer
    ' convert a BigInt to a integer
    Dim ans As Long
    Dim i As Integer

    ' we can't use anything beyond 2 bits
    lop = False
    For i = 2 To MAX_BYTE
        If x.n(i) <> 0 Then
            lop = True
        End If
    Next

    ' This is essentially a base 256 numbering system.
    For i = 0 To 1
        ans = ans + (x.n(i) * (256 ^ i))
    Next
    
    ' is it negative?
    If ans > INT_MAX Then
        ans = ans - (INT_MAX * 2&) - 2
    End If
    
    BigToInt = CInt(ans)
End Function
Public Function BigToBinStr(x As BigInt) As String
    ' binary pattern of a BigInt as a string
    Dim i As Integer
    Dim j As Integer
    Dim buf As String
    
    For i = MAX_BYTE To 0 Step -1
        For j = 7 To 0 Step -1
            If (2 ^ j) And x.n(i) Then
                buf = buf & "1"
            Else
                buf = buf & "0"
            End If
        Next
        buf = buf & " "
    Next
    
    BigToBinStr = buf
End Function
Public Function BigToDecStr(x As BigInt, Optional comp As Integer = 0) As String
    ' decimal pattern of a BigInt as a string
    Dim i As Integer
    Dim ten As BigInt
    Dim j As Integer
    Dim ending As Integer
    Dim buf As String
    Dim tempx As BigInt
    Dim one As BigInt
    Dim is_neg As Boolean
    
    ' The decimal system is the only one that has to deal with the 1's
    ' or 2's complement.  All other systems (octal, binary, hex) are
    ' inherently unsigned.
    tempx = x
    is_neg = False
    Select Case comp
        Case 0      ' unsigned
            ' do nothing
        Case 1      ' 1s complement
            If BigIsZero(BigAnd(x, sign_mask)) = False Then
                is_neg = True
                tempx = BigNot(x)
            End If
        Case 2      ' 2s complement
            If BigIsZero(BigAnd(x, sign_mask)) = False Then
                is_neg = True
                one.n(0) = 1
                tempx = BigAdd(BigNot(x), one)
            End If
    End Select
    
    ' how many digits are possible (not an exact science, this time)
    ending = CInt(word_size / 3.5)

    ' This normal stuff...  To get each digit, divide the number by 10
    ' and then get the remainder.  Repeat.
    ten.n(0) = 10
    For i = 0 To ending
        ' can't use my sneaky GetAt function... Ugh!
        j = BigToInt(BigMod(tempx, ten))
        buf = CStr(j) & buf
        tempx = BigDiv(tempx, ten)
    Next
    
    If is_neg Then
        buf = "-" & buf
    End If
    
    BigToDecStr = buf
End Function
Public Function BigToHexStr(x As BigInt) As String
    ' hex pattern of a BigInt as a string
    Dim i As Integer
    Dim j As Integer
    Dim ending As Integer
    Dim buf As String
    
    ' how many digits are possible?
    ending = word_size \ 4
    If (word_size Mod 4) = 0 Then
        ending = ending - 1
    End If
    
    ' Instead of using the normal method, we use the GetAt function to
    ' extract the digit's value directly from the BigInt structure.
    For i = 0 To ending
        ' sneaky trick to prevent multiplication and division
        j = GetAt(x, i * 4, 4)
        If j > 9 Then
            buf = Chr(Asc("a") + j - 10) & buf
        Else
            buf = CStr(j) & buf
        End If
    Next
    
    BigToHexStr = buf
End Function
Public Function BigToOctStr(x As BigInt) As String
    ' octal pattern of a BigInt as a string
    Dim i As Integer
    Dim j As Integer
    Dim ending As Integer
    Dim buf As String
    
    ' how many digits are possible
    ending = word_size \ 3
    If (word_size Mod 3) = 0 Then
        ending = ending - 1
    End If
    
    ' Instead of using the normal method, we use the GetAt function to
    ' extract the digit's value directly from the BigInt structure.
    For i = 0 To ending
        ' sneaky trick to prevent multiplication and division
        j = GetAt(x, i * 3, 3)
        buf = CStr(j) & buf
    Next
    
    BigToOctStr = buf
End Function
Public Function BigFromBinStr(buf As String) As BigInt
    ' convert a binary string to BigInt.  No input checking is done
    ' here, so it's up to you to do that elsewhere!
    Dim i As Integer
    Dim length As Integer
    Dim n_byte As Integer
    Dim n_bit As Integer
    Dim ans As BigInt
    Dim b As String

    length = Len(buf)
    For i = Len(buf) To 1 Step -1
        b = Mid(buf, i, 1)
        ' stop at first unknown character
        If InStr(" 01", b) = 0 Then
            Exit For
        End If
        ' purge any spaces
        If b = " " Then
            length = length - 1
        Else
            ' which byte and which bit?
            n_byte = (length - i) \ 8
            n_bit = (length - i) Mod 8
            If Mid(buf, i, 1) = "1" Then
                ans.n(n_byte) = ans.n(n_byte) + (2 ^ n_bit)
            End If
        End If
    Next
    
    BigFromBinStr = BigAnd(ans, word_mask)
End Function
Public Function BigFromDecStr(buf As String) As BigInt
    ' convert a decimal string to BigInt.  No input checking is done
    ' here, so it's up to you to do that elsewhere!
    Dim i As Integer
    Dim temp As Integer
    Dim ans As BigInt
    Dim mult As BigInt
    Dim ten As BigInt
    Dim one As BigInt
    Dim b As String
    
    ' This is normal stuff.  You multiply each digit times the value
    ' associated with that decimal place, then add 'em all up.
    mult.n(0) = 1
    ten.n(0) = 10
    For i = Len(buf) To 1 Step -1
        b = Mid(buf, i, 1)
        ' stop at first unknown character (including a negative sign)
        If InStr("0123456789", b) = 0 Then
            Exit For
        End If
        temp = CInt(b)
        If temp <> 0 Then
            ' can't use my sneaky PutAt function...
            ans = BigAdd(ans, BigMult(BigFromInt(temp), mult, True), True)
        End If
        ' decimal place
        mult = BigMult(mult, ten, True)
    Next
    
    ' is it negative?  I'm not treating a negative sign as an error in
    ' the unsigned mode... I wonder if that's wise.
    If Left(buf, 1) = "-" Then
        If comp = 1 Then
            ans = BigNot(ans)
        Else
            one.n(0) = 1
            ans = BigAdd(BigNot(ans), one)
        End If
    End If
    
    BigFromDecStr = BigAnd(ans, word_mask)
End Function
Public Function BigFromHexStr(buf As String) As BigInt
    ' convert a hex string to BigInt.  No input checking is done
    ' here, so it's up to you to do that elsewhere!
    Dim i As Integer
    Dim temp As Long
    Dim ans As BigInt
    Dim bits As Integer
    Dim b As String

    ' Instead of multiplying each digit with the value of that decimal
    ' place, we use the PutAt function to simply place the value at the
    ' correct "decimal place" within an BigInt structure.
    For i = Len(buf) To 1 Step -1
        b = Mid(buf, i, 1)
        ' stop at first unknown character
        If InStr("0123456789abcdefABCDEF", b) = 0 Then
            Exit For
        End If
        temp = CInt("&H" & b)
        If temp <> 0 Then
            ' sneaky trick to prevent multiplication or shifts
            ans = BigAdd(ans, PutAt(temp, bits), True)
        End If
        ' number of bits to shift
        bits = bits + 4
    Next
    
    BigFromHexStr = BigAnd(ans, word_mask)
End Function
Public Function BigFromOctStr(buf As String) As BigInt
    ' convert a hex string to BigInt.  No input checking is done
    ' here, so it's up to you to do that elsewhere!
    Dim i As Integer
    Dim temp As Long
    Dim ans As BigInt
    Dim bits As Integer
    Dim b As String

    ' Instead of multiplying each digit with the value of that decimal
    ' place, we use the PutAt function to simply place the value at the
    ' correct "decimal place" within an BigInt structure.
    For i = Len(buf) To 1 Step -1
        b = Mid(buf, i, 1)
        ' stop at first unknown character
        If InStr("01234567", b) = 0 Then
            Exit For
        End If
        temp = CInt("&0" & b)
        If temp <> 0 Then
            ' sneaky trick to prevent multiplication or shifts
            ans = BigAdd(ans, PutAt(temp, bits), True)
        End If
        ' number of bits to shift
        bits = bits + 3
    Next
    
    BigFromOctStr = BigAnd(ans, word_mask)
End Function
Public Function BigWordMask(ws As Integer) As BigInt
    ' Create a "word masks" to be used to truncate a BigInt to it's proper
    ' size
    Dim n_byte As Integer
    Dim n_bit As Integer
    Dim i As Integer
    Dim ans As BigInt
    
    ' sanity check... if out-of-bounds, return previous
    If ws < 1 Or ws > (MAX_BYTE + 1) * 8 Then
        BigWordMask = word_mask
        Exit Function
    End If
    
    ' a short cut if word size is the maximum
    If ws = (MAX_BYTE + 1) * 8 Then
        For i = 0 To MAX_BYTE
            ans.n(i) = 255
        Next
        BigWordMask = ans
        Exit Function
    End If
    
    ' create a word "mask"... all 1's for the length of the word
    n_byte = ws \ 8
    n_bit = ws Mod 8
    If n_byte > 0 Then
        For i = 0 To n_byte - 1
            ans.n(i) = 255
        Next
    End If
    ans.n(n_byte) = (2 ^ n_bit) - 1
    
    BigWordMask = ans
End Function
Public Function BigSignMask(ws As Integer) As BigInt
    ' create the sign bit "mask", used to determine if a number is negative
    Dim n_byte As Integer
    Dim n_bit As Integer
    Dim ans As BigInt
    
    ' sanity check... if out-of-bounds, return previous
    If ws < 1 Or ws > (MAX_BYTE + 1) * 8 Then
        BigSignMask = sign_mask
        Exit Function
    End If
        
    If ws = (MAX_BYTE + 1) * 8 Then
        ans.n(MAX_BYTE) = 128
        BigSignMask = ans
        Exit Function
    End If
    
    n_byte = (ws - 1) \ 8
    n_bit = (ws - 1) Mod 8
    ans.n(n_byte) = 2 ^ n_bit
    
    BigSignMask = ans
End Function
Public Sub BigClear(ByRef x As BigInt)
    ' Zero out an BigInt.  Notice the use of ByRef to keep my habitual C++
    ' coding style of using parenthesis from causing problems.
    Dim i As Integer
    
    For i = 0 To MAX_BYTE
        x.n(i) = 0
    Next
End Sub
Private Function PutAt(i As Long, bits As Integer) As BigInt
    ' Put the integer value (express as a long to prevent sign extension)
    ' at the specified starting bit location within a BigInt.  This prevents
    ' a lot of bit shifting or multiplication, when you can place the value
    ' where you want it.
    Dim ans As BigInt
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
Private Function GetAt(x As BigInt, bits As Integer, length As Integer) As Integer
    ' Get an integer value from the specified location (and length) within
    ' a BigInt.  Sorta like the Mid$ function for the BigInt structure.
    Dim ans As Integer
    Dim temp As Long
    Dim n_bit As Integer
    Dim n_byte As Integer
    
    n_byte = bits \ 8
    n_bit = bits Mod 8

    If n_byte < MAX_BYTE Then
        temp = CLng(x.n(n_byte)) + (CLng(x.n(n_byte + 1)) * 256)
    Else
        temp = x.n(n_byte)
    End If
    ans = (temp \ (2 ^ n_bit)) Mod (2 ^ length)
    
    GetAt = ans
End Function
Private Function NumDigits(x As BigInt) As Integer
    ' returns the location of the most significant digit (byte) in an BigInt
    Dim i As Integer
    
    For i = MAX_BYTE To 0 Step -1
        If x.n(i) <> 0 Then
            Exit For
        End If
    Next
    
    NumDigits = i
End Function

