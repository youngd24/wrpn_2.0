Attribute VB_Name = "Functions"
' WRPN is a public domain calculator for Win9x/NT that is modeled after the
' Hewlett-Packard HP-16c.  This is not a Hewlett-Packard product.
'
'   Emmet Gray
'   graye@hood-emh3.army.mil

Option Explicit

' In the real calculator, the sames registers are used to store both the
' floating point and integer numbers.  That's not practical here, so we
' have two sets.  This also allows for more functionality, since the
' contents are not destroyed during conversion of one mode to the other.

Public reg(0 To REG_SIZE - 1) As Double
Public reg64(0 To REG_SIZE - 1) As BigInt
Public stack(0 To STACK_SIZE - 1) As Double
Public stack64(0 To STACK_SIZE - 1) As BigInt
Public lastx As Double
Public lastx64 As BigInt

Public flag(0 To 5) As Boolean  ' 0-2=user, 3=leading zero, 4=carry, 5=overflow
Public word_size As Integer     ' word size
Public mode As Integer          ' operating mode: dec, oct, hex, bin or float
Public dpoints As Integer       ' number of decimal points
Public comp As Integer          ' 1s, 2s or unsigned complement
Public word_mask As BigInt      ' word mask
Public sign_mask As BigInt      ' sign mask
Public shift_window As Boolean  ' shift window to the left?
Public Sub calc_key(key As Integer)
    ' process the key strokes
    Dim store_it As Integer
    Dim overflow As Integer
    Dim i As Integer
    Dim j As Integer
    Dim dont_push As Boolean
    Dim op As Double
    Dim op64 As BigInt
    Dim zero As BigInt
    Dim current As Double
    Dim current64 As BigInt
    Dim screen_buf As String * SCREEN_SIZE
    Dim n_byte As Integer
    Dim n_bit As Integer
            
    Static input_buf As String
    Static pos As Integer
    Static two_part As Integer
    Static func_key As Integer
    Static got_decimal As Boolean
    Static got_exponent As Boolean
    Static high_reg As Integer
    
    ' is this the 2nd half of a compound (function key)?
    If func_key > 0 And key < CTRL Then
        key = key + func_key
        func_key = 0
        ' turn off the F or G in the display
        DispFlags DFLAG_FKEY, False
        DispFlags DFLAG_GKEY, False
    End If
    
    ' is this a two part function?
    If two_part > 0 Then
        ' in the "real" calculator, you can't backspace over the beginning
        ' of a two-part function key... but I'm not that restrictive
        If key = K_BSP Or key = F_CPRE Then
            two_part = 0
            high_reg = 0
            If mode = FLOAT_MODE Then
                Format_Display stack(0)
            Else
                Format_Display64 stack64(0), mode
            End If
            Exit Sub
        End If
        Select Case two_part
            Case K_STO          ' store
                If Chr(key) = "." Then
                    high_reg = 16
                    Exit Sub
                End If
                If InStr("0123456789abcdef", Chr(key)) > 0 Then
                    If key <= Asc("9") Then
                        i = Asc("0")
                    Else
                        i = Asc("a") + 10
                    End If
                    If mode = FLOAT_MODE Then
                        reg(key - i + high_reg) = pop()
                    Else
                        reg64(key - i + high_reg) = pop64()
                    End If
                    two_part = 0
                    high_reg = 0
                Else
                    Calc!Display.Text = "Error 3 - Improper register number"
                    Beep
                End If
                Exit Sub
            Case K_RCL          ' recall
                If Chr(key) = "." Then
                    high_reg = 16
                    Exit Sub
                End If
                If InStr("0123456789abcdef", Chr(key)) > 0 Then
                    If key <= Asc("9") Then
                        i = Asc("0")
                    Else
                        i = Asc("a") + 10
                    End If
                    If mode = FLOAT_MODE Then
                        current = reg(key - i + high_reg)
                        push current
                        Format_Display stack(0)
                    Else
                        current64 = reg64(key - i + high_reg)
                        push64 current64
                        Format_Display64 stack64(0), mode
                    End If
                    two_part = 0
                    high_reg = 0
                Else
                    Calc!Display.Text = "Error 3 - Improper register number"
                    Beep
                End If
                Exit Sub
            Case F_FLOAT        ' float digits
                If InStr("0123456789.", Chr(key)) > 0 Then
                    dpoints = key - Asc("0")
                    ' If you not already in float mode, then convert
                    ' everything from BigInt to double
                    If mode <> FLOAT_MODE Then
                        AllFrom64
                    End If
                    mode = FLOAT_MODE
                    two_part = 0
                    Format_Display stack(0)
                    
                    Calc.CM_VIEW_FLOAT.Checked = True
                    Calc.CM_VIEW_DEC.Checked = False
                    Calc.CM_VIEW_HEX.Checked = False
                    Calc.CM_VIEW_OCT.Checked = False
                    Calc.CM_VIEW_BIN.Checked = False
                Else
                    Calc!Display.Text = "Error 1 - Improper float number"
                    Beep
                End If
                Exit Sub
            Case G_SF           ' set flag
                If InStr("012345", Chr(key)) > 0 Then
                    i = key - Asc("0")
                    flag(i) = True
                    If i >= 4 Then
                        DispFlags i, True
                     End If
                    two_part = 0
                Else
                    Calc!Display.Text = "Error 1 - Improper flag number"
                    Beep
                End If
                Exit Sub
            Case G_CF           ' clear flag
                If InStr("012345", Chr(key)) > 0 Then
                    i = key - Asc("0")
                    flag(i) = False
                    If i >= 4 Then
                        DispFlags i, False
                     End If
                Else
                    Calc!Display.Text = "Error 1 - Improper flag number"
                    Beep
                End If
                Exit Sub
        End Select
    End If
            
    ' is this numeric input?
    If (key < CTRL) Or (key = F_EEX) Then
        store_it = False
        overflow = False

        Select Case mode
            Case FLOAT_MODE
                If InStr("0123456789.", Chr(key)) > 0 Then
                    store_it = True
                End If
                ' float mode allows more than just digits
                Select Case key
                    Case K_DP   ' decimal point
                        If (got_decimal Or got_exponent) Then
                            Beep
                            Exit Sub
                        End If
                        got_decimal = True
                    Case F_EEX  ' exponent
                        If (pos = 0 Or got_exponent) Then
                            Beep
                            Exit Sub
                        End If
                        got_exponent = True
                        key = Asc("e")
                        store_it = True
                    Case K_CHS  ' change sign
                        If pos > 0 Then
                            store_it = True
                            ' change sign of exponenet?
                            i = InStr(1, input_buf, "e")
                            If i <> 0 Then
                                If Mid(input_buf, i + 1, 1) = "-" Then
                                    input_buf = Left(input_buf, i) & Mid(input_buf, i + 2)
                                    pos = pos - 1
                                Else
                                    input_buf = Left(input_buf, i) & "-" & Mid(input_buf, i + 1)
                                    pos = pos + 1
                                End If
                            Else
                            ' ok, just toggle the sign
                                If Left(input_buf, 1) = "-" Then
                                    input_buf = Mid(input_buf, 2)
                                    pos = pos - 1
                                Else
                                    input_buf = "-" & input_buf
                                    pos = pos + 1
                                End If
                            End If
                        End If
                End Select
                If pos >= MAX_FLOAT_DIGITS Then
                    overflow = True
                End If
            Case DEC_MODE
                If InStr("0123456789", Chr(key)) > 0 Then
                    store_it = True
                End If
                If pos >= MAX_DECIMAL_DIGITS Then
                    overflow = True
                End If
            Case HEX_MODE
                If InStr("0123456789abcdef", Chr(key)) > 0 Then
                    store_it = True
                End If
                If pos >= MAX_HEX_DIGITS Then
                    overflow = True
                End If
            Case OCT_MODE
                If InStr("01234567", Chr(key)) > 0 Then
                    store_it = True
                End If
                If pos >= MAX_OCTAL_DIGITS Then
                    overflow = True
                End If
            Case BIN_MODE
                If InStr("01", Chr(key)) > 0 Then
                    store_it = True
                End If
                If pos >= MAX_BINARY_DIGITS Then
                    overflow = True
                End If
        End Select
    End If
    
    ' store numeric input
    If store_it = True Then
        If overflow = True Then
            Beep
            Exit Sub
        End If
        ' placement of the "-" is handled by K_CHS
        If (key <> K_CHS) Then
            pos = pos + 1
            input_buf = input_buf & Chr(key)
        End If
        Raw_Display input_buf, mode
        Exit Sub
    End If
    
    ' process "pseduo" operators (things that don't force the
    ' conversion of pending input)
    Select Case key
        Case K_BSP
            If pos > 0 Then
                input_buf = Left(input_buf, pos - 1)
                pos = pos - 1
                Raw_Display input_buf, mode
            Else    ' same as clear x
                If mode = FLOAT_MODE Then
                    stack(0) = 0#
                    Format_Display stack(0)
                Else
                    stack64(0) = zero
                    Format_Display64 stack64(0), mode
                End If
                dont_push = True 'me
            End If
            Exit Sub
        Case K_FKEY
            func_key = CTRL
            Exit Sub
        Case K_GKEY
            func_key = ALT
            Exit Sub
    End Select
                
    ' convert any pending numeric input
    If pos > 0 Then
        If mode = FLOAT_MODE Then
            current = Convert_Input(input_buf)
            ' This is a little weird... if the last operation was K_ENTER
            ' or G_CLX then replace (don't push) the value on the stack
            If dont_push Then
                stack(0) = current
                dont_push = False
            Else
                push current
            End If
        Else
            current64 = Convert_Input64(input_buf, mode)
            If dont_push Then
                stack64(0) = current64
                dont_push = False
            Else
                push64 current64
            End If
        End If
        pos = 0
        got_decimal = False
        got_exponent = False
        input_buf = ""
    End If
    
    ' process the function keys
    Select Case key
        Case 0                  ' just update the display
        Case K_DIV              ' division
            If mode = FLOAT_MODE Then
                lastx = stack(0)
                op = pop()
                If op = 0 Then
                    Calc!Display.Text = "Error 0 - Improper Math Operation"
                    Exit Sub
                End If
                current = pop() / op
                push (current)
            Else
                lastx64 = stack64(0)
                op64 = pop64()
                If BigIsZero(op64) Then
                    Calc!Display.Text = "Error 0 - Improper Math Operation"
                    Exit Sub
                End If
                current64 = BigDiv(pop64(), op64)
                push64 current64
                If carry_bit Then
                    DispFlags DFLAG_CARRYBIT, True
                Else
                    DispFlags DFLAG_CARRYBIT, True
                End If
            End If
        Case K_GSB              ' goto subroutine
            Calc!Display.Text = "GSB not implemented"
            Exit Sub
        Case K_GTO              ' goto label
            Calc!Display.Text = "GTO not implemented"
            Exit Sub
        Case K_HEX              ' hexadecimal mode
            Calc.CM_VIEW_FLOAT.Checked = False
            Calc.CM_VIEW_DEC.Checked = False
            Calc.CM_VIEW_HEX.Checked = True
            Calc.CM_VIEW_OCT.Checked = False
            Calc.CM_VIEW_BIN.Checked = False
            If mode = FLOAT_MODE Then
                AllTo64
            End If
            mode = HEX_MODE
        Case K_DEC              ' decimal mode
            Calc.CM_VIEW_FLOAT.Checked = False
            Calc.CM_VIEW_DEC.Checked = True
            Calc.CM_VIEW_HEX.Checked = False
            Calc.CM_VIEW_OCT.Checked = False
            Calc.CM_VIEW_BIN.Checked = False
            If mode = FLOAT_MODE Then
                AllTo64
            End If
            mode = DEC_MODE
        Case K_OCT              ' octal mode
            Calc.CM_VIEW_FLOAT.Checked = False
            Calc.CM_VIEW_DEC.Checked = False
            Calc.CM_VIEW_HEX.Checked = False
            Calc.CM_VIEW_OCT.Checked = True
            Calc.CM_VIEW_BIN.Checked = False
            If mode = FLOAT_MODE Then
                AllTo64
            End If
            mode = OCT_MODE
        Case K_BIN              ' binary mode
            Calc!CM_VIEW_FLOAT.Checked = False
            Calc!CM_VIEW_DEC.Checked = False
            Calc!CM_VIEW_HEX.Checked = False
            Calc!CM_VIEW_OCT.Checked = False
            Calc!CM_VIEW_BIN.Checked = True
            If mode = FLOAT_MODE Then
                AllTo64
            End If
            mode = BIN_MODE
        Case K_MULT             ' mulitplication
            If mode = FLOAT_MODE Then
                lastx = stack(0)
                current = pop() * pop() ' fizz, fizz, oh what a relief it is
                push (current)
            Else
                lastx64 = stack64(0)
                current64 = BigMult(pop64(), pop64())
                push64 current64
            End If
        Case K_RS               ' run/stop
            Calc!Display.Text = "RS not implemented"
            Exit Sub
        Case K_SST              ' single step
            Calc!Display.Text = "SST not implemented"
            Exit Sub
        Case K_ROL              ' roll the stack
            If mode = FLOAT_MODE Then
                op = stack(0)
                For i = 0 To STACK_SIZE - 2
                    stack(i) = stack(i + 1)
                Next
                stack(STACK_SIZE - 1) = op
            Else
                op64 = stack64(0)
                For i = 0 To STACK_SIZE - 2
                    stack64(i) = stack64(i + 1)
                Next
                stack64(STACK_SIZE - 1) = op64
            End If
        Case K_XY               ' exchange X and Y
            If mode = FLOAT_MODE Then
                op = pop()
                current = pop()
                push (op)
                push (current)
            Else
                op64 = pop64()
                current64 = pop64()
                push64 op64
                push64 current64
            End If
        Case K_ENTER            ' enter key
            ' clear overflow
            DispFlags DFLAG_OVERFLOW, False
            dont_push = True
        Case K_MINUS            ' subtraction
            If mode = FLOAT_MODE Then
                lastx = stack(0)
                op = pop()
                current = pop() - op
                push (current)
            Else
                lastx64 = stack64(0)
                op64 = pop64()
                current64 = BigSub(pop64(), op64)
                push64 current64
                If carry_bit Then
                    DispFlags DFLAG_CARRYBIT, True
                Else
                    DispFlags DFLAG_CARRYBIT, True
                End If
            End If
        Case K_STO              ' store into a register
            two_part = K_STO
        Case K_RCL              ' recall from a register
            two_part = K_RCL
        Case K_CHS              ' change sign
            If mode = FLOAT_MODE Then
                current = pop() * -1#
                push (current)
            Else
                lastx64 = stack64(0)
                ' It's not considered an error to "change the sign"
                ' while in the unsigned mode???
                If comp = 1 Then
                    current64 = BigNot(pop64())
                Else
                    current64 = BigAdd(BigNot(pop64()), BigFromInt(1))
                End If
                push64 current64
            End If
        Case K_PLUS             ' addition
            If mode = FLOAT_MODE Then
                lastx = stack(0)
                current = pop() + pop()
                push (current)
            Else
                lastx64 = stack64(0)
                current64 = BigAdd(pop64(), pop64())
                push64 current64
                If carry_bit Then
                    DispFlags DFLAG_CARRYBIT, True
                Else
                    DispFlags DFLAG_CARRYBIT, True
                End If
            End If
        '
        ' and now for the yellow/f (ctrl) functions
        '
        Case F_SL               ' shift left
            If mode = FLOAT_MODE Then
                Beep
                Exit Sub
            End If
            lastx64 = stack64(0)
            current64 = BigLeft(pop64())
            If carry_bit Then
                DispFlags DFLAG_CARRYBIT, True
            Else
                DispFlags DFLAG_CARRYBIT, False
            End If
            push64 current64
        Case F_SR               ' shift right
            If mode = FLOAT_MODE Then
                Beep
                Exit Sub
            End If
            lastx64 = stack64(0)
            current64 = BigRight(pop64())
            If carry_bit Then
                DispFlags DFLAG_CARRYBIT, True
            Else
                DispFlags DFLAG_CARRYBIT, False
            End If
            push64 current64
        Case F_RL               ' rotate left
            If mode = FLOAT_MODE Then
                Beep
                Exit Sub
            End If
            lastx64 = stack64(0)
            current64 = BigLeft(pop64())
            If carry_bit Then
                DispFlags DFLAG_CARRYBIT, True
                current64 = BigOr(current64, BigFromInt(1))
            Else
                DispFlags DFLAG_CARRYBIT, False
            End If
            push64 current64
        Case F_RR               ' rotate right
            If mode = FLOAT_MODE Then
                Beep
                Exit Sub
            End If
            lastx64 = stack64(0)
            current64 = BigRight(pop64())
            If carry_bit Then
                DispFlags DFLAG_CARRYBIT, True
                current64 = BigOr(current64, sign_mask)
            Else
                DispFlags DFLAG_CARRYBIT, False
            End If
            push64 current64
        Case F_RLN              ' rotate left "n" times
            If mode = FLOAT_MODE Then
                Beep
                Exit Sub
            End If
            lastx64 = stack64(0)
            ' the number of times to do it
            j = BigToInt(pop64())
            If j < 1 Or j > word_size Then
                Calc!Display.Text = "Error 2 - Improper Bit Number"
                Exit Sub
            End If
            current64 = pop64()
            For i = 1 To j
                current64 = BigLeft(current64)
                If carry_bit Then
                    current64 = BigOr(current64, BigFromInt(1))
                End If
            Next
            If carry_bit Then
                DispFlags DFLAG_CARRYBIT, True
            Else
                DispFlags DFLAG_CARRYBIT, False
            End If
            push64 current64
        Case F_RRN              ' rotate right "n" times
            If mode = FLOAT_MODE Then
                Beep
                Exit Sub
            End If
            lastx64 = stack64(0)
            ' the number of times to do it
            j = BigToInt(pop64())
            If j < 1 Or j > word_size Then
                Calc!Display.Text = "Error 2 - Improper Bit Number"
                Exit Sub
            End If
            current64 = pop64()
            For i = 1 To j
                current64 = BigRight(current64)
                If carry_bit Then
                    current64 = BigOr(current64, sign_mask)
                End If
            Next
            If carry_bit Then
                DispFlags DFLAG_CARRYBIT, True
            Else
                DispFlags DFLAG_CARRYBIT, False
            End If
            push64 current64
        Case F_MASKL            ' mask left
            If mode = FLOAT_MODE Then
                Beep
                Exit Sub
            End If
            lastx64 = stack64(0)
            i = BigToInt(pop64())
            If i < 1 Or i > word_size Then
                Calc!Display.Text = "Error 2 - Improper Bit Number"
                Exit Sub
            End If
            current64 = BigWordMask(i)
            push64 current64
        Case F_MASKR            ' mask right
            If mode = FLOAT_MODE Then
                Beep
                Exit Sub
            End If
            lastx64 = stack64(0)
            i = BigToInt(pop64())
            If i < 1 Or i > word_size Then
                Calc!Display.Text = "Error 2 - Improper Bit Number"
                Exit Sub
            End If
            current64 = BigNot(BigWordMask(word_size - i))
            push64 current64
        Case F_RMD              ' remainder (modulus)
            If mode = FLOAT_MODE Then
                Beep
                Exit Sub
            End If
            lastx64 = stack64(0)
            op64 = pop64()
            If BigIsZero(op64) Then
                Calc!Display.Text = "Error 0 - Improper Math Operation"
                Exit Sub
            End If
            current64 = BigMod(pop64(), op64)
            push64 current64
            If carry_bit Then
                DispFlags DFLAG_CARRYBIT, True
            Else
                DispFlags DFLAG_CARRYBIT, True
            End If
        Case F_XOR              ' exclusive or
            If mode = FLOAT_MODE Then
                Beep
                Exit Sub
            End If
            lastx64 = stack64(0)
            op64 = pop64()
            current64 = BigXor(pop64(), op64)
            push64 current64
            If carry_bit Then
                DispFlags DFLAG_CARRYBIT, True
            Else
                DispFlags DFLAG_CARRYBIT, True
            End If
        Case F_XIND             ' exchange x with contents of index
            Calc!Display.Text = "x:(i) not implemented"
            Exit Sub
        Case F_XI               ' exchange x with index
            Calc!Display.Text = "x:i not implemented"
            Exit Sub
        Case F_SHEX             ' show hex
            Format_Display64 BigFromDbl(stack(0)), HEX_MODE
            Exit Sub
        Case F_SDEC             ' show decimal
            Format_Display64 BigFromDbl(stack(0)), DEC_MODE
            Exit Sub
        Case F_SOCT             ' show octal
            Format_Display64 BigFromDbl(stack(0)), OCT_MODE
            Exit Sub
        Case F_SBIN             ' show binary
            Format_Display64 BigFromDbl(stack(0)), BIN_MODE
            Exit Sub
        Case F_SB               ' set bit
            If mode = FLOAT_MODE Then
                Beep
                Exit Sub
            End If
            lastx64 = stack64(0)
            i = BigToInt(pop64())
            If i < 0 Or i > word_size Then
                Calc!Display.Text = "Error 2 - Improper Bit Number"
                Exit Sub
            End If
            op64 = BigSignMask(i + 1)
            current64 = BigOr(pop64(), op64)
            push64 current64
        Case F_CB               ' clear bit
            If mode = FLOAT_MODE Then
                Beep
                Exit Sub
            End If
            lastx64 = stack64(0)
            i = BigToInt(pop64())
            If i < 0 Or i > word_size Then
                Calc!Display.Text = "Error 2 - Improper Bit Number"
                Exit Sub
            End If
            op64 = BigNot(BigSignMask(i + 1))
            current64 = BigAnd(pop64(), op64)
            push64 current64
        Case F_BSET             ' is bit set?
            Calc!Display.Text = "B? not implemented"
            Exit Sub
        Case F_AND              ' logical and
            If mode = FLOAT_MODE Then
                Beep
                Exit Sub
            End If
            lastx64 = stack64(0)
            current64 = BigAnd(pop64(), pop64())
            push64 current64
        Case F_IND              ' contents of index
            Calc!Display.Text = "(i) not implemented"
            Exit Sub
        Case F_I                ' index
            Calc!Display.Text = "I not implemented"
            Exit Sub
        Case F_CPGRM            ' clear program
            Calc!Display.Text = "CPGRM not implemented"
            Exit Sub
        Case F_CREG             ' clear registers
            If mode = FLOAT_MODE Then
                For i = 0 To REG_SIZE - 1
                    reg(i) = 0#
                Next
            Else
                For i = 0 To REG_SIZE - 1
                    reg64(i) = zero
                Next
            End If
        Case F_WIN              ' set window number
           ' I make a departure from the real calculator.  I silently
           ' ignore the WINDOW command, and use the "<" and ">" operators
           ' to move the entire 32bit window.
        Case F_S1S              ' 1s complement
            comp = 1
            Calc.CM_OPTION_1S.Checked = True
            Calc.CM_OPTION_2S.Checked = False
            Calc.CM_OPTION_UNS.Checked = False
        Case F_S2S              ' 2s complement
            comp = 2
            Calc.CM_OPTION_1S.Checked = False
            Calc.CM_OPTION_2S.Checked = True
            Calc.CM_OPTION_UNS.Checked = False
        Case F_SUNS             ' unsigned complement
            comp = 0
            Calc.CM_OPTION_1S.Checked = False
            Calc.CM_OPTION_2S.Checked = False
            Calc.CM_OPTION_UNS.Checked = True
        Case F_NOT              ' logical not
            If mode = FLOAT_MODE Then
                Beep
                Exit Sub
            End If
            lastx64 = stack64(0)
            current64 = BigNot(pop64())
            push64 current64
        Case F_WSIZE            ' word size
            If mode = FLOAT_MODE Then
                Beep
                Exit Sub
            End If
            lastx64 = stack64(0)
            i = BigToInt(pop64())
            If (i < 1 Or i > 64) Then
                Calc!Display.Text = "Error 2 - Improper Bit Number"
                Exit Sub
            End If
            word_mask = BigWordMask(i)
            sign_mask = BigSignMask(i)
            ' convert all existing number to fit into the smaller mask
            If word_size > i Then
                MaskAll
            End If
            word_size = i
            ' turn off all the checks (could be an "non-standard" size)
            Calc.CM_OPTION_8.Checked = False
            Calc.CM_OPTION_16.Checked = False
            Calc.CM_OPTION_32.Checked = False
            Calc.CM_OPTION_64.Checked = False
            Select Case word_size
                Case 8
                    Calc.CM_OPTION_8.Checked = True
                Case 16
                    Calc.CM_OPTION_16.Checked = True
                Case 32
                    Calc.CM_OPTION_32.Checked = True
                Case 64
                    Calc.CM_OPTION_64.Checked = True
            End Select
        Case F_FLOAT
            two_part = F_FLOAT
            Exit Sub
        Case F_MEM              ' display memory
            Calc!Display.Text = "MEM not implemented"
            Exit Sub
        Case F_STAT             ' display status
            RSet screen_buf = comp & "-" & Format(word_size, "00") & "-" & Right(Format(flag(3), "0"), 1) & Format(flag(2), "0") & Right(Format(flag(1), "0"), 1) & Right(Format(flag(0), "0"), 1)
            Calc!Display.Text = screen_buf
            Exit Sub
        Case F_OR               ' logical or
        '
        ' and now for the blue/g (alt) functions
        '
        Case G_LJ               ' left justify
            If mode = FLOAT_MODE Then
                Beep
                Exit Sub
            End If
            lastx64 = stack64(0)
            op64 = pop64()
            ' short cut for zero
            If BigIsZero(op64) Then
                current64 = op64
                push64 zero
            Else
                ' where is first bit?
                i = 0
                For n_byte = 7 To 0 Step -1
                    If op64.n(n_byte) <> 0 Then
                        Exit For
                    End If
                    i = i + 8
                Next
                For n_bit = 7 To 0 Step -1
                    If (op64.n(n_byte) And (2 ^ n_bit)) <> 0 Then
                        Exit For
                    End If
                    i = i + 1
                Next
                i = word_size - 64 + i
                ' now shift left that many times
                current64 = op64
                For j = 1 To i
                    current64 = BigLeft(current64)
                Next
            End If
            push64 current64
            push64 BigFromInt(i)
        Case G_ASR              ' arithemic shift right
            If mode = FLOAT_MODE Then
                Beep
                Exit Sub
            End If
            lastx64 = stack64(0)
            op64 = pop64()
            ' if negative
            If BigIsZero(BigAnd(op64, sign_mask)) = False Then
                current64 = BigOr(BigLeft(op64), sign_mask)
            Else
                current64 = BigLeft(op64)
            End If
            push64 current64
        Case G_RLC              ' rotate left with carry
            If mode = FLOAT_MODE Then
                Beep
                Exit Sub
            End If
            lastx64 = stack64(0)
            ' the previous carry bit
            If carry_bit Then
                current64 = BigOr(BigLeft(pop64()), BigFromInt(1))
            Else
                current64 = BigLeft(pop64())
            End If
            ' the new carry bit
            If carry_bit Then
                DispFlags DFLAG_CARRYBIT, True
            Else
                DispFlags DFLAG_CARRYBIT, False
            End If
            push64 current64
        Case G_RRC              ' rotate right with carry
            If mode = FLOAT_MODE Then
                Beep
                Exit Sub
            End If
            lastx64 = stack64(0)
            ' the previous carry bit
            If carry_bit Then
                current64 = BigOr(BigRight(pop64()), sign_mask)
            Else
                current64 = BigRight(pop64())
            End If
            ' the new carry bit
            If carry_bit Then
                DispFlags DFLAG_CARRYBIT, True
            Else
                DispFlags DFLAG_CARRYBIT, False
            End If
            push64 current64
        Case G_RLCN             ' rotate left with carry "n" times
            If mode = FLOAT_MODE Then
                Beep
                Exit Sub
            End If
            lastx64 = stack64(0)
            ' the number of times to do it
            j = BigToInt(pop64())
            If j < 1 Or j > word_size Then
                Calc!Display.Text = "Error 2 - Improper Bit Number"
                Exit Sub
            End If
            current64 = pop64()
            For i = 1 To j
                If carry_bit Then
                    current64 = BigOr(BigRight(pop64()), BigFromInt(1))
                Else
                    current64 = BigRight(pop64())
                End If
            Next
            If carry_bit Then
                DispFlags DFLAG_CARRYBIT, True
            Else
                DispFlags DFLAG_CARRYBIT, False
            End If
            push64 current64
        Case G_RRCN             ' rotate right with carry "n" times
            If mode = FLOAT_MODE Then
                Beep
                Exit Sub
            End If
            lastx64 = stack64(0)
            ' the number of times to do it
            j = BigToInt(pop64())
            If j < 1 Or j > word_size Then
                Calc!Display.Text = "Error 2 - Improper Bit Number"
                Exit Sub
            End If
            current64 = pop64()
            For i = 1 To j
                If carry_bit Then
                    current64 = BigOr(BigRight(pop64()), sign_mask)
                Else
                    current64 = BigRight(pop64())
                End If
            Next
            If carry_bit Then
                DispFlags DFLAG_CARRYBIT, True
            Else
                DispFlags DFLAG_CARRYBIT, False
            End If
            push64 current64
        Case G_BITS             ' sum the number of bits
            If mode = FLOAT_MODE Then
                Beep
                Exit Sub
            End If
            lastx64 = stack64(0)
            op64 = pop64()
            i = 0
            For n_byte = 0 To 7
                For n_bit = 0 To 7
                    If (op64.n(n_byte) And (2 ^ n_bit)) <> 0 Then
                        i = i + 1
                    End If
                Next
            Next
            current64 = BigFromInt(i)
            push64 current64
        Case G_ABS              ' absolute value
            If mode = FLOAT_MODE Then
                lastx = stack(0)
                current = Abs(pop())
                push current
            Else
                lastx64 = stack64(0)
                op64 = pop64()
                If BigIsZero(BigAnd(op64, sign_mask)) = False Then
                    If comp = 1 Then
                        current64 = BigNot(op64)
                    Else
                        current64 = BigAdd(BigNot(op64), BigFromInt(1))
                    End If
                Else
                    current64 = op64
                End If
                push64 current64
            End If
        Case G_DBLR             ' double remainder
            If mode = FLOAT_MODE Then
                Beep
                Exit Sub
            End If
            lastx64 = stack64(0)
            op64 = pop64()
            current64 = pop64()
            current64 = DblMod(op64, current64, pop64())
            push64 current64
        Case G_DDIV             ' double divide
            If mode = FLOAT_MODE Then
                Beep
                Exit Sub
            End If
            lastx64 = stack64(0)
            op64 = pop64()
            current64 = pop64()
            current64 = DblDiv(op64, current64, pop64())
            push64 current64
        Case G_RTN              ' return
            Calc!Display.Text = "RTN not implemented"
            Exit Sub
        Case G_LBL              ' label
            Calc!Display.Text = "LBL not implemented"
            Exit Sub
        Case G_DSZ              ' decrement skip on zero
            Calc!Display.Text = "DSZ not implemented"
            Exit Sub
        Case G_ISZ              ' increment skip on zero
            Calc!Display.Text = "ISZ not implemented"
            Exit Sub
        Case G_SQRT             ' square root
            If mode = FLOAT_MODE Then
                lastx = stack(0)
                op = pop()
                If op < 0 Then
                    Calc!Display.Text = "Error 0 - Improper Math Operation"
                    Exit Sub
                End If
                current = Sqr(op)
                push (current)
            Else
                lastx64 = stack64(0)
                op64 = pop64()
                If BigIsZero(op64) Then
                    Calc!Display.Text = "Error 0 - Improper Math Operation"
                    Exit Sub
                End If
                current64 = BigSqrt(op64)
                push64 current64
                If carry_bit Then
                    DispFlags DFLAG_CARRYBIT, True
                Else
                    DispFlags DFLAG_CARRYBIT, False
                End If
            End If
        Case G_INV              ' inverse
            If mode <> FLOAT_MODE Then
                Beep
                Exit Sub
            End If
            lastx = stack(0)
            op = pop()
            If op = 0 Then
                Calc!Display.Text = "Error 0 - Improper Math Operation"
                Exit Sub
            End If
            current = 1 / op
            push (current)
        Case G_SF               ' set flag
            two_part = G_SF
        Case G_CF               ' clear flag
            two_part = G_CF
        Case G_FSET             ' is flag set?
            Calc!Display.Text = "F? not implemented"
            Exit Sub
        Case G_DMUL             ' double multiply
            If mode = FLOAT_MODE Then
                Beep
                Exit Sub
            End If
            lastx64 = stack64(0)
            DblMult stack64(0), stack64(1)
            ' double multiply returns two values... directly to the stack!
        Case G_PR               ' program / run mode
            Calc!Display.Text = "PR not implemented"
            Exit Sub
        Case G_BST              ' back step
            Calc!Display.Text = "BST not implemented"
            Exit Sub
        Case G_RUP              ' roll stack up
            If mode = FLOAT_MODE Then
                op = stack(STACK_SIZE - 1)
                For i = STACK_SIZE - 1 To 1 Step -1
                    stack(i) = stack(i - 1)
                Next
                stack(0) = op
            Else
                op64 = stack64(STACK_SIZE - 1)
                For i = STACK_SIZE - 1 To 1 Step -1
                    stack64(i) = stack64(i - 1)
                Next
                stack64(0) = op64
            End If
        Case G_PSE              ' pause
            Calc!Display.Text = "PSE not implemented"
            Exit Sub
        Case G_CLX              ' clear x
            If mode = FLOAT_MODE Then
                stack(0) = 0#
            Else
                stack64(0) = zero
            End If
            dont_push = True
        Case G_LSTX             ' last x
            If mode = FLOAT_MODE Then
                push lastx
            Else
                push64 lastx64
            End If
        Case G_XGEY             ' x >= y
            Calc!Display.Text = "X>=Y not implemented"
            Exit Sub
        Case G_XL0              ' x < 0
            Calc!Display.Text = "X<0 not implemented"
            Exit Sub
        Case G_XGY              ' x > y
            Calc!Display.Text = "X>Y not implemented"
            Exit Sub
        Case G_XG0              ' x > 0
            Calc!Display.Text = "X>0 not implemented"
            Exit Sub
        Case G_WLEFT            ' move window to the left
            ' This is significantly different from the real calculator.
            ' Instead of moving the window 1 bit at a time, it gets
            ' moved 32 bits.
            If mode <> BIN_MODE Then
                Beep
                Exit Sub
            End If
            If word_size <= 32 Then
                ' ignore if nothing to do
                Exit Sub
            End If
            shift_window = True
        Case G_WRIGHT           ' move window to the right
            ' This is significantly different from the real calculator.
            ' Instead of moving the window 1 bit at a time, it gets
            ' moved 32 bits.
            If mode <> BIN_MODE Then
                Beep
                Exit Sub
            End If
            ' actually, since "right window" is the default, we don't
            ' have to do anything!
        Case G_XNEY             ' x <> y
            Calc!Display.Text = "X<>Y not implemented"
            Exit Sub
        Case G_XNE0             ' x <> 0
            Calc!Display.Text = "X<>0 not implemented"
            Exit Sub
        Case G_XEY              ' x = y
            Calc!Display.Text = "X=Y not implemented"
            Exit Sub
        Case G_XE0              ' x = 0
            Calc!Display.Text = "X=0 not implemented"
            Exit Sub
        Case Else
            Beep
    End Select
    
    ' display the answer
    If mode = FLOAT_MODE Then
        Format_Display stack(0)
    Else
        Format_Display64 stack64(0), mode
    End If
End Sub
Public Sub DispFlags(f As Integer, onoff As Boolean)
    ' display the annunciator flags
    Static dflag(0 To 3) As String
    
    Select Case f
        Case DFLAG_FKEY     ' f function key
            If onoff = True Then
                dflag(0) = Space(AN_WIDTH / 2) & "f" & Space(AN_WIDTH / 2)
            Else
                dflag(0) = Space(AN_WIDTH)
            End If
        Case DFLAG_GKEY     ' g function key
            If onoff = True Then
                dflag(1) = Space(AN_WIDTH / 2) & "g" & Space(AN_WIDTH / 2)
            Else
                dflag(1) = Space(AN_WIDTH)
            End If
        Case DFLAG_OVERFLOW ' overflow
            If onoff = True Then
                flag(5) = True
                Calc.CM_OPTION_FLAG5.Checked = True
                dflag(2) = Space(AN_WIDTH / 2) & "G" & Space(AN_WIDTH / 2)
            Else
                flag(5) = False
                Calc.CM_OPTION_FLAG5.Checked = False
                dflag(2) = Space(AN_WIDTH)
            End If
        Case DFLAG_CARRYBIT ' carry bit
            If onoff = True Then
                flag(4) = True
                Calc.CM_OPTION_FLAG4.Checked = True
                dflag(3) = Space(AN_WIDTH / 2) & "C" & Space(AN_WIDTH / 2)
            Else
                flag(4) = False
                Calc.CM_OPTION_FLAG4.Checked = False
                dflag(3) = Space(AN_WIDTH)
            End If
    End Select
    
    Calc!Annunciator.Text = dflag(0) & dflag(1) & dflag(2) & dflag(3)
End Sub
Public Sub push(val As Double)
    ' Push a value onto the stack.  If full, the top of the stack is lost
    Dim i As Integer
    
    For i = STACK_SIZE - 1 To 1 Step -1
        stack(i) = stack(i - 1)
    Next
    stack(0) = val
End Sub
Public Sub push64(val As BigInt)
    ' Push a value onto the stack.  If full, the top of the stack is lost
    Dim i As Integer
    
    For i = STACK_SIZE - 1 To 1 Step -1
        stack64(i) = stack64(i - 1)
    Next
    stack64(0) = val
End Sub
Private Function pop() As Double
    ' Pop a value off of the stack.  The top value is preserved
    Dim f As Double
    Dim i As Integer
    
    f = stack(0)
    lastx = f
    For i = 0 To STACK_SIZE - 2
        stack(i) = stack(i + 1)
    Next
    pop = f
End Function
Private Function pop64() As BigInt
    ' Pop a value off of the stack.  The top value is preserved
    Dim f As BigInt
    Dim i As Integer
    
    f = stack64(0)
    lastx64 = f
    For i = 0 To STACK_SIZE - 2
        stack64(i) = stack64(i + 1)
    Next
    pop64 = f
End Function
Public Function Convert_Input(str As String) As Double
    Dim f As Double
    
    On Error Resume Next
    f = CDbl(str)
    On Error GoTo 0
    
    Convert_Input = f
End Function
Public Function Convert_Input64(str As String, mode As Integer) As BigInt
    ' Convert ASCII input into a number
    Dim f As BigInt

    Select Case mode
        Case HEX_MODE
            f = BigFromHexStr(str)
        Case DEC_MODE
            f = BigFromDecStr(str)
        Case OCT_MODE
            f = BigFromOctStr(str)
        Case BIN_MODE
            f = BigFromBinStr(str)
    End Select
    
    Convert_Input64 = f
End Function
Public Sub Format_Display(f As Double)
    Dim i As Integer
    Dim buf As String
    Dim point_mask As String
    Dim screen_buf As String * SCREEN_SIZE
    
    ' scientific notation?
    If dpoints < 0 Then
        ' Note: dpoints is negative because I interpreted the "."
        ' character as an ASCII digit.
        point_mask = "0.000000000e-"
        screen_buf = Format(f, point_mask)
    Else
        point_mask = "0."
        For i = 1 To dpoints
            point_mask = point_mask & "0"
        Next
        buf = Format(f, point_mask)
        ' run out of screen space?  Try scientific notation
        If (Len(buf) >= SCREEN_SIZE) Or ((buf = point_mask) And (f <> 0#)) Then
            point_mask = point_mask & "e-"
            screen_buf = Format(f, point_mask)
        Else
            screen_buf = buf
        End If
    End If
    Calc!Display.Text = screen_buf
    
    ' every once and a while, we should check to see if there is
    ' something on the clipboard with the correct format
    If Clipboard.GetFormat(vbCFText) = True Then
        Calc.CM_EDIT_PASTE.Enabled = True
    Else
        Calc.CM_EDIT_PASTE.Enabled = False
    End If
End Sub
Public Sub Format_Display64(f As BigInt, mode As Integer)
    ' Format a number for the display
    Dim buf As String
    Dim start As Integer
    Dim screen_buf As String * SCREEN_SIZE
    
    Select Case mode
        Case HEX_MODE
            buf = LCase(BigToHexStr(f))
            If flag(3) = False Then
                buf = TrimZero(buf)
            End If
            RSet screen_buf = buf & " h "
        Case OCT_MODE
            buf = BigToOctStr(f)
            If flag(3) = False Then
                buf = TrimZero(buf)
            End If
            RSet screen_buf = buf & " o "
        Case DEC_MODE
            buf = BigToDecStr(f, comp)
            If flag(3) = False Then
                buf = TrimZero(buf)
            End If
            RSet screen_buf = buf & " d "
        Case BIN_MODE
            ' the binary mode is the only mode which can't always fit the
            ' answer on the screen.  So we sometimes display only half of
            ' the answer.
            If shift_window Then
                buf = Left(BigToBinStr(f), 36)
                start = SCREEN_SIZE - word_size - ((word_size - 32) / 8) + 31
                ' there is no need to trim leading zeros in binary mode.
                RSet screen_buf = Mid(buf, start) & "b."
                shift_window = False
            Else
                buf = Right(BigToBinStr(f), 36)
                If word_size <= 32 Then
                    ' mask the length of the bit pattern to the word size
                    start = SCREEN_SIZE - word_size - (word_size / 8) - 1
                    RSet screen_buf = Mid(buf, start) & "b "
                Else
                    RSet screen_buf = buf & ".b"
                End If
            End If
    End Select
    Calc!Display.Text = screen_buf
    
    ' every once and a while, we should check to see if there is
    ' something on the clipboard with the correct format
    If Clipboard.GetFormat(vbCFText) = True Then
        Calc.CM_EDIT_PASTE.Enabled = True
    Else
        Calc.CM_EDIT_PASTE.Enabled = False
    End If
End Sub
Sub Raw_Display(input_buf As String, mode As Integer)
    ' format the raw (text only) input for the screen
    Dim buf As String * SCREEN_SIZE
    
    Select Case mode
        Case FLOAT_MODE
            buf = input_buf
        Case HEX_MODE
            RSet buf = input_buf & " h "
        Case OCT_MODE
            RSet buf = input_buf & " o "
        Case DEC_MODE
            RSet buf = input_buf & " d "
        Case BIN_MODE
            RSet buf = input_buf & " b "
    End Select
    Calc!Display.Text = buf
End Sub
Function TrimZero(buf As String) As String
    Dim i As Integer
    Dim b As String
    Dim is_neg As Boolean
    
    is_neg = False
    If Left(buf, 1) = "-" Then
        i = 1
        is_neg = True
    End If
    
    Do
        i = i + 1
        b = Mid(buf, i, 1)
    Loop While b = "0"

    ' leave at least the last zero
    If i > Len(buf) Then
        i = i - 1
    End If
    
    If is_neg Then
        TrimZero = "-" & Mid(buf, i)
    Else
        TrimZero = Mid(buf, i)
    End If
End Function
Sub AllTo64()
    ' convert registers, stack, and lastx to BigInt.  Loss of precision
    ' is ignored.
    Dim i As Integer
    Dim zero As BigInt
        
    For i = 0 To REG_SIZE - 1
        If reg(i) = 0# Then
            reg64(i) = zero
        Else
            reg64(i) = BigFromDbl(reg(i))
        End If
    Next
    For i = 0 To STACK_SIZE - 1
        If stack(i) = 0# Then
            stack64(i) = zero
        Else
            stack64(i) = BigFromDbl(stack(i))
        End If
    Next
    If lastx = 0# Then
        lastx64 = zero
    Else
        lastx64 = BigFromDbl(lastx)
    End If
End Sub
Sub AllFrom64()
    ' convert registers, stack, and lastx from BigInt.  Loss of precision
    ' is ignored.
    Dim i As Integer
    
    For i = 0 To REG_SIZE - 1
        If BigIsZero(reg64(i)) Then
            reg(i) = 0#
        Else
            reg(i) = BigToDbl(reg64(i))
        End If
    Next
    For i = 0 To STACK_SIZE - 1
        If BigIsZero(stack64(i)) Then
            stack(i) = 0#
        Else
            stack(i) = BigToDbl(stack64(i))
        End If
    Next
    If BigIsZero(lastx64) Then
        lastx = 0#
    Else
        lastx = BigToDbl(lastx64)
    End If
End Sub
Sub MaskAll()
    ' Convert all registers, stack, and lastx to a new (smaller) word
    ' length.  When going the other way (to a larger word size), we
    ' don't do anything (so loss of sign bit will occur, yikes!).
    Dim i As Integer
    
    For i = 0 To REG_SIZE - 1
        reg64(i) = BigAnd(reg64(i), word_mask)
    Next
    For i = 0 To STACK_SIZE - 1
        stack64(i) = BigAnd(stack64(i), word_mask)
    Next
    lastx64 = BigAnd(lastx64, word_mask)
End Sub
