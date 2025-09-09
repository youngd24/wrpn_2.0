Attribute VB_Name = "Header"
Option Explicit

Public Const CTRL = 128
Public Const ALT = 256

' Key codes for function keys, numbers, and hex
' digits are their ASCII equivalents
Public Const K_A = 97       'a
Public Const K_B = 98       'b
Public Const K_C = 99       'c
Public Const K_D = 100      'd
Public Const K_E = 101      'e
Public Const K_F = 102      'f
Public Const K_7 = 55       '7
Public Const K_8 = 56       '8
Public Const K_9 = 57       '9
Public Const K_DIV = 47     '/
Public Const K_GSB = 71     'G
Public Const K_GTO = 116    't (gTo)
Public Const K_HEX = 104    'h
Public Const K_DEC = 68     'D
Public Const K_OCT = 111    'o
Public Const K_BIN = 66     'B
Public Const K_4 = 52       '4
Public Const K_5 = 53       '5
Public Const K_6 = 54       '6
Public Const K_MULT = 42    '*
Public Const K_RS = 82      'R
Public Const K_SST = 83     'S
Public Const K_ROL = 108    'l (roL)
Public Const K_XY = 120     'x
Public Const K_BSP = 8      '\b
Public Const K_ENTER = 13   '\r
Public Const K_1 = 49       '1
Public Const K_2 = 50       '2
Public Const K_3 = 51       '3
Public Const K_MINUS = 45   '-
Public Const K_ON = 113     'q (Quit)
Public Const K_FKEY = 1     ' ctrl key
Public Const K_GKEY = 2     ' alt key
Public Const K_STO = 115    's
Public Const K_RCL = 114    'r
Public Const K_0 = 48       '0
Public Const K_DP = 46      '.
Public Const K_CHS = 67     'C
Public Const K_PLUS = 43    '+
' Key codes for the yellow (cntrl) functions are created
' by adding 128 to the K_ codes above
Public Const F_SL = CTRL + K_A
Public Const F_SR = CTRL + K_B
Public Const F_RL = CTRL + K_C
Public Const F_RR = CTRL + K_D
Public Const F_RLN = CTRL + K_E
Public Const F_RRN = CTRL + K_F
Public Const F_MASKL = CTRL + K_7
Public Const F_MASKR = CTRL + K_8
Public Const F_RMD = CTRL + K_9
Public Const F_XOR = CTRL + K_DIV
Public Const F_XIND = CTRL + K_GSB
Public Const F_XI = CTRL + K_GTO
Public Const F_SHEX = CTRL + K_HEX
Public Const F_SDEC = CTRL + K_DEC
Public Const F_SOCT = CTRL + K_OCT
Public Const F_SBIN = CTRL + K_BIN
Public Const F_SB = CTRL + K_4
Public Const F_CB = CTRL + K_5
Public Const F_BSET = CTRL + K_6
Public Const F_AND = CTRL + K_MULT
Public Const F_IND = CTRL + K_RS
Public Const F_I = CTRL + K_SST
Public Const F_CPGRM = CTRL + K_ROL
Public Const F_CREG = CTRL + K_XY
Public Const F_CPRE = CTRL + K_BSP
Public Const F_WIN = CTRL + K_ENTER
Public Const F_S1S = CTRL + K_1
Public Const F_S2S = CTRL + K_2
Public Const F_SUNS = CTRL + K_3
Public Const F_NOT = CTRL + K_MINUS
' ON
' FKEY
' GKEY
Public Const F_WSIZE = CTRL + K_STO
Public Const F_FLOAT = CTRL + K_RCL
Public Const F_MEM = CTRL + K_0
Public Const F_STAT = CTRL + K_DP
Public Const F_EEX = CTRL + K_CHS
Public Const F_OR = CTRL + K_PLUS
' Key codes for the blue (alt) functions are created by
' adding 256 to the K_* values
Public Const G_LJ = ALT + K_A
Public Const G_ASR = ALT + K_B
Public Const G_RLC = ALT + K_C
Public Const G_RRC = ALT + K_D
Public Const G_RLCN = ALT + K_E
Public Const G_RRCN = ALT + K_F
Public Const G_BITS = ALT + K_7
Public Const G_ABS = ALT + K_8
Public Const G_DBLR = ALT + K_9
Public Const G_DDIV = ALT + K_DIV
Public Const G_RTN = ALT + K_BSP
Public Const G_LBL = ALT + K_GTO
Public Const G_DSZ = ALT + K_HEX
Public Const G_ISZ = ALT + K_DEC
Public Const G_SQRT = ALT + K_OCT
Public Const G_INV = ALT + K_BIN
Public Const G_SF = ALT + K_4
Public Const G_CF = ALT + K_5
Public Const G_FSET = ALT + K_6
Public Const G_DMUL = ALT + K_MULT
Public Const G_PR = ALT + K_RS
Public Const G_BST = ALT + K_SST
Public Const G_RUP = ALT + K_ROL
Public Const G_PSE = ALT + K_XY
Public Const G_CLX = ALT + K_BSP
Public Const G_LSTX = ALT + K_ENTER
Public Const G_XGEY = ALT + K_1
Public Const G_XL0 = ALT + K_2
Public Const G_XGY = ALT + K_3
Public Const G_XG0 = ALT + K_MINUS
' ON
' FKEY
' GKEY
Public Const G_WLEFT = ALT + K_STO
Public Const G_WRIGHT = ALT + K_RCL
Public Const G_XNEY = ALT + K_0
Public Const G_XNE0 = ALT + K_DP
Public Const G_XEY = ALT + K_CHS
Public Const G_XE0 = ALT + K_PLUS

Public Const STACK_SIZE = 4
Public Const REG_SIZE = 32
Public Const SCREEN_SIZE = 38
Public Const FLOAT_MODE = 0
Public Const HEX_MODE = 1
Public Const DEC_MODE = 2
Public Const OCT_MODE = 3
Public Const BIN_MODE = 4

Public Const MAX_FLOAT_DIGITS = SCREEN_SIZE - 1
Public Const MAX_HEX_DIGITS = 16
Public Const MAX_OCTAL_DIGITS = 23
Public Const MAX_DECIMAL_DIGITS = 19
Public Const MAX_BINARY_DIGITS = 64

Public Const AN_WIDTH = 21

Public Const DFLAG_FKEY = 0
Public Const DFLAG_GKEY = 1
Public Const DFLAG_CARRYBIT = 4
Public Const DFLAG_OVERFLOW = 5


