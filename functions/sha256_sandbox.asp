<%
' http://www.freevbcode.com/ShowCode.asp?ID=2565

' See the VB6 project that accompanies this sample for full code comments on how
' it works.
'
' ASP VBScript code for generating a SHA256 'digest' or 'signature' of a string. The
' MD5 algorithm is one of the industry standard methods for generating digital
' signatures. It is generically known as a digest, digital signature, one-way
' encryption, hash or checksum algorithm. A common use for SHA256 is for password
' encryption as it is one-way in nature, that does not mean that your passwords
' are not free from a dictionary attack. 
'
' If you are using the routine for passwords, you can make it a little more secure
' by concatenating some known random characters to the password before you generate
' the signature and on subsequent tests, so even if a hacker knows you are using
' SHA-256 for your passwords, the random characters will make it harder to dictionary
' attack.
'
' NOTE: Due to the way in which the string is processed the routine assumes a
' single byte character set. VB passes unicode (2-byte) character strings, the
' ConvertToWordArray function uses on the first byte for each character. This
' has been done this way for ease of use, to make the routine truely portable
' you could accept a byte array instead, it would then be up to the calling
' routine to make sure that the byte array is generated from their string in
' a manner consistent with the string type.

'
' Web Site:  http://www.frez.co.uk
' E-mail:    sales@frez.co.uk

Private m_lOnBits2(30)
Private m_l2Power2(30)
Private K2(63)

Private Const BITS_TO_A_BYTE2 = 8
Private Const BYTES_TO_A_WORD2 = 4
Private Const BITS_TO_A_WORD2 = 32

m_lOnBits2(0) = CLng(1)
m_lOnBits2(1) = CLng(3)
m_lOnBits2(2) = CLng(7)
m_lOnBits2(3) = CLng(15)
m_lOnBits2(4) = CLng(31)
m_lOnBits2(5) = CLng(63)
m_lOnBits2(6) = CLng(127)
m_lOnBits2(7) = CLng(255)
m_lOnBits2(8) = CLng(511)
m_lOnBits2(9) = CLng(1023)
m_lOnBits2(10) = CLng(2047)
m_lOnBits2(11) = CLng(4095)
m_lOnBits2(12) = CLng(8191)
m_lOnBits2(13) = CLng(16383)
m_lOnBits2(14) = CLng(32767)
m_lOnBits2(15) = CLng(65535)
m_lOnBits2(16) = CLng(131071)
m_lOnBits2(17) = CLng(262143)
m_lOnBits2(18) = CLng(524287)
m_lOnBits2(19) = CLng(1048575)
m_lOnBits2(20) = CLng(2097151)
m_lOnBits2(21) = CLng(4194303)
m_lOnBits2(22) = CLng(8388607)
m_lOnBits2(23) = CLng(16777215)
m_lOnBits2(24) = CLng(33554431)
m_lOnBits2(25) = CLng(67108863)
m_lOnBits2(26) = CLng(134217727)
m_lOnBits2(27) = CLng(268435455)
m_lOnBits2(28) = CLng(536870911)
m_lOnBits2(29) = CLng(1073741823)
m_lOnBits2(30) = CLng(2147483647)

m_l2Power2(0) = CLng(1)
m_l2Power2(1) = CLng(2)
m_l2Power2(2) = CLng(4)
m_l2Power2(3) = CLng(8)
m_l2Power2(4) = CLng(16)
m_l2Power2(5) = CLng(32)
m_l2Power2(6) = CLng(64)
m_l2Power2(7) = CLng(128)
m_l2Power2(8) = CLng(256)
m_l2Power2(9) = CLng(512)
m_l2Power2(10) = CLng(1024)
m_l2Power2(11) = CLng(2048)
m_l2Power2(12) = CLng(4096)
m_l2Power2(13) = CLng(8192)
m_l2Power2(14) = CLng(16384)
m_l2Power2(15) = CLng(32768)
m_l2Power2(16) = CLng(65536)
m_l2Power2(17) = CLng(131072)
m_l2Power2(18) = CLng(262144)
m_l2Power2(19) = CLng(524288)
m_l2Power2(20) = CLng(1048576)
m_l2Power2(21) = CLng(2097152)
m_l2Power2(22) = CLng(4194304)
m_l2Power2(23) = CLng(8388608)
m_l2Power2(24) = CLng(16777216)
m_l2Power2(25) = CLng(33554432)
m_l2Power2(26) = CLng(67108864)
m_l2Power2(27) = CLng(134217728)
m_l2Power2(28) = CLng(268435456)
m_l2Power2(29) = CLng(536870912)
m_l2Power2(30) = CLng(1073741824)
    
K2(0) = &H428A2F98
K2(1) = &H71374491
K2(2) = &HB5C0FBCF
K2(3) = &HE9B5DBA5
K2(4) = &H3956C25B
K2(5) = &H59F111F1
K2(6) = &H923F82A4
K2(7) = &HAB1C5ED5
K2(8) = &HD807AA98
K2(9) = &H12835B01
K2(10) = &H243185BE
K2(11) = &H550C7DC3
K2(12) = &H72BE5D74
K2(13) = &H80DEB1FE
K2(14) = &H9BDC06A7
K2(15) = &HC19BF174
K2(16) = &HE49B69C1
K2(17) = &HEFBE4786
K2(18) = &HFC19DC6
K2(19) = &H240CA1CC
K2(20) = &H2DE92C6F
K2(21) = &H4A7484AA
K2(22) = &H5CB0A9DC
K2(23) = &H76F988DA
K2(24) = &H983E5152
K2(25) = &HA831C66D
K2(26) = &HB00327C8
K2(27) = &HBF597FC7
K2(28) = &HC6E00BF3
K2(29) = &HD5A79147
K2(30) = &H6CA6351
K2(31) = &H14292967
K2(32) = &H27B70A85
K2(33) = &H2E1B2138
K2(34) = &H4D2C6DFC
K2(35) = &H53380D13
K2(36) = &H650A7354
K2(37) = &H766A0ABB
K2(38) = &H81C2C92E
K2(39) = &H92722C85
K2(40) = &HA2BFE8A1
K2(41) = &HA81A664B
K2(42) = &HC24B8B70
K2(43) = &HC76C51A3
K2(44) = &HD192E819
K2(45) = &HD6990624
K2(46) = &HF40E3585
K2(47) = &H106AA070
K2(48) = &H19A4C116
K2(49) = &H1E376C08
K2(50) = &H2748774C
K2(51) = &H34B0BCB5
K2(52) = &H391C0CB3
K2(53) = &H4ED8AA4A
K2(54) = &H5B9CCA4F
K2(55) = &H682E6FF3
K2(56) = &H748F82EE
K2(57) = &H78A5636F
K2(58) = &H84C87814
K2(59) = &H8CC70208
K2(60) = &H90BEFFFA
K2(61) = &HA4506CEB
K2(62) = &HBEF9A3F7
K2(63) = &HC67178F2

Private Function LShift(lValue, iShiftBits)
    If iShiftBits = 0 Then
        LShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And 1 Then
            LShift = &H80000000
        Else
            LShift = 0
        End If
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If
    
    If (lValue And m_l2Power2(31 - iShiftBits)) Then
        LShift = ((lValue And m_lOnBits2(31 - (iShiftBits + 1))) * m_l2Power2(iShiftBits)) Or &H80000000
    Else
        LShift = ((lValue And m_lOnBits2(31 - iShiftBits)) * m_l2Power2(iShiftBits))
    End If
End Function

Private Function RShift(lValue, iShiftBits)
    If iShiftBits = 0 Then
        RShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And &H80000000 Then
            RShift = 1
        Else
            RShift = 0
        End If
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If
    
    RShift = (lValue And &H7FFFFFFE) \ m_l2Power2(iShiftBits)
    
    If (lValue And &H80000000) Then
        RShift = (RShift Or (&H40000000 \ m_l2Power2(iShiftBits - 1)))
    End If
End Function

Private Function AddUnsigned(lX, lY)
    Dim lX4
    Dim lY4
    Dim lX8
    Dim lY8
    Dim lResult
 
    lX8 = lX And &H80000000
    lY8 = lY And &H80000000
    lX4 = lX And &H40000000
    lY4 = lY And &H40000000
 
    lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)
 
    If lX4 And lY4 Then
        lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
    ElseIf lX4 Or lY4 Then
        If lResult And &H40000000 Then
            lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
        Else
            lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
        End If
    Else
        lResult = lResult Xor lX8 Xor lY8
    End If
 
    AddUnsigned = lResult
End Function

Private Function Ch(x, y, z)
    Ch = ((x And y) Xor ((Not x) And z))
End Function

Private Function Maj(x, y, z)
    Maj = ((x And y) Xor (x And z) Xor (y And z))
End Function

Private Function S(x, n)
    S = (RShift(x, (n And m_lOnBits2(4))) Or LShift(x, (32 - (n And m_lOnBits2(4)))))
End Function

Private Function R(x, n)
    R = RShift(x, CInt(n And m_lOnBits2(4)))
End Function

Private Function Sigma0(x)
    Sigma0 = (S(x, 2) Xor S(x, 13) Xor S(x, 22))
End Function

Private Function Sigma1(x)
    Sigma1 = (S(x, 6) Xor S(x, 11) Xor S(x, 25))
End Function

Private Function Gamma0(x)
    Gamma0 = (S(x, 7) Xor S(x, 18) Xor R(x, 3))
End Function

Private Function Gamma1(x)
    Gamma1 = (S(x, 17) Xor S(x, 19) Xor R(x, 10))
End Function

Private Function ConvertToWordArray(sMessage)
    Dim lMessageLength
    Dim lNumberOfWords
    Dim lWordArray()
    Dim lBytePosition
    Dim lByteCount
    Dim lWordCount
    Dim lByte
    
    Const MODULUS_BITS = 512
    Const CONGRUENT_BITS = 448
    
    lMessageLength = Len(sMessage)
    
    lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE2)) \ (MODULUS_BITS \ BITS_TO_A_BYTE2)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD2)
    ReDim lWordArray(lNumberOfWords - 1)
    
    lBytePosition = 0
    lByteCount = 0
    Do Until lByteCount >= lMessageLength
        lWordCount = lByteCount \ BYTES_TO_A_WORD2
        
        lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD2)) * BITS_TO_A_BYTE2
        
        lByte = AscB(Mid(sMessage, lByteCount + 1, 1))
        
        lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(lByte, lBytePosition)
        lByteCount = lByteCount + 1
    Loop

    lWordCount = lByteCount \ BYTES_TO_A_WORD2
    lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD2)) * BITS_TO_A_BYTE2

    lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)

    lWordArray(lNumberOfWords - 1) = LShift(lMessageLength, 3)
    lWordArray(lNumberOfWords - 2) = RShift(lMessageLength, 29)
    
    ConvertToWordArray = lWordArray
End Function

Public Function SHA256(sMessage)
    Dim HASH(7)
    Dim M
    Dim W(63)
    Dim a
    Dim b
    Dim c
    Dim d
    Dim e
    Dim f
    Dim g
    Dim h
    Dim i
    Dim j
    Dim T1
    Dim T2
    
    HASH(0) = &H6A09E667
    HASH(1) = &HBB67AE85
    HASH(2) = &H3C6EF372
    HASH(3) = &HA54FF53A
    HASH(4) = &H510E527F
    HASH(5) = &H9B05688C
    HASH(6) = &H1F83D9AB
    HASH(7) = &H5BE0CD19
    
    M = ConvertToWordArray(sMessage)
    
    For i = 0 To UBound(M) Step 16
        a = HASH(0)
        b = HASH(1)
        c = HASH(2)
        d = HASH(3)
        e = HASH(4)
        f = HASH(5)
        g = HASH(6)
        h = HASH(7)
        
        For j = 0 To 63
            If j < 16 Then
                W(j) = M(j + i)
            Else
                W(j) = AddUnsigned(AddUnsigned(AddUnsigned(Gamma1(W(j - 2)), W(j - 7)), Gamma0(W(j - 15))), W(j - 16))
            End If
                
            T1 = AddUnsigned(AddUnsigned(AddUnsigned(AddUnsigned(h, Sigma1(e)), Ch(e, f, g)), K2(j)), W(j))
            T2 = AddUnsigned(Sigma0(a), Maj(a, b, c))
            
            h = g
            g = f
            f = e
            e = AddUnsigned(d, T1)
            d = c
            c = b
            b = a
            a = AddUnsigned(T1, T2)
        Next
        
        HASH(0) = AddUnsigned(a, HASH(0))
        HASH(1) = AddUnsigned(b, HASH(1))
        HASH(2) = AddUnsigned(c, HASH(2))
        HASH(3) = AddUnsigned(d, HASH(3))
        HASH(4) = AddUnsigned(e, HASH(4))
        HASH(5) = AddUnsigned(f, HASH(5))
        HASH(6) = AddUnsigned(g, HASH(6))
        HASH(7) = AddUnsigned(h, HASH(7))
    Next
    
    SHA256 = LCase(Right("00000000" & Hex(HASH(0)), 8) & Right("00000000" & Hex(HASH(1)), 8) & Right("00000000" & Hex(HASH(2)), 8) & Right("00000000" & Hex(HASH(3)), 8) & Right("00000000" & Hex(HASH(4)), 8) & Right("00000000" & Hex(HASH(5)), 8) & Right("00000000" & Hex(HASH(6)), 8) & Right("00000000" & Hex(HASH(7)), 8))
End Function
%>
