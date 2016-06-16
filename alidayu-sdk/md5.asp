<%
' http://www.yiit.cn/plugin/asp-hmac-md5-function.html
' asp hmac md5加密，支持中文，详细说明请查阅以上地址
' 调用方法：
'		HmacMd5(text,key)－加密内容支持中文，但key最好用非中文的。
'		md5(text)-简版，ASP_MD5(text)－标准版

Dim m_lOnBits(30)
Dim m_l2Power(30)
a = 0
For i = 0 To 30
a = a * 2 + 1
m_lOnBits(i) = a
m_l2Power(i) = 2 ^ (i)
Next

Function HmacMd5(txt,key)
Dim hkey
Dim ipad(63)
Dim opad(63)
Dim odata(79)
Dim idata()
text = UTF8bin(txt)
ReDim idata(64 + LenB(text) - 1)

If Len(key)>64 Then
hkey = ASP_MD5(key)
Else
hkey = key
End If

For x=0 To 63
idata(x) = &h36
odata(x) = &h5C
ipad(x) = &h36
opad(x) = &h5C
Next

For x=0 To Len(hkey)-1
ipad(x) = ipad(x) Xor Asc(CStr(Mid(hkey,x+1,1)))
opad(x) = opad(x) Xor Asc(CStr(Mid(hkey,x+1,1)))
idata(x) = ipad(x) And &hFF
odata(x) = opad(x) And &hFF
Next

For x = 1 To LenB(text)
idata(63 + x) = AscB(MidB(text, x, 1))
Next
ihas = binl2hex(coreMD5(ContArray(idata)))

For x=0 To 15
odata(64+x) = "&H" & Mid(ihas,x*2 + 1,2)
Next
HmacMd5 = LCase(binl2hex(coreMD5(ContArray(odata))))
End Function

Function UTF8bin(ByVal szInput)
inputLen = Len(szInput)
For x = 1 To inputLen
wch = Mid(szInput, x, 1)
nAsc = AscW(wch)
If nAsc < 0 Then nAsc = nAsc + 65536
If nAsc < 128 Then
szRet = szRet & ChrB(nAsc)
ElseIf nAsc < 4096 Then
uch = ChrB(((nAsc \ 2 ^ 6)) Or &HC0) & ChrB(nAsc And &H3F Or &H80)
szRet = szRet & uch
Else
uch = ChrB((nAsc \ 2 ^ 12) Or &HE0) & _
ChrB((nAsc \ 2 ^ 6) And &H3F Or &H80) & _
ChrB(nAsc And &H3F Or &H80)
szRet = szRet & uch
End If
Next
UTF8bin = szRet
End Function

Private Function str2bin(varstr) 
str2bin="" 
For i=1 To Len(varstr) 
varchar=mid(varstr,i,1) 
varasc = Asc(varchar) 
If Abs(varasc)>255 Then 
varlow = Left(Hex(varasc),2) 
varhigh = right(Hex(varasc),2) 
str2bin = str2bin & chrB("&H" & varlow) & chrB("&H" & varhigh) 
Else 
str2bin = str2bin & chrB(AscB(varchar)) 
End If 
Next 
End Function 

Function MD5(text)
Dim idata()
ReDim idata(Len(text) - 1)
For i = 1 To Len(text)
idata(i - 1) = Asc(Mid(text, i, 1))
Next
MD5 = LCase(binl2hex(coreMD5(ContArray(idata))))
End Function

Function ASP_MD5(sMessage)
Dim text, idata()
text = str2bin(sMessage)
ReDim idata(LenB(text) - 1)
For i = 1 To LenB(text)
idata(i - 1) = AscB(MidB(text, i, 1))
Next
ASP_MD5 = LCase(binl2hex(coreMD5(ContArray(idata))))
End Function

Function coreMD5(x)
Dim k
Dim AA
Dim BB
Dim CC
Dim DD
Dim a
Dim b
Dim c
Dim d

Const S11 = 7
Const S12 = 12
Const S13 = 17
Const S14 = 22
Const S21 = 5
Const S22 = 9
Const S23 = 14
Const S24 = 20
Const S31 = 4
Const S32 = 11
Const S33 = 16
Const S34 = 23
Const S41 = 6
Const S42 = 10
Const S43 = 15
Const S44 = 21

a = &H67452301
b = &HEFCDAB89
c = &H98BADCFE
d = &H10325476

For k = 0 To UBound(x)-1 Step 16
AA = a
BB = b
CC = c
DD = d

MD5_FF a, b, c, d, x(k + 0), S11, &HD76AA478
MD5_FF d, a, b, c, x(k + 1), S12, &HE8C7B756
MD5_FF c, d, a, b, x(k + 2), S13, &H242070DB
MD5_FF b, c, d, a, x(k + 3), S14, &HC1BDCEEE
MD5_FF a, b, c, d, x(k + 4), S11, &HF57C0FAF
MD5_FF d, a, b, c, x(k + 5), S12, &H4787C62A
MD5_FF c, d, a, b, x(k + 6), S13, &HA8304613
MD5_FF b, c, d, a, x(k + 7), S14, &HFD469501
MD5_FF a, b, c, d, x(k + 8), S11, &H698098D8
MD5_FF d, a, b, c, x(k + 9), S12, &H8B44F7AF
MD5_FF c, d, a, b, x(k + 10), S13, &HFFFF5BB1
MD5_FF b, c, d, a, x(k + 11), S14, &H895CD7BE
MD5_FF a, b, c, d, x(k + 12), S11, &H6B901122
MD5_FF d, a, b, c, x(k + 13), S12, &HFD987193
MD5_FF c, d, a, b, x(k + 14), S13, &HA679438E
MD5_FF b, c, d, a, x(k + 15), S14, &H49B40821

MD5_GG a, b, c, d, x(k + 1), S21, &HF61E2562
MD5_GG d, a, b, c, x(k + 6), S22, &HC040B340
MD5_GG c, d, a, b, x(k + 11), S23, &H265E5A51
MD5_GG b, c, d, a, x(k + 0), S24, &HE9B6C7AA
MD5_GG a, b, c, d, x(k + 5), S21, &HD62F105D
MD5_GG d, a, b, c, x(k + 10), S22, &H2441453
MD5_GG c, d, a, b, x(k + 15), S23, &HD8A1E681
MD5_GG b, c, d, a, x(k + 4), S24, &HE7D3FBC8
MD5_GG a, b, c, d, x(k + 9), S21, &H21E1CDE6
MD5_GG d, a, b, c, x(k + 14), S22, &HC33707D6
MD5_GG c, d, a, b, x(k + 3), S23, &HF4D50D87
MD5_GG b, c, d, a, x(k + 8), S24, &H455A14ED
MD5_GG a, b, c, d, x(k + 13), S21, &HA9E3E905
MD5_GG d, a, b, c, x(k + 2), S22, &HFCEFA3F8
MD5_GG c, d, a, b, x(k + 7), S23, &H676F02D9
MD5_GG b, c, d, a, x(k + 12), S24, &H8D2A4C8A
  
MD5_HH a, b, c, d, x(k + 5), S31, &HFFFA3942
MD5_HH d, a, b, c, x(k + 8), S32, &H8771F681
MD5_HH c, d, a, b, x(k + 11), S33, &H6D9D6122
MD5_HH b, c, d, a, x(k + 14), S34, &HFDE5380C
MD5_HH a, b, c, d, x(k + 1), S31, &HA4BEEA44
MD5_HH d, a, b, c, x(k + 4), S32, &H4BDECFA9
MD5_HH c, d, a, b, x(k + 7), S33, &HF6BB4B60
MD5_HH b, c, d, a, x(k + 10), S34, &HBEBFBC70
MD5_HH a, b, c, d, x(k + 13), S31, &H289B7EC6
MD5_HH d, a, b, c, x(k + 0), S32, &HEAA127FA
MD5_HH c, d, a, b, x(k + 3), S33, &HD4EF3085
MD5_HH b, c, d, a, x(k + 6), S34, &H4881D05
MD5_HH a, b, c, d, x(k + 9), S31, &HD9D4D039
MD5_HH d, a, b, c, x(k + 12), S32, &HE6DB99E5
MD5_HH c, d, a, b, x(k + 15), S33, &H1FA27CF8
MD5_HH b, c, d, a, x(k + 2), S34, &HC4AC5665

MD5_II a, b, c, d, x(k + 0), S41, &HF4292244
MD5_II d, a, b, c, x(k + 7), S42, &H432AFF97
MD5_II c, d, a, b, x(k + 14), S43, &HAB9423A7
MD5_II b, c, d, a, x(k + 5), S44, &HFC93A039
MD5_II a, b, c, d, x(k + 12), S41, &H655B59C3
MD5_II d, a, b, c, x(k + 3), S42, &H8F0CCC92
MD5_II c, d, a, b, x(k + 10), S43, &HFFEFF47D
MD5_II b, c, d, a, x(k + 1), S44, &H85845DD1
MD5_II a, b, c, d, x(k + 8), S41, &H6FA87E4F
MD5_II d, a, b, c, x(k + 15), S42, &HFE2CE6E0
MD5_II c, d, a, b, x(k + 6), S43, &HA3014314
MD5_II b, c, d, a, x(k + 13), S44, &H4E0811A1
MD5_II a, b, c, d, x(k + 4), S41, &HF7537E82
MD5_II d, a, b, c, x(k + 11), S42, &HBD3AF235
MD5_II c, d, a, b, x(k + 2), S43, &H2AD7D2BB
MD5_II b, c, d, a, x(k + 9), S44, &HEB86D391

a = AddUnsigned(a, AA)
b = AddUnsigned(b, BB)
c = AddUnsigned(c, CC)
d = AddUnsigned(d, DD)
Next

coreMD5 = Array(a,b,c,d)
End Function

'
' screwball MD5 functions
'
Sub MD5_FF(a, b, c, d, x, s, ac)
a = AddUnsigned(a, AddUnsigned(AddUnsigned(MD5_F(b, c, d), x), ac))
a = RotateLeft(a, s)
a = AddUnsigned(a, b)
End Sub

Sub MD5_GG(a, b, c, d, x, s, ac)
a = AddUnsigned(a, AddUnsigned(AddUnsigned(MD5_G(b, c, d), x), ac))
a = RotateLeft(a, s)
a = AddUnsigned(a, b)
End Sub

Sub MD5_HH(a, b, c, d, x, s, ac)
a = AddUnsigned(a, AddUnsigned(AddUnsigned(MD5_H(b, c, d), x), ac))
a = RotateLeft(a, s)
a = AddUnsigned(a, b)
End Sub

Sub MD5_II(a, b, c, d, x, s, ac)
a = AddUnsigned(a, AddUnsigned(AddUnsigned(MD5_I(b, c, d), x), ac))
a = RotateLeft(a, s)
a = AddUnsigned(a, b)
End Sub

Function MD5_F(x, y, z)
MD5_F = (x And y) Or ((Not x) And z)
End Function

Function MD5_G(x, y, z)
MD5_G = (x And z) Or (y And (Not z))
End Function

Function MD5_H(x, y, z)
MD5_H = (x Xor y Xor z)
End Function

Function MD5_I(x, y, z)
MD5_I = (y Xor (x Or (Not z)))
End Function

'
' utility functions
'
Function LShift(lValue, iShiftBits)
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

If (lValue And m_l2Power(31 - iShiftBits)) Then
LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
Else
LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
End If
End Function

Function RShift(lValue, iShiftBits)
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

RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)

If (lValue And &H80000000) Then
RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
End If
End Function

Function RotateLeft(lValue, iShiftBits)
RotateLeft = LShift(lValue, iShiftBits) Or RShift(lValue, (32 - iShiftBits))
End Function

Function AddUnsigned(lX, lY)
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

Function WordToHex(lValue)
For l = 0 To 3
lByte = RShift(lValue, l * 8) And &HFF
WordToHex = WordToHex & Right("0" & Hex(lByte), 2)
Next
End Function

Function ContArray(barray)
Dim nblk
nblk = ((UBound(barray) + 8) \ 64 + 1) * 16 - 1

Dim blks()
ReDim blks(nblk)

For x = 0 To UBound(barray)
blks(x \ 4) = blks(x \ 4) Or LShift(barray(x), (x Mod 4) * 8)
Next

blks(x \ 4) = blks(x \ 4) Or LShift(&H80, ((x Mod 4) * 8))
blks(nblk - 1) = (UBound(barray) + 1) * 8
ContArray = blks
End Function

Function binl2hex(r)
binl2hex = WordToHex(r(0)) & WordToHex(r(1)) & WordToHex(r(2)) & WordToHex(r(3))
End Function
%>