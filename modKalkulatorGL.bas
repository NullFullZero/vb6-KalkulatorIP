Attribute VB_Name = "modKalkulatorGL"

Function BinNaDz(LiczbaBinarna As String) As Long
  Dim n As Integer
     n = Len(LiczbaBinarna) - 1
     a = n
     Do While n > -1
        x = Mid(LiczbaBinarna, ((a + 1) - n), 1)
        BinNaDz = IIf((x = "1"), BinNaDz + (2 ^ (n)), BinNaDz)
        n = n - 1
     Loop
End Function
'Aliquet 1.0
'Autor: Adam Shaman Czana
'Nazwa: Kalkulator IP
Public Function DzNaBin(ByVal WartDz As Long, Optional ByVal NrBitu As Integer = 8) As String

  Dim i As Integer
  
  Do While WartDz > (2 ^ NrBitu) - 1
    NrBitu = NrBitu + 8
  Loop

  DzNaBin = vbNullString

  For i = 0 To (NrBitu - 1)
      DzNaBin = CStr((WartDz And 2 ^ i) / 2 ^ i) & DzNaBin
  Next i

End Function

Public Function aAND(ByVal x As Byte, ByVal y As Byte) As Byte
On Error Resume Next

aAND = 0

If (x = 1 And y = 1) Then
    aAND = 1
Else
    aAND = 0
End If


End Function

Public Function LiczJedynki(Maska As String)
On Error Resume Next
Dim IloscJedynek As Byte
Dim i As Byte
Dim x As Byte
Const Ile = 32

IloscJedynek = 0

For i = 1 To Ile + 9 'kropki
    x = Mid(Maska, i, 1)
    
    If x = 1 Then IloscJedynek = IloscJedynek + 1
    x = 0
Next i
LiczJedynki = IloscJedynek

End Function

Public Function LiczAdr(CIDR As Integer) As Long
'On Error Resume Next
Dim Ilosc As Long
Dim potega As Integer
'Call MsgBox(CIDR, vbOKOnly)

potega = 32 - CIDR
'Call MsgBox(potega, vbOKOnly)

Ilosc = 2 ^ potega '- 2
'Call MsgBox(Ilosc, vbOKOnly)
LiczAdr = Ilosc
End Function

Public Function MaskaZCIDR(CIDRS As Byte) As String
On Error Resume Next
Dim Maska As String
Dim Ile, i As Byte
Dim Dlugosc As Byte
Dim WskDL As Byte
Dim WlMaska As String

Ile = CIDRS
Maska = ""

For i = 1 To Ile

    Maska = Maska & "1"

Next i
Dlugosc = Len(Maska)
WskDL = 32 - Dlugosc


'Call MsgBox(Maska & " d:" & WskDL, vbOKOnly)

For i = 1 To WskDL

    Maska = Maska & "0"

Next i

'Call MsgBox(Maska, vbOKOnly)
MaskaZCIDR = Maska
End Function

Public Function Bin2Dec(Binarna As String) As Long
On Error Resume Next
  Dim lTemp As Long, i As Long
  
  lTemp = 0
  For i = Len(Binarna) To 1 Step -1
    If Mid(Binarna, i, 1) = "1" Then
      lTemp = lTemp + 2 ^ (Len(Binarna) - i)
    End If
  Next i
  Bin2Dec = lTemp
End Function

Public Function CalyAND(czIP, czMa As String) As String
On Error Resume Next
Dim dlIP, dlMaski As Byte
Dim i As Byte
Dim temp As String

dlIP = Len(czIP)
dlMaski = Len(czMa)
CalyAND = ""
temp = ""

For i = 1 To dlIP

    temp = temp & aAND(Mid(czIP, i, 1), Mid(czMa, i, 1))

Next i

CalyAND = temp

End Function

Public Function nNeg(x As Byte) As Byte
On Error Resume Next
If x = 1 Then
    nNeg = 0
ElseIf x = 0 Then
    nNeg = 1
End If
    
End Function

Public Function oOR(ByVal x As Byte, ByVal y As Byte) As Byte
On Error Resume Next
oOR = 0

If (x = 0 And y = 0) Then
    oOR = 0
ElseIf (x = 1 And y = 1) Then
    oOR = 1
ElseIf (x = 0 And y = 1) Then
    oOR = 1
ElseIf (x = 1 And y = 0) Then
    oOR = 1
End If

End Function

Public Function LiczAdrUzytkowe(xCIDR As Byte) As Long
On Error Resume Next
Dim Ilosc As Long
Dim potega As Byte
'Call MsgBox(CIDR, vbOKOnly)
potega = 32 - xCIDR

Ilosc = (2 ^ potega) - 2

LiczAdrUzytkowe = Ilosc

End Function

Public Function PierwszyAdres(AIPb As String, MSb As String, Ink As Boolean) As String
On Error Resume Next
Dim IP, Maska As String
Dim i As Byte
Dim Wynik As String
Dim WynikDZ(5) As String
Dim Czlon As Integer

IP = Replace(AIPb, ".", "")
IP = Replace(IP, " ", "")

Maska = Replace(MSb, ".", "")
Maska = Replace(Maska, " ", "")

'Call MsgBox(IP & " " & Maska)


For i = 1 To 32

    Wynik = Wynik & aAND(Mid(IP, i, 1), Mid(Maska, i, 1))
    

Next i

WynikDZ(1) = Bin2Dec(Mid(Wynik, 1, 8))
WynikDZ(2) = Bin2Dec(Mid(Wynik, 9, 8))
WynikDZ(3) = Bin2Dec(Mid(Wynik, 17, 8))
WynikDZ(4) = Bin2Dec(Mid(Wynik, 25, 8))

If Ink = True Then
    
    Czlon = Int(WynikDZ(4))
    Czlon = Czlon + 1
    WynikDZ(4) = Czlon

End If



'Call MsgBox(Wynik)
'Call MsgBox(WynikDZ(1))
'Call MsgBox(WynikDZ(2))
'Call MsgBox(WynikDZ(3))
'Call MsgBox(WynikDZ(4))

PierwszyAdres = WynikDZ(1) & "." & WynikDZ(2) & "." & WynikDZ(3) & "." & WynikDZ(4)
End Function
