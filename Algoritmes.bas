Attribute VB_Name = "Module1"
Option Explicit

'***********************************************
' omzetten van decimaal naar hexadecimaal
'***********************************************
Public Function Dec2Hex(n As Integer) As String
Dim s As String
Dim getal As Integer, rest As Integer

getal = n
Do
rest = getal Mod 16
Select Case rest
    Case 0 To 9: s = Right(Str(rest), 1) & s
    Case 10 To 15: s = Chr(Asc("A") + rest - 10) + s
    Case Else: Debug.Print "else"
End Select

getal = (getal - rest) / 16

Loop Until getal = 0
If Len(s) < 2 Then
    s = "0" & s
End If
Dec2Hex = s
End Function

Private Function Minimum(rR As Double, rG As Double, rB As Double) As Single
  If (rR < rG) Then
  If (rR < rB) Then
    Minimum = rR
  Else
    Minimum = rB
  End If
  Else
  If (rB < rG) Then
    Minimum = rB
  Else
    Minimum = rG
  End If
End If
End Function

Private Function Maximum(rR As Double, rG As Double, rB As Double) As Single
  If (rR > rG) Then
    If (rR > rB) Then
        Maximum = rR
    Else
        Maximum = rB
    End If
  Else
    If (rB > rG) Then
        Maximum = rB
    Else
        Maximum = rG
    End If
  End If
End Function

'********************************
' RGB to HSV conversie
'********************************
'// RGB, each 0 to 255, to HSV.
'// H = 0.0 to 360.0 (corresponding to 0..360.0 degrees around hexcone)
'// S = 0.0 (shade of gray) to 1.0 (pure color)
'// V = 0.0 (black) to 1.0 {white)

'// Based on C Code in "Computer Graphics -- Principles and Practice,"
'// Foley et al, 1996, p. 592.

Public Sub RGB2HSV(ByVal r As Double, ByVal g As Double, ByVal b As Double, _
                          h As Double, s As Double, v As Double)

Dim delta As Double
Dim min As Double

min = Minimum(r, g, b)
v = Maximum(r, g, b)

delta = v - min

' // Calculate saturation: saturation is 0 if r, g and b are all 0
If v = 0# Then
    s = 0: h = 0
 Else
    s = delta / v
End If
If s = 0# Then
    h = 0 'NaN   // Achromatic: When s = 0, h is undefined
 Else     '      // Chromatic
   If r = v Then
      'between yellow and magenta [degrees]
       h = 60# * (g - b) / delta
   ElseIf g = v Then
       ' between cyan and yellow
       h = 120# + 60# * (b - r) / delta
   Else
          ' between magenta and cyan
         h = 240# + 60# * (r - g) / delta
   End If

End If
If h < 0# Then
        h = h + 360#
End If
v = v / 255
End Sub '{RGB2HSV};

'********************************
' HSV to RGB conversie
'********************************
'// Based on C Code in "Computer Graphics -- Principles and Practice,"
'// Foley et al, 1996, p. 593.
'//
'// H = 0.0 to 360.0 (corresponding to 0..360 degrees around hexcone)
'// NaN (undefined) for S = 0
'// S = 0.0 (shade of gray) to 1.0 (pure color)
'// V = 0.0 (black) to 1.0 (white)

Public Sub HSV2RGB(ByVal h As Double, ByVal s As Double, ByVal v As Double, _
     r As Double, g As Double, b As Double)

   
   Dim f As Double
   Dim i As Integer
   Dim hTemp As Double '// since H is CONST parameter
   Dim p As Double, q As Double, t As Double

 If s = 0# Then        '// color is on black-and-white center line
         r = v         '// achromatic: shades of gray
         g = v
         b = v
    
 Else            '// chromatic color
   If h >= 360# Then              '// 360 degrees same as 0 degrees
        hTemp = h - 360#
   Else
        hTemp = h
       hTemp = hTemp / 60    '// h is now IN [0,6)
       i = Int(hTemp)       '// largest integer <= h
       f = hTemp - i        '// fractional part of h

        p = v * (1# - s)
        q = v * (1# - (s * f))
        t = v * (1# - (s * (1# - f)))

   Select Case i
     Case 0:
            r = v: g = t: b = p
     Case 1:
            r = q: g = v: b = p
     Case 2:
            r = p: g = v: b = t
     Case 3:
            r = p: g = q: b = v
     Case 4:
            r = t: g = p: b = v
     Case 5:
            r = v: g = p: b = q
   End Select
   End If
End If
End Sub '{HSVtoRGB};


Public Sub RGBtoHSV(ByVal r As Double, ByVal g As Double, ByVal b As Double, _
                    h As Double, s As Double, v As Double)
Dim max As Double
Dim min As Double

max = Maximum(r, g, b)
min = Minimum(r, g, b)
'// V berechnen
v = max / 255#
'// S berechnen
If max <> 0 Then
    s = (max - min) / max
Else
    s = 0
End If

'// H berechnen
If s = 0 Then
    h = 0   '// keine Farbe!
Else
    Dim delta As Double
    delta = max - min
    If r = max Then
        h = 60 * (((g - b)) / delta) '// Farbe liegt zwischen Gelb und Magenta
    ElseIf g = max Then
        h = 60 * (((b - r)) / delta + 2) ' // Farbe liegt zwischen Cyan und Gelb
    ElseIf b = max Then
        h = 60 * (((r - g)) / delta + 4) ' // Farbe liegt zwischen Magenta und Cyan
    End If

    If (h < 0) Then '// H darf nich negativ sein
        h = h + 360
        h = h / 360 ' // H in den Bereich [0,1] bringen
    End If
End If
End Sub


