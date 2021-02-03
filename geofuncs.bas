Attribute VB_Name = "geofuncs"
Option Explicit
Option Base 1

Public Function bearing_to_azimuth(s As String) As Single

    Dim ss As Variant, d As Single
    
    ss = Split(s, " ")
    
    If IsNumeric(ss(1)) And IsNumeric(ss(2)) Then
        d = CSng(ss(1)) + CSng(ss(2)) / 60
        
        Select Case (ss(0) & ss(3))
            Case "NE": d = d
            Case "NW": d = 360 - d
            Case "SE": d = 180 - d
            Case "SW": d = 180 + d
            Case Else: MsgBox ("Error")
        End Select
    Else
        MsgBox ("Error")
    End If

    bearing_to_azimuth = d
    
End Function


Public Function azimuth_to_bearing(d As Single) As String
    Dim a As Single, b As Single
    Dim ans As String
    
    If d < 90 Then
        a = Fix(d)
        b = Round((d - a) * 60, 0)
        a = a Mod 360
        ans = "N " & CStr(a) & " " & CStr(b) & " E"
    ElseIf d < 180 Then
        d = 180 - d
        a = Fix(d)
        b = Round((d - a) * 60, 0)
        a = a Mod 360
        ans = "S " & CStr(a) & " " & CStr(b) & " E"
    ElseIf d < 270 Then
        d = d - 180
        a = Fix(d)
        b = Round((d - a) * 60, 0)
        a = a Mod 360
        ans = "S " & CStr(a) & " " & CStr(b) & " W"
    Else
        d = 360 - d
        a = Fix(d)
        b = Round((d - a) * 60, 0)
        a = a Mod 360
        ans = "N " & CStr(a) & " " & CStr(b) & " W"
    End If
    
    azimuth_to_bearing = ans
    
End Function

Public Function pi() As Single

  pi = 4 * Atn(1)
  
End Function

Public Function radians_to_azimuth(r As Single) As Single
    
    Dim a As Single
    a = (450 - r * 180 / pi()) * 100000
    a = (a Mod 36000000) / 100000
    radians_to_azimuth = a
    
End Function

Public Function radians_to_bearing(r As Single) As String

    radians_to_bearing = bearing_from_azimuth(radians_to_azimuth(r))

End Function

Public Function azimuth_to_radians(d As Single) As Single
  
    azimuth_to_radians = (450 - d) * pi() / 180

End Function

Public Function bearing_to_radians(s As String) As Single
  
  bearing_to_radians = azimuth_to_radians(bearing_to_azimuth(s))

End Function

Public Function distance(a As Single, b As Single) As Single
  
  distance = (a * a + b * b) ^ (1 / 2)

End Function

Public Function tek_to_dxdy(bearing As String, distance As Single) As Variant
    
    Dim r: r = bearing_to_radians(bearing)
    Dim dx: dx = distance * Cos(r)
    Dim dy: dy = distance * Sin(r)
    tek_to_dxdy = Array(dx, dy)

End Function

Public Function dtor() As Single
    
    dtor = pi() / 180

End Function

Public Function rtod() As Single
    
    rtod = 180 / pi()

End Function
