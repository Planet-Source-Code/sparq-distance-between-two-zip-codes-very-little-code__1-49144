Attribute VB_Name = "Module1"
Public vLat1, vLat2, vLong1, vLong2 As Double

Const pi = 3.14159265358979

Function distance(lat1, lon1, lat2, lon2, unit)
    Dim theta, dist
    
    theta = lon1 - lon2
    dist = Sin(deg2rad(lat1)) * Sin(deg2rad(lat2)) + Cos(deg2rad(lat1)) * Cos(deg2rad(lat2)) * Cos(deg2rad(theta))
    dist = acos(dist)
    dist = rad2deg(dist)
    distance = dist * 60 * 1.1515
    
    Select Case UCase(unit)
        Case "K"
            distance = distance * 1.609344
        Case "N"
            distance = distance * 0.8684
    End Select
End Function

Function acos(Rad)
    If Abs(Rad) <> 1 Then
        acos = pi / 2 - Atn(Rad / Sqr(1 - Rad * Rad))
    ElseIf Rad = -1 Then
        acos = pi
    End If
End Function

Function RoundNum(Num As Double) As Double
  Dim Modular As Single
    Modular = Num - Int(Num)
    If Modular = 0 Then GoTo Ending
    If Modular >= 0.5 Then
        Num = Int(Num) + 1
    Else
        Num = Int(Num)
    End If

Ending:
    RoundNum = Num
End Function

Function deg2rad(Deg)
    deg2rad = CDbl(Deg * pi / 180)
End Function

Function rad2deg(Rad)
    rad2deg = CDbl(Rad * 180 / pi)
End Function

Sub SetCoords(Zip As String, Index As Integer)
  Dim TempDB As Database
  Dim TempTable As Recordset
  
  Set TempDB = OpenDatabase(App.Path & "\gps.mdb")
  Set TempTable = TempDB.OpenRecordset("Select Lat, Lon from Zips where zip = '" & Zip & "'")
  
  If Index = 1 Then
    vLat1 = TempTable!lat
    vLong1 = TempTable!lon
  Else
    vLat2 = TempTable!lat
    vLong2 = TempTable!lon
  End If
End Sub
