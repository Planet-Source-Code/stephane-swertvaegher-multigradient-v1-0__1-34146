Attribute VB_Name = "MGMod2"
'This is the actual sub to make a multigradient
'on every objects who supports the line method

Public Function GetFirst(GFbegin%)
Dim Gf%
For Gf = GFbegin To 9
If MGPct(Gf) <> -1 Then
GetFirst = Gf
Exit Function
End If
Next Gf
End Function

Public Function GetLast(GFbegin)
If GFbegin = 9 Then
GetLast = 9
Exit Function
End If
Dim Gl%
For Gl = GFbegin To 9
If MGPct(Gl) <> -1 Then
GetLast = Gl
Exit Function
End If
Next Gl
End Function

Public Sub MultiGrad(Ob As Object, Style%, Rev As Boolean)
Dim Mgx%, Mgy%
Dim Mgf%, Mgl%, Mgr1!, Mgg1!, Mgb1!, Mgr2&, Mgg2&, Mgb2&
Dim Mgh1!, Mgh2!, Mgh%, Mgw%
Dim Mgsr!, Mgsg!, Mgsb!

    If Style = 0 Then Mgw = Ob.ScaleHeight
    If Style = 1 Then Mgw = Ob.ScaleWidth
    If Style = 2 Then Mgw = Ob.ScaleWidth + Ob.ScaleHeight
For Mgy = 0 To 9
Mgf = GetFirst(Mgy)
Mgl = GetLast(Mgf + 1)
'cut to R, G and B
Mgr1 = MGrad(Mgf) Mod 256&
Mgg1 = ((MGrad(Mgf) And &HFF00) / 256&) Mod 256&
Mgb1 = (MGrad(Mgf) And &HFF0000) / 65536
Mgr2 = MGrad(Mgl) Mod 256&
Mgg2 = ((MGrad(Mgl) And &HFF00) / 256&) Mod 256&
Mgb2 = (MGrad(Mgl) And &HFF0000) / 65536
Mgh1 = Int(MGPct(Mgf) * Mgw)
Mgh2 = Int(MGPct(Mgl) * Mgw)
'get distance
Mgh = Mgh2 - Mgh1
If Mgh < 1 Then Mgh = 1
Mgsr = (Mgr2 - Mgr1) / Mgh
Mgsg = (Mgg2 - Mgg1) / Mgh
Mgsb = (Mgb2 - Mgb1) / Mgh
'do gradient
    If Style = 0 Then 'hor
    For Mgx = Mgh1 To Mgh2
    If Rev = False Then
    Ob.Line (0, Mgx)-(Ob.Width, Mgx), RGB(Mgr1, Mgg1, Mgb1)
    Else
    Ob.Line (0, Ob.Height - Mgx)-(Ob.Width, Ob.Height - Mgx), RGB(Mgr1, Mgg1, Mgb1)
    End If
    Mgr1 = Mgr1 + Mgsr
    Mgg1 = Mgg1 + Mgsg
    Mgb1 = Mgb1 + Mgsb
    Next Mgx
    End If
    If Style = 1 Then 'vert
    For Mgx = Mgh1 To Mgh2
    If Rev = False Then
    Ob.Line (Mgx, 0)-(Mgx, Ob.Height), RGB(Mgr1, Mgg1, Mgb1)
    Else
    Ob.Line (Ob.ScaleWidth - Mgx, 0)-(Ob.ScaleWidth - Mgx, Ob.Height), RGB(Mgr1, Mgg1, Mgb1)
    End If
    Mgr1 = Mgr1 + Mgsr
    Mgg1 = Mgg1 + Mgsg
    Mgb1 = Mgb1 + Mgsb
    Next Mgx
    End If
    If Style = 2 Then '45°
    For Mgx = Mgh1 To Mgh2
    If Rev = False Then
    Ob.Line (Mgx - Ob.Height, 0)-(Mgx, Ob.Height), RGB(Mgr1, Mgg1, Mgb1)
    Else
    Ob.Line (Ob.Width - Mgx, 0)-(Ob.Width + Ob.Height - Mgx, Ob.Height), RGB(Mgr1, Mgg1, Mgb1)
    End If
    Mgr1 = Mgr1 + Mgsr
    Mgg1 = Mgg1 + Mgsg
    Mgb1 = Mgb1 + Mgsb
    Next Mgx
    End If
If Mgl = 9 Then Exit Sub
Next Mgy
End Sub
'---------------------- end of sub

