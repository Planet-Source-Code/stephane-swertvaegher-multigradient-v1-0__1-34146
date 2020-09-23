Attribute VB_Name = "MGMod1"
Public xx%, yy%, C1%, C2%, C3%, qq%, Temp$
Public MGrad&(9), MGPos!(9), MGPct!(9), Mgpct2&(9)
Public MGrad1&(9), MGPct1!(9), Mgpct3&(9)
Public MGFile$, ff%, LoadFile As Boolean, TempCol&, Start As Boolean, KillFile As Boolean
Public R1&, G1&, B1&, MGTitle$
Public Enum T3dFill
T3dF0
T3dF1
End Enum

Public Enum Borderstyle
T3dRaiseRaise
T3dRaiseInset
T3dInsetRaise
T3dInsetInset
T3dNone
End Enum
'API for translating system colors to 'normal' colors
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long
   
Public Function T3D(Obj0 As Object, Obj As Object, Bev%, Optional Style3D As Borderstyle, Optional T3dFilled As T3dFill)
Dim R%, G%, B%, R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%
Dim FC&, T3Dxx%, SM%
On Error Resume Next

'global things
SM = Obj0.ScaleMode 'save scalemode
Obj0.ScaleMode = 3 'pixel
Obj.Borderstyle = 0 'no border
If IsMissing(Style3D) Then Style3D = 0
If Style3D > 4 Then Style3D = 3

'get formcolor
FC = Obj0.BackColor
'in case formcolor = systemcolor --> call the function RealColor
FC = RealColor(FC)
' convert to RGB
R = FC And &HFF
G = Int((FC And &HFF00&) / 256)
B = Int((FC And &HFF0000) / 65536)
'-------------------
If Style3D = 0 Then 'RaiseRaise
    R1 = R + 64
    If R1 > 255 Then R1 = 255
    R2 = R - 64
    If R2 < 0 Then R2 = 0
    R3 = R1
    R4 = R2
    G1 = G + 64
    If G1 > 255 Then G1 = 255
    G2 = G - 64
    If G2 < 0 Then G2 = 0
    G3 = G1
    G4 = G2
    B1 = B + 64
    If B1 > 255 Then B1 = 255
    B2 = B - 64
    If B2 < 0 Then B2 = 0
    B3 = B1
    B4 = B2
End If
'-------------------
If Style3D = 1 Then 'RaiseInset
    R1 = R + 64
    If R1 > 255 Then R1 = 255
    R2 = R - 64
    If R2 < 0 Then R2 = 0
    R4 = R1
    R3 = R2
    G1 = G + 64
    If G1 > 255 Then G1 = 255
    G2 = G - 64
    If G2 < 0 Then G2 = 0
    G4 = G1
    G3 = G2
    B1 = B + 64
    If B1 > 255 Then B1 = 255
    B2 = B - 64
    If B2 < 0 Then B2 = 0
    B4 = B1
    B3 = B2
End If
If Style3D = 2 Then 'InsetRaise
    R2 = R + 64
    If R2 > 255 Then R2 = 255
    R1 = R - 64
    If R1 < 0 Then R1 = 0
    R4 = R1
    R3 = R2
    G2 = G + 64
    If G2 > 255 Then G2 = 255
    G1 = G - 64
    If G1 < 0 Then G1 = 0
    G4 = G1
    G3 = G2
    B2 = B + 64
    If B2 > 255 Then B2 = 255
    B1 = B - 64
    If B1 < 0 Then B1 = 0
    B4 = B1
    B3 = B2
End If
If Style3D = 3 Then 'InsetInset
    R2 = R + 64
    If R2 > 255 Then R2 = 255
    R1 = R - 64
    If R1 < 0 Then R1 = 0
    R3 = R1
    R4 = R2
    G2 = G + 64
    If G2 > 255 Then G2 = 255
    G1 = G - 64
    If G1 < 0 Then G1 = 0
    G3 = G1
    G4 = G2
    B2 = B + 64
    If B2 > 255 Then B2 = 255
    B1 = B - 64
    If B1 < 0 Then B1 = 0
    B3 = B1
    B4 = B2
End If
If Style3D = 4 Then 'No Border
R1 = R: R2 = R: R3 = R: R4 = R
G1 = G: G2 = G: G3 = G: G4 = G
B1 = B: B2 = B: B3 = B: B4 = B
End If
Bev = Bev + 1
T3Dxx = Bev 'just in case Filled = 1

'Outer
If IsMissing(T3dFilled) Or T3dFilled = 0 Then
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left - Bev, Obj.Top + Obj.Height + Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top - Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left + Obj.Width + Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top + Obj.Height + Bev), RGB(R2, G2, B2)
    Obj0.Line (Obj.Left - Bev, Obj.Top + Obj.Height + Bev)-(Obj.Left + Obj.Width + Bev + 1, Obj.Top + Obj.Height + Bev), RGB(R2, G2, B2)
Else
For Bev = T3Dxx To 1 Step -1 'in case T3DF1 (filled)
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left - Bev, Obj.Top + Obj.Height + Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top - Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left + Obj.Width + Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top + Obj.Height + Bev), RGB(R2, G2, B2)
    Obj0.Line (Obj.Left - Bev, Obj.Top + Obj.Height + Bev)-(Obj.Left + Obj.Width + Bev + 1, Obj.Top + Obj.Height + Bev), RGB(R2, G2, B2)
Next Bev
End If
'Inner
    Obj0.Line (Obj.Left - 1, Obj.Top - 1)-(Obj.Left - 1, Obj.Top + Obj.Height + 1), RGB(R3, G3, B3)
    Obj0.Line (Obj.Left - 1, Obj.Top - 1)-(Obj.Left + Obj.Width + 1, Obj.Top - 1), RGB(R3, G3, B3)
    Obj0.Line (Obj.Left + Obj.Width + 1, Obj.Top - 1)-(Obj.Left + Obj.Width + 1, Obj.Top + Obj.Height + 1), RGB(R4, G4, B4)
    Obj0.Line (Obj.Left - 1, Obj.Top + Obj.Height + 1)-(Obj.Left + Obj.Width + 2, Obj.Top + Obj.Height + 1), RGB(R4, G4, B4)

Obj0.ScaleMode = SM 'restore original scalemode
End Function
  
  ' if System Color then translate to 'normal color'
  ' else, do nothing
  Public Function RealColor(ByVal Color As OLE_COLOR) As Long
     Dim Col As Long
     Col = TranslateColor(Color, 0, RealColor)
  End Function


Public Sub NewGrad()
With Form1
LoadFile = True
For xx = 0 To 9
.Check1(xx).Value = 1
Next xx
.Label6 = "Unknown"
MGrad(0) = 0
MGrad(1) = &HFFFF00
MGrad(2) = &H80FF00
MGrad(3) = &HFF8000
MGrad(4) = &HFFFFFF
MGrad(5) = &HA0A0A0
MGrad(6) = &H40FF40
MGrad(7) = &HFF00FF
MGrad(8) = &HFF&
MGrad(9) = &HFFFF&
For xx = 0 To 9
.Sli(xx).Visible = False
.Sli(xx).Left = (xx * 55) - 2
MGPos(xx) = (xx * 55) - 2
MGPct(xx) = (MGPos(xx) + 2) / (.Pic1.ScaleWidth)
.Label4(xx) = Format(MGPct(xx), "0.0000")
.Sli(xx).BackColor = MGrad(xx)
.Label3(xx).BackColor = MGrad(xx)
Next xx
.Sli(9).Left = 497
MGPos(9) = 497
MGPct(9) = (MGPos(9) + 2) / (.Pic1.ScaleWidth)
.Label4(9) = Format(MGPct(9), "0.0000")
MGPct(1) = -1
MGPct(2) = -1
MGPct(4) = -1
MGPct(5) = -1
MGPct(7) = -1
MGPct(8) = -1
For xx = 0 To 9
If MGPct(xx) = -1 Then
.Check1(xx).Value = 0
.Sli(xx).Visible = False
.Label4(xx) = "***"
MGPct(xx) = -1
Else
.Sli(xx).Visible = True
End If
Next xx
.Image1.Top = .Label2(0).Top + 1
End With
LoadFile = False
End Sub

Public Sub MultiGrad2(Ob As Object)
Dim Mgx%, Mgy%
Dim Mgf%, Mgl%, Mgr1!, Mgg1!, Mgb1!, Mgr2&, Mgg2&, Mgb2&
Dim Mgh1!, Mgh2!, Mgh%, Mgw%
Dim Mgsr!, Mgsg!, Mgsb!

Mgw = Ob.ScaleWidth
For Mgy = 0 To 9
Mgf = GetFirst(Mgy)
Mgl = GetLast(Mgf + 1)
Mgr1 = MGrad(Mgf) Mod 256&
Mgg1 = ((MGrad(Mgf) And &HFF00) / 256&) Mod 256&
Mgb1 = (MGrad(Mgf) And &HFF0000) / 65536
Mgr2 = MGrad(Mgl) Mod 256&
Mgg2 = ((MGrad(Mgl) And &HFF00) / 256&) Mod 256&
Mgb2 = (MGrad(Mgl) And &HFF0000) / 65536
Mgh1 = Int(MGPct(Mgf) * Mgw)
Mgh2 = Int(MGPct(Mgl) * Mgw)
Mgh = Mgh2 - Mgh1
If Mgh < 1 Then Mgh = 1
Mgsr = (Mgr2 - Mgr1) / Mgh
Mgsg = (Mgg2 - Mgg1) / Mgh
Mgsb = (Mgb2 - Mgb1) / Mgh
For Mgx = Mgh1 To Mgh2
Ob.Line (Mgx, 0)-(Mgx, Ob.Height), RGB(Mgr1, Mgg1, Mgb1)
Mgr1 = Mgr1 + Mgsr
Mgg1 = Mgg1 + Mgsg
Mgb1 = Mgb1 + Mgsb
Next Mgx
If Mgl = 9 Then Exit Sub
Next Mgy
End Sub

Public Sub MultiGrad3(Ob As Object)
Dim Mgx%, Mgy%
Dim Mgf%, Mgl%, Mgr1!, Mgg1!, Mgb1!, Mgr2&, Mgg2&, Mgb2&
Dim Mgh1!, Mgh2!, Mgh%, Mgw%
Dim Mgsr!, Mgsg!, Mgsb!

Mgw = Ob.ScaleWidth
For Mgy = 0 To 9
Mgf = GetFirst1(Mgy)
Mgl = GetLast1(Mgf + 1)
Mgr1 = MGrad1(Mgf) Mod 256&
Mgg1 = ((MGrad1(Mgf) And &HFF00) / 256&) Mod 256&
Mgb1 = (MGrad1(Mgf) And &HFF0000) / 65536
Mgr2 = MGrad1(Mgl) Mod 256&
Mgg2 = ((MGrad1(Mgl) And &HFF00) / 256&) Mod 256&
Mgb2 = (MGrad1(Mgl) And &HFF0000) / 65536
Mgh1 = Int(MGPct1(Mgf) * Mgw)
Mgh2 = Int(MGPct1(Mgl) * Mgw)
Mgh = Mgh2 - Mgh1
If Mgh < 1 Then Mgh = 1
Mgsr = (Mgr2 - Mgr1) / Mgh
Mgsg = (Mgg2 - Mgg1) / Mgh
Mgsb = (Mgb2 - Mgb1) / Mgh
For Mgx = Mgh1 To Mgh2
Ob.Line (Mgx, 0)-(Mgx, Ob.Height), RGB(Mgr1, Mgg1, Mgb1)
Mgr1 = Mgr1 + Mgsr
Mgg1 = Mgg1 + Mgsg
Mgb1 = Mgb1 + Mgsb
Next Mgx
If Mgl = 9 Then Exit Sub
Next Mgy
End Sub

Public Function GetFirst1(GFbegin%)
Dim Gf%
For Gf = GFbegin To 9
If MGPct1(Gf) <> -1 Then
GetFirst1 = Gf
Exit Function
End If
Next Gf
End Function

Public Function GetLast1(GFbegin)
If GFbegin = 9 Then
GetLast1 = 9
Exit Function
End If
Dim Gl%
For Gl = GFbegin To 9
If MGPct1(Gl) <> -1 Then
GetLast1 = Gl
Exit Function
End If
Next Gl
End Function


