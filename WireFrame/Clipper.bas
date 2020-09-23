Attribute VB_Name = "Clipper"
Option Explicit
Sub ClipLine(RX1!, RY1!, RX2!, RY2!, X1!, Y1!, X2!, Y2!, OutX1!, OutY1!, OutX2!, OutY2!)

 'Liang Barsky Line Clipping Algorithm (1984):
 '(Parametric clipping but special case for rectangular clipping regions)
 '
 'Note that this routine is *Very* optimized!
 'Realy, it is two modules that i have it transformed in only one procedure!!!!!

 Dim PX1!, PY1!, PX2!, PY2!
 Dim TX1!, TY1!, TX2!, TY2!
 Dim U1!, U2!, Dx!, Dy!, Temp!
 Dim P!, Q!, UU1!, UU2!, R!, CT As Byte

 If (RX1 > RX2) Then Temp = RX1: RX1 = RX2: RX2 = Temp
 If (RY1 > RY2) Then Temp = RY1: RY1 = RY2: RY2 = Temp

 PX1 = X1: PY1 = Y1: PX2 = X2: PY2 = Y2
 U1 = 0: U2 = 1: Dx = (PX2 - PX1)

 P = (-1 * Dx): Q = (PX1 - RX1): UU1 = U1: UU2 = U1: CT = 1
 If P < 0 Then
  R = Q / P: If R > U2 Then CT = 0 Else If R > U1 Then U1 = R
 Else
  If P > 0 Then
   R = Q / P: If R < U1 Then CT = 0 Else If R < U2 Then U2 = R
  ElseIf Q < 0 Then
   CT = 0
  End If
 End If

 If CT = 1 Then
  P = Dx: Q = (RX2 - PX1): UU1 = U1: UU2 = U2: CT = 1
  If P < 0 Then
   R = Q / P: If R > U2 Then CT = 0 Else If R > U1 Then U1 = R
  Else
   If P > 0 Then
    R = Q / P: If R < U1 Then CT = 0 Else If R < U2 Then U2 = R
   ElseIf Q < 0 Then
    CT = 0
   End If
  End If
  If CT = 1 Then
   Dy = (PY2 - PY1): P = (-1 * Dy): Q = (PY1 - RY1): UU1 = U1: UU2 = U2: CT = 1
   If P < 0 Then
    R = Q / P: If R > U2 Then CT = 0 Else If R > U1 Then U1 = R
   Else
    If P > 0 Then
     R = Q / P: If R < U1 Then CT = 0 Else If R < U2 Then U2 = R
    ElseIf Q < 0 Then
     CT = 0
    End If
   End If
   If CT = 1 Then
    P = Dy: Q = (RY2 - PY1): UU1 = U1: UU2 = U2: CT = 1
    If P < 0 Then
     R = Q / P: If R > U2 Then CT = 0 Else If R > U1 Then U1 = R
    Else
     If P > 0 Then
      R = Q / P: If R < U1 Then CT = 0 Else If R < U2 Then U2 = R
     ElseIf Q < 0 Then
      CT = 0
     End If
    End If
    If CT = 1 Then
     If U2 < 1 Then PX2 = PX1 + (U2 * Dx): PY2 = PY1 + (U2 * Dy)
     If U1 > 0 Then PX1 = PX1 + (U1 * Dx): PY1 = PY1 + (U1 * Dy)
     OutX1 = PX1: OutY1 = PY1: OutX2 = PX2: OutY2 = PY2
    End If
   End If
  End If
 End If

End Sub
Function Accept(RX1!, RY1!, RX2!, RY2!, X1!, Y1!, X2!, Y2!) As Boolean

 'Cohen-Sutherland Trivial Accept (with codes):

 Dim Code1(3) As Boolean
 Dim Code2(3) As Boolean
 Dim Temp As Single

 If (RX1 > RX2) Then Temp = RX1: RX1 = RX2: RX2 = Temp
 If (RY1 > RY2) Then Temp = RY1: RY1 = RY2: RY2 = Temp

 If X1 < RX1 Then Code1(0) = True Else Code1(0) = False
 If X1 > RX2 Then Code1(1) = True Else Code1(1) = False
 If Y1 < RY1 Then Code1(2) = True Else Code1(2) = False
 If Y1 > RY2 Then Code1(3) = True Else Code1(3) = False

 If X2 < RX1 Then Code2(0) = True Else Code2(0) = False
 If X2 > RX2 Then Code2(1) = True Else Code2(1) = False
 If Y2 < RY1 Then Code2(2) = True Else Code2(2) = False
 If Y2 > RY2 Then Code2(3) = True Else Code2(3) = False

 If (Code1(0) Or Code2(0)) Then Accept = True
 If (Code1(1) Or Code2(1)) Then Accept = True
 If (Code1(2) Or Code2(2)) Then Accept = True
 If (Code1(3) Or Code2(3)) Then Accept = True

End Function
Function Reject(RX1!, RY1!, RX2!, RY2!, X1!, Y1!, X2!, Y2!) As Boolean

 'Cohen-Sutherland Trivial Reject (with codes):

 Dim Code1(3) As Boolean
 Dim Code2(3) As Boolean
 Dim Temp As Single

 If (RX1 > RX2) Then Temp = RX1: RX1 = RX2: RX2 = Temp
 If (RY1 > RY2) Then Temp = RY1: RY1 = RY2: RY2 = Temp

 If X1 < RX1 Then Code1(0) = True Else Code1(0) = False
 If X1 > RX2 Then Code1(1) = True Else Code1(1) = False
 If Y1 < RY1 Then Code1(2) = True Else Code1(2) = False
 If Y1 > RY2 Then Code1(3) = True Else Code1(3) = False

 If X2 < RX1 Then Code2(0) = True Else Code2(0) = False
 If X2 > RX2 Then Code2(1) = True Else Code2(1) = False
 If Y2 < RY1 Then Code2(2) = True Else Code2(2) = False
 If Y2 > RY2 Then Code2(3) = True Else Code2(3) = False

 If (Code1(0) And Code2(0)) Then Reject = True
 If (Code1(1) And Code2(1)) Then Reject = True
 If (Code1(2) And Code2(2)) Then Reject = True
 If (Code1(3) And Code2(3)) Then Reject = True

End Function
