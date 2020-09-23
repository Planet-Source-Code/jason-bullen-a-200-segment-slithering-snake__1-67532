Attribute VB_Name = "modMisc"
Option Explicit
Option Base 0

Public Const PI As Single = 3.14285714285714
Public Const PI2 As Single = 6.28571428571429


' StepAngle: Steps (step amount) closer towards a desired angle by the quickest direction
Public Function StepAngle(current As Single, desired As Single, step As Single) As Single
    Dim diff As Single
    diff = desired - current
    If Abs(diff) < PI Then
        If diff > 0# Then
            StepAngle = current + step
        Else
            StepAngle = current - step
        End If
    Else
        If diff > 0# Then
            StepAngle = BoundDirection(current - step)
        Else
            StepAngle = BoundDirection(current + step)
        End If
    End If
End Function

' BoundDirection: Keeps the angle value within the circle
Public Function BoundDirection(angle As Single) As Single
    If angle > PI Then
        angle = angle - PI2
    End If
    If angle < -PI Then
        angle = angle + PI2
    End If
    BoundDirection = angle
End Function



Public Function Atan2(ByVal y As Double, ByVal x As Double) As Double
    Dim theta As Double

    If (Abs(x) < 0.0000001) Then
        If (Abs(y) < 0.0000001) Then
            theta = 0#
        ElseIf (y > 0#) Then
            theta = 1.5707963267949
        Else
            theta = -1.5707963267949
        End If
    Else
        theta = Atn(y / x)
        If (x < 0) Then
            If (y >= 0#) Then
                theta = PI + theta
            Else
                theta = theta - PI
            End If
        End If
    End If
    Atan2 = theta
End Function

