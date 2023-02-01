Public Function CubicSpline(X As Double, X As Variant, Y As Variant) As Double
' Returns an estimated value for y for a given x using cubic spline interpolation.
' x - independent variable for which y is to be estimated
' X, Y - arrays of x and y values for known data points

Dim n As Integer
Dim h As Variant
Dim b As Variant
Dim u As Variant
Dim v As Variant
Dim spline As Variant
Dim i As Integer

n = UBound(X) - LBound(X) + 1
ReDim h(1 To n - 1), b(1 To n), u(1 To n), v(1 To n)
ReDim spline(1 To 4, 1 To n - 1)

' Calculate h values
For i = 1 To n - 1
    h(i) = X(i + 1) - X(i)
Next i

' Calculate b values
For i = 2 To n - 1
    b(i) = (6 / h(i)) * ((Y(i + 1) - Y(i)) / h(i) - (Y(i) - Y(i - 1)) / h(i - 1))
Next i

' Calculate u and v values
u(2) = 2 * (X(3) - X(1)) - h(1) * h(1) / 3
v(2) = b(2) - h(1) * b(1) / 3

For i = 3 To n - 1
    u(i) = 2 * (X(i + 1) - X(i - 1)) - h(i - 1) * h(i - 1) / 3
    v(i) = b(i) - h(i - 1) * b(i - 1) / 3
Next i

' Back substitution to calculate spline coefficients
For i = n - 2 To 2 Step -1
    spline(2, i) = (v(i) - h(i) * spline(2, i + 1)) / u(i)
Next i

For i = 1 To n - 2
    spline(1, i) = (Y(i + 1) - Y(i)) / h(i) - h(i) * (spline(2, i + 1) + 2 * spline(2, i)) / 3
    spline(3, i) = (spline(2, i + 1) - spline(2, i)) / (3 * h(i))
    spline(4, i) = Y(i)
Next i

' Interpolate for given x value
For i = 1 To n - 2
    If X >= X(i) And X <= X(i + 1) Then
        CubicSpline = spline(1, i) * (X - X(i)) ^ 3 + spline(2, i) * (X - X(i)) ^ 2 + _
                      spline(3, i) * (X - X(i)) + spline(4, i)
        Exit Function
    End If
Next i
End Function






Public Function coef(a As Double, b As Double, c As Double, d As Double) As Variant

    coef = Array(0.2, -3, 0.5, -3)
End Function



Public Function poly(x As Double) As Double

    Dim coefs As Variant
    coefs = coef(3)

    poly = coefs(0) * x ^ 3 - coefs(1) * x ^ 2 - coefs(2) * x + coefs(3)

End Function


Public Function MySpline(x As Double) As Double

    If x < 0 Then
        MySpline = poly(x)
    Else
        MySpline = poly(x) * -1
    End If

End Function

