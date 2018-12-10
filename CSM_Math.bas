Attribute VB_Name = "CSM_Math"
Option Explicit

Public Const PI As Double = 3.14159265358979

'*******************************
'Secant
Public Function Sec(ByVal Number As Double) As Double
    Sec = 1 / Cos(Number)
End Function

'*******************************
'Cosecant
Public Function Cosec(ByVal Number As Double) As Double
    Cosec = 1 / Sin(Number)
End Function

'*******************************
'Cotangent
Public Function Cotangent(ByVal Number As Double) As Double
    Cotangent = 1 / Tan(Number)
End Function

'*******************************
'Inverse Sine
Public Function Arcsine(ByVal Number As Double) As Double
    Arcsine = Atn(Number / Sqr(-Number * Number + 1))
End Function

'*******************************
'Inverse Cosine
Public Function Arccos(ByVal Number As Double) As Double
    Arccos = Atn(-Number / Sqr(-Number * Number + 1)) + 2 * Atn(1)
End Function

'*******************************
'Inverse Secant
Public Function Arcsec(ByVal Number As Double) As Double
    Arcsec = Atn(Number / Sqr(Number * Number - 1)) + Sgn((Number) - 1) * (2 * Atn(1))
End Function

'*******************************
'Inverse Cosecant
Public Function Arccosec(ByVal Number As Double) As Double
    Arccosec = Atn(Number / Sqr(Number * Number - 1)) + (Sgn(Number) - 1) * (2 * Atn(1))
End Function

'*******************************
'Inverse Cotangent
Public Function Arccotan(ByVal Number As Double) As Double
    Arccotan = Atn(Number) + 2 * Atn(1)
End Function

'Hyperbolic Sine HSin(X) = (Exp(X) – Exp(-X)) / 2
'Hyperbolic Cosine HCos(X) = (Exp(X) + Exp(-X)) / 2
'Hyperbolic Tangent HTan(X) = (Exp(X) – Exp(-X)) / (Exp(X) + Exp(-X))
'Hyperbolic Secant HSec(X) = 2 / (Exp(X) + Exp(-X))
'Hyperbolic Cosecant HCosec(X) = 2 / (Exp(X) – Exp(-X))
'Hyperbolic Cotangent HCotan(X) = (Exp(X) + Exp(-X)) / (Exp(X) – Exp(-X))
'Inverse Hyperbolic Sine HArcsin(X) = Log(X + Sqr(X * X + 1))
'Inverse Hyperbolic Cosine HArccos(X) = Log(X + Sqr(X * X – 1))
'Inverse Hyperbolic Tangent HArctan(X) = Log((1 + X) / (1 – X)) / 2
'Inverse Hyperbolic Secant HArcsec(X) = Log((Sqr(-X * X + 1) + 1) / X)
'Inverse Hyperbolic Cosecant HArccosec(X) = Log((Sgn(X) * Sqr(X * X + 1) + 1) / X)
'Inverse Hyperbolic Cotangent HArccotan(X) = Log((X + 1) / (X – 1)) / 2

'*******************************
'Logarithm to base N
Public Function LogN(ByVal Number As Double, ByVal Base As Double) As Double
    LogN = Log(Number) / Log(Base)
End Function

'////////////////////////////////////////////////////////////
'// Autor: Tomás A. Cardoner
'// Creación: 2014-02-14
'// Modificación: 2014-02-14
'// Descripción: Esta función convierte la parte decimal (o fraccional) de un número Double a un valor entero Long
'//              Ejemplo: 145.6548 -> 6548
'////////////////////////////////////////////////////////////
Public Function ConvertDecimalPartToInteger(ByVal Value As Double) As Long
    Value = Value - Fix(Value)
    Do While Value <> Fix(Value)
        Value = Value * 10
    Loop
    ConvertDecimalPartToInteger = CLng(Value)
End Function
