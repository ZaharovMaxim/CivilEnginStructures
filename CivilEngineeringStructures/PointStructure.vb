'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Imports Topomatic.Cad.Foundation

Public Class PointStructure
    Private _xOrigin As Double
    Private _yOrigin As Double
    Private _zOrigin As Double
    Private _dxOrigin As Double
    Private _dyOrigin As Double
    Private _dzOrigin As Double
    Private _dhOrigin As Double
    Private _xProj As Double
    Private _yProj As Double
    Private _zProj As Double
    Private _code As String
    Public Sub New()
        _xOrigin = 0
        _yOrigin = 0
        _zOrigin = 0
        _dxOrigin = 0
        _dyOrigin = 0
        _dzOrigin = 0
        _dhOrigin = 0
        _xProj = 0
        _yProj = 0
        _zProj = 0
        _code = ""
    End Sub
    Public Sub New(X1 As Double, Y1 As Double, Z1 As Double, dx As Double, dy As Double, dz As Double, DH As Double, Code As String)
        _xOrigin = Math.Round(X1, 3)
        _yOrigin = Math.Round(Y1, 3)
        _zOrigin = Math.Round(Z1, 3)
        If dx <> 0 Then
            _dxOrigin = Math.Round(dx, 3)
        Else
            _dxOrigin = 0
        End If
        If dy <> 0 Then
            _dyOrigin = Math.Round(dy, 3)
        Else
            _dyOrigin = 0
        End If
        If dz <> 0 Then
            _dzOrigin = Math.Round(dz, 3)
        Else
            _dzOrigin = 0
        End If
        _dhOrigin = DH
        _code = Code
    End Sub
    Public Sub New(point As Vector3D, Code As String)
        _xOrigin = Math.Round(point.X, 3)
        _yOrigin = Math.Round(point.Y, 3)
        _zOrigin = Math.Round(point.Z, 3)
        _code = Code
    End Sub
    Public Property X() As Double
        Get
            Return _xOrigin
        End Get
        Set(value As Double)
            _xOrigin = value
        End Set
    End Property

    Public Property Y() As Double
        Get
            Return _yOrigin
        End Get
        Set(value As Double)
            _yOrigin = value
        End Set
    End Property

    Public Property Z() As Double
        Get
            Return _zOrigin
        End Get
        Set(value As Double)
            _zOrigin = value
        End Set
    End Property

    Public Property dx() As Double
        Get
            Return _dxOrigin
        End Get
        Set(value As Double)
            _dxOrigin = value
        End Set
    End Property

    Public Property dy() As Double
        Get
            Return _dyOrigin
        End Get
        Set(value As Double)
            _dyOrigin = value
        End Set
    End Property

    Public Property dz() As Double
        Get
            Return _dzOrigin
        End Get
        Set(value As Double)
            _dzOrigin = value
        End Set
    End Property

    Public Property dz2() As Double
        Get
            Return _dhOrigin
        End Get
        Set(value As Double)
            _dhOrigin = value
        End Set
    End Property

    Public Property Code() As String
        Get
            Return _code
        End Get
        Set(value As String)
            _code = value
        End Set
    End Property

    Public Property Xproj() As Double
        Get
            Return _xProj
        End Get
        Set(value As Double)
            _xProj = value
        End Set
    End Property

    Public Property Yproj() As Double
        Get
            Return _yProj
        End Get
        Set(value As Double)
            _yProj = value
        End Set
    End Property

    Public Property Zproj() As Double
        Get
            Return _zProj
        End Get
        Set(value As Double)
            _zProj = value
        End Set
    End Property
End Class
