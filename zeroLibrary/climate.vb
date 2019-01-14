''' <summary>
''' 기상 자료를 
''' </summary>
Public Class ClimateData
#Region "Instance Variables"
    '외기온도
    Private _Temperature(11) As Double
    Private _Irradiance(8, 11) As Double '(방위, 월)
    Private _ClimateString As String

#End Region

#Region "Properties"
    Public ReadOnly Property Temperature(ByVal Month As Integer) As Double
        Get
            Return _Temperature(Month)
        End Get
    End Property
    Public ReadOnly Property Irradiance(ByVal Orientation As Integer, ByVal Month As Integer) As Double
        Get
            Return _Irradiance(Orientation, Month)
        End Get
    End Property

    Private _Test As Double

    Public Property Test As Double
        Get
            Return _Test
        End Get
        Set(value As Double)
            If value < 0 Or value > 11 Then
                value = 0
            End If
            _Test = value
        End Set
    End Property



#End Region

#Region "Methods"
    Public Sub Import(ByVal WeatherFile As String)

    End Sub
    Public Sub Export(ByVal WeatherFile As String)

    End Sub
#End Region

End Class
