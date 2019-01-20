'인터페이스는 계약이다.

''' <summary>
''' energy balance를 계산하는 인터페이스 
''' </summary>
Public Interface IEnergyBalance

    ''' <summary>
    ''' 열획득을 계산한다.
    ''' </summary>
    ''' <returns>열획득(heat source)</returns>
    Function GetHeatSource() As Double

    ''' <summary>
    ''' 열손실을 계산한다.
    ''' </summary>
    ''' <returns>열손실(heat sink)</returns>
    Function GetHeatSink() As Double
End Interface
