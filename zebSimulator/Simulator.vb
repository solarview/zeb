Imports ZEB.Core
Imports Microsoft.Office.Interop

Public Class Simulator

#Region “Instance Variables”
    'Model-View-Controller Pattern
    'Private _Model As zebSimulator.Campus
    Private _View As ExcelTester '[View]

#End Region
#Region “Constructors”
    Public Sub New(ByVal anyView As ExcelTester)
        ' '_Model = myModel
        _View = anyView 'New MainForm(Me) '[create the View]
    End Sub
#End Region
#Region “Properties”
    Public ReadOnly Property View As ExcelTester
        Get
            Return _View
        End Get
    End Property
#End Region
#Region “Methods”
    Sub Test()
        'Dim theBuilding As building
    End Sub
#End Region
#Region “Shared Functions”
#End Region
#Region "Main Application Runs."
    Public Shared Sub Main()
        '---- 분석을 위한 개체 생성
        Dim theView As ExcelTester = New ExcelTester()
        Dim theSimulator As Simulator = New Simulator(theView) '[create the Controller]
        '
        '---- 프로그램을 실행 -> 폼이 닫히면 프로그램은 종료된다.
        Application.Run(theSimulator.View)
        ''
    End Sub
#End Region
End Class
