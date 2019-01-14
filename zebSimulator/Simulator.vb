Imports zebSimulator

Public Class Simulator

#Region “Instance Variables”
    'Model-View-Controller Pattern
    'Private _Model As zebSimulator.Campus
    'Private _View As MainForm '[View]
    'Private _Model As SV_Diagram '
#End Region
#Region “Constructors”
    ' Public Sub New(ByVal myModel As IDiagramModel)
    ' '_Model = myModel
    ' _View = New SV_MainForm(Me, myModel) '[create the View]
    ' End Sub
#End Region
#Region “Properties”
#End Region
#Region “Methods”
#End Region
#Region “Shared Functions”
#End Region
#Region "Main Application Runs."
    Public Shared Sub Main()
        '---- 분석을 위한 개체 생성
        'Dim myDiagram As IDiagramModel = New SV_Diagram()
        'Dim mySimulator As Simulator = New Simulator(myDiagram) '[create the Controller]
        '
        '---- 프로그램을 실행 -> 폼이 닫히면 프로그램은 종료된다.
        'Application.Run(mySimulator.View)
        ''
    End Sub
#End Region
End Class
