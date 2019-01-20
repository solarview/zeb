Imports ZEB.Core '우리
Imports Microsoft.Office.Interop '엑셀 객체 라이브러리
'
Imports System.IO
Imports System.Xml
Imports System.Xml.Schema '유효성 검증용
Imports System.Text '//Encoding.GetEncoding()함수용

Public Class ExcelTester
    Dim objApp As Excel.Application
    Dim oldBook As Excel._Workbook

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'txtDisplay.Text = "Hello, World!"

        Call Simulator.Main()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim objBooks As Excel.Workbooks
        Dim objSheets As Excel.Sheets
        Dim objSheet As Excel._Worksheet
        Dim range As Excel.Range
        '

        '
        'Dim InputFileName As String = "c:\Test.XLS"
        ''<파일존재여부>
        'Dim CheckFile As New FileInfo(InputFileName) '파일 존재 여부를 검토하는 객체
        'If Not CheckFile.Exists Then '-- 파일의 존재 여부 검토
        '    Exit Sub
        'End If
        ''</파일존재여부>

        ' Create a new instance of Excel and start a new workbook.
        objApp = New Excel.Application()
        '<이전방법>
        objBooks = objApp.Workbooks
        '</이전방법>
        '<새방법>
        'objBooks = objApp.Workbooks.Open(My.Settings.BuildingDataFile)
        '</새방법>
        '
        '
        oldBook = objBooks.Add
        objSheets = oldBook.Worksheets
        objSheet = objSheets(1)

        'Get the range where the starting cell has the address
        'm_sStartingCell and its dimensions are m_iNumRows x m_iNumCols.
        range = objSheet.Range("A1", Reflection.Missing.Value)
        range = range.Resize(5, 5)

        If (Me.FillWithStrings.Checked = False) Then
            'Create an array.
            Dim saRet(5, 5) As Double

            'Fill the array.
            Dim iRow As Long
            Dim iCol As Long
            For iRow = 0 To 5
                For iCol = 0 To 5

                    'Put a counter in the cell.
                    saRet(iRow, iCol) = iRow * iCol
                Next iCol
            Next iRow

            'Set the range value to the array.
            range.Value = saRet

        Else
            'Create an array.
            Dim saRet(5, 5) As String

            'Fill the array.
            Dim iRow As Long
            Dim iCol As Long
            For iRow = 0 To 5
                For iCol = 0 To 5

                    'Put the row and column address in the cell.
                    saRet(iRow, iCol) = iRow.ToString() + "|" + iCol.ToString()
                Next iCol
            Next iRow

            'Set the range value to the array.
            range.Value = saRet
        End If

        'Return control of Excel to the user.
        objApp.Visible = True
        objApp.UserControl = True

        'Clean up a little.
        range = Nothing
        objSheet = Nothing
        objSheets = Nothing
        objBooks = Nothing
    End Sub

    Private Sub CmdImportExcel_Click(sender As Object, e As EventArgs) Handles cmdImportExcel.Click
        '
        Dim oldSheets As Excel.Sheets
        Dim oldSheet As Excel._Worksheet
        Dim oldRange As Excel.Range

        Dim newBook As Excel.Workbook
        Dim newSheets As Excel.Sheets
        Dim newSheet As Excel._Worksheet
        '
        '<엑셀파일 읽어오기>
        ' Create a new instance of Excel and start a new workbook.
        objApp = New Excel.Application()
        '<이전방법>
        'objBooks = objApp.Workbooks
        '</이전방법>
        Dim filename As String
        filename = GetFileName()
        '<새방법>
        oldBook = objApp.Workbooks.Open(filename) '("c:\Temp\ZeroEnergyTool2Test.xlsx") '(My.Settings.BuildingDataFile)
        '</새방법>
        '</엑셀파일 읽어오기>

        'Get a reference to the first sheet of the workbook.
        On Error GoTo ExcelNotRunning
        oldSheets = oldBook.Worksheets
        oldSheet = oldSheets("건물") ' oldSheets.Item("건물")
        '
        newBook = objApp.Workbooks.Add '새 엑셀 파일을 연다
        newSheets = newBook.Worksheets
        newSheet = newSheets(1)

ExcelNotRunning:
        If (Not (Err.Number = 0)) Then
            MessageBox.Show("Cannot find the Excel workbook.  Try clicking Button1 to " +
            "create an Excel workbook with data before running Button2.",
            "Missing Workbook?")

            'We cannot automate Excel if we cannot find the data we created, 
            'so leave the subroutine.
            Exit Sub
        End If

        For i As Long = 0 To 66 ' 지역
            'Get a range of data.
            'oldRange = oldSheet.Range("D8", "M19")
            newSheet.Cells(1 + i, 1).Value = oldSheet.Cells(4 + (17 * i), 2).Value '번호
            newSheet.Cells((1 + i), 2).Value = oldSheet.Cells(4 + (17 * i), 3).Value '지역이름

            Dim oldRowCounter As Long
            Dim colCounter As Long
            '
            Dim newRowCounter As Long
            Dim newColCounter As Long
            '
            Dim j As Long = 3
            For oldRowCounter = 0 To 11
                For colCounter = 0 To 9
                    newSheet.Cells((1 + i), j).Value = oldSheet.Cells(8 + (17 * i) + oldRowCounter, 4 + colCounter).Value
                    j += 1
                Next colCounter
            Next oldRowCounter
        Next i
        '
        'txtDisplay.Text = valueString
        '
        'Return control of Excel to the user.
        objApp.Visible = True
        objApp.UserControl = True

        'Clean up a little.
        oldRange = Nothing
        oldSheet = Nothing
        oldSheets = Nothing
        '
        'newRange = Nothing
        newSheet = Nothing
        newSheets = Nothing
    End Sub
    Sub test()
        Dim theZone As Building
        Dim otherZone As Building
        '
        theZone = New Building()
        otherZone = New Building(theZone)
    End Sub

    Function GetFileName() As String
        Dim filename As String = ""
        '0) 파일 열기를 위한 다이얼로그박스
        Using dlgFile As New OpenFileDialog()
            If My.Settings.BuildingDataFile.Length <> 0 Then
                dlgFile.InitialDirectory = My.Computer.FileSystem.GetParentPath(My.Settings.BuildingDataFile) 'Application.StartupPath()
            Else
                dlgFile.InitialDirectory = "c:\Temp\"
            End If
            dlgFile.FileName = My.Settings.BuildingDataFile
            dlgFile.Filter = "Microsoft Excel File(*.xlsx)|*.xlsx"
            dlgFile.DefaultExt = ".xlsx"
            '
            If dlgFile.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then     'If .FileName.Length <> 0 Then '새 값을 선택한 경우
                '<파일 이름 설정>
                My.Settings.BuildingDataFile = dlgFile.FileName '--선택한 파일명 저장

                '</파일 이름 설정>
                '
                filename = dlgFile.FileName
                '<파일 읽기>
                'm_Simulator.Read()
                '</파일 읽기>
                '
                '<화면 출력>
                'Me.Text = "SolarView - " & My.Computer.FileSystem.GetName(dlgFile.FileName)
                'Me.Simulator.SetViewModeChanged(m_ViewMode)

                '</화면 출력>
                '
            End If
            dlgFile.Dispose() '삭제
        End Using
        '
        Return filename
    End Function
End Class