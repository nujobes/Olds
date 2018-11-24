Imports Oracle.DataAccess.Client
Imports vb01.GetBasedata
Imports System.IO
Imports System.Text

Public Class Form1
    Public g_excel As Object
    Public g_excelsheet As Worksheet
    Public g_excelbook As Workbook
    'Public pos As Integer

    Public g_basedata As New Basedata '오

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            g_excel = GetObject(, "Excel.Application")
        Catch ex As Exception
            Dim msg As String = ex.Message
        End Try
        g_excelbook = g_excel.ActiveWorkbook
        g_excelsheet = g_excelbook.ActiveSheet

        TextBox1.Clear()

        For Each wkSht In g_excelbook.Worksheets
            TextBox1.Text += wkSht.Name + vbCrLf
        Next wkSht

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        OpenFileDialog1.Title = "Please select a file to open"
        OpenFileDialog1.Filter = "Excel files (*.xlsx; *.xls)|*.xlsx;*.xls"

        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            TextBox1.Clear()
            ProgressBar1.Minimum = 0
            ProgressBar1.Maximum = OpenFileDialog1.FileNames.Length
            ProgressBar1.Value = 0

            For Each file_path In OpenFileDialog1.FileNames
                TextBox1.Text += file_path + vbCrLf
                g_excel = CreateObject("Excel.Application")
                g_excel.Visible = False
                g_excelbook = g_excel.Workbooks.Open(file_path)
                g_excelbook = g_excel.ActiveWorkbook
                g_excelsheet = g_excelbook.ActiveSheet

                For Each wkSht In g_excelbook.Worksheets
                    If (wkSht.Name = "겉표지" OrElse wkSht.Name = "개정이력" OrElse wkSht.Name = "산출물작성지침") Then

                    Else
                        TextBox1.Text += wkSht.Name + vbCrLf
                    End If

                Next wkSht
                ProgressBar1.Value += 1
                g_excelbook.Close()
                g_excel.Quit()
            Next file_path
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        OpenFileDialog1.Title = "Please select a file to open"
        OpenFileDialog1.Filter = "Excel files (*.xlsx; *.xls)|*.xlsx;*.xls"

        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            TextBox1.Clear()
            TextBox1.Text += "개방번호," & "공공데이터명," & "서비스명," & "레이어명," & "레이어ID," & "서비스명-시트" + vbCrLf
            ProgressBar1.Minimum = 0
            ProgressBar1.Maximum = OpenFileDialog1.FileNames.Length
            ProgressBar1.Value = 0

            For Each file_path In OpenFileDialog1.FileNames
                TextBox1.Text += file_path + vbCrLf
                g_excel = CreateObject("Excel.Application")
                g_excel.Visible = False
                g_excelbook = g_excel.Workbooks.Open(file_path)
                g_excelbook = g_excel.ActiveWorkbook
                g_excelsheet = g_excelbook.ActiveSheet

                For Each wkSht In g_excelbook.Worksheets
                    If (wkSht.Name = "겉표지" OrElse wkSht.Name = "개정이력" OrElse wkSht.Name = "산출물작성지침") Then

                    Else
                        TextBox1.Text += wkSht.Range("E5").Value & "," & wkSht.Range("C4").Value & "," & wkSht.Range("C5").Value & ","
                        TextBox1.Text += wkSht.Range("E6").Value & "," & wkSht.Range("C6").Value & ","
                        TextBox1.Text += wkSht.Name + vbCrLf
                    End If
                Next wkSht

                ProgressBar1.Value += 1
                g_excelbook.Close()
                g_excel.Quit()
            Next file_path
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        OpenFileDialog1.Title = "Please select a file to open"
        OpenFileDialog1.Filter = "Excel files (*.xlsx; *.xls)|*.xlsx;*.xls"

        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            TextBox1.Clear()
            TextBox1.Text += "개방번호," & "공공데이터명," & "서비스명," & "레이어명," & "레이어ID," & "서비스명-시트" + vbCrLf
            ProgressBar1.Minimum = 0
            ProgressBar1.Maximum = OpenFileDialog1.FileNames.Length
            ProgressBar1.Value = 0
            For Each file_path In OpenFileDialog1.FileNames
                Dim file_name As String
                file_name = System.IO.Path.GetFileName(file_path)
                TextBox1.Text += file_name + vbCrLf

                g_excel = CreateObject("Excel.Application")
                g_excel.Visible = False
                g_excelbook = g_excel.Workbooks.Open(file_path)
                g_excelbook = g_excel.ActiveWorkbook

                For Each wkSht In g_excelbook.Worksheets
                    If (wkSht.Name = "겉표지" OrElse wkSht.Name = "개정이력" OrElse wkSht.Name = "산출물작성지침") Then
                    Else
                        TextBox1.Text += wkSht.cells(5, 5).Value & "," & wkSht.cells(4, 3).Value & "," & wkSht.cells(5, 3).Value & ","
                        TextBox1.Text += wkSht.cells(6, 5).Value & "," & wkSht.cells(6, 3).Value & ","
                        TextBox1.Text += wkSht.Name + vbCrLf
                    End If
                Next wkSht

                ProgressBar1.Value += 1
                g_excelbook.Close()
                g_excel.Quit()
            Next file_path
        End If
        TextBox1.Text += vbCrLf + "작업완료."
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        OpenFileDialog1.Title = "Please select a file to open"
        OpenFileDialog1.Filter = "Excel files (*.xlsx; *.xls)|*.xlsx;*.xls"

        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            TextBox1.Clear()
            TextBox1.Text += "파일명|" & "테이블ID|" & "컬럼ID|" & "데이터명|" & "타입|" & "길이|" & "NULL|" & "J" + vbCrLf
            ProgressBar1.Minimum = 0
            ProgressBar1.Maximum = OpenFileDialog1.FileNames.Length
            ProgressBar1.Value = 0
            For Each file_path In OpenFileDialog1.FileNames
                Dim file_name As String
                file_name = System.IO.Path.GetFileName(file_path)
                'TextBox1.Text += file_name + vbCrLf

                g_excel = CreateObject("Excel.Application")
                g_excel.Visible = False
                g_excelbook = g_excel.Workbooks.Open(file_path)
                g_excelbook = g_excel.ActiveWorkbook

                For Each wkSht In g_excelbook.Worksheets
                    Label1.Text = "[" & ProgressBar1.Value & "/" & OpenFileDialog1.FileNames.Length & "] " & file_name
                    If (wkSht.Name = "개방테이블") Then
                        Dim pos As Integer = 1
                        Dim wkSht_lastrow As Long
                        wkSht_lastrow = wkSht.UsedRange.Rows.Count + 1
                        pos = 1
                        If (wkSht.cells(5, 7).Value = "데이터설명") Then
                            Do While (wkSht_lastrow > pos)
                                If (wkSht.cells(pos, 10).Value <> "" And wkSht.cells(pos, 10).Value <> "J") Then
                                    TextBox1.Text += file_name & "|" & wkSht.cells(3, 3).Value & "|"
                                    TextBox1.Text += wkSht.cells(pos, 2).Value & "|" & wkSht.cells(pos, 3).Value & "," & wkSht.cells(pos, 4).Value & "|"
                                    TextBox1.Text += wkSht.cells(pos, 5).Value & "|"
                                    If (wkSht.cells(pos, 6).Value = "N" OrElse wkSht.cells(pos, 6).Value = "n" OrElse wkSht.cells(pos, 6).Value = "NotNull") Then
                                        TextBox1.Text += "NotNull|"
                                    ElseIf (wkSht.cells(pos, 6).Value = "" OrElse wkSht.cells(pos, 6).Value = " ") Then
                                        TextBox1.Text += "Null|"
                                    Else
                                        TextBox1.Text += wkSht.cells(pos, 6).Value & "|"
                                    End If
                                    TextBox1.Text += wkSht.cells(pos, 10).Value + vbCrLf
                                End If
                                pos = pos + 1
                            Loop
                        ElseIf (wkSht.cells(5, 7).Value = "PK참여정보") Then
                            Do While (wkSht_lastrow > pos)
                                If (wkSht.cells(pos, 12).Value <> "" And wkSht.cells(pos, 12).Value <> "J") Then
                                    TextBox1.Text += file_name & "|" & wkSht.cells(3, 3).Value & "|"
                                    TextBox1.Text += wkSht.cells(pos, 2).Value & "|" & wkSht.cells(pos, 3).Value & "|" & wkSht.cells(pos, 4).Value & "|"
                                    TextBox1.Text += wkSht.cells(pos, 5).Value & "|"
                                    If (wkSht.cells(pos, 6).Value = "N" OrElse wkSht.cells(pos, 6).Value = "n" OrElse wkSht.cells(pos, 6).Value = "NotNull") Then
                                        TextBox1.Text += "NotNull|"
                                    ElseIf (wkSht.cells(pos, 6).Value = "" OrElse wkSht.cells(pos, 6).Value = " ") Then
                                        TextBox1.Text += "Null|"
                                    Else
                                        TextBox1.Text += wkSht.cells(pos, 6).Value & "|"
                                    End If
                                    TextBox1.Text += wkSht.cells(pos, 12).Value + vbCrLf
                                End If
                                pos = pos + 1
                            Loop
                        Else : TextBox1.Text += "알수없는정의서형식!" + vbCrLf
                        End If
                    End If
                Next wkSht

                ProgressBar1.Value += 1
                g_excelbook.Close()
                g_excel.Quit()
            Next file_path
        End If
        Label1.Text = "[" & ProgressBar1.Value & "/" & OpenFileDialog1.FileNames.Length & "] " & "[작업완료]"
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        OpenFileDialog1.Title = "Please select a file to open"
        OpenFileDialog1.Filter = "Excel files (*.xlsx; *.xls)|*.xlsx;*.xls"

        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            TextBox1.Clear()
            TextBox1.Text += "파일명|" & "테이블ID|" & "컬럼ID|" & "데이터명|" & "타입|" & "길이|" & "NULL|" & "J" + vbCrLf
            ProgressBar1.Minimum = 0
            ProgressBar1.Maximum = OpenFileDialog1.FileNames.Length
            ProgressBar1.Value = 0
            For Each file_path In OpenFileDialog1.FileNames
                Dim file_name As String
                file_name = System.IO.Path.GetFileName(file_path)
                'TextBox1.Text += file_name + vbCrLf

                g_excel = CreateObject("Excel.Application")
                g_excel.Visible = False
                g_excelbook = g_excel.Workbooks.Open(file_path)
                g_excelbook = g_excel.ActiveWorkbook

                For Each wkSht In g_excelbook.Worksheets
                    Label1.Text = "[" & ProgressBar1.Value & "/" & OpenFileDialog1.FileNames.Length & "] " & file_name
                    If (wkSht.Name = "개방테이블") Then
                        Dim pos As Integer = 1
                        Dim wkSht_lastrow As Long
                        wkSht_lastrow = wkSht.UsedRange.Rows.Count + 1
                        pos = 1
                        If (wkSht.cells(5, 7).Value = "데이터설명") Then
                            Do While (wkSht_lastrow > pos)
                                If (wkSht.cells(pos, 10).Value <> "" And wkSht.cells(pos, 10).Value <> "J") Then
                                    TextBox1.Text += file_name & "|" & wkSht.cells(3, 3).Value & "|"
                                    TextBox1.Text += wkSht.cells(pos, 2).Value & "|" & wkSht.cells(pos, 3).Value & "|" & wkSht.cells(pos, 4).Value & "|"
                                    TextBox1.Text += wkSht.cells(pos, 5).Value & "|"
                                    If (wkSht.cells(pos, 6).Value = "N" OrElse wkSht.cells(pos, 6).Value = "n" OrElse wkSht.cells(pos, 6).Value = "NotNull") Then
                                        TextBox1.Text += "NotNull|"
                                    ElseIf (wkSht.cells(pos, 6).Value = "" OrElse wkSht.cells(pos, 6).Value = " ") Then
                                        TextBox1.Text += "Null|"
                                    Else
                                        TextBox1.Text += wkSht.cells(pos, 6).Value & "|"
                                    End If
                                    TextBox1.Text += wkSht.cells(pos, 10).Value + vbCrLf
                                End If
                                pos = pos + 1
                            Loop
                        ElseIf (wkSht.cells(5, 7).Value = "PK참여정보") Then
                            Do While (wkSht_lastrow > pos)
                                If (wkSht.cells(pos, 12).Value <> "" And wkSht.cells(pos, 12).Value <> "J") Then
                                    TextBox1.Text += file_name & "|" & wkSht.cells(3, 3).Value & "|"
                                    TextBox1.Text += wkSht.cells(pos, 2).Value & "|" & wkSht.cells(pos, 3).Value & "|" & wkSht.cells(pos, 4).Value & "|"
                                    TextBox1.Text += wkSht.cells(pos, 5).Value & "|"
                                    If (wkSht.cells(pos, 6).Value = "N" OrElse wkSht.cells(pos, 6).Value = "n" OrElse wkSht.cells(pos, 6).Value = "NotNull") Then
                                        TextBox1.Text += "NotNull|"
                                    ElseIf (wkSht.cells(pos, 6).Value = "" OrElse wkSht.cells(pos, 6).Value = " ") Then
                                        TextBox1.Text += "Null|"
                                    Else
                                        TextBox1.Text += wkSht.cells(pos, 6).Value & ","
                                    End If
                                    TextBox1.Text += wkSht.cells(pos, 12).Value + vbCrLf
                                End If
                                pos = pos + 1
                            Loop
                        Else : TextBox1.Text += "알수없는정의서형식!" + vbCrLf
                        End If
                    End If
                Next wkSht

                ProgressBar1.Value += 1
                'g_excelbook.Close(False)
                g_excelbook.Close(False)
                g_excel.Quit()
            Next file_path
        End If
        Label1.Text = "[" & ProgressBar1.Value & "/" & OpenFileDialog1.FileNames.Length & "] " & "작업완료"
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        OpenFileDialog1.Title = "Please select a file to open"
        OpenFileDialog1.Filter = "Excel files (*.xlsx; *.xls)|*.xlsx;*.xls"

        Dim DataGrid_temp(1, 1) As String
        Dim Row_count As Integer

        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            TextBox1.Clear()
            TextBox1.Text += "파일명|" & "테이블ID|" & "컬럼ID|" & "데이터명|" & "타입|" & "길이|" & "NULL|" & "J" + vbCrLf

            '데이터그리드선언, 반가변으로, 시트당20줄씩할당, 열은8개
            Dim DataGrid_rowsCount As Integer = OpenFileDialog1.FileNames.Length * 20 - 1
            Row_count = DataGrid_rowsCount
            ReDim DataGrid_temp(DataGrid_rowsCount, 7)
            'Dim DataGrid_temp(DataGrid_rowsCount, 7) As String

            'DataGrid_temp = {{"파일명", "테이블ID", "컬럼ID", "데이터명", "타입", "길이", "NULL", "J"}}
            DataGrid_temp(0, 0) = "파일명"
            DataGrid_temp(0, 1) = "테이블ID"
            DataGrid_temp(0, 2) = "컬럼ID"
            DataGrid_temp(0, 3) = "데이터명"
            DataGrid_temp(0, 4) = "타입"
            DataGrid_temp(0, 5) = "길이"
            DataGrid_temp(0, 6) = "NULL"
            DataGrid_temp(0, 7) = "J"
            'For i = 1 To 8
            'DataGrid_temp(1, i) = ""
            'Next
            Dim DataGrid_rowPos As Integer = 1

            ProgressBar1.Minimum = 0
            ProgressBar1.Maximum = OpenFileDialog1.FileNames.Length
            ProgressBar1.Value = 0

            For Each file_path In OpenFileDialog1.FileNames
                Dim file_name As String
                file_name = System.IO.Path.GetFileName(file_path)
                'TextBox1.Text += file_name + vbCrLf

                g_excel = CreateObject("Excel.Application")
                g_excel.Visible = False
                g_excelbook = g_excel.Workbooks.Open(file_path)
                g_excelbook = g_excel.ActiveWorkbook

                For Each wkSht In g_excelbook.Worksheets
                    Label1.Text = "[" & ProgressBar1.Value & "/" & OpenFileDialog1.FileNames.Length & "] " & file_name
                    If (wkSht.Name = "개방테이블") Then
                        Dim pos As Integer = 1
                        Dim wkSht_lastrow As Long
                        wkSht_lastrow = wkSht.UsedRange.Rows.Count + 1
                        pos = 1
                        If (wkSht.cells(5, 7).Value = "데이터설명") Then
                            Do While (wkSht_lastrow > pos)
                                If (wkSht.cells(pos, 10).Value <> "" And wkSht.cells(pos, 10).Value <> "J") Then
                                    TextBox1.Text += file_name & "|" & wkSht.cells(3, 3).Value & "|"
                                    TextBox1.Text += wkSht.cells(pos, 2).Value & "|" & wkSht.cells(pos, 3).Value & "|" & wkSht.cells(pos, 4).Value & "|"
                                    TextBox1.Text += wkSht.cells(pos, 5).Value & "|"

                                    DataGrid_temp(DataGrid_rowPos, 0) = file_name
                                    DataGrid_temp(DataGrid_rowPos, 1) = wkSht.cells(3, 3).Value
                                    DataGrid_temp(DataGrid_rowPos, 2) = wkSht.cells(pos, 2).Value
                                    DataGrid_temp(DataGrid_rowPos, 3) = wkSht.cells(pos, 3).Value
                                    DataGrid_temp(DataGrid_rowPos, 4) = wkSht.cells(pos, 4).Value
                                    DataGrid_temp(DataGrid_rowPos, 5) = wkSht.cells(pos, 5).Value

                                    If (wkSht.cells(pos, 6).Value = "N" OrElse wkSht.cells(pos, 6).Value = "n" OrElse wkSht.cells(pos, 6).Value = "NotNull") Then
                                        TextBox1.Text += "NotNull|"
                                        DataGrid_temp(DataGrid_rowPos, 6) = "NotNull"
                                    ElseIf (wkSht.cells(pos, 6).Value = "" OrElse wkSht.cells(pos, 6).Value = " ") Then
                                        TextBox1.Text += "Null|"
                                        DataGrid_temp(DataGrid_rowPos, 6) = "Null"
                                    Else
                                        TextBox1.Text += wkSht.cells(pos, 6).Value & "|"
                                        DataGrid_temp(DataGrid_rowPos, 6) = wkSht.cells(pos, 6).Value
                                    End If
                                    TextBox1.Text += wkSht.cells(pos, 10).Value + vbCrLf
                                    DataGrid_temp(DataGrid_rowPos, 7) = wkSht.cells(pos, 10).Value
                                    DataGrid_rowPos = DataGrid_rowPos + 1
                                End If
                                pos = pos + 1
                            Loop
                        ElseIf (wkSht.cells(5, 7).Value = "PK참여정보") Then
                            Do While (wkSht_lastrow > pos)
                                If (wkSht.cells(pos, 12).Value <> "" And wkSht.cells(pos, 12).Value <> "J") Then
                                    TextBox1.Text += file_name & "|" & wkSht.cells(3, 3).Value & "|"
                                    TextBox1.Text += wkSht.cells(pos, 2).Value & "|" & wkSht.cells(pos, 3).Value & "|" & wkSht.cells(pos, 4).Value & "|"
                                    TextBox1.Text += wkSht.cells(pos, 5).Value & "|"

                                    DataGrid_temp(DataGrid_rowPos, 0) = file_name
                                    DataGrid_temp(DataGrid_rowPos, 1) = wkSht.cells(3, 3).Value
                                    DataGrid_temp(DataGrid_rowPos, 2) = wkSht.cells(pos, 2).Value
                                    DataGrid_temp(DataGrid_rowPos, 3) = wkSht.cells(pos, 3).Value
                                    DataGrid_temp(DataGrid_rowPos, 4) = wkSht.cells(pos, 4).Value
                                    DataGrid_temp(DataGrid_rowPos, 5) = wkSht.cells(pos, 5).Value

                                    If (wkSht.cells(pos, 6).Value = "N" OrElse wkSht.cells(pos, 6).Value = "n" OrElse wkSht.cells(pos, 6).Value = "NotNull") Then
                                        TextBox1.Text += "NotNull|"
                                        DataGrid_temp(DataGrid_rowPos, 6) = "NotNull"
                                    ElseIf (wkSht.cells(pos, 6).Value = "" OrElse wkSht.cells(pos, 6).Value = " ") Then
                                        TextBox1.Text += "Null|"
                                        DataGrid_temp(DataGrid_rowPos, 6) = "Null"
                                    Else
                                        TextBox1.Text += wkSht.cells(pos, 6).Value & "|"
                                        DataGrid_temp(DataGrid_rowPos, 6) = wkSht.cells(pos, 6).Value
                                    End If
                                    TextBox1.Text += wkSht.cells(pos, 12).Value + vbCrLf
                                    DataGrid_temp(DataGrid_rowPos, 7) = wkSht.cells(pos, 12).Value
                                    DataGrid_rowPos = DataGrid_rowPos + 1
                                End If
                                pos = pos + 1
                            Loop
                        Else : TextBox1.Text += "알수없는정의서형식!" + vbCrLf
                        End If
                    End If
                Next wkSht

                ProgressBar1.Value += 1
                'g_excelbook.Close(False)
                g_excelbook.Close(False)
                g_excel.Quit()
            Next file_path
        Else
            Exit Sub
        End If
        Label1.Text = "[" & ProgressBar1.Value & "/" & OpenFileDialog1.FileNames.Length & "] " & "작업완료"

        'Start a new workbook in Excel
        g_excel = CreateObject("Excel.Application")
        g_excel.Visible = False
        'g_excelbook = g_excel.ActiveWorkbook
        'g_excelsheet = g_excelbook.ActiveSheet
        g_excelbook = g_excel.Workbooks.Add

        'Add data to cells of the first worksheet in the new workbook
        g_excelsheet = g_excelbook.Worksheets(1)
        'g_excelsheet.Range("A1").Value = "Last Name"
        'g_excelsheet.Range("B1").Value = "First Name"
        'g_excelsheet.Range("A1:B1").Font.Bold = True
        'g_excelsheet.Range("A2").Value = "Doe"
        'g_excelsheet.Range("B2").Value = "John"
        For i As Integer = 0 To Row_count
            If DataGrid_temp(i, 0) = Nothing Then
                Exit For
            End If
            g_excelsheet.Cells(i + 1, 1).Value = DataGrid_temp(i, 0)
            g_excelsheet.Cells(i + 1, 2).Value = DataGrid_temp(i, 1)
            g_excelsheet.Cells(i + 1, 3).Value = DataGrid_temp(i, 2)
            g_excelsheet.Cells(i + 1, 4).Value = DataGrid_temp(i, 3)
            g_excelsheet.Cells(i + 1, 5).Value = DataGrid_temp(i, 4)
            g_excelsheet.Cells(i + 1, 6).Value = DataGrid_temp(i, 5)
            g_excelsheet.Cells(i + 1, 7).Value = DataGrid_temp(i, 6)
            g_excelsheet.Cells(i + 1, 8).Value = DataGrid_temp(i, 7)
        Next

        'Save the Workbook and Quit Excel
        Dim SaveAs_filePath As String = "C:\RDB\" & "서비스정의서" & DateTime.Now.ToString("-yyMMddHHmmss") & ".xlsx"
        g_excelbook.SaveAs(SaveAs_filePath)
        'g_excelbook.SaveAs("C:\RDB\서비스정의서" & saveas_filename & ".xls")
        g_excel.Quit()
        Label1.Text = "[" & ProgressBar1.Value & "/" & OpenFileDialog1.FileNames.Length & "] " & "작업완료 [" & SaveAs_filePath & "] 저장완료"
        'Erase DataGrid_temp
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        TextBox1.Clear()
        Dim oradb As String = "Data Source=(DESCRIPTION=" _
           + "(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=${HOSTIP})(PORT=1521)))" _
           + "(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=ORCL)));" _
           + "User Id=SYSTEM;Password=1234567;"
        Dim conn As New OracleConnection(oradb)

        Try
            conn.Open() '연결
            Dim sql_count As Integer = 0

            'Dim sql As String = "SELECT * FROM TABLE_INFO"
            Dim sql As String = "SELECT TABLE_ID FROM TABLE_INFO GROUP BY TABLE_ID"
            'Dim sql As String = "SELECT * FROM TABLE_INFO where Table_Id='새올-154-01'
            Dim cmd As New OracleCommand(sql, conn)

            cmd.CommandType = CommandType.Text

            Dim data_reader_count As OracleDataReader = cmd.ExecuteReader() ' VB.NET
            Try
                While data_reader_count.Read()
                    sql_count = sql_count + 1
                End While
            Catch ex As Exception
                Label1.Text = ex.Message.ToString()
                TextBox1.Text += ex.Message.ToString()
            Finally
                data_reader_count.Close()
            End Try

            Dim data_reader As OracleDataReader = cmd.ExecuteReader() ' VB.NET
            Try
                ProgressBar1.Minimum = 0
                ProgressBar1.Maximum = sql_count
                ProgressBar1.Value = 0

                While data_reader.Read()
                    TextBox1.Text += data_reader.Item("TABLE_ID") ' retrieve by column name
                    'TextBox1.Text += data_reader.Item("COLUMN_ID")
                    'TextBox1.Text += data_reader.Item("DATA_NAME")
                    'TextBox1.Text += data_reader.Item("DATA_TYPE")
                    'TextBox1.Text += data_reader.Item("DATA_LENGTH")
                    'TextBox1.Text += data_reader.Item("IS_NULL")
                    'TextBox1.Text += data_reader.Item("INFO")
                    TextBox1.Text += vbCrLf
                    ProgressBar1.Value += 1
                    Label1.Text = "[" & ProgressBar1.Value & "/" & sql_count & "]"
                End While
            Catch ex As Exception
                'Label1.Text = ex.Message.ToString()
                'TextBox1.Text = ex.Message.ToString()
            Finally
                data_reader.Close()
            End Try
            sql_count = 0
        Catch ex As Exception
            Label1.Text = ex.Message.ToString()
            TextBox1.Text += ex.Message.ToString()
        Finally
            conn.Dispose()
        End Try
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        TextBox1.Clear()
        Dim oradb As String = "Data Source=(DESCRIPTION=" _
   + "(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=${HOSTIP})(PORT=1521)))" _
   + "(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=ORCL)));" _
   + "User Id=SYSTEM;Password=1234567;"
        Dim conn As New OracleConnection(oradb)

        Try
            conn.Open() '연결
            Dim sql_count As Integer = 0

            'Dim sql As String = "SELECT * FROM TABLE_INFO"
            Dim sql As String = "SELECT TABLE_ID FROM TABLE_INFO GROUP BY TABLE_ID"
            'Dim sql As String = "SELECT * FROM TABLE_INFO where Table_Id='새올-154-01'
            Dim cmd As New OracleCommand(sql, conn)

            cmd.CommandType = CommandType.Text

            Dim data_reader_count As OracleDataReader = cmd.ExecuteReader()
            Try
                While data_reader_count.Read()
                    sql_count = sql_count + 1
                End While
            Catch ex As Exception
                Label1.Text = ex.Message.ToString()
                TextBox1.Text = ex.Message.ToString()
            Finally
                data_reader_count.Close()
            End Try

            Dim data_reader As OracleDataReader = cmd.ExecuteReader()
            Try
                ProgressBar1.Minimum = 0
                ProgressBar1.Maximum = sql_count
                ProgressBar1.Value = 0

                While data_reader.Read()
                    Label1.Text = "[" & ProgressBar1.Value & "/" & sql_count & "]"

                    ''TextBox1.Text += "SELECT "
                    'TextBox1.Text += vbCrLf
                    Dim sql_table_id As String = "SELECT COLUMN_ID FROM TABLE_INFO WHERE TABLE_ID='" & data_reader.Item("TABLE_ID") & "'"
                    Dim cmd_table_id As New OracleCommand(sql_table_id, conn)
                    Dim data_reader_table_id As OracleDataReader = cmd_table_id.ExecuteReader() ' VB.NET
                    Dim sql_column_id_temp As String = ""
                    Try
                        While data_reader_table_id.Read()
                            sql_column_id_temp += data_reader_table_id.Item("COLUMN_ID")
                            sql_column_id_temp += ", "
                            ''TextBox1.Text += data_reader_table_id.Item("COLUMN_ID")
                            ''TextBox1.Text += ", "
                        End While
                        sql_column_id_temp = sql_column_id_temp.Remove(sql_column_id_temp.Length - 2)
                    Catch ex As Exception
                        TextBox1.Text = ex.Message.ToString()
                    End Try
                    Dim sql_column_id As String = "SELECT "
                    sql_column_id += sql_column_id_temp
                    sql_column_id += " FROM " & data_reader.Item("TABLE_ID")
                    TextBox1.Text += sql_column_id
                    ''TextBox1.Text += "FROM '" & data_reader.Item("TABLE_ID") & "'"
                    TextBox1.Text += vbCrLf
                    ProgressBar1.Value += 1
                    Label1.Text = "[" & ProgressBar1.Value & "/" & sql_count & "]"
                End While
            Catch ex As Exception
                'Label1.Text = ex.Message.ToString()
                TextBox1.Text = ex.Message.ToString()
            Finally
                data_reader.Close()
            End Try
            sql_count = 0
        Catch ex As Exception
            'Label1.Text = ex.Message.ToString()
            TextBox1.Text = ex.Message.ToString()
        Finally
            conn.Dispose()
        End Try
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        TextBox1.Clear()
        Dim oradb As String = "Data Source=(DESCRIPTION=" _
+ "(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=${HOSTIP})(PORT=1521)))" _
+ "(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=ORCL)));" _
+ "User Id=SYSTEM;Password=1234567;"
        Dim conn As New OracleConnection(oradb)

        Try
            conn.Open() '연결
            Dim sql_count As Integer = 0

            'Dim sql As String = "SELECT * FROM TABLE_INFO"
            Dim sql As String = "SELECT TABLE_ID FROM TABLE_INFO GROUP BY TABLE_ID"
            'Dim sql As String = "SELECT * FROM TABLE_INFO where Table_Id='새올-154-01'
            Dim cmd As New OracleCommand(sql, conn)

            cmd.CommandType = CommandType.Text

            Dim data_reader_count As OracleDataReader = cmd.ExecuteReader()
            Try
                While data_reader_count.Read()
                    sql_count = sql_count + 1
                End While
            Catch ex As Exception
                Label1.Text = ex.Message.ToString()
                TextBox1.Text = ex.Message.ToString()
            Finally
                data_reader_count.Close()
            End Try

            Dim data_reader As OracleDataReader = cmd.ExecuteReader()
            Try
                ProgressBar1.Minimum = 0
                ProgressBar1.Maximum = sql_count
                ProgressBar1.Value = 0

                While data_reader.Read()
                    Label1.Text = "[" & ProgressBar1.Value & "/" & sql_count & "]"

                    ''TextBox1.Text += "SELECT "
                    'TextBox1.Text += vbCrLf
                    Dim sql_table_id As String = "SELECT COLUMN_ID, DATA_TYPE, DATA_LENGTH, IS_NULL FROM TABLE_INFO WHERE TABLE_ID='" & data_reader.Item("TABLE_ID") & "'"
                    Dim cmd_table_id As New OracleCommand(sql_table_id, conn)
                    Dim data_reader_table_id As OracleDataReader = cmd_table_id.ExecuteReader() ' VB.NET
                    Dim sql_table_context_temp As String = ""
                    Try
                        While data_reader_table_id.Read()
                            sql_table_context_temp += data_reader_table_id.Item("COLUMN_ID") & " "
                            sql_table_context_temp += data_reader_table_id.Item("DATA_TYPE") & "(" & data_reader_table_id.Item("DATA_LENGTH") & ") "
                            If (data_reader_table_id.Item("IS_NULL") = "NotNull") Then
                                sql_table_context_temp += "Not Null, "
                            ElseIf (data_reader_table_id.Item("IS_NULL") = "Null") Then
                                sql_table_context_temp += data_reader_table_id.Item("IS_NULL") & ", "
                            Else
                                sql_table_context_temp += "--exception--, "
                            End If
                            ''TextBox1.Text += data_reader_table_id.Item("COLUMN_ID")
                            ''TextBox1.Text += ", "
                        End While
                        sql_table_context_temp = sql_table_context_temp.Remove(sql_table_context_temp.Length - 2)
                    Catch ex As Exception
                        TextBox1.Text = ex.Message.ToString()
                    End Try
                    Dim sql_column_id As String = "CREATE TABLE " & data_reader.Item("TABLE_ID") & " ("
                    sql_column_id += sql_table_context_temp & ")"
                    TextBox1.Text += sql_column_id
                    ''TextBox1.Text += "FROM '" & data_reader.Item("TABLE_ID") & "'"
                    TextBox1.Text += vbCrLf
                    ProgressBar1.Value += 1
                    Label1.Text = "[" & ProgressBar1.Value & "/" & sql_count & "]"
                End While
            Catch ex As Exception
                'Label1.Text = ex.Message.ToString()
                TextBox1.Text = ex.Message.ToString()
            Finally
                data_reader.Close()
            End Try
            sql_count = 0
        Catch ex As Exception
            'Label1.Text = ex.Message.ToString()
            TextBox1.Text = ex.Message.ToString()
        Finally
            conn.Dispose()
        End Try
    End Sub
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        TextBox1.Clear()
        Dim oradb As String = "Data Source=(DESCRIPTION=" _
+ "(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=${HOSTIP})(PORT=1521)))" _
+ "(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=ORCL)));" _
+ "User Id=SYSTEM;Password=1234567;"
        Dim conn As New OracleConnection(oradb)

        Try
            conn.Open() '연결
            Dim sql_count As Integer = 0

            'Dim sql As String = "SELECT * FROM TABLE_INFO"
            Dim sql As String = "SELECT TABLE_ID FROM TABLE_INFO GROUP BY TABLE_ID"
            'Dim sql As String = "SELECT * FROM TABLE_INFO where Table_Id='새올-154-01'
            Dim cmd As New OracleCommand(sql, conn)

            cmd.CommandType = CommandType.Text

            Dim data_reader_count As OracleDataReader = cmd.ExecuteReader()
            Try
                While data_reader_count.Read()
                    sql_count = sql_count + 1
                End While
            Catch ex As Exception
                Label1.Text = ex.Message.ToString()
                TextBox1.Text = ex.Message.ToString()
            Finally
                data_reader_count.Close()
            End Try

            Dim data_reader As OracleDataReader = cmd.ExecuteReader()
            Try
                ProgressBar1.Minimum = 0
                ProgressBar1.Maximum = sql_count
                ProgressBar1.Value = 0

                While data_reader.Read()
                    Label1.Text = "[" & ProgressBar1.Value & "/" & sql_count & "]"

                    ''TextBox1.Text += "SELECT "
                    'TextBox1.Text += vbCrLf
                    Dim sql_table_id As String = "SELECT COLUMN_ID, DATA_TYPE, DATA_LENGTH, IS_NULL FROM TABLE_INFO WHERE TABLE_ID='" & data_reader.Item("TABLE_ID") & "'"
                    Dim cmd_table_id As New OracleCommand(sql_table_id, conn)
                    Dim data_reader_table_id As OracleDataReader = cmd_table_id.ExecuteReader() ' VB.NET
                    Dim sql_table_context_temp As String = ""
                    Try
                        While data_reader_table_id.Read()
                            sql_table_context_temp += data_reader_table_id.Item("COLUMN_ID") & " "
                            If data_reader_table_id.Item("DATA_TYPE") = "VARCHAR2" Then
                                sql_table_context_temp += "VARCHAR" & "(" & data_reader_table_id.Item("DATA_LENGTH") & ") "
                            Else
                                sql_table_context_temp += data_reader_table_id.Item("DATA_TYPE") & "(" & data_reader_table_id.Item("DATA_LENGTH") & ") "
                            End If
                            If (data_reader_table_id.Item("IS_NULL") = "NotNull") Then
                                sql_table_context_temp += "Not Null, "
                            ElseIf (data_reader_table_id.Item("IS_NULL") = "Null") Then
                                sql_table_context_temp += data_reader_table_id.Item("IS_NULL") & ", "
                            Else
                                sql_table_context_temp += "--exception--, "
                            End If
                            ''TextBox1.Text += data_reader_table_id.Item("COLUMN_ID")
                            ''TextBox1.Text += ", "
                        End While
                        sql_table_context_temp = sql_table_context_temp.Remove(sql_table_context_temp.Length - 2)
                    Catch ex As Exception
                        TextBox1.Text = ex.Message.ToString()
                    End Try
                    Dim sql_column_id As String = "CREATE TABLE " & data_reader.Item("TABLE_ID") & " ("
                    sql_column_id += sql_table_context_temp & ");"
                    TextBox1.Text += sql_column_id
                    ''TextBox1.Text += "FROM '" & data_reader.Item("TABLE_ID") & "'"
                    TextBox1.Text += vbCrLf
                    ProgressBar1.Value += 1
                    Label1.Text = "[" & ProgressBar1.Value & "/" & sql_count & "]"
                End While
            Catch ex As Exception
                'Label1.Text = ex.Message.ToString()
                TextBox1.Text = ex.Message.ToString()
            Finally
                data_reader.Close()
            End Try
            sql_count = 0
        Catch ex As Exception
            'Label1.Text = ex.Message.ToString()
            TextBox1.Text = ex.Message.ToString()
        Finally
            conn.Dispose()
        End Try
    End Sub
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        TextBox1.Clear()
        Dim oradb As String = "Data Source=(DESCRIPTION=" _
+ "(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=${HOSTIP})(PORT=1521)))" _
+ "(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=ORCL)));" _
+ "User Id=SYSTEM;Password=1234567;"
        Dim conn_get_colsname As New OracleConnection(oradb)
        Dim conn_get_value As New OracleConnection(oradb)
        Dim conn_get_colsname_temp As New OracleConnection(oradb)
        Dim sql_table_id_selected As String = "TABLE_INFO_184" '지정테이블 일단1테이블
        Dim sql_insert_head As String = "INSERT INTO " & sql_table_id_selected '지정테이블
        Dim sql_insert_tail As String = ""
        Dim sql_colsname As String = ""
        Dim sql_get_colsname As String = "SELECT COLUMN_NAME FROM COLS WHERE TABLE_NAME = '" & sql_table_id_selected & "'" '지정테이블컬럼읽기
        Dim sql_get_value = ""

        Try
            conn_get_colsname.Open() '연결
            Dim cmd As New OracleCommand(sql_get_colsname, conn_get_colsname)

            cmd.CommandType = CommandType.Text

            Dim data_reader As OracleDataReader = cmd.ExecuteReader()
            Try
                While data_reader.Read()
                    sql_colsname += data_reader.Item("COLUMN_NAME") & ", "
                End While
            Catch ex As Exception
                TextBox1.Text = ex.Message.ToString()
            Finally
                data_reader.Close()
            End Try
            sql_colsname = sql_colsname.Remove(sql_colsname.Length - 2)
        Catch ex As Exception
            TextBox1.Text = ex.Message.ToString()
        Finally
            conn_get_colsname.Dispose()
        End Try
        '이상, 오라클컬럼명긁어오기
        sql_get_value += "SELECT " & sql_colsname & " FROM " & sql_table_id_selected
        'TextBox1.Text = "SELECT " & sql_colsname & " FROM " & sql_table_id_selected + vbCrLf

        Try
            conn_get_value.Open() '연결
            Dim cmd As New OracleCommand(sql_get_value, conn_get_value)
            cmd.CommandType = CommandType.Text

            Dim data_reader As OracleDataReader = cmd.ExecuteReader()
            While data_reader.Read() '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Dim values_temp As String = ""
                Dim col_temp As String
                Try
                    conn_get_colsname_temp.Open() '연결
                    Dim cmd_temp As New OracleCommand(sql_get_colsname, conn_get_colsname_temp)

                    cmd_temp.CommandType = CommandType.Text

                    Dim data_reader_temp As OracleDataReader = cmd_temp.ExecuteReader()
                    Try
                        While data_reader_temp.Read()
                            col_temp = data_reader_temp.Item("COLUMN_NAME")
                            values_temp += "'" & data_reader.Item(col_temp) & "', "
                        End While
                    Catch ex As Exception
                        TextBox1.Text += ex.Message.ToString()
                    Finally
                        data_reader_temp.Close()
                    End Try
                Catch ex As Exception
                    TextBox1.Text += ex.Message.ToString()
                Finally
                    conn_get_colsname_temp.Close()
                End Try
                values_temp = values_temp.Remove(values_temp.Length - 2)

                TextBox1.Text += sql_insert_head & " (" & sql_colsname & ") VALUES (" & values_temp & ");" & vbCrLf

                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            End While

        Catch ex As Exception
            TextBox1.Text = ex.Message.ToString()
        Finally
            conn_get_value.Dispose()
        End Try

    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        '포스트그레스
        TextBox1.Clear()
        Dim postgresdb As String = "Data Source=(DESCRIPTION=" _
           + "(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=${HOSTIP})(PORT=5432)))" _
           + "(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=ORCL)));" _
           + "User Id=postgres;Password=geoserver;"
        Dim conn As New OracleConnection(postgresdb)

        Try
            conn.Open() '연결
            Dim sql_count As Integer = 0

            'Dim sql As String = "SELECT * FROM TABLE_INFO"
            Dim sql As String = "SELECT TABLE_ID FROM TABLE_INFO GROUP BY TABLE_ID"
            'Dim sql As String = "SELECT * FROM TABLE_INFO where Table_Id='새올-154-01'
            Dim cmd As New OracleCommand(sql, conn)

            cmd.CommandType = CommandType.Text

            Dim data_reader_count As OracleDataReader = cmd.ExecuteReader() ' VB.NET
            Try
                While data_reader_count.Read()
                    sql_count = sql_count + 1
                End While
            Catch ex As Exception
                Label1.Text = ex.Message.ToString()
                TextBox1.Text = ex.Message.ToString()
            Finally
                data_reader_count.Close()
            End Try

            Dim data_reader As OracleDataReader = cmd.ExecuteReader() ' VB.NET
            Try
                ProgressBar1.Minimum = 0
                ProgressBar1.Maximum = sql_count
                ProgressBar1.Value = 0

                While data_reader.Read()
                    TextBox1.Text += data_reader.Item("TABLE_ID") ' retrieve by column name
                    'TextBox1.Text += data_reader.Item("COLUMN_ID")
                    'TextBox1.Text += data_reader.Item("DATA_NAME")
                    'TextBox1.Text += data_reader.Item("DATA_TYPE")
                    'TextBox1.Text += data_reader.Item("DATA_LENGTH")
                    'TextBox1.Text += data_reader.Item("IS_NULL")
                    'TextBox1.Text += data_reader.Item("INFO")
                    TextBox1.Text += vbCrLf
                    ProgressBar1.Value += 1
                    Label1.Text = "[" & ProgressBar1.Value & "/" & sql_count & "]"
                End While
            Catch ex As Exception
                'Label1.Text = ex.Message.ToString()
                'TextBox1.Text = ex.Message.ToString()
            Finally
                data_reader.Close()
            End Try
            sql_count = 0
        Catch ex As Exception
            Label1.Text = ex.Message.ToString()
            TextBox1.Text = ex.Message.ToString()
        Finally
            conn.Dispose()
        End Try
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        ''Basedata.SetBaseData()
        'g_basedata.SetBaseData()
        'For Each g_basedata.g_Layer_info In g_basedata.g_Layer_info
        '    TextBox1.Text += "Select ?"
        '    'TextBox1.Text += g_basedata.g_Layer_info(Name).ToString
        '    'TextBox1.Text += g_basedata.GetLayerHGO2("A00_ROT_RODMANTCE")
        '    'TextBox1.Text += Basedata.GetLayerHGO2("A00_ROT_RODMANTCE") '아ㅇㅋㅇㅋ
        '    TextBox1.Text += vbCrLf
        'Next g_basedata.g_Layer_info

        TextBox1.Clear()
        Dim oradb As String = "Data Source=(DESCRIPTION=" _
   + "(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=${HOSTIP})(PORT=1521)))" _
   + "(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=ORCL)));" _
   + "User Id=SYSTEM;Password=1234567;"
        Dim conn As New OracleConnection(oradb)
        Dim conn_layer_id As New OracleConnection(oradb)

        Try
            conn.Open() '연결
            Dim sql_count As Integer = 0
            Dim sql As String = "SELECT LAYERINFO_ID FROM TABLE_INFO_HGO GROUP BY LAYERINFO_ID"
            Dim cmd As New OracleCommand(sql, conn)

            cmd.CommandType = CommandType.Text
            Dim data_reader_count As OracleDataReader = cmd.ExecuteReader()
            Try
                While data_reader_count.Read()
                    sql_count = sql_count + 1
                End While
            Catch ex As Exception
                Label1.Text += ex.Message.ToString()
                TextBox1.Text += ex.Message.ToString()
            Finally
                data_reader_count.Close()
            End Try

            Dim data_reader As OracleDataReader = cmd.ExecuteReader()
            Try
                ProgressBar1.Minimum = 0
                ProgressBar1.Maximum = sql_count
                ProgressBar1.Value = 0

                While data_reader.Read()
                    Dim sql_field_names As String = ""

                    Dim sql_field_nm As String = "select fieldinfo_name from table_info_hgo where layerinfo_id='" & data_reader.Item("layerinfo_id") & "'"
                    Dim cmd_field_nm As New OracleCommand(sql_field_nm, conn)
                    cmd_field_nm.CommandType = CommandType.Text
                    Dim data_reader_field_nm As OracleDataReader = cmd_field_nm.ExecuteReader() ' vb.net
                    Try
                        While data_reader_field_nm.Read()
                            sql_field_names += data_reader_field_nm("fieldinfo_name")
                            sql_field_names += ", "
                        End While
                    Catch ex As Exception
                        '.Text += ex.Message.ToString()
                        TextBox1.Text += ex.Message.ToString()
                    Finally
                        data_reader_field_nm.Close()
                    End Try
                    Dim sql_output As String = "SELECT " & sql_field_names & "ASTEXT(G2_SPATIAL) SHAPE FROM " & data_reader.Item("LAYERINFO_ID") & " WHERE HGO_LAST_MOD_YMD > ?;"
                    'TextBox1.Text += "SELECT " & sql_field_names & "ASTEXT(G2_SPATIAL) SHAPE FROM " & data_reader.Item("LAYERINFO_ID") & " WHERE HGO_LAST_MOD_YMD > ?;" & vbCrLf
                    TextBox1.Text += sql_output & vbCrLf

                    ''여기부터
                    Dim SaveAs_filePath As String = "C:\RDB\" & data_reader.Item("LAYERINFO_ID") & ".sql"
                    Dim fs As FileStream = File.Create(SaveAs_filePath)
                    Dim info As Byte() = New UTF8Encoding(True).GetBytes(sql_output)
                    fs.Write(info, 0, info.Length)
                    fs.Close()
                    ''여기까지

                    ProgressBar1.Value += 1
                    Label1.Text = "[" & ProgressBar1.Value & "/" & sql_count & "]"
                End While
            Catch ex As Exception
                TextBox1.Text += ex.Message.ToString()
            Finally
                TextBox1.Text += "작업완료"
                data_reader.Close()
            End Try
            sql_count = 0
        Catch ex As Exception
            TextBox1.Text += ex.Message.ToString()
        Finally
            conn.Dispose()
        End Try
    End Sub

    Private Sub Button4_MouseHover(sender As Object, e As EventArgs) Handles Button4.MouseHover
        Label_help.Text = "테이블정의서를 긁어서.. 보완필요"
    End Sub
    Private Sub Button4_MouseLeave(sender As Object, e As EventArgs) Handles Button4.MouseLeave
        Label_help.Text = ""
    End Sub
    Private Sub Button7_MouseHover(sender As Object, e As EventArgs) Handles Button7.MouseHover
        Label_help.Text = "여러개의 서비스정의서의 J열을 긁어 취합하여 엑셀로 저장. (C\RDB)"
    End Sub
    Private Sub Button7_MouseLeave(sender As Object, e As EventArgs) Handles Button7.MouseLeave
        Label_help.Text = ""
    End Sub
    Private Sub Button9_MouseHover(sender As Object, e As EventArgs) Handles Button9.MouseHover
        Label_help.Text = "오라클DB(TABLE_INFO)에 접속하여 테이블 속성을 긁어, SELECT 쿼리 작성"
    End Sub
    Private Sub Button9_MouseLeave(sender As Object, e As EventArgs) Handles Button9.MouseLeave
        Label_help.Text = ""
    End Sub

    Private Sub Button10_MouseHover(sender As Object, e As EventArgs) Handles Button10.MouseHover
        Label_help.Text = "오라클DB(TABLE_INFO)에 접속하여 테이블 속성을 긁어, CREATE 쿼리 작성. 오라클ver"
    End Sub
    Private Sub Button10_MouseLeave(sender As Object, e As EventArgs) Handles Button10.MouseLeave
        Label_help.Text = ""
    End Sub


    Private Sub Button11_MouseHover(sender As Object, e As EventArgs) Handles Button11.MouseHover
        Label_help.Text = "오라클DB(TABLE_INFO)에 접속하여 테이블 속성을 긁어, CREATE 쿼리 작성. PostGresql. 항목이 많으면 토함ㅠㅠ (ex 167-01, 313-01)"
    End Sub

    Private Sub Button11_MouseLeave(sender As Object, e As EventArgs) Handles Button11.MouseLeave
        Label_help.Text = ""
    End Sub

    Private Sub Button12_MouseHover(sender As Object, e As EventArgs) Handles Button12.MouseHover
        Label_help.Text = "오라클DB(TABLE_INFO)에 접속하여 데이터트 값을 긁어, INSERT 쿼리 작성. PostGresql"
    End Sub

    Private Sub Button12_MouseLeave(sender As Object, e As EventArgs) Handles Button12.MouseLeave
        Label_help.Text = ""
    End Sub
    Private Sub Button14_MouseHover(sender As Object, e As EventArgs) Handles Button14.MouseHover
        Label_help.Text = "오라클DB(TABLE_INFO_HGO:행정주제도)에 접속하여 쿼리 작성하여 *.sql파일로 각각 저장. (C\RDB)"
    End Sub

    Private Sub Button14_MouseLeave(sender As Object, e As EventArgs) Handles Button14.MouseLeave
        Label_help.Text = ""
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class