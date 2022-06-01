Imports System.Data.OleDb
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Public Class Form1
#Region "variables"
    Dim con As New OleDbConnection
    Dim db_path As String = ""
    Dim con_string As String = ""
    Dim dt As New DataTable

#End Region

#Region "Functions"
    Function AddLog(ByVal s As String)
        Dim file_path As String = "Loger.txt"
        Dim fileExists As Boolean = File.Exists(file_path)
        Using sw As New StreamWriter(File.Open(file_path, FileMode.Append))
            sw.WriteLine("[" + DateTime.Now + "]=> " + s)
        End Using
    End Function
    Function gettabels()
        ComboBox1.Items.Clear()
        Using con As New OleDbConnection(con_string)
            con.Open()
            Dim dt As DataTable = con.GetSchema("TABLES", {Nothing, Nothing, Nothing, "TABLE"})
            For Each dr As DataRow In dt.Rows
                ComboBox1.Items.Add(dr("TABLE_NAME"))
            Next
        End Using

    End Function

    Function showitems(ByVal tabel_name As String)

        dt = New DataTable
        Dim da As New OleDbDataAdapter
        da = New OleDbDataAdapter("SELECT * FROM " + tabel_name, con)
        da.Fill(dt)
    End Function

    Function save_as_excel(ByVal dir_path As String)
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer
        Dim j As Integer

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")


        'FOR HEADERS
        For i = 1 To DataGridView1.ColumnCount
            xlWorkSheet.Cells(1, i) = DataGridView1.Columns(i - 1).HeaderText
            'FOR ITEMS
            For j = 1 To DataGridView1.RowCount
                xlWorkSheet.Cells(j + 1, i) = DataGridView1(i - 1, j - 1).Value.ToString()
            Next
        Next


        xlWorkSheet.SaveAs(dir_path)
        xlWorkBook.Close()
        xlApp.Quit()

        MsgBox("Saved Successfully To " + dir_path)
    End Function

    Function save_csv(ByVal dir_path As String)

        Dim i As Integer
        Dim j As Integer

        Using sw As New IO.StreamWriter(dir_path)
            For i = 1 To DataGridView1.ColumnCount
                If i = DataGridView1.ColumnCount Then
                    sw.Write(DataGridView1.Columns(i - 1).HeaderText)
                Else
                    sw.Write(DataGridView1.Columns(i - 1).HeaderText + ",")
                End If
            Next
            sw.WriteLine("")




            For i = 1 To DataGridView1.RowCount
                For j = 1 To DataGridView1.ColumnCount


                    If j = DataGridView1.ColumnCount Then
                        sw.Write(DataGridView1(j - 1, i - 1).Value.ToString())
                        sw.WriteLine("")
                    Else
                        sw.Write(DataGridView1(j - 1, i - 1).Value.ToString() + ",")
                    End If
                Next
            Next

        End Using

        MsgBox("Saved Successfully To " + dir_path)
    End Function

    Function gettabels_type(ByVal tabel_name As String, ByVal dir_path As String) As String
        Using con As New OleDbConnection(con_string)
            con.Open()
            Dim dt As DataTable = con.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, {Nothing, Nothing, tabel_name, Nothing})
            'Dim dt As DataTable = con.GetSchema("TABLES", {Nothing, Nothing, Nothing, "TABLE"})
            Using sw As New IO.StreamWriter(dir_path)
                For Each dr As DataRow In dt.Rows
                    sw.WriteLine(dr("COLUMN_NAME").ToString() + "|" + get_TYPE(CType(dr("DATA_TYPE"), OleDb.OleDbType).ToString))
                Next
            End Using
        End Using

    End Function

    Function get_TYPE(ByVal datatype As String) As String
        Dim val As String
        If datatype = "WChar" Then
            val = "Char"
        ElseIf datatype = "Integer" Then
            val = "INTEGER"
        ElseIf datatype = "Date" Then
            val = "Date"
        End If
        Return val

    End Function

    Function get_line_datatype(ByVal str As String, ByVal filepath As String) As String
        Using reader As New StreamReader(filepath)
            While Not reader.EndOfStream
                Dim line As String = reader.ReadLine()
                If line.Contains(str) Then
                    Return line
                End If
            End While
        End Using
    End Function

    Function create_tabel(ByVal tabel_name As String, ByVal csv_file As String, ByVal data_type_path_file As String)


        Dim lines() As String = IO.File.ReadAllLines(csv_file)
        Dim arr() As String = lines(0).Split(",")
        Dim query As String = ""

        Dim Q_forInsert As String = ""
        Dim Q_forInsert2 As String = ""
        Dim listofparams As List(Of String)
        listofparams = New List(Of String)

        For Each vaal In arr
            query += "[" + vaal + "] " + get_line_datatype(vaal, data_type_path_file).Split("|")(1) + ","
            Q_forInsert += "[" + vaal + "], "
            Q_forInsert2 += "@" + vaal + ", "
            listofparams.Add("@" + vaal)
        Next


        'remove last char
        query = query.Substring(0, query.Length - 1)


        Dim final_q As String = "CREATE TABLE " + tabel_name + " (" + query + ");"

        Using con As New OleDbConnection(con_string)

            con.Open()
            Using cmd As New OleDbCommand()
                cmd.Connection = con
                cmd.CommandText = final_q


                Try
                    cmd.ExecuteNonQuery()
                    MsgBox("Table created. will insert data now")

#Region "insert Data"


                    'remove last 2 char
                    Q_forInsert = Q_forInsert.Substring(0, Q_forInsert.Length - 2)
                    Q_forInsert2 = Q_forInsert2.Substring(0, Q_forInsert2.Length - 2)
                    Q_forInsert = "(" + Q_forInsert + ")"
                    Q_forInsert2 = "(" + Q_forInsert2 + ")"


                    Dim array_of_values() As String


                    For i = 1 To lines.Length - 1

                        'MsgBox(lines(i))


                        Dim qu As String = "INSERT INTO [" + tabel_name + "] " + Q_forInsert + " VALUES " + Q_forInsert2 + ";"

                        If i = 1 Then
                        Else
                            Array.Clear(array_of_values, 0, array_of_values.Length)
                        End If

                        array_of_values = lines(i).Split(",")
                        Using comd As New OleDb.OleDbCommand()
                            comd.Connection = con
                            comd.CommandText = qu


                            For ii = 0 To listofparams.Count() - 1
                                comd.Parameters.AddWithValue(listofparams(ii), array_of_values(ii))
                            Next



                            Try
                                comd.ExecuteNonQuery()

                            Catch ex As Exception
                                MsgBox(ex.ToString())
                            End Try
                        End Using
                    Next
                    MsgBox("Record Appended", MsgBoxStyle.Information, "Successfully Added!")

#End Region

                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End Using
            con.Close()

        End Using

    End Function



#End Region


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Btn_browse.Click
        OpenFileDialog1.Filter = "ACCESS DATABAES Files (*.accdb)|*.accdb"
        OpenFileDialog1.Title = "Open Your ACCESS DB File"
        OpenFileDialog1.RestoreDirectory = True
        OpenFileDialog1.FileName = "*.accdb"

retry:
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then

            db_path = OpenFileDialog1.FileName
        Else
            GoTo retry
        End If
        Label1.Text = "Database :  " + db_path

        con_string = "Provider=Microsoft.ACE.OLEDB.12.0;data source=" + db_path + ";"
        con.ConnectionString = con_string
        con.Open()
        gettabels()
        con.Close()
        AddLog("Database " + db_path + " Selected")
    End Sub



    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        showitems(ComboBox1.SelectedItem.ToString)
        DataGridView1.DataSource = dt
    End Sub

    Private Sub btn_save_excel_Click(sender As Object, e As EventArgs) Handles btn_save_excel.Click
        If DataGridView1.Rows.Count = 0 Then
            MsgBox("There is no Data in DataGridView")
            Return
        End If


        Dim prefared_name = ComboBox1.SelectedItem.ToString

        Dim saveFileDialog1 As New SaveFileDialog()
        Dim filepath As String = ""
        saveFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx"
        saveFileDialog1.Title = "Save Your Excel File to"
        saveFileDialog1.RestoreDirectory = True
        saveFileDialog1.FileName = prefared_name + ".xlsx"
retry:
        If saveFileDialog1.ShowDialog() = DialogResult.OK Then

            filepath = saveFileDialog1.FileName
        Else
            GoTo retry
        End If



        save_as_excel(filepath)
        AddLog("Excel File Exported To" + filepath)
    End Sub


    Private Sub Btn_save_csv_Click(sender As Object, e As EventArgs) Handles Btn_save_csv.Click
        If DataGridView1.Rows.Count = 0 Then
            MsgBox("There is no Data in DataGridView")
            Return
        End If


        Dim prefared_name = ComboBox1.SelectedItem.ToString

        Dim saveFileDialog1 As New SaveFileDialog()
        Dim filepath As String = ""
        saveFileDialog1.Filter = "CSV files (*.csv)|*.csv"
        saveFileDialog1.Title = "Save Your CSV File to"
        saveFileDialog1.RestoreDirectory = True
        saveFileDialog1.FileName = prefared_name + ".csv"
retry:
        If saveFileDialog1.ShowDialog() = DialogResult.OK Then

            filepath = saveFileDialog1.FileName
        Else
            GoTo retry
        End If

        Dim types_path As String = saveFileDialog1.FileName.Replace(".csv", "_type.txt")
        save_csv(filepath)
        gettabels_type(prefared_name, types_path)
        AddLog("Csv File Exported To" + filepath)
        AddLog("Csv_type File Exported To" + types_path)
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles btn_import.Click
        Me.WindowState = FormWindowState.Minimized
        Form2.ShowDialog()
        Me.WindowState = FormWindowState.Normal
    End Sub



#Region "datagridview custom hide the first <"
    Private Sub DataGridView1_RowPrePaint_1(sender As Object, e As DataGridViewRowPrePaintEventArgs) Handles DataGridView1.RowPrePaint
        e.PaintCells(e.ClipBounds, DataGridViewPaintParts.All)
        e.PaintHeader(DataGridViewPaintParts.Background Or DataGridViewPaintParts.Border Or DataGridViewPaintParts.Focus Or DataGridViewPaintParts.SelectionBackground Or DataGridViewPaintParts.ContentForeground)
        e.Handled = True
    End Sub

    Private Sub DataGridView1_CellFormatting_1(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        Me.DataGridView1.Rows(e.RowIndex).HeaderCell.Value = (e.RowIndex + 1).ToString()
    End Sub
#End Region

End Class
