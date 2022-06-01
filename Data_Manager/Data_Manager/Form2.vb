Imports System.Data.OleDb
Imports System.IO
Public Class Form2
#Region "variables"
    Dim dt_csv As New DataTable
    Dim con As New OleDbConnection
    Dim db_path As String = ""
    Dim con_string As String = ""
    Dim dt As New DataTable
#End Region

#Region "Functions"
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

    Function insert_data(ByVal tabel_name As String, ByVal column As String, ByVal values As List(Of String))

        Dim query As String = "INSERT INTO [" + tabel_name + "] ([" + column + "]) VALUES ([@" + column + "])"
        Try
            For Each value In values
                Using conn As New OleDbConnection(con_string),
                    cmd As New OleDbCommand(query, conn)
                    conn.Open()
                    cmd.Parameters.AddWithValue("@" + column, value)
                    cmd.ExecuteNonQuery()

                End Using
            Next
            MsgBox("Record Appended", MsgBoxStyle.Information, "Successfully Added!")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Function



    Function update_data(ByVal tabel_name As String, ByVal column As String, ByVal values As List(Of String))
        Dim query As String

        If CheckBox1.Checked Then
            query = "UPDATE (SELECT TOP 1 [" + column + "] from [" + tabel_name + "] where [" + column + "] IS NULL Or " + column + "=0) As a SET a." + column + " = [@" + column + "] "
        Else
            query = "UPDATE (SELECT TOP 1 [" + column + "] from [" + tabel_name + "] where [" + column + "] IS NULL) As a SET a." + column + " = [@" + column + "] "
        End If


        For Each value In values
            Try

                Using conn As New OleDbConnection(con_string),
                cmd As New OleDbCommand(query, conn)
                    conn.Open()
                    cmd.Parameters.AddWithValue("@" + column, value)
                    cmd.ExecuteNonQuery()
                    conn.Close()
                End Using
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try


        Next
        MsgBox("Record Appended", MsgBoxStyle.Information, "Successfully Added!")

    End Function


    ''remove last 2 char
    'Dim query As String = "INSERT INTO [" + tabel_name + "] ([" + column + "]) VALUES ([@" + column + "])"
    ''UPDATE [TABLE1] SET [datefield] = '4/28/1949' where [datefield] = null
    ''Dim query As String = "UPDATE [" + tabel_name + "] SET [" + column + "] = [@" + column + "] where [" + column + "] IS NULL"
    'Try
    'Using conn As New OleDbConnection(con_string),
    '      cmd As New OleDbCommand(query, conn)
    '            conn.Open()
    '            'cmd.Parameters.Add("@" + column, OleDbType.WChar)
    '            'cmd.Parameters.Add("@" + column, OleDbType.Variant)
    '            For Each value In values
    '                ' cmd.Parameters("@" + column).Value = value
    '                cmd.Parameters.AddWithValue("@" + column, value)
    '                cmd.ExecuteNonQuery()
    '            Next
    'End Using
    '        MsgBox("Record Appended", MsgBoxStyle.Information, "Successfully Added!")
    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try


#End Region
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Normal
        Dim csv_file As String = ""

        OpenFileDialog1.Filter = "CSV Files (*.csv)|*.csv"
        OpenFileDialog1.Title = "Open Your CSV File"
        OpenFileDialog1.RestoreDirectory = True
        OpenFileDialog1.FileName = "*.csv"

retry:
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then

            csv_file = OpenFileDialog1.FileName
        Else
            GoTo retry
        End If


        Dim CnStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Path.GetDirectoryName(csv_file) & ";Extended Properties='text;HDR=Yes;FMT=Delimited(,)';"

        Using adapter As New OleDbDataAdapter("select * from " + Path.GetFileName(OpenFileDialog1.FileName) + ";", CnStr)
            adapter.Fill(dt_csv)
        End Using

        DataGridView1.DataSource = dt_csv


        For Each c As DataGridViewColumn In DataGridView1.Columns
            c.SortMode = DataGridViewColumnSortMode.NotSortable
            c.Selected = False
        Next

        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullColumnSelect
        DataGridView1.Columns(0).Selected = True







    End Sub



    Private Sub Btn_browse_Click(sender As Object, e As EventArgs) Handles Btn_browse.Click
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
        con_string = "Provider=Microsoft.ACE.OLEDB.12.0;data source=" + db_path + ";"
        con.ConnectionString = con_string
        con.Open()
        gettabels()
        con.Close()


    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        DataGridView2.SelectionMode = DataGridViewSelectionMode.CellSelect

        dt = New DataTable()
        showitems(ComboBox1.SelectedItem.ToString)
        DataGridView2.DataSource = dt

        For Each c As DataGridViewColumn In DataGridView2.Columns
            c.SortMode = DataGridViewColumnSortMode.NotSortable
            c.Selected = False
        Next

        DataGridView2.SelectionMode = DataGridViewSelectionMode.FullColumnSelect
        DataGridView2.Columns(0).Selected = True
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim columndata As List(Of String)
        columndata = New List(Of String)

        For Each row As DataRow In dt_csv.Rows
            columndata.Add(row(DataGridView1.Columns(DataGridView1.CurrentCell.ColumnIndex).HeaderText))
        Next

        Dim tabel_name As String = ComboBox1.SelectedItem.ToString
        Dim col As String = DataGridView2.Columns(DataGridView2.CurrentCell.ColumnIndex).HeaderText
        insert_data(tabel_name, col, columndata)

        gettabels()
        ComboBox1.SelectedIndex = ComboBox1.FindStringExact(tabel_name)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim columndata As List(Of String)
        columndata = New List(Of String)

        For Each row As DataRow In dt_csv.Rows
            columndata.Add(row(DataGridView1.Columns(DataGridView1.CurrentCell.ColumnIndex).HeaderText))
        Next

        Dim tabel_name As String = ComboBox1.SelectedItem.ToString
        Dim col As String = DataGridView2.Columns(DataGridView2.CurrentCell.ColumnIndex).HeaderText
        update_data(tabel_name, col, columndata)
        gettabels()
        ComboBox1.SelectedIndex = ComboBox1.FindStringExact(tabel_name)
    End Sub
End Class