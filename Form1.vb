
Imports Microsoft.Office.Interop
Imports System.Data.OleDb

Public Class Form1
    Dim Excel As New Excel.Application
    Dim WorkBook As Excel.Workbook
    Dim WorkSheet As Excel.Worksheet
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.Filter = "Archivos Excel(*.xlsx)|*.xlsx|Excel (97-2003) files(*.xls)|*.xls"
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            Dim WorkBooks = Excel.Workbooks.Open(OpenFileDialog1.FileName)
            WorkBook = WorkBooks
            ListBox1.Items.Clear()
            For Each WorkSheet In Excel.Sheets
                ListBox1.Items.Add(WorkSheet.Name)
            Next
        End If
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If Not IsNothing(WorkBook) Then
            CloseWorkbook()
            WorkBook = Nothing
        End If

        Excel.Quit()
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        WorkSheet = WorkBook.Worksheets(ListBox1.SelectedItem)
        Button2.Visible = True
        DataGridView1.DataSource = DataSetCreate()
    End Sub

    Dim row As Integer
    Dim column As Integer
    Private Sub DataGridView1_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DataGridView1.CellBeginEdit
        column = DataGridView1.CurrentCell.ColumnIndex + 1
        row = DataGridView1.CurrentCell.RowIndex + 2
    End Sub
    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        WorkSheet.Cells(row, column) = DataGridView1.CurrentCell.Value
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        WorkBook.Save()
    End Sub

    Private Function DataSetCreate()
        Dim pathto = (WorkBook.Path + "\" + WorkBook.Name)
        Dim connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathto + ";Extended Properties='Excel 12.0 Xml;HDR=YES';"
        Dim connection = New OleDb.OleDbConnection(connString)
        connection.Open()
        Dim sheetTable As DataTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
        Dim dataSet As New Data.DataSet
        If sheetTable IsNot Nothing Then
            For Each row As DataRow In sheetTable.Rows
                Dim sheetName As String = row("TABLE_NAME").ToString()

                ' Query each sheet
                Dim query As String = $"SELECT * FROM [{sheetName}]"
                Using command As New OleDbCommand(query, connection)
                    Using adapter As New OleDbDataAdapter(command)
                        Dim dataTable As New DataTable(sheetName)
                        adapter.Fill(dataTable)
                        dataSet.Tables.Add(dataTable)
                    End Using
                End Using
            Next
        End If
        Return dataSet.Tables
    End Function

    Private Sub CloseWorkbook()
        If Not WorkBook.Saved Then
            Dim savestate = (MsgBox("Queres Guardar antes de salir?", vbYesNo) = vbYes)
            WorkBook.Close(SaveChanges:=savestate)
        End If
    End Sub

End Class
