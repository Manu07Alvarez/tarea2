Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Imports System.Runtime.InteropServices

Public Class Form1
    Dim Excel As New Excel.Application
    Dim WorkBook As Excel.Workbook
    Dim rech As Excel.Research
    Dim WorkSheet As Excel.Worksheet
    Dim range As Excel.Range
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.Filter = "Archivos Excel(*.xlsx)|*.xlsx|Excel (97-2003) files(*.xls)|*.xls"
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            Dim WorkBooks = Excel.Workbooks.Open(OpenFileDialog1.FileName)
            WorkBook = WorkBooks
            For Each WorkSheet In Excel.Sheets
                ListBox1.Items.Add(WorkSheet.Name)
            Next
        End If
    End Sub



    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If Not IsNothing(WorkBook) Then
            Debug.WriteLine("a1")
            WorkBook.Close()
            WorkBook = Nothing
        End If
        Excel.Quit()
    End Sub


    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        WorkBook.Worksheets(ListBox1.SelectedItem).activate()
        WorkSheet = WorkBook.Worksheets(ListBox1.SelectedItem)
        Button2.Visible = True
        range = WorkSheet.UsedRange
        DataGridView1.DataSource = DataSetCreate()
    End Sub

    Private Function DataSetCreate()
        Dim pathto = (WorkBook.Path + "\" + WorkBook.Name)
        Dim connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathto + ";Extended Properties='Excel 12.0 Xml;HDR=YES';"
        Dim conn = New OleDb.OleDbConnection(connString)
        Dim sqlstr = "Select * from " + "[" + WorkSheet.Name + "$" + "]"
        Dim command = New OleDb.OleDbDataAdapter(sqlstr, conn)
        Dim table As New Data.DataSet
        command.Fill(table)
        Return table.Tables(0)
    End Function

    Dim row As Integer
    Dim column As Integer
    Private Sub DataGridView1_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DataGridView1.CellBeginEdit
        column = DataGridView1.CurrentCell.ColumnIndex + 1
        row = DataGridView1.CurrentCell.RowIndex + 2
        Debug.WriteLine(DataGridView1.CurrentCell.Value)
    End Sub
    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        WorkSheet.Cells(row, column) = DataGridView1.CurrentCell.Value
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        WorkBook.Save()
    End Sub


End Class
