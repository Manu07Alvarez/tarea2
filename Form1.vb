
Imports Microsoft.Office.Interop
Imports System.Runtime

Public Class Form1
    Dim Excel As New Excel.Application
    Dim WorkBook As Excel.Workbook
    Dim WorkSheet As Excel.Worksheet
    Dim cache As Caching.ObjectCache = Caching.MemoryCache.Default
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Not IsNothing(WorkBook) Then
            CloseWorkbook()
        End If
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

    Private Sub CloseWorkbook()
        Dim savestate = (MsgBox("Queres Guardar antes de salir?", vbYesNo) = vbYes)
        WorkBook.Close(SaveChanges:=savestate)
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
        If IsNothing(cache(WorkSheet.Name)) Then
            DataGridView1.DataSource = DataSetCreate()
            Debug.WriteLine("C1")
        Else
            DataGridView1.DataSource = cache(WorkSheet.Name)
        End If
    End Sub

    Private Function DataSetCreate()
        Dim pathto = (WorkBook.Path + "\" + WorkBook.Name)
        Dim connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathto + ";Extended Properties='Excel 12.0 Xml;HDR=YES';"
        Dim conn = New OleDb.OleDbConnection(connString)
        Dim sqlstr = "Select * from " + "[" + WorkSheet.Name + "$" + "]"
        Dim command = New OleDb.OleDbDataAdapter(sqlstr, conn)
        Dim table As New Data.DataSet
        command.Fill(table)
        CachingSheet(table)
        Return table.Tables(0)
    End Function

    Private Sub CachingSheet(ByVal table As Data.DataSet)
        Dim policy = New Caching.CacheItemPolicy
        cache.Set(WorkSheet.Name, table.Tables(0), policy)

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


End Class
