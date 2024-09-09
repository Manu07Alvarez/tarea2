﻿Imports Microsoft.Office.Interop

Imports System.Runtime.InteropServices

Public Class Form1
    Dim Excel As New Excel.Application
    Dim WorkBook As Excel.Workbook
    Dim WorkSheet As Excel.Worksheet
    Dim range As Excel.Range
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.Filter = "Archivos Excel(*.xlsx)|*.xlsx|Excel (97-2003) files(*.xls)|*.xls"
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            WorkBook = Excel.Workbooks.Open(OpenFileDialog1.FileName)
            For Each WorkSheet In Excel.Sheets
                ListBox1.Items.Add(WorkSheet.Name)
            Next
        End If
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If Not IsNothing(Excel) Then
            If Not IsNothing(WorkBook) Then
                Debug.WriteLine("a1")
                WorkBook.Close()
                Excel.Quit()
                Marshal.ReleaseComObject(WorkBook)
                If Not IsNothing(WorkSheet) Then Marshal.ReleaseComObject(WorkSheet)
                If Not IsNothing(range) Then Marshal.ReleaseComObject(range)
            End If
            Marshal.ReleaseComObject(Excel)
        End If
        Debug.WriteLine("se cerro")

    End Sub

    Private Function DataSetCreate()
        Dim table As New Data.DataTable
        Dim Row As DataRow
        Dim rowCount = range.Rows.Count
        Dim columnCount = range.Columns.Count
        For i = 1 To range.Columns.Count
            table.Columns.Add(range.Cells(1, i).Value2.ToString)
        Next

        Dim rowCounter As Integer
        For i = 2 To rowCount
            Row = table.NewRow()
            rowCounter = 0
            For j = 1 To columnCount
                If (Not IsNothing(range.Cells(i, j)) And Not IsNothing(range.Cells(i, j).value2)) Then
                    Row(rowCounter) = range.Cells(i, j).Value2.ToString
                Else
                    Row(i) = ""
                End If
                rowCounter += 1
            Next
            table.Rows.Add(Row)
        Next
        Return table
    End Function

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        WorkBook.Worksheets(ListBox1.SelectedItem).activate()
        WorkSheet = WorkBook.Worksheets(ListBox1.SelectedItem)
        Button2.Visible = True
        Debug.WriteLine(WorkSheet.Name)
        range = WorkSheet.UsedRange
        DataGridView1.DataSource = DataSetCreate()
    End Sub

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

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
