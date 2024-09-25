Imports Microsoft.Office.Interop
Imports System.Runtime
Imports System.Security.Cryptography
Imports System.Text
Imports System.Threading
Public Class Form1
    Dim WithEvents Excel As New Excel.Application
    Dim WorkBook As Excel.Workbook
    Dim WorkSheets As Excel.Worksheets
    Dim WorkSheet As Excel.Worksheet
    Dim cacheWorkbook As Caching.ObjectCache = Caching.MemoryCache.Default
    Dim currentDataGridHash As Dictionary(Of Integer, String) ' Hash del DataGridView

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.Filter = "Archivos Excel(*.xlsx)|*.xlsx|Excel (97-2003) files(*.xls)|*.xls"
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            If Not IsNothing(WorkBook) Then
                CloseWorkbook()
                For Each i In cacheWorkbook
                    cacheWorkbook.Remove(i.Key)
                Next
            End If
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
        If IsNothing(cacheWorkbook(WorkSheet.Name)) Then
            DataGridView1.DataSource = DataSetCreate(WorkSheet)
            currentDataGridHash = CreateHashDictionary(DataGridView1, WorkSheet.Name + "Hash")
            Debug.WriteLine("C1")
        Else
            DataGridView1.DataSource = cacheWorkbook(WorkSheet.Name)
            currentDataGridHash = cacheWorkbook(WorkSheet.Name + "Hash")
        End If
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

    Private Function DataSetCreate(ByVal worksheet As Excel.Worksheet)
        Dim pathto = (WorkBook.Path + "\" + WorkBook.Name)
        Dim connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathto + ";Extended Properties='Excel 12.0 Xml;HDR=YES';"
        Using conn = New OleDb.OleDbConnection(connString)
            Dim sqlstr = "Select * from " + "[" + worksheet.Name + "$" + "]"
            Dim command = New OleDb.OleDbDataAdapter(sqlstr, conn)
            Dim table As New Data.DataSet
            command.Fill(table)
            CachingF(table.Tables(0), worksheet.Name)
            Return table.Tables(0)
        End Using
    End Function

    Private Sub CachingF(value, name)
        Dim policy = New Caching.CacheItemPolicy
        cacheWorkbook.Set(name, value, policy)
    End Sub

    Private Sub CloseWorkbook()
        If Not WorkBook.Saved Then
            Dim savestate = (MsgBox("Queres Guardar antes de salir?", vbYesNo) = vbYes)
            WorkBook.Close(SaveChanges:=savestate)
        End If
    End Sub

    Private Function GetRowHash(row As Object) As String
        ' Serializar el array de objetos de la fila
        Dim rawData As String = String.Join(",", row)
        Using md5 As MD5 = MD5.Create()
            Dim bytes = Encoding.UTF8.GetBytes(rawData)
            Dim hashBytes = md5.ComputeHash(bytes)
            Return BitConverter.ToString(hashBytes).Replace("-", "").ToLower()
        End Using
    End Function

    Private Sub CompareTablesAndUpdate(oldTableHashes As Dictionary(Of Integer, String), newTable As Data.DataTable)
        ' Comparar con la nueva tabla
        For rowIndex As Integer = 0 To newTable.Rows.Count - 1
            Dim newRowHash As String = GetRowHash(newTable.Rows(rowIndex))

            ' Verificar si la fila ha cambiado comparando el hash
            If oldTableHashes.ContainsKey(rowIndex) Then
                If oldTableHashes(rowIndex) <> newRowHash Then
                    ' Actualizar solo las celdas de la fila que ha cambiado
                    For columnIndex As Integer = 0 To newTable.Columns.Count - 1
                        DataGridView1.Rows(rowIndex).Cells(columnIndex).Value = newTable.Rows(rowIndex)(columnIndex)
                    Next
                End If
            End If
        Next
    End Sub

    Private Function CreateHashDictionary(dgv As DataGridView, name As String) As Dictionary(Of Integer, String)
        Dim hashDict As New Dictionary(Of Integer, String)
        For rowIndex As Integer = 0 To dgv.Rows.Count - 1
            Dim rowHash As String = GetRowHash(dgv.Rows(rowIndex).Cells.Cast(Of DataGridViewCell).Select(Function(cell) cell.Value).ToArray())
            hashDict(rowIndex) = rowHash
        Next
        CachingF(hashDict, name)
        Return hashDict
    End Function

    Private Async Sub WorkBookSheetChangesAsync(sh As Object) Handles Excel.SheetCalculate
        Dim worksheetnew As Excel.Worksheet = TryCast(sh, Excel.Worksheet)
        Dim newTable As Data.DataTable = Await Task.Run(DataSetCreate(worksheetnew))

        If WorkSheet.Name = worksheetnew.Name Then
            ' Comparar solo si estamos en la tabla actual
            CompareTablesAndUpdate(currentDataGridHash, newTable)
        End If
    End Sub

End Class