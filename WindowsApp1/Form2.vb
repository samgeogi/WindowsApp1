Imports System.IO
Imports ExcelDataReader

Public Class Form2
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim FileName = My.Application.Info.DirectoryPath & "\resources\Book 2 1.xlsx"
        Using stream = File.Open(FileName, FileMode.Open, FileAccess.Read)
            Using reader As IExcelDataReader = ExcelReaderFactory.CreateReader(stream)
                Dim result As DataSet = reader.AsDataSet(New ExcelDataSetConfiguration() With {
                                                             .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration() With {
                                                             .UseHeaderRow = True}})
                Dim tables = result.Tables

                For Each table As DataTable In tables
                    DataGridView1.DataSource = table
                Next
            End Using
        End Using
    End Sub
End Class