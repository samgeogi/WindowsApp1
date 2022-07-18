Imports System.IO
Imports ExcelDataReader
Imports OfficeOpenXml
Imports OfficeOpenXml.Style

Public Class FrmMain

    Dim Atrbt = "None"
    Dim ID
    Private Sub FrmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Panel3_Paint(sender As Object, e As PaintEventArgs) Handles Panel3.Paint

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        ID = TextBox1.Text
        Debug.Print(ID)
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress



        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If

    End Sub

    Private Sub Atr1_Click(sender As Object, e As EventArgs) Handles Atr1.Click
        Atrbt = Atr1.Text

    End Sub
    Private Sub Atr2_Click(sender As Object, e As EventArgs) Handles Atr2.Click
        Atrbt = Atr2.Text

    End Sub
    Private Sub Atr3_Click(sender As Object, e As EventArgs) Handles Atr3.Click
        Atrbt = Atr3.Text

    End Sub
    Private Sub Atr4_Click(sender As Object, e As EventArgs) Handles Atr4.Click
        Atrbt = Atr4.Text

    End Sub
    Private Sub Atr5_Click(sender As Object, e As EventArgs) Handles Atr5.Click
        Atrbt = Atr5.Text

    End Sub
    Private Sub Atr6_Click(sender As Object, e As EventArgs) Handles Atr6.Click
        Atrbt = Atr6.Text

    End Sub
    Private Sub Atr7_Click(sender As Object, e As EventArgs) Handles Atr7.Click
        Atrbt = Atr7.Text

    End Sub
    Private Sub Atr8_Click(sender As Object, e As EventArgs) Handles Atr8.Click
        Atrbt = Atr8.Text

    End Sub
    Private Sub svbtn_Click(sender As Object, e As EventArgs) Handles svbtn.Click
        Debug.Print(ID)
        Debug.Print(Atrbt)
        If ID.Length = 4 And Atrbt <> "None" Then

            WriteExcel(ID, Atrbt)

        Else
            Debug.Print("Error")
            MessageBox.Show("The ID should be 4 digits and an attribute should be selected for saving")
        End If

    End Sub









    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Dim result(7)
        Dim i = 0
        Dim reader As IExcelDataReader
        Dim stream = File.Open(My.Application.Info.DirectoryPath & "\resources\Book 1.xlsx", FileMode.Open, FileAccess.Read)
        reader = ExcelReaderFactory.CreateReader(stream)
        While reader.Read()

            result(i) = reader.GetValue(0).ToString
            Debug.Print(result(i))

            i = i + 1
        End While


        reader.Close()
        Atr1.Text = result(0)
        Atr2.Text = result(1)
        Atr3.Text = result(2)
        Atr4.Text = result(3)
        Atr5.Text = result(4)
        Atr6.Text = result(5)
        Atr7.Text = result(6)
        Atr8.Text = result(7)




    End Sub



    Function WriteExcel(ByVal IDe As Integer, atrb As String) As Double
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial
        Dim path = My.Application.Info.DirectoryPath & "\resources\Book 2 1.xlsx"
        Using package = New ExcelPackage(path)


            Dim worksheet As ExcelWorksheet
            worksheet = package.Workbook.Worksheets.FirstOrDefault()


            Dim colCount = worksheet.Dimension.End.Column  'get Column Count
            Dim rowCount = worksheet.Dimension.End.Row     'get row count
            worksheet.Cells(rowCount + 1, 1).Value = IDe
            worksheet.Cells(rowCount + 1, 2).Value = atrb
            Debug.Print(rowCount)
            package.Save()
        End Using
    End Function

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Dim oForm As New Form2

        oForm.Show()
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Me.Close()
    End Sub
End Class
