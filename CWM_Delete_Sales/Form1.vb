Imports System.Data.SqlClient

Public Class Form1

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        If CheckBox1.Checked Then
            Dim com21 As New SqlCommand("delete from invoiceHeader  WHERE CompanyCode='2000' and divcode='01' and RepID='" & TextBox2.Text & "'  AND DailySalesID=" & TextBox3.Text & " AND StockiestID='" & TextBox1.Text & "'  ", SFAHeadoffice)
            Dim a As Integer = com21.ExecuteNonQuery()
            Dim com211 As New SqlCommand("delete from invoiceDetails   WHERE CompanyCode='2000' and divcode='01' and RepID='" & TextBox2.Text & "'  AND DailySalesID=" & TextBox3.Text & " AND StockiestID='" & TextBox1.Text & "'  ", SFAHeadoffice)
            Dim a1 As Integer = com211.ExecuteNonQuery()
            Dim com2111 As New SqlCommand("delete from SalesSummary WHERE CompanyCode='2000' and divcode='01' and RepID='" & TextBox2.Text & "'  AND DailySalesID=" & TextBox3.Text & "  AND DisCode='" & TextBox1.Text & "'", SFAHeadoffice)
            Dim a2 As Integer = com2111.ExecuteNonQuery()

            MessageBox.Show("Web Server -  " & a & " / " & a1 & " / " & a2, "Delete Sales", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)

        End If

        If CheckBox2.Checked Then
            Dim com21 As New SqlCommand("delete from invoiceHeader  WHERE CompanyCode='2000' and divcode='01' and RepID='" & TextBox2.Text & "'  AND DailySalesID=" & TextBox3.Text & " AND StockiestID='" & TextBox1.Text & "'  ", SFAHeadofficeLoal)
            Dim a As Integer = com21.ExecuteNonQuery()
            Dim com211 As New SqlCommand("delete from invoiceDetails   WHERE CompanyCode='2000' and divcode='01' and RepID='" & TextBox2.Text & "'  AND DailySalesID=" & TextBox3.Text & " AND StockiestID='" & TextBox1.Text & "'  ", SFAHeadofficeLoal)
            Dim a1 As Integer = com211.ExecuteNonQuery()
            Dim com2111 As New SqlCommand("delete from SalesSummary WHERE CompanyCode='2000' and divcode='01' and RepID='" & TextBox2.Text & "'  AND DailySalesID=" & TextBox3.Text & "  AND DisCode='" & TextBox1.Text & "'", SFAHeadofficeLoal)
            Dim a2 As Integer = com2111.ExecuteNonQuery()

            MessageBox.Show("Remote Server -  " & a & " / " & a1 & " / " & a2, "Delete Sales", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)

        End If



    End Sub

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        If Today > "2018-03-06" Then
            Application.Exit()
        End If

        SFAHeadoffice.Open()
        SFAHeadofficeLoal.Open()

    End Sub
End Class
