Imports Inventor
Imports System.Runtime.InteropServices

Public Class BulkPDF_Form

    Public Property DocumentList As ArrayList
    Private _OverwriteAll As Boolean
    Public Property OverwriteAll As Boolean
        Get
            Return _OverwriteAll
        End Get
        Private Set(value As Boolean)
            _OverwriteAll = value
        End Set
    End Property

    Private Sub BulkPDF_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        OverwriteAll = CheckBox1.Checked
        For Each D As DrawingDocument In DocumentList
            DataGridView1.Rows.Add(D.DisplayName)
        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OverwriteAll = CheckBox1.Checked
    End Sub

    Public Function getFilenameOf(ByRef Drawing As DrawingDocument) As String

        If CheckBox2.Checked Then
            Return ""
        End If



    End Function

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged, TextBox2.TextChanged, TextBox3.TextChanged, TextBox4.TextChanged
        Try
            TextBox4.Text = getName(TextBox1.Text, Int(TextBox2.Text), Int(TextBox3.TextLength))
        Catch ex As Exception
            TextBox4.Text = "-"
        End Try

    End Sub

    Private Function getName(initial As String, index As Integer, padLength As Integer) As String

        Dim ret As String = initial
        ret = ret & index.ToString.PadLeft(padLength, "0"c) & ".pdf"
        Return ret

    End Function


End Class