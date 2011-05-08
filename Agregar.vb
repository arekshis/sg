Imports System.Data.SqlClient




Public Class Agregar
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then
            MsgBox("debe ingresar un nombre", MsgBoxStyle.Information, "Error")
            TextBox1.Focus()
            Exit Sub
        End If

        Dim strSql As String
        Dim db As New basededatos

        '"insert into equipos (nombre) values ('notebook')"
        strSql = "insert into equipos (nombre) values ('" & TextBox1.Text & "')"
        'db.SetSql(strSql)
        db.EjecutaSQLServer(strSql)
        MsgBox("Se ha ingresado correctamente", MsgBoxStyle.Information, Principal.Text)
        TextBox1.Text = ""
        cargar_grilla()

    End Sub

    Private Sub Agregar_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cargar_grilla()
    End Sub

    Private Sub cargar_grilla()
        Dim db As New basededatos
        Dim txtSql As String
        txtSql = "Select * from equipos"

        'DataGridView1.DataSource = db.getData(txtSql)

        Dim ds As New DataSet
        ds = db.ObtieneDatosSQLServer(txtSql, "datos")
        DataGridView1.DataSource = ds.Tables("datos")
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim txtSql As String

        Dim columna, fila As Integer
        Dim db As New basededatos

        columna = 0
        fila = DataGridView1.CurrentCell.RowIndex
        txtSql = "delete from equipos where id=" & DataGridView1.Item(columna, fila).Value.ToString
        db.setSql(txtSql)
        MsgBox("Se a Eliminado correctamente", MsgBoxStyle.Information)
        cargar_grilla()
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class
