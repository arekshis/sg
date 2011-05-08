Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlServerCe
Public Class basededatos
    Dim ConexionSQLServer As String = "Data Source=ALEX-FA92C6D960\SQLEXPRESS;Initial Catalog=peluca;Integrated Security=True"
    Public Sub setSql(ByVal sqlStr As String)
        Dim conn As New SqlCeConnection
        Dim cmd As New SqlCeCommand
        Dim sExe As String = System.Reflection.Assembly.GetEntryAssembly().Location
        Dim sPathDir As String = System.IO.Path.GetDirectoryName(sExe)

        Try
            conn.ConnectionString = "datasource = " & sPathDir & "\Database1.sdf"
            cmd.Connection = conn
            cmd.CommandText = sqlStr
            conn.Open()
            cmd.ExecuteNonQuery()
            conn.Close()

        Catch ex As Exception
            MsgBox("Error en la comunicación con el archivo Database1.sdf" + vbCrLf + ex.Message.ToString, MsgBoxStyle.Critical, Principal.Text)
        End Try
    End Sub

    Public Function getData(ByVal sqlStr As String) As DataTable
        Dim conn As New SqlCeConnection
        Dim cmd As New SqlCeCommand
        Dim dt As New DataTable
        Dim da As New SqlCeDataAdapter(cmd)
        Dim sExe As String = System.Reflection.Assembly.GetEntryAssembly().Location
        Dim sPathDir As String = System.IO.Path.GetDirectoryName(sExe)

        Try
            conn.ConnectionString = "datasource = " & sPathDir & "\Database1.sdf"
            cmd.Connection = conn
            cmd.CommandText = sqlStr
            conn.Open()
            da.Fill(dt)
            conn.Close()

            Return dt

        Catch ex As Exception
            MsgBox("Error en la comunicación con el archivo bus.sdf" + vbCrLf + ex.Message.ToString, MsgBoxStyle.Critical, "titulo")
            Return dt
        End Try
    End Function

    Public Function ObtieneDatosSQLServer(ByVal strSql As String, ByVal strTable As String) As System.Data.DataSet
        Dim ds As New System.Data.DataSet()
        Try
            Dim MyConnection As New SqlConnection(ConexionSQLServer)
            Dim MyComman As New SqlDataAdapter(strSql, MyConnection)

            MyComman.Fill(ds, strTable)
            MyConnection.Close() 'importante
            Return ds


        Catch ex As Exception
            Return ds
        End Try


    End Function

    Public Sub EjecutaSQLServer(ByVal strSql As String)
        Dim MyConnection As New SqlConnection(ConexionSQLServer)
        Dim MyCommand As New SqlCommand()
        MyConnection.Open()
        MyCommand.Connection = MyConnection
        MyCommand.CommandText = strSql
        MyCommand.ExecuteNonQuery()
        MyConnection.Close()
    End Sub

End Class
