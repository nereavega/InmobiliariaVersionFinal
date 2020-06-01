Imports System.Data.SQLite

Public Class Balance
    Private Const cs As String = "Data Source = C:\Users\nerea\Desktop\database.db;Version=3"
    Dim respuesta As MsgBoxResult
    Dim dt As New DataTable
    Dim dt2 As New DataTable

    Dim myconnection As New SQLiteConnection(cs)
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Carga las tablas
        LoadViviendasRecords()
        LoadContratosRecords()
    End Sub

    Public Sub LoadViviendasRecords()

        myconnection.Open()

        Dim cmd As New SQLiteCommand
        cmd.Connection = myconnection

        cmd.CommandText = "SELECT* FROM Viviendas WHERE enVenta='VENDIDO' ORDER BY id DESC"

        Dim rdr As SQLite.SQLiteDataReader = cmd.ExecuteReader

        dt.Load(rdr)
        rdr.Close()
        myconnection.Close()
        DataGridView1.DataSource = dt

    End Sub
    Public Sub LoadContratosRecords()
        myconnection.Open()

        Dim cmd As New SQLiteCommand
        cmd.Connection = myconnection

        cmd.CommandText = "SELECT idVivi,fecha,idComprador,precio FROM Contratos ORDER BY fecha"

        Dim rdr As SQLite.SQLiteDataReader = cmd.ExecuteReader

        dt2.Load(rdr)
        rdr.Close()
        myconnection.Close()
        DataGridView2.DataSource = dt2

    End Sub

    Private Sub BunifuImageButton1_Click(sender As Object, e As EventArgs) Handles BunifuImageButton1.Click
        Principal.Show()
        Me.Hide()
    End Sub

    Private Sub BunifuImageButton2_Click(sender As Object, e As EventArgs) Handles BunifuImageButton2.Click
        Visitas.Show()
        Me.Hide()
    End Sub

    Private Sub BunifuImageButton4_Click(sender As Object, e As EventArgs) Handles BunifuImageButton4.Click
        Ventas.Show()
        Me.Hide()
    End Sub

    Private Sub BunifuImageButton5_Click(sender As Object, e As EventArgs) Handles BunifuImageButton5.Click
        Clientes.Show()
        Me.Hide()
    End Sub

    Private Sub BunifuImageButton7_Click(sender As Object, e As EventArgs) Handles BunifuImageButton7.Click
        Principal.Close()
    End Sub
End Class