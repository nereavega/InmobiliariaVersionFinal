Imports System.Data.SQLite

Public Class Ventas
    Private Const cs As String = "Data Source = C:\DATOS\database.db;Version=3"
    Dim respuesta As MsgBoxResult


    Dim myconnection As New SQLiteConnection(cs)
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Carga la tabla
        LoadVentasRecords()

        Dim conn As New Principal.ConectarSQLite
        'Carga los combos
        Principal.LoadCbViviendas()
        Principal.LoadCbClientes()

    End Sub

    Private Sub LoadVentasRecords()

        myconnection.Open()
        Dim dt As New DataTable
        Dim cmd As New SQLiteCommand
        cmd.Connection = myconnection

        cmd.CommandText = "SELECT* FROM Contratos ORDER BY fecha DESC"

        Dim rdr As SQLite.SQLiteDataReader = cmd.ExecuteReader

        dt.Load(rdr)
        rdr.Close()
        myconnection.Close()
        DataGridView1.DataSource = dt

    End Sub

    'BOTÓN NUEVA VENTA
    Private Sub BunifuFlatButton1_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton1.Click
        '1.-Crea el registro
        myconnection.Open()

        Dim cmd As New SQLiteCommand
        cmd.Connection = myconnection

        cmd.CommandText = "INSERT INTO Contratos(idComprador,idVivi,fecha,precio) VALUES (@idComprador,@idVivi,@fecha,@precio)"

        cmd.Parameters.AddWithValue("@idComprador", ComboBox1.Text)
        cmd.Parameters.AddWithValue("@idVivi", ComboBox2.Text)
        cmd.Parameters.AddWithValue("@fecha", Convert.ToString(BunifuDatepicker1.Value))
        cmd.Parameters.AddWithValue("@precio", BunifuMaterialTextbox1.Text)

        BunifuMaterialTextbox1.Text = ""

        Dim recsadded As Integer = cmd.ExecuteNonQuery()
        myconnection.Close()

        '2.-Actualiza la casa como VENDIDA
        myconnection.Open()
        cmd.Connection = myconnection

        cmd.CommandText = "UPDATE Viviendas SET enVenta='VENDIDO' where id=@idVivi"
        cmd.Parameters.AddWithValue("@idVivi", ComboBox2.Text)

        Dim recsadded2 As Integer = cmd.ExecuteNonQuery()
        myconnection.Close()

        '3.- Actualiza al cliente el campo haComprado
        myconnection.Open()
        cmd.Connection = myconnection

        cmd.CommandText = "UPDATE Clientes SET haComprado='SI' where id=@idComprador"
        cmd.Parameters.AddWithValue("@idVivi", ComboBox2.Text)

        Dim recsadded3 As Integer = cmd.ExecuteNonQuery()
        myconnection.Close()

        Principal.LoadCbViviendas() 'Desaparece del cb de viviendas disponibles
        LoadVentasRecords() 'tabla de ventas. Se añade el registro
        Principal.LoadViviendasRecords() 'tabla principal viviendas, desaparece de la tabla
        Visitas.LoadViviendasRecords2() 'tabla visitas, la casa ya no se puede visitar
        Visitas.LoadClientesRecords2() 'tabla Clientes, ahora "ha comprado"
    End Sub


    'BOTONES DEL MENÚ
    Private Sub BunifuImageButton1_Click(sender As Object, e As EventArgs) Handles BunifuImageButton1.Click
        Principal.Refresh()
        Principal.Show()
        Me.Hide()
    End Sub

    Private Sub BunifuImageButton2_Click(sender As Object, e As EventArgs) Handles BunifuImageButton2.Click
        Visitas.Refresh()
        Visitas.Show()
        Me.Hide()
    End Sub
    Private Sub BunifuImageButton5_Click(sender As Object, e As EventArgs) Handles BunifuImageButton5.Click
        Clientes.Refresh()
        Clientes.Show()
        Me.Hide()
    End Sub

    'BOTÓN CERRAR PROGRAMA
    Private Sub BunifuImageButton7_Click(sender As Object, e As EventArgs) Handles BunifuImageButton7.Click
        Principal.Close()
    End Sub


End Class