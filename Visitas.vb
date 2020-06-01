Imports System.Data.SQLite

Public Class Visitas
    Private Const cs As String = "Data Source = C:\DATOS\database.db;Version=3"
    Dim respuesta As MsgBoxResult

    Dim myconnection As New SQLiteConnection(cs)
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadVisitasRecords()
        LoadViviendasRecords2()

        'Cargar combo comerciales
        Dim conn As New Principal.ConectarSQLite
        Dim SqlComerciales As String = "SELECT * FROM Comerciales"
        Dim DtComerciales As New DataTable
        DtComerciales = conn.consulta(SqlComerciales)
        ComboBox6.DataSource = DtComerciales
        ComboBox6.DisplayMember = "id"


    End Sub


    Private Sub LoadVisitasRecords()
        myconnection.Open()
        Dim dt As New DataTable
        Dim cmd As New SQLiteCommand
        cmd.Connection = myconnection

        cmd.CommandText = "SELECT* FROM Visitas ORDER BY fecha, hora"

        Dim rdr As SQLite.SQLiteDataReader = cmd.ExecuteReader

        dt.Load(rdr)
        rdr.Close()
        myconnection.Close()
        DataGridView1.DataSource = dt

    End Sub

    Sub LoadViviendasRecords2()

        myconnection.Open()
        Dim dt2 As New DataTable
        Dim cmd As New SQLiteCommand
        cmd.Connection = myconnection

        cmd.CommandText = "SELECT id, direccion FROM Viviendas WHERE enVenta='EN VENTA' ORDER BY direccion"

        Dim rdr As SQLite.SQLiteDataReader = cmd.ExecuteReader

        dt2.Load(rdr)
        rdr.Close()
        myconnection.Close()
        DataGridView2.DataSource = dt2

    End Sub
    Public Sub LoadClientesRecords2() 'Carga la tabla de clientes de CLIENTES
        myconnection.Open()
        Dim dt As New DataTable
        Dim cmd As New SQLiteCommand
        cmd.Connection = myconnection

        cmd.CommandText = "SELECT* FROM Clientes ORDER BY nombre"

        Dim rdr As SQLite.SQLiteDataReader = cmd.ExecuteReader

        dt.Load(rdr)
        rdr.Close()
        myconnection.Close()
        Clientes.DataGridView1.DataSource = dt

    End Sub

    'NUEVA VISITA
    Private Sub BunifuFlatButton1_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton1.Click
        myconnection.Open()

        Dim cmd As New SQLiteCommand
        cmd.Connection = myconnection

        cmd.CommandText = "INSERT INTO Visitas(idComercial,idVivienda,nombreCliente,tlfCliente,direccion,fecha,hora) VALUES (@idComercial,@idVivienda,@nombreCliente,@tlfCliente,@direccion,@fecha,@hora)"
        Dim idViv As Integer = DataGridView2.SelectedRows(0).Cells("id").Value

        'No puede pasar de DataGridView a String
        cmd.Parameters.AddWithValue("@idComercial", ComboBox6.Text)
        cmd.Parameters.AddWithValue("@idVivienda", idViv)
        cmd.Parameters.AddWithValue("@nombreCliente", (BunifuMaterialTextbox1.Text))
        cmd.Parameters.AddWithValue("@tlfCliente", Convert.ToInt32(BunifuMaterialTextbox2.Text))
        cmd.Parameters.AddWithValue("@direccion", (BunifuMaterialTextbox3.Text))
        cmd.Parameters.AddWithValue("@fecha", Convert.ToString(BunifuDatepicker1.Value))
        cmd.Parameters.AddWithValue("@hora", TextBox3.Text)

        Dim recsadded As Integer = cmd.ExecuteNonQuery()
        myconnection.Close()

        'generar un registro de cliente tomando el nombre contacto fecha y añadiendo dni XXX y que no ha comprado
        myconnection.Open()
        Dim cmd2 As New SQLiteCommand
        cmd2.Connection = myconnection

        cmd2.CommandText = "INSERT INTO Clientes (nombre,contacto,dni,haComprado,fechIncorp) VALUES (@nombre,@contacto,@dni,@haComprado,@fechaIncorp)"
        cmd2.Parameters.AddWithValue("@nombre", BunifuMaterialTextbox1.Text)
        cmd2.Parameters.AddWithValue("@contacto", Convert.ToInt32(BunifuMaterialTextbox2.Text))
        cmd2.Parameters.AddWithValue("@dni", "XXXXXX")
        cmd2.Parameters.AddWithValue("haComprado", "NO")
        cmd2.Parameters.AddWithValue("fechaIncorp", Convert.ToString(BunifuDatepicker1.Value))

        cmd2.ExecuteNonQuery()
        myconnection.Close()

        BunifuMaterialTextbox1.Text = ""
        BunifuMaterialTextbox2.Text = ""
        BunifuMaterialTextbox3.Text = ""
        TextBox3.Text = ""


        MsgBox("Cita añadida", MsgBoxStyle.Information)
        LoadVisitasRecords()
        LoadClientesRecords2()
        Principal.LoadCbClientes()

    End Sub




    'BOTÓN BORRAR CITA SELECCIONADA
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim borrar As Boolean
        respuesta = CType(MessageBox.Show("¿Está seguro de que desea eliminar la cita seleccionada? Los cambios serán permanentes.", "Advertencia", MessageBoxButtons.YesNo, MessageBoxIcon.Question), MsgBoxResult)
        If respuesta = MsgBoxResult.No Then
            borrar = False
        Else
            borrar = True
        End If

        If (borrar) Then

            Dim thisId As Integer = DataGridView1.SelectedRows(0).Cells("id").Value

            myconnection.Open()

            Dim cmd As New SQLiteCommand
            cmd.Connection = myconnection

            cmd.CommandText = "DELETE FROM Visitas WHERE id=@id"
            cmd.Parameters.AddWithValue("@id", thisId)

            Dim recsadded As Integer = cmd.ExecuteNonQuery()

            myconnection.Close()

            MsgBox("Cita Eliminada", MsgBoxStyle.Information)
            LoadVisitasRecords()

        End If
    End Sub

    'MODIFICAR CITA
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim thisId As Integer = DataGridView1.SelectedRows(0).Cells("id").Value

        myconnection.Open()

        Dim cmd As New SQLiteCommand
        cmd.Connection = myconnection

        cmd.CommandText = "UPDATE Visitas SET fecha=@fechaNueva, hora=@horaNueva WHERE id=@id"
        cmd.Parameters.AddWithValue("@fechaNueva", Convert.ToString(BunifuDatepicker2.Value))
        cmd.Parameters.AddWithValue("@horaNueva", TextBox4.Text)
        cmd.Parameters.AddWithValue("@id", thisId)

        TextBox4.Text = ""
        MsgBox("Registro número " + thisId.ToString + " modificado", MsgBoxStyle.Information)

        cmd.ExecuteNonQuery()

        myconnection.Close()

        LoadVisitasRecords()
    End Sub

    'BOTONES DEL MENÚ
    Private Sub BunifuImageButton1_Click(sender As Object, e As EventArgs) Handles BunifuImageButton1.Click
        Principal.Refresh()
        Principal.Show()
        Me.Hide()

    End Sub

    Private Sub BunifuImageButton4_Click(sender As Object, e As EventArgs) Handles BunifuImageButton4.Click
        Ventas.Refresh()
        Ventas.Show()
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

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class