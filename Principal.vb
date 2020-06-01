Imports System.Data.SQLite
Imports System.Net.Mime.MediaTypeNames

Public Class Principal
    Private Const cs As String = "Data Source = C:\DATOS\database.db;Version=3"
    Dim respuesta As MsgBoxResult
    'Dim dt As New DataTable
    Dim myconnection As New SQLiteConnection(cs)

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Carga la tabla
        LoadViviendasRecords()

        'Carga los combos
        LoadCbComerciales()
        LoadCbClientes()

    End Sub
    Sub LoadCbClientes()
        Dim conn As New ConectarSQLite
        Dim SqlClientes As String = "SELECT * FROM Clientes"
        Dim DtClientes As New DataTable
        DtClientes = conn.consulta(SqlClientes)
        ComboBox4.DataSource = DtClientes
        ComboBox4.DisplayMember = "id"
        Ventas.ComboBox1.DataSource = DtClientes
        Ventas.ComboBox1.DisplayMember = "id"
    End Sub
    Sub LoadCbComerciales()
        Dim conn As New ConectarSQLite
        Dim SqlComerciales As String = "SELECT * FROM Comerciales"
        Dim DtComerciales As New DataTable
        DtComerciales = conn.consulta(SqlComerciales)
        ComboBox7.DataSource = DtComerciales
        ComboBox7.DisplayMember = "id"
        ComboBox8.DataSource = DtComerciales
        ComboBox8.DisplayMember = "id"
    End Sub
    Sub LoadCbViviendas()
        'Carga el combo de Viviendas en venta en el form VENTAS
        Dim conn As New ConectarSQLite
        Dim SqlClientes As String = "SELECT * FROM Viviendas WHERE enVenta='EN VENTA'"
        Dim DtClientes As New DataTable
        DtClientes = conn.consulta(SqlClientes)
        Ventas.ComboBox2.DataSource = DtClientes
        Ventas.ComboBox2.DisplayMember = "id"
    End Sub
    Public Class ConectarSQLite
        ':::Nos conectamos a la base de datos por medio del objeto SQLiteConnection
        Private con As New SQLiteConnection(cs)

        Function consulta(ByVal Sql As String)
            ':::Creamos un objeto DataTable donde almacenaremos nuestros datos
            Dim Dt As New DataTable
            ':::Usamos un capturador para los posibles errores
            Try
                ':::Creamos un objeto SQLiteDataAdapter para realizar la consulta a la base de datos
                ':::Le pasamos 2 parametros la consulta SQL y la conexion
                Dim Da As New SQLiteDataAdapter(Sql, con)
                ':::Le pasamos los datos recibidos a nuestro DataTable
                Da.Fill(Dt)
            Catch ex As Exception
                MessageBox.Show("No se completo la consulta por: " & ex.ToString)
            End Try
            ':::Retornamos el resultado obtenido
            Return Dt
        End Function
    End Class

    Public Sub LoadViviendasRecords() 'Recarga la tabla de viviendas PRINCIPAL
        myconnection.Open()
        Dim dt As New DataTable
        Dim cmd As New SQLiteCommand
        cmd.Connection = myconnection

        cmd.CommandText = "SELECT id,precio,barrio,metros,habitaciones,banios,ascensor,estado,direccion,idPropietario,idComercial FROM Viviendas WHERE enVenta='EN VENTA' ORDER BY id DESC"

        Dim rdr As SQLite.SQLiteDataReader = cmd.ExecuteReader

        dt.Load(rdr)
        rdr.Close()
        myconnection.Close()
        DataGridView1.DataSource = dt

    End Sub
    Sub LoadViviendasRecords2() 'Recarga la tabla de viviendas de VISITAS 

        myconnection.Open()
        Dim dt2 As New DataTable
        Dim cmd As New SQLiteCommand
        cmd.Connection = myconnection

        cmd.CommandText = "SELECT id, direccion FROM Viviendas WHERE enVenta='EN VENTA' ORDER BY direccion"

        Dim rdr As SQLite.SQLiteDataReader = cmd.ExecuteReader

        dt2.Load(rdr)
        rdr.Close()
        myconnection.Close()
        Visitas.DataGridView2.DataSource = dt2

    End Sub

    'EVENTO DE LA TABLA
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim i As Integer = DataGridView1.CurrentRow.Index

        TextBox1.Text = ""
        TextBox1.Text = DataGridView1.Item(1, i).Value() 'precio
        ComboBox2.SelectedText = ""
        ComboBox2.SelectedText = DataGridView1.Item(2, i).Value() 'barrio
        TextBox2.Text = ""
        TextBox2.Text = DataGridView1.Item(3, i).Value()  'metros
        ComboBox1.Text = ""
        ComboBox1.Text = DataGridView1.Item(4, i).Value()   'habitaciones
        ComboBox3.Text = ""
        ComboBox3.Text = DataGridView1.Item(5, i).Value()  'banios
        ComboBox6.Text = ""
        ComboBox6.Text = DataGridView1.Item(6, i).Value()  'ascensor
        ComboBox5.Text = ""
        ComboBox5.Text = DataGridView1.Item(7, i).Value()  'estado
        TextBox4.Text = ""
        TextBox4.Text = DataGridView1.Item(8, i).Value() 'Direccion
        ComboBox4.Text = ""
        ComboBox4.Text = DataGridView1.Item(9, i).Value() 'idPropietario
        ComboBox7.Text = ""
        ComboBox7.Text = DataGridView1.Item(10, i).Value() 'idComercial

    End Sub

    'CAMBIAR PRECIO'
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim thisId As Integer = DataGridView1.SelectedRows(0).Cells("id").Value

        myconnection.Open()

        Dim cmd As New SQLiteCommand
        cmd.Connection = myconnection

        cmd.CommandText = "Update Viviendas set precio=@precioNuevo where id=@id"
        cmd.Parameters.AddWithValue("@precioNuevo", TextBox3.Text)
        cmd.Parameters.AddWithValue("@id", thisId)

        TextBox3.Text = ""
        MsgBox("Registro número " + thisId.ToString + " modificado", MsgBoxStyle.Information)

        cmd.ExecuteNonQuery()

        myconnection.Close()

        LoadViviendasRecords()
        LoadViviendasRecords2()

    End Sub

    'CAMBIAR DE COMERCIAL
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim thisId As Integer = DataGridView1.SelectedRows(0).Cells("id").Value

        myconnection.Open()

        Dim cmd As New SQLiteCommand
        cmd.Connection = myconnection

        cmd.CommandText = "UPDATE Viviendas set idComercial=@comercialNuevo where id=@id"
        cmd.Parameters.AddWithValue("@comercialNuevo", ComboBox8.Text)
        cmd.Parameters.AddWithValue("@id", thisId)

        TextBox3.Text = ""
        MsgBox("Registro número " + thisId.ToString + " modificado", MsgBoxStyle.Information)

        cmd.ExecuteNonQuery()
        myconnection.Close()

        LoadViviendasRecords()
        LoadViviendasRecords2()
    End Sub

    'BORRAR REGISTRO'
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim borrar As Boolean
        respuesta = CType(MessageBox.Show("¿Está seguro de que desea borrar el registro número seleccionado? Los cambios serán permanentes.", "Advertencia", MessageBoxButtons.YesNo, MessageBoxIcon.Question), MsgBoxResult)
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

            cmd.CommandText = "DELETE FROM Viviendas WHERE id=@id"
            cmd.Parameters.AddWithValue("@id", thisId)

            Dim recsadded As Integer = cmd.ExecuteNonQuery()
            myconnection.Close()
            MsgBox("Registro número " + thisId.ToString + " eliminado", MsgBoxStyle.Information)

            LoadViviendasRecords()
            LoadViviendasRecords2()
            LoadCbViviendas()

        End If


    End Sub


    'AÑADIR REGISTRO'
    Private Sub BunifuFlatButton1_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton1.Click

        myconnection.Open()

        Dim cmd As New SQLiteCommand
        cmd.Connection = myconnection

        cmd.CommandText = "INSERT INTO Viviendas(precio,barrio,metros,habitaciones,banios,ascensor,estado,direccion,idPropietario,idComercial) VALUES (@precio,@barrio,@metros,@habitaciones,@banios,@ascensor,@estado,@direccion,@idProp,@idCom)"

        cmd.Parameters.AddWithValue("@precio", Convert.ToInt32(TextBox1.Text))
        cmd.Parameters.AddWithValue("@barrio", ComboBox2.SelectedItem)
        cmd.Parameters.AddWithValue("@metros", Convert.ToInt32(TextBox2.Text))
        cmd.Parameters.AddWithValue("@habitaciones", Convert.ToInt32(ComboBox1.SelectedItem))
        cmd.Parameters.AddWithValue("@banios", Convert.ToInt32(ComboBox3.SelectedItem))
        cmd.Parameters.AddWithValue("@ascensor", ComboBox6.SelectedItem)
        cmd.Parameters.AddWithValue("@estado", ComboBox5.SelectedItem)
        cmd.Parameters.AddWithValue("@direccion", TextBox4.Text)
        cmd.Parameters.AddWithValue("@idProp", ComboBox4.Text)
        cmd.Parameters.AddWithValue("@idCom", ComboBox7.Text)

        TextBox1.Text = ""
        ComboBox2.Text = ""
        TextBox2.Text = ""
        ComboBox1.Text = ""
        ComboBox3.Text = ""
        ComboBox4.Text = ""
        ComboBox5.Text = ""
        ComboBox6.Text = ""
        TextBox4.Text = ""
        ComboBox4.Text = ""
        ComboBox7.Text = ""


        Dim recsadded As Integer = cmd.ExecuteNonQuery()
        myconnection.Close()

        MsgBox("Registro añadido", MsgBoxStyle.Information)
        LoadViviendasRecords()
        LoadViviendasRecords2()
        LoadCbViviendas()

    End Sub

    'BOTÓN CERRAR APLICACIÓN'
    Private Sub BunifuImageButton7_Click(sender As Object, e As EventArgs) Handles BunifuImageButton7.Click
        Me.Close()

    End Sub

    Private Sub Principal_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        respuesta = CType(MessageBox.Show("¿Desea salir de la aplicación?", "Advertencia", MessageBoxButtons.YesNo, MessageBoxIcon.Question), MsgBoxResult)
        If respuesta = MsgBoxResult.No Then
            e.Cancel = True
        Else
            e.Cancel = False
            Visitas.Close()
            Ventas.Close()
            Clientes.Close()

        End If
    End Sub

    'BOTONES DEL MENÚ
    Private Sub BunifuImageButton2_Click(sender As Object, e As EventArgs) Handles BunifuImageButton2.Click
        Visitas.Refresh()
        Visitas.Show()
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

End Class

