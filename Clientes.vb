Imports System.Data.SQLite

Public Class Clientes

    Dim myconnection As New SQLiteConnection("Data Source = C:\DATOS\database.db;Version=3")
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Carga la tabla
        Visitas.LoadClientesRecords2()

    End Sub


    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim i As Integer
        i = DataGridView1.CurrentRow.Index

        TextBox1.Text = DataGridView1.Item(1, i).Value()
        TextBox2.Text = DataGridView1.Item(2, i).Value()
        TextBox3.Text = DataGridView1.Item(3, i).Value()
        TextBox4.Text = DataGridView1.Item(4, i).Value()
        TextBox5.Text = DataGridView1.Item(5, i).Value()

    End Sub

    'MODIFICAR REGISTRO CLIENTE
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim thisId As Integer = DataGridView1.SelectedRows(0).Cells("id").Value
        If (IsDBNull(DataGridView1.SelectedRows(0))) Then
            MsgBox("Error. Seleccione una fila para modificar")

        Else
            myconnection.Open()

            Dim cmd As New SQLiteCommand
            cmd.Connection = myconnection

            cmd.CommandText = "Update Clientes set nombre=@nombre,contacto=@contacto,dni=@dni,haComprado=@haComprado,fechIncorp=@fechaIncorp where id=@id"
            cmd.Parameters.AddWithValue("@nombre", TextBox1.Text)
            cmd.Parameters.AddWithValue("@contacto", TextBox2.Text)
            cmd.Parameters.AddWithValue("@dni", TextBox3.Text)
            cmd.Parameters.AddWithValue("haComprado", TextBox4.Text)
            cmd.Parameters.AddWithValue("fechaIncorp", TextBox5.Text)

            cmd.Parameters.AddWithValue("@id", thisId)

            cmd.ExecuteNonQuery()

            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = ""

            MsgBox("Registro número " + thisId.ToString + " modificado", MsgBoxStyle.Information)

            myconnection.Close()

            Visitas.LoadClientesRecords2() 'recarga la tabla de clientes
        End If
    End Sub

    'AÑADIR NUEVO CLIENTE
    Private Sub BunifuFlatButton1_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton1.Click
        myconnection.Open()

        Dim cmd As New SQLiteCommand
        cmd.Connection = myconnection

        cmd.CommandText = "INSERT INTO Clientes (nombre,contacto,dni,haComprado,fechIncorp) VALUES (@nombre,@contacto,@dni,@haComprado,@fechaIncorp)"
        cmd.Parameters.AddWithValue("@nombre", TextBox1.Text)
        cmd.Parameters.AddWithValue("@contacto", TextBox2.Text)
        cmd.Parameters.AddWithValue("@dni", TextBox3.Text)
        cmd.Parameters.AddWithValue("haComprado", "NO")
        cmd.Parameters.AddWithValue("fechaIncorp", TextBox5.Text)

        cmd.ExecuteNonQuery()

        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""

        MsgBox("Cliente añadido", MsgBoxStyle.Information)

        myconnection.Close()

        Visitas.LoadClientesRecords2() 'añade el registro a la tabla
        Principal.LoadCbClientes() 'añade el cliente al cb de la pantalla principal y de ventas
    End Sub

    'BORRAR UN CLIENTE
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim thisId As Integer = DataGridView1.SelectedRows(0).Cells("id").Value
        Dim borrar As Boolean
        Dim respuesta As MsgBoxResult
        respuesta = CType(MessageBox.Show("¿Está seguro de que desea eliminar el cliente seleccionado? Los cambios serán permanentes.", "Advertencia", MessageBoxButtons.YesNo, MessageBoxIcon.Question), MsgBoxResult)
        If respuesta = MsgBoxResult.No Then
            borrar = False
        Else
            borrar = True
        End If
        If (borrar) Then
            myconnection.Open()

            Dim cmd As New SQLiteCommand
            cmd.Connection = myconnection

            cmd.CommandText = "DELETE FROM Clientes WHERE id=@id"

            cmd.Parameters.AddWithValue("@id", thisId)

            cmd.ExecuteNonQuery()

            MsgBox("Cliente número " + thisId.ToString + " eliminado", MsgBoxStyle.Information)

            myconnection.Close()
        End If

        Visitas.LoadClientesRecords2() 'añade el registro a la tabla
        Principal.LoadCbClientes() 'añade el cliente al cb de la pantalla principal y de ventas
    End Sub

    'BOTONES MENÚ
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

    Private Sub BunifuImageButton4_Click(sender As Object, e As EventArgs) Handles BunifuImageButton4.Click
        Ventas.Refresh()
        Ventas.Show()
        Me.Hide()
    End Sub


    'BOTÓN CERRAR PROGRAMA
    Private Sub BunifuImageButton7_Click(sender As Object, e As EventArgs) Handles BunifuImageButton7.Click
        Principal.Close()
    End Sub

    Private Sub BunifuCustomLabel2_Click(sender As Object, e As EventArgs) Handles BunifuCustomLabel2.Click

    End Sub
End Class