Imports System.Data.SQLite

Public Class Principal
    Dim respuesta As MsgBoxResult


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadRecords()

    End Sub
    Private Sub LoadRecords()
        Dim myconnection As New SQLiteConnection("Data Source = C:\Users\nerea\Desktop\database.db;Version=3")
        myconnection.Open()

        Dim cmd As New SQLiteCommand
        cmd.Connection = myconnection

        cmd.CommandText = "Select* from enVenta order by id desc"

        Dim rdr As SQLite.SQLiteDataReader = cmd.ExecuteReader
        Dim dt As New DataTable
        dt.Load(rdr)
        rdr.Close()
        myconnection.Close()
        DataGridView1.DataSource = dt

    End Sub


    'CAMBIAR PRECIO'
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim thisId As Integer = DataGridView1.SelectedRows(0).Cells("id").Value

        Dim myconnection As New SQLiteConnection("Data Source = C:\Users\nerea\Desktop\database.db;Version=3")
        myconnection.Open()

        Dim cmd As New SQLiteCommand
        cmd.Connection = myconnection

        cmd.CommandText = "Update enVenta set precio=@precioNuevo where id=@id"
        cmd.Parameters.AddWithValue("@precioNuevo", TextBox3.Text)
        cmd.Parameters.AddWithValue("@id", thisId)

        TextBox3.Text = ""
        MsgBox("Registro número " + thisId.ToString + " modificado", MsgBoxStyle.Information)

        cmd.ExecuteNonQuery()

        myconnection.Close()

        LoadRecords()

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

            Dim myconnection As New SQLiteConnection("Data Source = C:\Users\nerea\Desktop\database.db;Version=3")
            myconnection.Open()

            Dim cmd As New SQLiteCommand
            cmd.Connection = myconnection
            cmd.Connection = myconnection

            cmd.CommandText = "Delete from enVenta where id=@id"
            cmd.Parameters.AddWithValue("@id", thisId)

            TextBox3.Text = ""
            MsgBox("Registro número " + thisId.ToString + " eliminado", MsgBoxStyle.Information)

            cmd.ExecuteNonQuery()

            myconnection.Close()

            LoadRecords()
        End If
    End Sub

    Private Sub BunifuImageButton4_Click(sender As Object, e As EventArgs) Handles BunifuImageButton4.Click

    End Sub


    'AÑADIR REGISTRO'
    Private Sub BunifuFlatButton1_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton1.Click
        If TextBox1.Text = "" Or ComboBox2.Text = "" Or TextBox2.Text = "" Or ComboBox1.Text = "" Or ComboBox3.Text = "" Or ComboBox4.Text = "" Or ComboBox5.Text = "" Then
            MsgBox("Por favor introduzca datos en todos los campos si quiere añadir un registro.", MsgBoxStyle.Exclamation, "Error")
            Exit Sub
        End If

        Dim myconnection As New SQLiteConnection("Data Source = C:\Users\nerea\Desktop\database.db;Version=3")
            myconnection.Open()

        Dim cmd As New SQLiteCommand
        cmd.Connection = myconnection

        cmd.CommandText = "INSERT INTO enVenta(precio,barrio,metros,habitaciones,banios,ascensor,estado) VALUES (@precio,@barrio,@metros,@habitaciones,@banios,@ascensor,@estado)"
        cmd.Parameters.AddWithValue("@precio", TextBox1.Text)
        cmd.Parameters.AddWithValue("@barrio", ComboBox2.Text)
        cmd.Parameters.AddWithValue("@metros", TextBox2.Text)
        cmd.Parameters.AddWithValue("@habitaciones", ComboBox1.Text)
        cmd.Parameters.AddWithValue("@banios", ComboBox3.Text)
        cmd.Parameters.AddWithValue("@ascensor", ComboBox4.Text)
        cmd.Parameters.AddWithValue("@estado", ComboBox5.Text)

        TextBox1.Text = ""
        ComboBox2.Text = ""
        TextBox2.Text = ""
        ComboBox1.Text = ""
        ComboBox3.Text = ""
        ComboBox4.Text = ""
        ComboBox5.Text = ""


        Dim recsadded As Integer = cmd.ExecuteNonQuery()
        myconnection.Close()

        MsgBox("Registro añadido", MsgBoxStyle.Information)

        LoadRecords()
    End Sub

    Private Sub BunifuImageButton7_Click(sender As Object, e As EventArgs) Handles BunifuImageButton7.Click
        Me.Close()
    End Sub

    Private Sub Principal_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        respuesta = CType(MessageBox.Show("¿Desea salir de la aplicación?", "Advertencia", MessageBoxButtons.YesNo, MessageBoxIcon.Question), MsgBoxResult)
        If respuesta = MsgBoxResult.No Then
            e.Cancel = True
        Else
            e.Cancel = False

        End If
    End Sub

    'CONSULTA DINÁMICA'
    Sub consultaDinamica(ByVal id As String, ByVal dgv As DataGridView)
        Dim myconnection As New SQLiteConnection("Data Source = C:\Users\nerea\Desktop\database.db;Version=3")
        myconnection.Open()
        Try
            Dim cmd As New SQLiteCommand
            cmd.Connection = myconnection

            cmd.CommandText = ("SELECT * FROM enVenta WHERE barrio LIKE '&@barrio+%&'")
            cmd.Parameters.AddWithValue("@barrio", BunifuMaterialTextbox1.Text)

            Dim recsadded As Integer = cmd.ExecuteNonQuery()
            myconnection.Close()
            LoadRecords()

        Catch ex As Exception

        End Try


    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub
End Class

