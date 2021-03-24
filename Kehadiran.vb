Imports System.Data
Imports System.Data.OleDb
Imports System.IO

Public Class Kehadiran
    Dim cnnOLEDB As New OleDbConnection
    Dim cmdOLEDB As New OleDbCommand
    Dim cmdInsert As New OleDbCommand
    Dim cmdUpdate As New OleDbCommand
    Dim cmdDelete As New OleDbCommand
    Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = Akademik.accdb"
    Public ADP As OleDbDataAdapter
    Public DS As New DataSet

    Private Sub Kehadiran_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        cnnOLEDB.ConnectionString = strConnectionString
        cnnOLEDB.Open()
        TampilData()
        ListBox1.Visible = False
        ButtonEnable()
    End Sub

    Sub TampilData()
        ADP = New OleDbDataAdapter("SELECT * FROM Kehadiran ORDER BY NIM", cnnOLEDB)
        DS = New DataSet
        ADP.Fill(DS, "Tabel1")
        DataGridView1.DataSource = DS.Tables("Tabel1")
    End Sub

    Sub ButtonEnable()
        Button1.Enabled = True
        Button2.Enabled = False
        Button4.Enabled = False
        TextBox1.Enabled = True
    End Sub

    Sub ButtonDisable()
        Button1.Enabled = False
        Button2.Enabled = True
        Button4.Enabled = True
        TextBox1.Enabled = False
    End Sub

    Sub IsiList()
        Dim query As String
        query = "SELECT Nama_Mhs FROM Master_Mahasiswa WHERE Nama_Mhs LIKE '" & TextBox2.Text & "%' ORDER BY NIM "
        ADP = New OleDbDataAdapter(query, cnnOLEDB)
        DS = New DataSet
        ADP.Fill(DS, "List")
        ListBox1.Items.Clear()
        For i = 0 To DS.Tables("List").Rows.Count - 1
            ListBox1.Items.Add(DS.Tables("List").Rows(i).Item("Nama_Mhs").ToString)
        Next
    End Sub

    Sub IsiNim()
        Dim query As String
        query = "SELECT Nama_Mhs FROM Master_Mahasiswa WHERE NIM = '" & TextBox1.Text & "'"
        ADP = New OleDbDataAdapter(query, cnnOLEDB)
        DS = New DataSet
        ADP.Fill(DS, "NIM")
        For i = 0 To DS.Tables("NIM").Rows.Count - 1
            TextBox2.Text = (DS.Tables("NIM").Rows(i).Item("Nama_Mhs").ToString)
        Next
        ListBox1.Visible = False
    End Sub

    Sub IsiKelas()
        Dim query As String
        query = "SELECT Kelas FROM Kelas  WHERE NIM = '" & TextBox1.Text & "'"
        ADP = New OleDbDataAdapter(query, cnnOLEDB)
        DS = New DataSet
        ADP.Fill(DS, "Kelas")
        TextBox3.Text = ""
        For i = 0 To DS.Tables("Kelas").Rows.Count - 1
            TextBox3.Text = (DS.Tables("Kelas").Rows(i).Item("Kelas").ToString)
        Next
        ListBox1.Visible = False
    End Sub

    Sub Bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        DateTimePicker1.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
    End Sub

    Sub GetData(ByVal e)
        Dim NIM As Object = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        Dim Semester As Object = DataGridView1.Rows(e.RowIndex).Cells(1).Value
        Dim TA As Object = DataGridView1.Rows(e.RowIndex).Cells(2).Value
        Dim Tanggal As Object = DataGridView1.Rows(e.RowIndex).Cells(3).Value
        Dim Ijin As Object = DataGridView1.Rows(e.RowIndex).Cells(4).Value
        Dim Sakit As Object = DataGridView1.Rows(e.RowIndex).Cells(5).Value
        Dim Alpa As Object = DataGridView1.Rows(e.RowIndex).Cells(6).Value

        TextBox1.Text = CType(NIM, String)
        TextBox4.Text = CType(Semester, String)
        TextBox5.Text = CType(TA, String)
        DateTimePicker1.Text = CType(Tanggal, String)
        TextBox6.Text = CType(Ijin, String)
        TextBox7.Text = CType(Sakit, String)
        TextBox8.Text = CType(Alpa, String)
    End Sub


    Private Sub TextBox2_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles TextBox2.TextChanged
        ListBox1.Location = New Point(TextBox2.Location.X, TextBox2.Location.Y + 20)
        ListBox1.BringToFront()
        ListBox1.Visible = True
        IsiList()
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ListBox1.SelectedIndexChanged
        TextBox2.Text = ListBox1.SelectedItem
        ListBox1.Visible = False
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        GetData(e)
        ButtonDisable()
    End Sub


    Private Sub DataGridView1_CellContextMenuStripChanged(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles DataGridView1.CellContextMenuStripChanged
        GetData(e)
        ButtonDisable()
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text.Length() = 12 Then
            IsiNim()
            IsiKelas()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button1.Click
        If TextBox1.Text <> "" And TextBox2.Text <> "" And TextBox3.Text <> "" And DateTimePicker1.Text <> "" And TextBox4.Text <> "" And TextBox5.Text <> "" And TextBox6.Text <> "" And TextBox7.Text <> "" And TextBox8.Text <> "" Then
            Try
                cmdInsert.CommandText = "INSERT INTO Kehadiran " & "(NIM, Semester, Tahun_Akademik, Tanggal, Ijin, Sakit, Alpa)" &
                    "VALUES(@NIM, @Semester, @TA, @Tgl, @Ijin, @Sakit, @Alpa)"

                cmdInsert.Parameters.AddWithValue("@NIM", Me.TextBox1.Text)
                cmdInsert.Parameters.AddWithValue("@Semester", Me.TextBox4.Text)
                cmdInsert.Parameters.AddWithValue("@TA", Me.TextBox5.Text)
                cmdInsert.Parameters.AddWithValue("@Tgl", Me.DateTimePicker1.Value.Date)
                cmdInsert.Parameters.AddWithValue("@Ijin", Me.TextBox6.Text)
                cmdInsert.Parameters.AddWithValue("@Sakit", Me.TextBox7.Text)
                cmdInsert.Parameters.AddWithValue("@Alpa", Me.TextBox8.Text)

                cmdInsert.CommandType = CommandType.Text
                cmdInsert.Connection = cnnOLEDB
                cmdInsert.ExecuteNonQuery()
                MsgBox("DATA DIMASUKKAN")
                Bersih()
                TampilData()
                ListBox1.Visible = False
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        Else
            MsgBox("Masukkan data secara lengkap")
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button2.Click
        If TextBox1.Text <> "" And TextBox2.Text <> "" And TextBox3.Text <> "" And DateTimePicker1.Text <> "" And TextBox4.Text <> "" And TextBox5.Text <> "" And TextBox6.Text <> "" And TextBox7.Text <> "" And TextBox8.Text <> "" Then
            Try
                ADP = New OleDbDataAdapter("UPDATE Kehadiran SET NIM = '" & TextBox1.Text &
                    "', Semester = '" & TextBox4.Text & "', Tahun_Akademik = '" & TextBox5.Text & "', Tanggal ='" & DateTimePicker1.Value.ToString & "', Ijin = '" & TextBox6.Text &
                    "', Sakit = '" & TextBox6.Text & "', Alpa = '" & TextBox7.Text & "' WHERE NIM = '" & TextBox1.Text & "' ", cnnOLEDB)

                ADP.Fill(DS, "Kehadiran")
                MsgBox("DATA DIPERBARUI")
                TampilData()
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button4.Click
        Try
            cmdDelete.CommandText = "DELETE FROM Kehadiran WHERE NIM=@NIM"
            cmdDelete.Parameters.AddWithValue("@NIM", Me.TextBox1.Text)
            cmdDelete.CommandType = CommandType.Text
            cmdDelete.Connection = cnnOLEDB
            cmdDelete.ExecuteNonQuery()
            MsgBox("DATA DIHAPUS")
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        cmdDelete.Dispose()
        TampilData()
        Bersih()
    End Sub

    Private Sub Button3_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button3.Click
        Bersih()
        ButtonEnable()
        ListBox1.Visible = False
    End Sub



End Class
