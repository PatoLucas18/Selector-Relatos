Imports System.IO
Imports System.Text
Imports Microsoft.VisualBasic.FileIO

Public Class Form1
    Dim ejecutable As String = "pes6.exe" 'Nombre del exe

    'Mostra previa
    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        'El siguiente código mostrará la imagen del pulsar un boton segun la opcion seleccionada
        Try
            PictureBox1.Image = Image.FromFile(list_comm(ComboBox1.SelectedIndex, 3)) 'Si selecciono la opcion 4
        Catch ex As Exception

        End Try

    End Sub

    'Guardar relato
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'El siguiente código escribe el nuevo nombre para el relato segun la opcion elegida

        'Abrir el exe para escribir en el
        Dim fs As New FileStream(ejecutable, FileMode.Open)
        Dim bw As New BinaryWriter(fs)

        Dim offset As Integer = &H77E108 'offset nombre de relato en español

        'Escribir nombre de relato
        Dim bytes2(11) As Byte
        Dim nombreAFS As String = list_comm(ComboBox1.SelectedIndex, 2)
        
        bytes2 = Encoding.UTF8.GetBytes(nombreAFS)
        If bytes2.Length > 11 Then
            'MsgBox("Nombre de afs muy largo, colocar uno con 11 letras")
            MsgBox(String.Format("Nombre de afs muy largo, colocar uno con 11 letras, Debes Restar {0} letras en Nombre afs de la opcion {1}", bytes2.Length - 11, list_comm(ComboBox1.SelectedIndex, 1)))
            Exit Sub
        End If
        bw.Seek(offset, SeekOrigin.Begin)
        bw.Write(bytes2, 0, 11)
        fs.Close()
        bw.Close()
        MsgBox(ComboBox1.Text & " Instalado")

    End Sub

    'abrir juego
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'abrir juego
        'Shell(ejecutable)
        Shell("CMD.EXE /r " & ejecutable, AppWinStyle.Hide)
    End Sub


    'al carga formulario
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Leer el archivo exe (pes6.exe)
        Dim fs As New FileStream(ejecutable, FileMode.Open)
        Dim br As New BinaryReader(fs)
        'Leer la posicion del nombre del relato en español offset = 77E108
        fs.Seek(&H77E108, SeekOrigin.Begin)
        Dim bytes() As Byte = br.ReadBytes(&HB)
        Dim s_sound As String = Encoding.UTF8.GetString(bytes, 0, bytes.Length)

        fs.Close()
        br.Close()

        csv_read()
        For i = 0 To 50
            If s_sound = list_comm(i, 2) Then
                ComboBox1.SelectedIndex = list_comm(i, 0)
            End If
        Next

        'MsgBox(commentaries(1, 1))

    End Sub
    Dim list_comm(50, 50) As String
    Public Function csv_read() As String

        Dim tfp As New TextFieldParser("dat\map.csv")
        tfp.Delimiters = New String() {","}
        tfp.TextFieldType = FieldType.Delimited


        tfp.ReadLine() ' skip header
        While tfp.EndOfData = False
            Dim fields = tfp.ReadFields()
            'id = fields(0)
            'title = fields(1)
            'name_afs = fields(2)
            'file_preview = fields(3)

            list_comm(fields(0), 0) = fields(0)
            list_comm(fields(0), 1) = fields(1)
            list_comm(fields(0), 2) = fields(2)
            list_comm(fields(0), 3) = fields(3)
            ComboBox1.Items.Add(fields(1))
            'MsgBox(String.Format("{0} - {1} - {2}", title, name_afs, file_preview))
        End While
        Return False
    End Function
End Class
