Imports System.IO

Public Class CopyingDirectory
    Public Sub Copying(ByVal srcDirectory As String, ByVal destDirectory As String)
        Dim excludeType As String
        excludeType = TextBox3.Text
        Dim TestArray() As String = Split(excludeType, ",")


        For Each foundDirectory As String In My.Computer.FileSystem.GetDirectories(srcDirectory)
            Dim dirName As String = New DirectoryInfo(foundDirectory).Name
            'MessageBox.Show(dirName)
            Dim fp As String = My.Computer.FileSystem.CombinePath(destDirectory, dirName)
            'Dim dirlist() As String = System.IO.Directory.GetDirectories(srcDirectory)

            'For Each dire In dirlist
            'Dim dir() = foundDirectory.Split("\")
            'Dim fp = destDirectory & dir.Last
            Copying(foundDirectory, fp)
            ' Next

        Next
        For Each foundFile As String In _
            My.Computer.FileSystem.GetFiles(srcDirectory)
            Dim check As String = _
            System.IO.Path.GetExtension(foundFile)
            If (TestArray.Contains(Path.GetExtension(foundFile))) Then
                Continue For
            Else
                Dim fullPath As String
                fullPath = My.Computer.FileSystem.CombinePath(destDirectory, My.Computer.FileSystem.GetName(foundFile))
                'My.Computer.FileSystem.CreateDirectory(fullPath)
                My.Computer.FileSystem.CopyFile(foundFile, fullPath, showUI:=FileIO.UIOption.AllDialogs)
                ProgressBar1.Visible = True
                Timer1.Enabled = True
            End If

        Next

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Dim srcDirectory, destDirectory As String
        srcDirectory = TextBox1.Text
        destDirectory = TextBox2.Text



        If (My.Computer.FileSystem.DirectoryExists(destDirectory)) Then
            If (My.Computer.FileSystem.DirectoryExists(srcDirectory)) Then
                If (My.Computer.FileSystem.DirectoryExists(destDirectory)) Then

                    Copying(srcDirectory, destDirectory)

                End If

            Else
                MessageBox.Show(srcDirectory + "Does not exist")
                ProgressBar1.Visible = False
                ProgressBar1.Value = ProgressBar1.Maximum - 1
            End If
        Else
            My.Computer.FileSystem.CreateDirectory(destDirectory)
            If (My.Computer.FileSystem.DirectoryExists(destDirectory)) Then
                Copying(srcDirectory, destDirectory)
            End If

        End If
    End Sub
    Private Sub DisplayMessageBoxText()

        MessageBox.Show("Successfully Copied.")
    End Sub
    Private Sub TableLayoutPanel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles TableLayoutPanel1.Paint


    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ProgressBar1.Visible = False
    End Sub
    Private Function CompareMethod(ByVal TestArray As String, ByVal p2 As String) As Integer
        Throw New NotImplementedException
    End Function
    Private Sub ProgressBar1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProgressBar1.Click
        ProgressBar1.Value = 0
        ProgressBar1.Maximum = 100
        ProgressBar1.Minimum = 0
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Using fld As New FolderBrowserDialog()
            If fld.ShowDialog() = Windows.Forms.DialogResult.OK Then
                TextBox1.Text = fld.SelectedPath

            End If
        End Using
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Using fld As New FolderBrowserDialog()
            If fld.ShowDialog() = Windows.Forms.DialogResult.OK Then
                TextBox2.Text = fld.SelectedPath

            End If
        End Using
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        ProgressBar1.Value += 1

        If ProgressBar1.Value >= ProgressBar1.Maximum Then
            Timer1.Enabled = False
            DisplayMessageBoxText()
            ProgressBar1.Visible = False
        End If
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub
End Class

