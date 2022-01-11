Imports System.IO
Imports System.Text.RegularExpressions
Public Class Main

#Region " Declare"
    Private ReadOnly myLocalDb As String = Environment.CurrentDirectory & "\" & "database.db"
    Private passwordok As Boolean = False
    Private modify As Boolean = False
    Private ReadOnly randomLenght As Integer = 50
#End Region


#Region " Msgbox info alert error"
    ' cred
    ReadOnly passsecretalert As String = "You didn't insert the secret key for encoding."
    ReadOnly descriptionalert As String = "You didn't insert the description."
    ReadOnly emailalert As String = "Make sure that you correctly insert your e-mail."
    ReadOnly useralert As String = "You didn't insert your username."
    ReadOnly passalert As String = "You didn't insert your password."
    ReadOnly decriptalert As String = "The secret key For decode must be equal To that used for encoding."
    ReadOnly deletealert As String = "Are you sure to want to delete the selected items?"
    ReadOnly deltitle As String = "Delete"
    ' txt
    ReadOnly txtkeyalert As String = "Please generate a random key, or insert your key."
    ReadOnly txtempityalert As String = "Please insert the text to encode."
    ReadOnly txtempitydecodealert As String = "Please insert the encoded text."
    ReadOnly txtsavedalert As String = "The text was successfully saved."
    ReadOnly txtsaveerror As String = "Error to save the encoded text."
    ReadOnly txtloaderror As String = "Error to load the encoded text."
    ReadOnly nottext As String = "Not a text file (.txt)"
    ' file
    ReadOnly filecripted As String = "The file was encoded successfully."
    ReadOnly filedecrypted As String = "The file was decoded successfully."

#End Region

#Region " Load close me"
    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Try
            ' load from the file
            If File.Exists(myLocalDb) Then
                Dim myPswFileLines() As String = IO.File.ReadAllLines(myLocalDb)
                For Each line As String In myPswFileLines
                    Dim lineArray() As String = line.Split(CChar("#"))
                    Dim newItem As New ListViewItem(lineArray(0), 0)
                    newItem.SubItems.AddRange(New String() {lineArray(1), lineArray(2), lineArray(3), lineArray(4)})
                    LVPsw.Items.Add(newItem)
                Next
            End If
        Catch ex As Exception
            Me.TopMost = False
            MsgBox("No data found!", MsgBoxStyle.Information)
            Me.TopMost = True
        End Try

        If LVPsw.Items.Count > 0 Then
            SetListViewColumnSizes(LVPsw, -1)
        End If
        cbPasswordType.SelectedIndex = 0
        pbFile.AllowDrop = True
        Me.MinimumSize = New Size(Me.Width, Me.Height) 'add min form size
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        SaveAllData()
        FadeOut()
        Process.GetCurrentProcess.Kill()
    End Sub

    Public Sub SaveAllData()
        Try
            ' on form close save all data
            Dim myWriter As New IO.StreamWriter(myLocalDb)
            For Each myItem As ListViewItem In LVPsw.Items
                myWriter.WriteLine(myItem.Text & "#" &
                                   myItem.SubItems(1).Text & "#" &
                                   myItem.SubItems(2).Text & "#" &
                                   myItem.SubItems(3).Text & "#" &
                                   myItem.SubItems(4).Text & "#")
                'write Item and SubItem.
            Next
            myWriter.Close()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub CloseMeForm()
        Process.GetCurrentProcess.Kill()
    End Sub

    Private Sub FadeOut()
        'Just some simple opacity animation.
        Dim cntr As Integer
        For cntr = 90 To 10 Step -10
            Me.Opacity = cntr / 100
            Me.Refresh()
            Threading.Thread.Sleep(50)
        Next cntr
        Me.Dispose()
    End Sub
#End Region

#Region " Resize colum"
    Private Sub SetListViewColumnSizes(lvw As ListView, width As Integer)
        For Each col As ColumnHeader In lvw.Columns
            col.Width = width
        Next
    End Sub
#End Region

#Region " Clipboard get set"

    'set
    Private Sub btnClipUrl_Click(sender As Object, e As EventArgs) Handles btnClipUrl.Click
        If Not txtUrl.Text = "" Then Clipboard.SetText(txtUrl.Text)
    End Sub

    Private Sub btnClipDescription_Click(sender As Object, e As EventArgs) Handles btnClipDescription.Click
        If Not txtDescription.Text = "" Then Clipboard.SetText(txtDescription.Text)
    End Sub

    Private Sub btnClipEmail_Click(sender As Object, e As EventArgs) Handles btnClipEmail.Click
        If Not txtEmail.Text = "" Then Clipboard.SetText(txtEmail.Text)
    End Sub

    Private Sub btnClipUsername_Click(sender As Object, e As EventArgs) Handles btnClipUsername.Click
        If Not txtUsername.Text = "" Then Clipboard.SetText(txtUsername.Text)
    End Sub

    Private Sub btnClipPassword_Click(sender As Object, e As EventArgs) Handles btnClipPassword.Click
        If Not txtPassword.Text = "" Then Clipboard.SetText(txtPassword.Text)
    End Sub

    Private Sub btnClipSecret_Click(sender As Object, e As EventArgs) Handles btnClipSecret.Click
        If Not txtSecret.Text = "" Then Clipboard.SetText(txtSecret.Text)
    End Sub

    Private Sub CopiaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles tsmCopy.Click
        If Not txtUrl.Text = "" Then Clipboard.SetText(txtUrl.Text)
    End Sub

    'get
    Private Sub tsmPast_Click(sender As Object, e As EventArgs) Handles tsmPast.Click
        If Not txtUrl.Text = "" Then txtUrl.Text = Clipboard.GetText
    End Sub

#End Region

#Region " Clean data"
    Private Sub btnCleanDescription_Click(sender As Object, e As EventArgs) Handles btnCleanDescription.Click
        txtDescription.Text = ""
    End Sub

    Private Sub btnCleanEmail_Click(sender As Object, e As EventArgs) Handles btnCleanEmail.Click
        txtEmail.Text = ""
    End Sub

    Private Sub btnCleanUrl_Click(sender As Object, e As EventArgs) Handles btnCleanUrl.Click
        txtUrl.DataBindings.Clear()
        txtUrl.Text = String.Empty
        txtUrl.Text = ""
    End Sub

    Private Sub btnCleanUsername_Click(sender As Object, e As EventArgs) Handles btnCleanUsername.Click
        txtUsername.Text = ""
    End Sub

    Private Sub btnCleanPassword_Click(sender As Object, e As EventArgs) Handles btnCleanPassword.Click
        txtPassword.Text = ""
    End Sub

#End Region

#Region " Secret random generator"
    Private Sub btnRandomSecret_Click(sender As Object, e As EventArgs) Handles btnRandomSecret.Click
        txtSecret.Text = RandomKeyGenerator.RandomStringNumber(randomLenght)
    End Sub

    Private Sub txtSecret_TextChanged(sender As Object, e As EventArgs) Handles txtSecret.TextChanged
        Dim s As String = txtSecret.Text
        txtSecret.UseSystemPasswordChar = False
        If s.Length < randomLenght Then
            txtSecret.UseSystemPasswordChar = True
        End If
    End Sub

#End Region

#Region " Btn encode add"

    Private Sub btnAddEncode_Click(sender As System.Object, e As System.EventArgs) Handles btnAddEncode.Click
        Try
            ' encode text
            If txtSecret.Text = "" Then
                Me.TopMost = False
                MsgBox(passsecretalert, MsgBoxStyle.Exclamation)
                Me.TopMost = True
                Exit Sub
            ElseIf txtDescription.Text = "" Then
                Me.TopMost = False
                MsgBox(descriptionalert, MsgBoxStyle.Exclamation)
                Me.TopMost = True
                Exit Sub
            ElseIf txtUsername.Text = "" Then
                Me.TopMost = False
                MsgBox(useralert, MsgBoxStyle.Exclamation)
                Me.TopMost = True
                Exit Sub
            ElseIf txtPassword.Text = "" Then
                Me.TopMost = False
                MsgBox(passalert, MsgBoxStyle.Exclamation)
                Me.TopMost = True
                Exit Sub
            Else

                If Not txtEmail.Text = "" Then
                    If Not validateEmail(txtEmail.Text) Then
                        Me.TopMost = False
                        MsgBox(emailalert, MsgBoxStyle.Exclamation)
                        Me.TopMost = True
                        Exit Sub
                    End If
                End If

                If modify Then
                    'modify item
                    txtUsername.Text = Cryptology.Rijndaelcrypt(txtUsername.Text, txtSecret.Text)
                    txtPassword.Text = Cryptology.Rijndaelcrypt(txtPassword.Text, txtSecret.Text)
                    LVPsw.SelectedItems(0).SubItems(3).Text = txtUsername.Text
                    LVPsw.SelectedItems(0).SubItems(4).Text = txtPassword.Text
                Else
                    'add new item
                    txtUsername.Text = Cryptology.Rijndaelcrypt(txtUsername.Text, txtSecret.Text)
                    txtPassword.Text = Cryptology.Rijndaelcrypt(txtPassword.Text, txtSecret.Text)
                    LVPsw.Items.Add(txtDescription.Text, 0)
                    LVPsw.Items(LVPsw.Items.Count - 1).SubItems.Add(txtEmail.Text)
                    LVPsw.Items(LVPsw.Items.Count - 1).SubItems.Add(txtUrl.Text)
                    LVPsw.Items(LVPsw.Items.Count - 1).SubItems.Add(txtUsername.Text)
                    LVPsw.Items(LVPsw.Items.Count - 1).SubItems.Add(txtPassword.Text)
                    Reset()
                End If


                If LVPsw.Items.Count > 0 Then
                    SetListViewColumnSizes(LVPsw, -1)
                End If

            End If
        Catch ex As Exception
        End Try
    End Sub
    Public Function validateEmail(ByVal emailAddress As String) As Boolean
        Dim email As New Regex("([\w-+]+(?:\.[\w-+]+)*@(?:[\w-]+\.)+[a-zA-Z]{2,7})")
        If email.IsMatch(emailAddress) Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Sub Reset()
        'clean text 
        txtDescription.Text = ""
        txtEmail.Text = ""
        txtUrl.Text = ""
        txtUsername.Text = ""
        txtPassword.Text = ""
    End Sub
#End Region

#Region " Search items"
    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Try
            Dim found As Boolean = False

            LVPsw.SelectedIndices.Clear()
            For Each lvi As ListViewItem In LVPsw.Items
                lvi.BackColor = Color.White
                lvi.ForeColor = Color.Black
                For Each lvisub As ListViewItem.ListViewSubItem In lvi.SubItems
                    If lvisub.Text = txtSearchList.Text Then
                        LVPsw.SelectedIndices.Add(lvi.Index)
                        lvi.BackColor = Color.LightYellow
                        lvi.ForeColor = Color.Black
                        found = True
                        Exit For
                    End If
                Next
            Next
            LVPsw.Focus()

            If Not found Then
                Me.TopMost = False
                MsgBox("Not found!", MsgBoxStyle.Critical)
                Me.TopMost = True
            End If

        Catch ex As Exception
        End Try
    End Sub
#End Region

#Region " Modify"
    Private Sub cbModifica_CheckedChanged(sender As Object, e As EventArgs) Handles cbModifica.CheckedChanged

        If LVPsw.SelectedItems.Count > 0 Then

            If cbModifica.Checked Then
                ' decode the seleceted from the list
                Try
                    If LVPsw.SelectedItems.Count > 0 Then
                        If txtSecret.Text = "" Then
                            Me.TopMost = False
                            MsgBox(decriptalert, MsgBoxStyle.Exclamation)
                            Me.TopMost = True
                            passwordok = False
                            cbModifica.Checked = False
                            Exit Sub
                        Else
                            passwordok = True

                            Dim decryptedusername As String = Cryptology.RijndaelDecrypt(LVPsw.FocusedItem.SubItems(3).Text, txtSecret.Text)
                            Dim decryptedpassword As String = Cryptology.RijndaelDecrypt(LVPsw.FocusedItem.SubItems(4).Text, txtSecret.Text)

                            txtUsername.Text = decryptedusername
                            txtPassword.Text = decryptedpassword

                            LVPsw.SelectedItems(0).SubItems(3).Text = txtUsername.Text
                            LVPsw.SelectedItems(0).SubItems(4).Text = txtPassword.Text

                            modifyTextBox(True)
                            modify = True
                            txtUrl.DetectUrls = False
                            LVPsw.Enabled = False
                        End If
                    End If
                Catch ex As Exception
                    MsgBox("Err: " + ex.Message)
                End Try
            Else
                ' decode the selected
                Try
                    If passwordok = True Then
                        If LVPsw.SelectedItems.Count > 0 Then
                            If txtSecret.Text = "" Then
                                Me.TopMost = False
                                MsgBox(decriptalert, MsgBoxStyle.Exclamation)
                                Me.TopMost = True
                                Exit Sub
                            Else
                                'modify item
                                txtUsername.Text = Cryptology.Rijndaelcrypt(txtUsername.Text, txtSecret.Text)
                                txtPassword.Text = Cryptology.Rijndaelcrypt(txtPassword.Text, txtSecret.Text)

                                LVPsw.SelectedItems(0).SubItems(3).Text = txtUsername.Text
                                LVPsw.SelectedItems(0).SubItems(4).Text = txtPassword.Text

                                modifyTextBox(False)
                                modify = False
                                txtUrl.DetectUrls = True
                                LVPsw.Enabled = True
                            End If
                        End If
                    End If

                Catch ex As Exception
                    MsgBox("Err: " + ex.Message)
                End Try

            End If
        End If
    End Sub
    Private Sub modifyTextBox(m As Boolean)
        Select Case m
            Case True
                setColor(Color.LightYellow, Color.Black)
            Case False
                setColor(Color.White, Color.Black)
        End Select
        ' enable / disable
        btnDelItem.Enabled = Not m
        btnAddEncode.Enabled = Not m
        btnSearch.Enabled = Not m
        btnRandomSecret.Enabled = Not m
        txtSecret.Enabled = Not m
    End Sub

    Private Sub setColor(backColor As Color, foreColor As Color)
        ' backColor
        txtDescription.BackColor = backColor
        txtEmail.BackColor = backColor
        txtUrl.BackColor = backColor
        txtUsername.BackColor = backColor
        txtPassword.BackColor = backColor
        ' ForeColor
        txtDescription.ForeColor = foreColor
        txtEmail.ForeColor = foreColor
        txtUrl.ForeColor = foreColor
        txtUsername.ForeColor = foreColor
        txtPassword.ForeColor = foreColor
    End Sub

    Private Sub txtdescrizione_TextChanged(sender As Object, e As EventArgs) Handles txtDescription.TextChanged
        If LVPsw.SelectedItems.Count > 0 Then
            If modify Then LVPsw.SelectedItems(0).Text = txtDescription.Text
        End If
    End Sub
    Private Sub txtEmail_TextChanged(sender As Object, e As EventArgs) Handles txtEmail.TextChanged
        If LVPsw.SelectedItems.Count > 0 Then
            If modify Then LVPsw.SelectedItems(0).SubItems(1).Text = txtEmail.Text
        End If
    End Sub
    Private Sub txtWebSite_TextChanged(sender As Object, e As EventArgs) Handles txtUrl.TextChanged
        If LVPsw.SelectedItems.Count > 0 Then
            If modify Then LVPsw.SelectedItems(0).SubItems(2).Text = txtUrl.Text
        End If
    End Sub
    Private Sub txtUsername_TextChanged(sender As Object, e As EventArgs) Handles txtUsername.TextChanged
        If LVPsw.SelectedItems.Count > 0 Then
            If modify Then LVPsw.SelectedItems(0).SubItems(3).Text = txtUsername.Text
        End If
    End Sub
    Private Sub txtPassword_TextChanged(sender As Object, e As EventArgs) Handles txtPassword.TextChanged
        If LVPsw.SelectedItems.Count > 0 Then
            If modify Then LVPsw.SelectedItems(0).SubItems(4).Text = txtPassword.Text
        End If
    End Sub
#End Region

#Region " Delete"
    Private Sub btnDelItem_Click(sender As Object, e As EventArgs) Handles btnDelItem.Click
        Try
            If LVPsw.SelectedItems.Count > 0 Then
                Dim result = MessageBox.Show(deletealert, deltitle, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If result = DialogResult.Yes Then
                    For Each item As ListViewItem In LVPsw.SelectedItems
                        LVPsw.Items.Remove(item)
                    Next
                Else : Exit Sub
                End If
            End If
            txtDescription.Text = ""
            txtEmail.Text = ""
            txtUrl.Text = ""
            txtUsername.Text = ""
            txtPassword.Text = ""
        Catch ex As Exception
        End Try
    End Sub

#End Region

#Region " Listview manage"
    Private Sub LVPsw_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LVPsw.SelectedIndexChanged
        If LVPsw.SelectedItems.Count > 0 Then
            If cbModifica.Checked Then
                cbModifica.Checked = False
            End If

            Dim description As String = LVPsw.SelectedItems(0).Text
            Dim email As String = LVPsw.SelectedItems(0).SubItems(1).Text
            Dim web As String = LVPsw.SelectedItems(0).SubItems(2).Text
            Dim user As String = LVPsw.SelectedItems(0).SubItems(3).Text
            Dim pass As String = LVPsw.SelectedItems(0).SubItems(4).Text

            txtDescription.Text = description
            txtEmail.Text = email

            txtUrl.DataBindings.Clear()
            txtUrl.Text = String.Empty
            txtUrl.Text = web

            txtUsername.Text = user
            txtPassword.Text = pass

            If txtSecret.Text IsNot "" Then
                ' try to decode
                Try
                    If LVPsw.SelectedItems.Count > 0 Then
                        If txtSecret.Text = "" Then
                            Me.TopMost = False
                            MsgBox(decriptalert, MsgBoxStyle.Exclamation)
                            Me.TopMost = True
                            Exit Sub
                        Else

                            Dim decryptedusername As String = ""
                            Dim decryptedpassword As String = ""

                            decryptedusername = Cryptology.RijndaelDecrypt(LVPsw.FocusedItem.SubItems(3).Text, txtSecret.Text)
                            decryptedpassword = Cryptology.RijndaelDecrypt(LVPsw.FocusedItem.SubItems(4).Text, txtSecret.Text)

                            txtUsername.Text = decryptedusername
                            txtPassword.Text = decryptedpassword

                            If modify Then
                                LVPsw.SelectedItems(0).SubItems(3).Text = txtUsername.Text
                                LVPsw.SelectedItems(0).SubItems(4).Text = txtPassword.Text
                            End If

                        End If
                    End If
                Catch ex As Exception
                    MsgBox("Err: " + ex.Message)
                End Try
            End If
        End If
    End Sub

    Private Sub txtWebSite_LinkClicked(sender As Object, e As LinkClickedEventArgs) Handles txtUrl.LinkClicked
        If modify Then Return
        Try
            Process.Start(e.LinkText)
        Catch ex As Exception
        End Try
    End Sub
#End Region


#Region " .txt encode decode"
    Private Sub btnTextLoad_Click(sender As Object, e As EventArgs) Handles btnTextLoad.Click
        Try
            Dim o As New OpenFileDialog With {.Filter = "Text Files (*.txt)|*.txt"}
            With o
                If .ShowDialog = Windows.Forms.DialogResult.OK Then
                    Try
                        Dim txt_enc As String = .FileName
                        Dim objreader As New IO.StreamReader(txt_enc)
                        txtEncodeDecodeText.Text = objreader.ReadToEnd
                        objreader.Close()
                    Catch
                        Me.TopMost = False
                        MsgBox(txtloaderror, MsgBoxStyle.Critical)
                        Me.TopMost = True
                    End Try
                End If
            End With
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnRandom_Click(sender As Object, e As EventArgs) Handles btnTextRandomSecret.Click
        txtTextSecret.Text = RandomKeyGenerator.RandomStringNumber(randomLenght)
    End Sub

    Private Sub txtTextSecret_TextChanged(sender As Object, e As EventArgs) Handles txtTextSecret.TextChanged
        Dim s As String = txtTextSecret.Text
        If s.Length < randomLenght Then
            txtTextSecret.UseSystemPasswordChar = True
        Else
            txtTextSecret.UseSystemPasswordChar = False
        End If
    End Sub

    Private Sub btnClipText_Click(sender As Object, e As EventArgs) Handles btnClipText.Click
        If Not txtTextSecret.Text = "" Then Clipboard.SetText(txtTextSecret.Text)
    End Sub

    Private Sub btnEncodeText_Click(sender As Object, e As EventArgs) Handles btnEncodeText.Click
        Try
            Dim s As String = txtEncodeDecodeText.Text

            Dim txt As String = Nothing
            If txtTextSecret.Text = "" Then
                Me.TopMost = False
                MsgBox(txtkeyalert, MsgBoxStyle.Exclamation)
                Me.TopMost = True
                Exit Sub
            ElseIf txtEncodeDecodeText.Text = "" Then
                Me.TopMost = False
                MsgBox(txtempityalert, MsgBoxStyle.Exclamation)
                Me.TopMost = True
                Exit Sub
            End If
            txt = Cryptology.Rijndaelcrypt(txtEncodeDecodeText.Text, txtTextSecret.Text) : txtEncodeDecodeText.Text = txt
        Catch ex As Exception
        End Try
    End Sub

    Private Sub bbtnDecodeText_Click(sender As Object, e As EventArgs) Handles btnDecodeText.Click
        Try
            Dim txt As String = Nothing
            If txtTextSecret.Text = "" Then
                Me.TopMost = False
                MsgBox(decriptalert, MsgBoxStyle.Exclamation)
                Me.TopMost = True
                Exit Sub
            ElseIf txtEncodeDecodeText.Text = "" Then
                Me.TopMost = False
                MsgBox(txtempitydecodealert, MsgBoxStyle.Exclamation)
                Me.TopMost = True
                Exit Sub
            End If
            txt = Cryptology.RijndaelDecrypt(txtEncodeDecodeText.Text, txtTextSecret.Text) : txtEncodeDecodeText.Text = txt
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnSaveText_Click(sender As Object, e As EventArgs) Handles btnSaveText.Click
        Try
            If txtEncodeDecodeText.Text = "" Then
                Exit Sub
            End If
            Dim Filename As String
            Dim save As New SaveFileDialog
            With save
                .DefaultExt = "txt"
                .FileName = "Encoded" & String.Format("{0:_dd-M-yyyy_hh-mm-ss}", DateTime.Now)
                .Filter = "txt files (*.txt)|*.txt"
                .FilterIndex = 1
                .OverwritePrompt = True
                .Title = "Save File Dialog"
            End With

            If save.ShowDialog = Windows.Forms.DialogResult.OK Then
                Filename = save.FileName
                Try
                    Dim sw As StreamWriter = File.CreateText(Filename)
                    For i As Integer = 0 To txtEncodeDecodeText.Lines.Length - 1
                        sw.WriteLine(txtEncodeDecodeText.Lines(i))
                    Next
                    sw.Flush()
                    sw.Close()
                    Me.TopMost = False
                    MsgBox(txtsavedalert, MsgBoxStyle.Information)
                    Me.TopMost = True
                Catch ex As Exception
                    Me.TopMost = False
                    MsgBox(txtsaveerror, MsgBoxStyle.Critical)
                    Me.TopMost = True
                End Try
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnCleanText_Click(sender As Object, e As EventArgs) Handles btnCleanText.Click
        txtEncodeDecodeText.Text = ""
    End Sub
#End Region

#Region " .txt drag and drop"
    Private Sub txtEncodeDecodeText_DragEnter(sender As Object, e As DragEventArgs) Handles txtEncodeDecodeText.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        End If
    End Sub

    Private Sub txtEncodeDecodeText_DragDrop(sender As Object, e As DragEventArgs) Handles txtEncodeDecodeText.DragDrop
        Dim file As String() = DirectCast(e.Data.GetData(DataFormats.FileDrop, True), String())
        Try
            Dim extension As String
            extension = Path.GetExtension(file(0))
            If (extension <> ".txt") Then
                Me.TopMost = False
                MsgBox(nottext, MsgBoxStyle.Critical)
                Me.TopMost = True
                Exit Sub
            Else
                Try
                    Dim txt_enc As String = file(0)
                    Dim objreader As New IO.StreamReader(txt_enc)
                    txtEncodeDecodeText.Text = objreader.ReadToEnd
                    objreader.Close()
                Catch
                    Me.TopMost = False
                    MsgBox(txtloaderror, MsgBoxStyle.Critical)
                    Me.TopMost = True
                End Try
            End If
        Catch ex As Exception
        End Try
    End Sub
#End Region


#Region " .file encode decode"
    Private Sub btnFileRandomSecret_Click(sender As Object, e As EventArgs) Handles btnFileRandomSecret.Click
        txtFileSecret.Text = RandomKeyGenerator.RandomStringNumber(randomLenght)
    End Sub

    Private Sub txtFileSecret_TextChanged(sender As Object, e As EventArgs) Handles txtFileSecret.TextChanged
        Dim s As String = txtFileSecret.Text
        If s.Length < randomLenght Then
            txtFileSecret.UseSystemPasswordChar = True
        Else
            txtFileSecret.UseSystemPasswordChar = False
        End If
    End Sub

    Private Sub btnClipFile_Click(sender As Object, e As EventArgs) Handles btnClipFile.Click
        If Not txtFileSecret.Text = "" Then Clipboard.SetText(txtFileSecret.Text)
    End Sub

    Private Sub btnEncodeFile_Click(sender As Object, e As EventArgs) Handles btnEncodeFile.Click
        If txtFileSecret.Text = Nothing Then
            Me.TopMost = False
            MsgBox(passsecretalert, MsgBoxStyle.Exclamation)
            Me.TopMost = True
            Exit Sub
        Else
            Dim open As New OpenFileDialog
            If open.ShowDialog <> DialogResult.Cancel Then
                Dim file As String = open.FileName
                Dim fpath As String = Path.GetDirectoryName(file)
                Dim ext As String = Path.GetExtension(file)
                Dim filen As String = Path.GetFileNameWithoutExtension(file)
                Dim infoReader As FileInfo = My.Computer.FileSystem.GetFileInfo(file)
                Dim filesiz As Long = infoReader.Length
                Cryptology.EncryptFile(file, fpath & "\" & filen & "#Encoded" & ext, txtFileSecret.Text)
                If Cryptology.Encoded Then
                    Me.TopMost = False
                    MsgBox(filecripted, MsgBoxStyle.Information)
                    Me.TopMost = True
                    'Process.Start(fpath)
                End If
            End If

        End If
    End Sub

    Private Sub btnDecodeFile_Click(sender As Object, e As EventArgs) Handles btnDecodeFile.Click
        If txtFileSecret.Text = Nothing Then
            Me.TopMost = False
            MsgBox(decriptalert, MsgBoxStyle.Critical)
            Me.TopMost = True
            Exit Sub
        Else
            Dim open As New OpenFileDialog
            If open.ShowDialog <> DialogResult.Cancel Then
                Dim file As String = open.FileName
                Dim fpath As String = Path.GetDirectoryName(file)
                Dim ext As String = Path.GetExtension(file)
                Dim filen As String = Path.GetFileNameWithoutExtension(file)
                Dim splitfn As String() = filen.Split(CChar("#"))
                Cryptology.DecryptFile(file, fpath & "\" & splitfn(0) & "#Decoded" & ext, txtFileSecret.Text)
                If Cryptology.Decoded Then
                    Me.TopMost = False
                    MsgBox(filedecrypted, MsgBoxStyle.Information)
                    Me.TopMost = True
                    'Process.Start(fpath)
                End If
            End If
        End If
    End Sub
#End Region

#Region " .file drag and drop"
    Private Function GetFileName(path As String) As String
        Return IO.Path.GetFileNameWithoutExtension(path)
    End Function

    Private Sub pbFile_DragEnter(sender As Object, e As DragEventArgs) Handles pbFile.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop, False) Then
            e.Effect = DragDropEffects.All
        End If
    End Sub
    Private Sub pbFile_DragDrop(sender As Object, e As DragEventArgs) Handles pbFile.DragDrop
        If txtFileSecret.Text = Nothing Then
            Me.TopMost = False
            MsgBox(passsecretalert, MsgBoxStyle.Exclamation)
            Me.TopMost = True
            Exit Sub
        Else
            Dim droppedfile As String() = CType(e.Data.GetData(DataFormats.FileDrop), String())
            For Each files In droppedfile
                Dim file As String = droppedfile(0)
                Dim filename As String = GetFileName(file)
                Dim fpath As String = Path.GetDirectoryName(file)
                Dim ext As String = Path.GetExtension(file)
                Dim filen As String = Path.GetFileNameWithoutExtension(file)
                Dim infoReader As FileInfo = My.Computer.FileSystem.GetFileInfo(file)
                Dim filesiz As Long = infoReader.Length
                Cryptology.EncryptFile(file, fpath & "\" & filen & "#Encoded" & ext, txtFileSecret.Text)
                If Cryptology.Encoded Then
                    Me.TopMost = False
                    MsgBox(filecripted, MsgBoxStyle.Information)
                    Me.TopMost = True
                    'Process.Start(fpath)
                End If
            Next
        End If
    End Sub
#End Region


#Region " Random password generator with params"
    Private Sub btnParamsPwdGenerate_Click(sender As Object, e As EventArgs) Handles btnParamsPwdGenerate.Click
        txtParamsPassword.Text = RandomKeyGenerator.RandomStringParams(CInt(nchar.Value))
    End Sub
    Private Sub btnClipParamsPwd_Click(sender As Object, e As EventArgs) Handles btnClipParamsPwd.Click
        If Not txtParamsPassword.Text = "" Then Clipboard.SetText(txtParamsPassword.Text)
    End Sub

#End Region

#Region " Random pool"
    Private Sub RandomPool3_CharacterSelection(s As Object, c As Char) Handles RandomPool.CharacterSelection
        txtRandomPassword.AppendText(c)
    End Sub
    Private Sub btnCleanRandomPwd_Click(sender As Object, e As EventArgs) Handles btnCleanRandomPwd.Click
        txtRandomPassword.Text = ""
    End Sub
    Private Sub btnClipRandomPwd_Click(sender As Object, e As EventArgs) Handles btnClipRandomPwd.Click
        If Not txtRandomPassword.Text = "" Then Clipboard.SetText(txtRandomPassword.Text)
    End Sub
#End Region

#Region " Random Password generator list"
    Private Sub btnPasswordListGenerate_Click(sender As Object, e As EventArgs) Handles btnPasswordListGenerate.Click
        Select Case cbPasswordType.SelectedIndex
            Case 0
                For i As Integer = 1 To CInt(nupPwdCount.Value)
                    txtPwdLists.Text = txtPwdLists.Text & RandomKeyGenerator.RandomStringParams(CInt(pll.Value)) & vbNewLine
                Next
            Case 1
                For i As Integer = 1 To CInt(nupPwdCount.Value)
                    txtPwdLists.Text = txtPwdLists.Text & RandomKeyGenerator.RandomStringNumberSpecialChar(CInt(pll.Value)) & vbNewLine
                Next
        End Select
    End Sub
    Private Sub btnCleanPasswordList_Click(sender As Object, e As EventArgs) Handles btnCleanPasswordList.Click
        txtPwdLists.Text = ""
    End Sub

    Private Sub txtPwdLists_TextChanged(sender As Object, e As EventArgs) Handles txtPwdLists.TextChanged
        txtPwdLists.SelectionStart = txtPwdLists.TextLength : txtPwdLists.ScrollToCaret()
    End Sub

    Private Sub btnSavePasswordList_Click(sender As Object, e As EventArgs) Handles btnSavePasswordList.Click
        Try
            If txtPwdLists.Text = "" Then
                Exit Sub
            End If
            Dim Filename As String
            Dim save As New SaveFileDialog
            With save
                .DefaultExt = "txt"
                .FileName = "PasswordLists" & String.Format("{0:_dd-M-yyyy_hh-mm-ss}", DateTime.Now)
                .Filter = "txt files (*.txt)|*.txt"
                .FilterIndex = 1
                .OverwritePrompt = True
                .Title = "Save File Dialog"
            End With

            If save.ShowDialog = Windows.Forms.DialogResult.OK Then
                Filename = save.FileName
                Try
                    Dim sw As StreamWriter = File.CreateText(Filename)
                    For i As Integer = 0 To txtPwdLists.Lines.Length - 1
                        sw.WriteLine(txtPwdLists.Lines(i))
                    Next
                    sw.Flush()
                    sw.Close()
                    Me.TopMost = False
                    MsgBox(txtsavedalert, MsgBoxStyle.Information)
                    Me.TopMost = True
                Catch ex As Exception
                    Me.TopMost = False
                    MsgBox(txtsaveerror, MsgBoxStyle.Critical)
                    Me.TopMost = True
                End Try
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnClipPasswordList_Click(sender As Object, e As EventArgs) Handles btnClipPasswordList.Click
        If Not txtPwdLists.Text = "" Then Clipboard.SetText(txtPwdLists.Text)
    End Sub
#End Region


#Region " About"
    Private Sub RichTextBox1_LinkClicked(sender As Object, e As LinkClickedEventArgs) Handles RichTextBox1.LinkClicked
        Process.Start(e.LinkText)
    End Sub
#End Region

End Class
