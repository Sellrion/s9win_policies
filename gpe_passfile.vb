Imports System.IO
Imports System.Text
Imports System.Object

Public Class gpe_passfile

    Private ReadOnly CRLF As String = Environment.NewLine   'Represents new line symbol
    Private ReadOnly MAX_PASSFILE_SIZE As Integer = 100024  'Maximum size of passwords file
    Public ReadOnly MAX_PASSPHRASE_LENGTH As Integer = 14   'Maximum length of passphrase string
    Public ReadOnly MAX_USERNAME_LENGTH As Integer = 20     'Maximum length of username string
    Private PASSFILE As FileStream                          'Object that contains currently opened file
    Public ReadOnly PASSFILE_BUFFER() As Byte               'Buffer that contains file data
    Public objectstate As Integer                           'Object state flag
    Public Exeption As String                               'Last error string representation

    Sub New(ByVal fmode As FileMode, Optional ByVal path As String = "")
        'We have to know what we going to do with a file
        Select Case fmode
            Case FileMode.Open
                If path = "" Or Not File.Exists(path) Then
                    Dim ofd As New OpenFileDialog()
                    ofd.FileName = ""
                    ofd.Filter = "Текстовый документ (*.txt)|*.txt"
                    If ofd.ShowDialog() = DialogResult.OK Then
                        path = ofd.FileName
                    Else
                        Me.objectstate = 1
                        Exit Sub
                    End If
                End If

                'Here we have to check file size.
                'Just for them who likes to joke
                Dim fileinfo As New FileInfo(path)
                If fileinfo.Length > Me.MAX_PASSFILE_SIZE Then
                    Me.objectstate = 1
                    Me.Exeption = "Файл имеет слишком большой размер. Максимальный размер файла паролей - " & Me.MAX_PASSFILE_SIZE / 1024 & "Кб."
                    Exit Sub
                End If

                'Open file
                Try
                    Me.PASSFILE = File.Open(path, fmode, FileAccess.Read, FileShare.None)
                Catch ex As Exception
                    Me.objectstate = 1
                    Me.Exeption = "Не удается открыть файл: " & Convert.ToChar(10) & Convert.ToChar(13) & ex.Message()
                    Exit Sub
                End Try

                ReDim Me.PASSFILE_BUFFER(fileinfo.Length)
                Dim result As Integer = Me.PASSFILE.Read(Me.PASSFILE_BUFFER, 0, fileinfo.Length)

                'File is placed to the memory so we don't need FileStream object anymore
                Me.PASSFILE.Close()
            Case FileMode.Create
                'Here we prepare a file for saving
                If path = "" Then
                    Dim sfd As New SaveFileDialog()
                    sfd.Title = "Сохранить"
                    sfd.Filter = "Текстовый документ (*.txt)|*.txt"
                    sfd.FileName = ""
                    If sfd.ShowDialog() = DialogResult.OK Then
                        path = sfd.FileName
                    Else
                        Me.objectstate = 1
                        Exit Sub
                    End If
                End If

                'Try to create a file
                Try
                    Me.PASSFILE = File.Open(path, fmode, FileAccess.Write, FileShare.None)
                Catch ex As Exception
                    Me.Exeption = "Не удается создать файл файл: " & Convert.ToChar(10) & Convert.ToChar(13) & ex.Message()
                    Me.objectstate = 1
                    Exit Sub
                End Try
        End Select

        'All was fine
        Me.objectstate = 0
    End Sub

    Public Function parsePassFile() As Object
        If Me.objectstate = 1 Then Return Nothing

        'Get string
        Dim _sbuff As String = System.Text.Encoding.Default.GetSt<?xml version="1.0" encoding="utf-8"?>
<PendingCommit>
  <CommitComment />
  <WorkItems />
  <PinnedBranches />
  <PublishPrompt Enabled="True" />
</PendingCommit>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        If username = "" Then Return False

        Dim stopchars() As Char = {"""", "/", "\", "[", "]", ":", "|", "<", ">", "+", "=", ";", "?", "*"}
        Dim j As Integer

        If username.Length > Me.MAX_USERNAME_LENGTH Or username.Length = 0 Or username = "" Then Return False


        For j = 0 To stopchars.Length - 1
            If username.IndexOf(stopchars(j)) > -1 Then Return False
        Next j

        For j = 1 To 31
            If username.IndexOf(Convert.ToChar(j)) > -1 Then Return False
        Next j

        Return True
    End Function

    Private Function strRepeat(ByVal symbol As Char, ByVal length As Integer) As String
        Dim outstr As String = ""
        Dim i As Integer
        For i = 0 To length - 1
            outstr += symbol
        Next i
        Return outstr
    End Function

End Class
