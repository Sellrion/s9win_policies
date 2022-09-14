Public Class gpe_main

    Public policies As New gpe_policies
    Public profileApplying As Boolean = False
    Public loading As Boolean = True

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim p As Boolean = policies.shutdown_Privileges()
        If Not p Then Environment.Exit(1) Else Environment.Exit(0)
    End Sub

    Private Sub ControlBox_Close(sender As System.Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Dim p As Boolean = policies.shutdown_Privileges()
        If Not p Then Environment.Exit(1) Else Environment.Exit(0)
    End Sub

    Sub New()
        'Important method call
        InitializeComponent()

        'Handle errors during policies scheme building
        If policies.objectstatus = 1 Then
            MessageBox.Show(policies.Exeption, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Environment.Exit(1)
        End If

        'Fill the userlist interface
        Dim i As Integer = 0
        If policies.USERS.Length > 0 Then
            Do Until i = policies.USERS.Length
                Me.ComboBox1.Items.Add(policies.USERS(i).username)
                i = i + 1
            Loop
            'If we have more than one users, let's allow to edit policies to all of them at once
            If policies.USERS.Length > 1 Then Me.ComboBox1.Items.Add("Все пользователи")
        Else
            'If no users were found, then we can say that program is running under super-administrator account.
            'And there is no other accounts on that machine.
            'We can't allow user to edit policies of that acount. It's too dangerous.
            'But we can allow him to create more users.
            Me.ComboBox1.Items.Add(policies.CurrentUser)
            Me.ComboBox1.Enabled = False
            Me.TabControl1.Enabled = False
        End If

        'Fill the profiles list
        Me.ComboBox3.Items.Clear()
        i = 0
        Do Until i = policies.PROFILES.Length
            Me.ComboBox3.Items.Add(policies.PROFILES(i).p_name)
            i = i + 1
        Loop
        Me.ComboBox3.Items.Add("Другое")

        'Select a proper user in userlist
        If Me.ComboBox1.Items.Contains(policies.CurrentUser) Then
            Me.ComboBox1.SelectedItem = policies.CurrentUser
        Else
            Me.ComboBox1.SelectedIndex = 0
        End If

        'OS version compabillity
        'For users
        Dim SETTINGS As New Dictionary(Of String, gpe_policies.setting)
        SETTINGS = policies.getSettings()
        For i = 0 To Me.Panel1.Controls.Count - 1
            If SETTINGS.ContainsKey(Me.Panel1.Controls(i).Name) Then Me.Panel1.Controls(i).Enabled = policies.osverison_Compare(SETTINGS.Item(Me.Panel1.Controls(i).Name).minosversion)
        Next i

        'And for system
        For i = 0 To policies.SYS_POLICIES.Length - 1
            If Me.Panel2.Controls.ContainsKey(policies.SYS_POLICIES(i).controlname) Then Me.Panel2.Controls.Item(policies.SYS_POLICIES(i).controlname).Enabled = policies.osverison_Compare(policies.SYS_POLICIES(i).minosversion)
        Next i

        'Figure out current situation with a super administrator account on the system
        If policies.AdministratorState = "1" Then
            Me.ToolTip1.SetToolTip(Me.Button3, "Разблокировать встроенную учетную запись администратора.")
            Me.Button3.Enabled = True
            Me.Button3.BackColor = Color.LightCoral
        ElseIf policies.AdministratorState = "2" Then
            Me.ToolTip1.SetToolTip(Me.Button3, "Изменить пароль встроенной учетной записи администратора.")
            Me.Button3.Enabled = True
            Me.Button3.BackColor = Color.LightGreen
        Else
            Me.CheckBox21.Enabled = False
            Me.Button3.Enabled = False
        End If

        'Tooltips and other staff
        Me.ToolTip2.SetToolTip(Me.Button4, "Позволяет создать большое количество локальных пользователей в автоматическом режиме")
        Me.ToolTip3.SetToolTip(Me.Button5, "Позволяет изменить или удалить пароль большому количеству локальных пользователей в автоматическом режиме")
        Me.ToolTip4.SetToolTip(Me.Button6, "Позволяет восстановить пароли большому количеству локальных пользователей на основе ранее созданного файла паролей")
        Me.ToolTip5.SetToolTip(Me.Button7, "Позволяет пересортировать файлы паролей относительно имен пользователей")
        Me.ToolTip6.SetToolTip(Me.Button8, "Позволяет производить очистку учетных записей пользователей (удаление файлов, восстановление параметров по умолчанию)")

        'The apply button
        Me.Button2.Enabled = False

        'Version and copyrights
        Me.Label5.Text = "build " & My.Application.Info.Version.ToString() & " @ WinAPI" & Convert.ToChar(10) & Convert.ToChar(13) & My.Application.Info.Copyright

        'All done
        Me.loading = False
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If Me.ComboBox1.SelectedItem = "Все пользователи" Then
            Exit Sub
        End If

        Dim i As Integer = 0
        Dim j As Integer = 0
        Do Until i = policies.USERS.Length
            If policies.USERS(i).username = Me.ComboBox1.SelectedItem Then
                For j = 0 To policies.USERS(i).POLICIES.Length - 1
                    Select Case policies.USERS(i).POLICIES(j).controltype
                        Case 0
                            'Support for three state checkboxes
                            If Not Me.Panel1.Controls.Item(policies.USERS(i).POLICIES(j).controlname).ThreeState Then
                                Me.Panel1.Controls.Item(policies.USERS(i).POLICIES(j).controlname).Checked = policies.vBool(policies.USERS(i).POLICIES(j))
                            Else
                                Select Case policies.vBool(policies.USERS(i).POLICIES(j))
                                    Case True
                                        Me.Panel1.Controls.Item(policies.USERS(i).POLICIES(j).controlname).CheckState = CheckState.Checked
                                    Case False
                                        Me.Panel1.Controls.Item(policies.USERS(i).POLICIES(j).controlname).CheckState = CheckState.Unchecked
                                End Select
                            End If
                        Case 2
                            If Not policies.USERS(i).POLICIES(j).value = Nothing Then
                                Me.Panel1.Controls.Item(policies.USERS(i).POLICIES(j).controlname).SelectedIndex = Convert.ToInt32(policies.USERS(i).POLICIES(j).value)
                                'Me.ComboBox2.SelectedIndex = Convert.ToInt32(policies.USERS(i).POLICIES(j).value)
                            Else
                                Me.Panel1.Controls.Item(policies.USERS(i).POLICIES(j).controlname).SelectedIndex = 0
                            End If
                        Case 1
                            If Not policies.USERS(i).POLICIES(j).value = Nothing Then
                                Me.Panel1.Controls.Item(policies.USERS(i).POLICIES(j).controlname).Text = policies.USERS(i).POLICIES(j).value
                            Else
                                Me.Panel1.Controls.Item(policies.USERS(i).POLICIES(j).controlname).Text = ""
                            End If
                    End Select
                Next j
                Me.Button2.Enabled = False
                Exit Do
            End If
            i = i + 1
        Loop

        'Display current settings for system
        'This will run once on application startup
        If Me.loading Then
            For i = 0 To policies.SYS_POLICIES.Length - 1
                Select Case policies.SYS_POLICIES(i).controltype
                    Case 0
                        'Support for three state checkboxes
                        If Not Me.Panel2.Controls.Item(policies.SYS_POLICIES(i).controlname).ThreeState Then
                            Me.Panel2.Controls.Item(policies.SYS_POLICIES(i).controlname).Checked = policies.vBool(policies.SYS_POLICIES(i))
                        Else
                            Select Case policies.vBool(policies.SYS_POLICIES(i))
                                Case True
                                    Me.Panel2.Controls.Item(policies.SYS_POLICIES(i).controlname).CheckState = CheckState.Checked
                                Case Else
                                    Me.Panel2.Controls.Item(policies.SYS_POLICIES(i).controlname).CheckState = CheckState.Unchecked
                            End Select
                        End If
                    Case 2
                        If Not policies.SYS_POLICIES(i).value = Nothing Then
                            Me.Panel2.Controls.Item(policies.SYS_POLICIES(i).controlname).SelectedIndex = Convert.ToInt32(policies.SYS_POLICIES(i).value)
                        Else
                            Me.Panel2.Controls.Item(policies.SYS_POLICIES(i).controlname).SelectedIndex = 0
                        End If
                    Case 1
                        If Not policies.SYS_POLICIES(i).value = Nothing Then
                            Me.Panel2.Controls.Item(policies.SYS_POLICIES(i).controlname).Text = policies.SYS_POLICIES(i).value
                        Else
                            Me.Panel2.Controls.Item(policies.SYS_POLICIES(i).controlname).Text = ""
                        End If
                End Select
            Next i
        End If

        'Try to detect current applied profile
        Dim cp As Integer = Me.profileMatch_Check()
        If (cp = -1) Then Me.ComboBox3.SelectedIndex = Me.ComboBox3.Items.Count - 1 Else Me.ComboBox3.SelectedIndex = cp
        Me.TabControl1.SelectedIndex = 0
        Me.Button2.Enabled = False
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox1.CheckedChanged
        Me.Button2.Enabled = True
        If Not profileApplying Then Me.ComboBox3.SelectedIndex = Me.ComboBox3.Items.Count - 1
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox2.CheckedChanged
        Me.Button2.Enabled = True
        If Not profileApplying Then Me.ComboBox3.SelectedIndex = Me.ComboBox3.Items.Count - 1
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox3.CheckedChanged
        Me.Button2.Enabled = True
        If Not profileApplying Then Me.ComboBox3.SelectedIndex = Me.ComboBox3.Items.Count - 1
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox4.CheckedChanged
        Me.Button2.Enabled = True
        If Not profileApplying Then Me.ComboBox3.SelectedIndex = Me.ComboBox3.Items.Count - 1
    End Sub

    Private Sub CheckBox11_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox11.CheckedChanged
        Me.Button2.Enabled = True
        If Not profileApplying Then Me.ComboBox3.SelectedIndex = Me.ComboBox3.Items.Count - 1
    End Sub

    Private Sub CheckBox10_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox10.CheckedChanged
        Me.Button2.Enabled = True
        If Not profileApplying Then Me.ComboBox3.SelectedIndex = Me.ComboBox3.Items.Count - 1
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox5.CheckedChanged
        Me.Button2.Enabled = True
        If Not profileApplying Then Me.ComboBox3.SelectedIndex = Me.ComboBox3.Items.Count - 1
    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox6.CheckedChanged
        Me.Button2.Enabled = True
    End Sub

    Private Sub CheckBox7_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox7.CheckedChanged
        Me.Button2.Enabled = True
        If Not profileApplying Then Me.ComboBox3.SelectedIndex = Me.ComboBox3.Items.Count - 1
    End Sub

    Private Sub CheckBox8_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox8.CheckedChanged
        Me.Button2.Enabled = True
        If Not profileApplying Then Me.ComboBox3.SelectedIndex = Me.ComboBox3.Items.Count - 1
    End Sub

    Private Sub CheckBox9_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox9.CheckedChanged
        Me.Button2.Enabled = True
        If Not profileApplying Then Me.ComboBox3.SelectedIndex = Me.ComboBox3.Items.Count - 1
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Me.Button2.Enabled = True
        If Not profileApplying Then Me.ComboBox3.SelectedIndex = Me.ComboBox3.Items.Count - 1
    End Sub

    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged
        Me.Button2.Enabled = True
    End Sub

    Private Sub TextBox2_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox2.TextChanged
        Me.Button2.Enabled = True
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Dim SETTINGS As New Dictionary(Of String, gpe_policies.setting)
        Dim SYS_SETTINGS As New Dictionary(Of String, gpe_policies.setting)
        SETTINGS = policies.getSettings()
        SYS_SETTINGS = policies.getSystemSettings()
        Dim newpolicies As New Dictionary(Of String, String)
        Dim newsys_policies As New Dictionary(Of String, String)
        Dim c_item As KeyValuePair(Of String, gpe_policies.setting)

        'This for users
        For Each c_item In SETTINGS
            Select Case c_item.Value.controltype
                Case 0
                    Dim chbox As New System.Windows.Forms.CheckBox
                    chbox = Me.Panel1.Controls.Item(c_item.Key)
                    'Support for three state checkboxes
                    If Not chbox.ThreeState Then
                        newpolicies.Add(c_item.Key, Convert.ToInt32(chbox.Checked).ToString())
                    Else
                        newpolicies.Add(c_item.Key, Convert.ToInt32(chbox.CheckState).ToString())
                    End If
                Case 1
                    Dim tbox As New System.Windows.Forms.TextBox
                    tbox = Me.Panel1.Controls.Item(c_item.Key)
                    newpolicies.Add(c_item.Key, tbox.Text)
                Case 2
                    Dim cobox As New System.Windows.Forms.ComboBox
                    cobox = Me.Panel1.Controls.Item(c_item.Key)
                    newpolicies.Add(c_item.Key, cobox.SelectedIndex)
            End Select
        Next c_item

        'This for the system
        For Each c_item In SYS_SETTINGS
            Select Case c_item.Value.controltype
                Case 0
                    Dim chbox As New System.Windows.Forms.CheckBox
                    chbox = Me.Panel2.Controls.Item(c_item.Key)
                    'Support fot three state checkboxes
                    If Not chbox.ThreeState Then
                        newsys_policies.Add(c_item.Key, Convert.ToInt32(chbox.Checked).ToString())
                    Else
                        newsys_policies.Add(c_item.Key, Convert.ToInt32(chbox.CheckState).ToString())
                    End If
                Case 1
                    Dim tbox As New System.Windows.Forms.TextBox
                    tbox = Me.Panel2.Controls.Item(c_item.Key)
                    newsys_policies.Add(c_item.Key, tbox.Text)
                Case 2
                    Dim cobox As New System.Windows.Forms.ComboBox
                    cobox = Me.Panel2.Controls.Item(c_item.Key)
                    newsys_policies.Add(c_item.Key, cobox.SelectedIndex)
            End Select
        Next c_item

        Me.Enabled = False
        If Me.ComboBox1.SelectedItem = "Все пользователи" Then
            Dim i As Integer = 0
            Dim p As Boolean = True
            Dim fusers As String = ""
            'User policies
            For i = 0 To Me.ComboBox1.Items.Count - 1
                If Not Me.ComboBox1.Items(i) = "Все пользователи" Then
                    If Not policies.apply_Policies(Me.ComboBox1.Items(i), newpolicies) Then
                        p = False
                        If fusers = "" Then fusers = Me.ComboBox1.Items(i) Else fusers &= ", " & Me.ComboBox1.Items(i)
                    End If
                End If
            Next i

            'System policies
            If Not policies.apply_Policies(newsys_policies) Then
                p = False
                If fusers = "" Then fusers = "[СИСТЕМА]" Else fusers &= ", [СИСТЕМА]"
            End If

            If Not p Then
                MessageBox.Show("Не удалось выполнить изменения для пользователей: " & fusers & "." & Convert.ToChar(10) & Convert.ToChar(13) & policies.Exeption, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                MessageBox.Show("Изменения были успешно выполнены. Для применения некоторых параметров все пользователи должены перезапустить свой сеанс работы в Windows.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.Button2.Enabled = False
            End If
        Else
            If policies.apply_Policies(Me.ComboBox1.SelectedItem, newpolicies) And policies.apply_Policies(newsys_policies) Then
                MessageBox.Show("Изменения были успешно выполнены. Для применения некоторых параметров пользователь " & Me.ComboBox1.SelectedItem & " должен перезапустить свой сеанс работы в Windows.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.Button2.Enabled = False
            Else
                MessageBox.Show("Не удалось выполнить необходимые изменения." & Convert.ToChar(10) & Convert.ToChar(13) & policies.Exeption, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
        Me.Enabled = True
    End Sub

    Private Sub CheckBox12_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox12.CheckedChanged
        Me.Button2.Enabled = True
        If Not profileApplying Then Me.ComboBox3.SelectedIndex = Me.ComboBox3.Items.Count - 1
    End Sub

    Private Sub CheckBox13_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox13.CheckedChanged
        Me.Button2.Enabled = True
    End Sub

    Private Sub Panel1_Paint(sender As System.Object, e As System.EventArgs) Handles Panel1.Click
        Me.Panel1.Focus()
    End Sub

    Private Sub CheckBox14_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox14.CheckedChanged
        Me.Button2.Enabled = True
        If Not profileApplying Then Me.ComboBox3.SelectedIndex = Me.ComboBox3.Items.Count - 1
    End Sub

    Private Sub CheckBox15_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox15.CheckedChanged
        Me.Button2.Enabled = True
        If Not profileApplying Then Me.ComboBox3.SelectedIndex = Me.ComboBox3.Items.Count - 1
    End Sub

    Private Sub CheckBox16_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox16.CheckedChanged
        Me.Button2.Enabled = True
        If Not profileApplying Then Me.ComboBox3.SelectedIndex = Me.ComboBox3.Items.Count - 1
    End Sub

    Private Sub CheckBox17_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox17.CheckedChanged
        Me.Button2.Enabled = True
        If Not profileApplying Then Me.ComboBox3.SelectedIndex = Me.ComboBox3.Items.Count - 1
    End Sub

    Private Sub CheckBox18_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox18.CheckedChanged
        Me.Button2.Enabled = True
        If Not profileApplying Then Me.ComboBox3.SelectedIndex = Me.ComboBox3.Items.Count - 1
    End Sub

    Private Sub CheckBox19_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox19.CheckedChanged
        Me.Button2.Enabled = True
        If Not profileApplying Then Me.ComboBox3.SelectedIndex = Me.ComboBox3.Items.Count - 1
    End Sub

    Private Sub CheckBox20_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox20.CheckedChanged
        Me.Button2.Enabled = True
        If Not profileApplying Then Me.ComboBox3.SelectedIndex = Me.ComboBox3.Items.Count - 1
    End Sub

    Private Sub CheckBox21_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox21.CheckedChanged
        Me.Button2.Enabled = True
        If Not profileApplying Then Me.CheckBox23.CheckState = CheckState.Indeterminate
    End Sub

    Private Sub CheckBox22_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox22.CheckedChanged
        Me.Button2.Enabled = True
    End Sub

    Private Sub CheckBox23_CheckStateChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox23.CheckStateChanged
        Me.Button2.Enabled = True
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Not policies.AdministratorState = "2" Then
            Dim Q As DialogResult = MessageBox.Show("Программа попытается активировать встроенную учетную запись администратора. Продолжить?", Me.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If Q = Windows.Forms.DialogResult.Yes Then
                If Not policies.activateAdminAccount() Then
                    MessageBox.Show("Попытка активации встроенной учетной записи администратора не удалась. Возможно учетная запись удалена или переименована.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    MessageBox.Show("Встроенная учетная запись администратора успешно активирована.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Me.ToolTip1.SetToolTip(Me.Button3, "Изменить пароль встроенной учетной записи администратора.")
                    Me.Button3.BackColor = Color.LightGreen
                End If
            End If
        Else
            Dim Q As DialogResult = MessageBox.Show("Изменение пароля учетной записи может повлечь за собой потерю доступа ко всем файлам, принадлежащим данному пользователю. Продолжить?", Me.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If Q = Windows.Forms.DialogResult.Yes Then
                If Not policies.activateAdminAccount() Then
                    MessageBox.Show("Не удалось задать пароль.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    MessageBox.Show("Пароль был задан успешно.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If
        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        If Not Me.ComboBox3.SelectedItem = "Другое" Then Me.profileMatch_Apply(Me.ComboBox3.SelectedItem)
    End Sub

    Private Sub CheckBox24_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox24.CheckedChanged
        Me.Button2.Enabled = True
        If Not profileApplying Then Me.ComboBox3.SelectedIndex = Me.ComboBox3.Items.Count - 1
    End Sub

    Private Sub CheckBox25_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox25.CheckedChanged
        Me.Button2.Enabled = True
        If Not profileApplying Then Me.ComboBox3.SelectedIndex = Me.ComboBox3.Items.Count - 1
    End Sub

    Private Function profileMatch_Check() As Integer
        Dim i As Integer
        Dim p As Boolean = True
        Dim c_item As KeyValuePair(Of String, String)
        Dim SETTINGS As New Dictionary(Of String, gpe_policies.setting)
        SETTINGS = policies.getSettings()
        For i = 0 To policies.PROFILES.Length - 1
            For Each c_item In policies.PROFILES(i).p_settings
                Select Case SETTINGS(c_item.Key).controltype
                    Case 0
                        Dim chbox As New System.Windows.Forms.CheckBox
                        chbox = Me.Panel1.Controls.Item(SETTINGS(c_item.Key).controlname)
                        If Not chbox.Checked = CBool(c_item.Value) Then p = False
                    Case 1
                        Dim tbox As New System.Windows.Forms.TextBox
                        tbox = Me.Panel1.Controls.Item(SETTINGS(c_item.Key).controlname)
                        If Not tbox.Text = c_item.Value Then p = False
                    Case 2
                        Dim cobox As New System.Windows.Forms.ComboBox
                        cobox = Me.Panel1.Controls.Item(SETTINGS(c_item.Key).controlname)
                        If Not cobox.SelectedIndex = Int32.Parse(c_item.Value) Then p = False
                End Select
                If Not p Then Exit For
            Next c_item
            If p Then Return i Else p = True
        Next
        Return -1
    End Function

    Private Sub profileMatch_Apply(ByVal p_name As String)
        profileApplying = True
        Dim i As Integer = 0
        Dim c_item As KeyValuePair(Of String, String)
        Dim SETTINGS As New Dictionary(Of String, gpe_policies.setting)
        SETTINGS = policies.getSettings()
        For i = 0 To policies.PROFILES.Length - 1
            If policies.PROFILES(i).p_name = p_name Then
                For Each c_item In policies.PROFILES(i).p_settings
                    Select Case SETTINGS(c_item.Key).controltype
                        Case 0
                            Dim chbox As New System.Windows.Forms.CheckBox
                            chbox = Me.Panel1.Controls.Item(SETTINGS(c_item.Key).controlname)
                            'Support for three state checkboxes
                            If Not chbox.ThreeState Then
                                chbox.Checked = CBool(c_item.Value)
                            Else
                                Select Case c_item.Value
                                    Case "1"
                                        chbox.CheckState = CheckState.Checked
                                    Case "2"
                                        chbox.CheckState = CheckState.Indeterminate
                                    Case Else
                                        chbox.CheckState = CheckState.Unchecked
                                End Select
                            End If
                        Case 1
                            Dim tbox As New System.Windows.Forms.TextBox
                            tbox = Me.Panel1.Controls.Item(SETTINGS(c_item.Key).controlname)
                            tbox.Text = c_item.Value
                        Case 2
                            Dim cobox As New System.Windows.Forms.ComboBox
                            cobox = Me.Panel1.Controls.Item(SETTINGS(c_item.Key).controlname)
                            cobox.SelectedIndex = Int32.Parse(c_item.Value)
                    End Select
                Next c_item
                profileApplying = False
                Exit Sub
            End If
        Next
        profileApplying = False
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim ulAdd As New userlist_Add(policies.MAX_USERNAME_LENGTH)
        ulAdd.NumericUpDown1.Enabled = False
        ulAdd.Label2.Enabled = False
        ulAdd.NumericUpDown1.Maximum = policies.MAX_PASSPHRASE_LENGTH
        ulAdd.NumericUpDown1.Value = 8
        ulAdd.TextBox1.Text = "5А-1;5А-2;5Б-1;5Б-2;6А-1;6А-2;6Б-1;6Б-2;7А-1;7А-2;7Б-1;7Б-2;8А-1;8А-2;8Б-1;8Б-2;9А-1;9А-2;9Б-1;9Б-2;10А-1;10А-2;11А-1;11А-2"
        ulAdd.ShowDialog()

        If ulAdd.DialogResult = Windows.Forms.DialogResult.OK Then
            Me.Enabled = False
            Dim usernames() As String = ulAdd.TextBox1.Text.Split(";")
            Dim pasl As Integer
            If ulAdd.CheckBox1.Checked = False Then pasl = 0 Else pasl = ulAdd.NumericUpDown1.Value
            Dim u_list As Dictionary(Of String, String) = policies.createUser_List(usernames, pasl)
            If Not u_list Is Nothing Then
                Dim urep As New userlist_Report(0, u_list, policies.machine_name)
                urep.ShowDialog()
                MessageBox.Show("Чтобы управлять новыми пользователями с помощью данного средства, необходимо произвести первый вход в каждую из вновь созданных учетных записей.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                If policies.Exeption = "" Then
                    MessageBox.Show("Учетные записи были успешно созданы. Чтобы управлять новыми пользователями с помощью данного средства, необходимо произвести первый вход в каждую из вновь созданных учетных записей.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If
        End If
        Me.Enabled = True
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim uichp As New userlist_Changepass
        uichp.CheckBox1.Checked = False
        uichp.Label2.Enabled = False
        uichp.NumericUpDown1.Enabled = False
        uichp.CheckedListBox1.Items.Clear()
        uichp.Button1.Enabled = False
        uichp.NumericUpDown1.Maximum = policies.MAX_PASSPHRASE_LENGTH
        uichp.NumericUpDown1.Value = 8

        Dim i As Integer
        For i = 0 To policies.USERS.Length - 1
            uichp.CheckedListBox1.Items.Add(policies.USERS(i).username)
        Next i
        uichp.ShowDialog()

        If uichp.DialogResult = Windows.Forms.DialogResult.OK Then
            Me.Enabled = False
            Dim passlength As Integer
            If uichp.CheckBox1.Checked = True Then passlength = uichp.NumericUpDown1.Value Else passlength = 0
            Dim usernames(uichp.CheckedListBox1.CheckedItems.Count - 1) As String
            For i = 0 To uichp.CheckedListBox1.CheckedItems.Count - 1
                usernames(i) = uichp.CheckedListBox1.CheckedItems(i)
            Next i
            Dim p_list As Dictionary(Of String, String) = policies.changePassword_List(usernames, passlength)
            If Not p_list Is Nothing Then
                Dim urep As New userlist_Report(0, p_list, policies.machine_name)
                urep.ShowDialog()
                MessageBox.Show("Все пароли были успешно изменены.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                If policies.Exeption = "" Then
                    MessageBox.Show("Все пароли были успешно удалены.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If
            Me.Enabled = True
        End If
    End Sub

    Private Sub CheckBox28_CheckedChanged(sender As Object, e As EventArgs)
        Me.Button2.Enabled = True
        If Not profileApplying Then Me.ComboBox3.SelectedIndex = Me.ComboBox3.Items.Count - 1
    End Sub

    Private Sub CheckBox29_CheckedChanged(sender As Object, e As EventArgs)
        Me.Button2.Enabled = True
        If Not profileApplying Then Me.ComboBox3.SelectedIndex = Me.ComboBox3.Items.Count - 1
    End Sub

    Private Sub CheckBox30_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox30.CheckedChanged
        Me.Button2.Enabled = True
        If Not profileApplying Then Me.ComboBox3.SelectedIndex = Me.ComboBox3.Items.Count - 1
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim pfile As New gpe_passfile(IO.FileMode.Open)
        If pfile.objectstate = 1 Then
            If Not pfile.Exeption = "" Then MessageBox.Show(pfile.Exeption, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Dim pfileParsed As Object = pfile.parsePassFile()
        If pfileParsed Is Nothing Then
            If Not pfile.Exeption = "" Then MessageBox.Show("Во время выполнения произошла ошибка: " & pfile.Exeption, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        If Not pfileParsed(0) = policies.machine_name Then
            Dim Q As DialogResult = MessageBox.Show("Имя компьютера, указанное в файле паролей (" & pfileParsed(0) & ") не совпадает с именем этого компьютера (" & policies.machine_name & "). Возможно вы используете не тот файл паролей." & Convert.ToChar(10) & Convert.ToChar(13) & Convert.ToChar(10) & Convert.ToChar(13) & "Продолжить?", Me.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If Q = DialogResult.No Then Exit Sub
        End If

        Dim utoset As New Dictionary(Of String, String)
        Dim fprep As New userlist_Report(1, pfileParsed(1), pfileParsed(0), utoset)
        If fprep.ShowDialog() = Windows.Forms.DialogResult.OK Then
            If policies.setPassword_List(utoset) Then
                MessageBox.Show("Все пароли были успешно изменены.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("Во время выполнения произошла ошибка: " & policies.Exeption, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub

    Private Sub CheckBox28_CheckedChanged_1(sender As Object, e As EventArgs) Handles CheckBox28.CheckedChanged
        Me.Button2.Enabled = True
        If Not profileApplying Then Me.ComboBox3.SelectedIndex = Me.ComboBox3.Items.Count - 1
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        Me.Button2.Enabled = True
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        'We need to know what files need to be resorted
        Me.OpenFileDialog1.FileName = ""
        Me.OpenFileDialog1.Filter = "Текстовый документ (*.txt)|*.txt"
        If Me.OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Me.Enabled = False
            Dim i As Integer = 0
            Dim pfiles(OpenFileDialog1.FileNames.Length - 1) As Object
            For i = 0 To OpenFileDialog1.FileNames.Length - 1
                'Parse passfiles
                Dim pfile As New gpe_passfile(IO.FileMode.Open, OpenFileDialog1.FileNames(i))
                If pfile.objectstate = 0 Then pfiles(i) = pfile.parsePassFile()
            Next i

            Dim pfiles_resorted(-1) As Object
            Dim j, k As Integer

            'Firstly we need to know what passfile contains the most count of username/password pairs
            j = pfiles(0)(1).Count
            For i = 1 To pfiles.Length - 1
                If pfiles(i)(1).Count > j Then
                    k = i
                    j = pfiles(i)(1).Count
                End If
            Next i

            'Make resort
            Dim c_item As KeyValuePair(Of String, String)
            Dim pfilesrejected As Integer = 0
            j = 0
            For Each c_item In pfiles(k)(1)
                Dim _class As String = c_item.Key
                Dim _clpass As New Dictionary(Of String, String)
                For i = 0 To pfiles.Length - 1
                    If Not pfiles(i) Is Nothing Then
                        'If we have more then one file from the same machine, only the first file aplies
                        If _clpass.ContainsKey(pfiles(i)(0)) Then
                            pfilesrejected += 1
                            Continue For
                        End If
                        If pfiles(i)(1).ContainsKey(_class) Then _clpass.Add(pfiles(i)(0), pfiles(i)(1).Item(_class))
                    End If
                Next i

                ReDim Preserve pfiles_resorted(j)
                pfiles_resorted(j) = {_class, _clpass}
                j += 1
            Next c_item

            'Now save resorted data
            'Get the files directory
            Dim wd() As String = OpenFileDialog1.FileNames(0).Split("\")
            Dim wds As String = wd(0)
            For i = 1 To wd.Length - 2
                wds &= "\" & wd(i)
            Next i
            wds &= "\Resorted"

            'Create directory to save data
            IO.Directory.CreateDirectory(wds)

            'And save data
            For i = 0 To pfiles_resorted.Length - 1
                Dim sfile As New gpe_passfile(IO.FileMode.Create, wds & "\" & pfiles_resorted(i)(0) & ".txt")
                If sfile.objectstate = 0 Then
                    sfile.printPassFile(pfiles_resorted(i), True)
                    sfile.closePassFile()
                End If
            Next

            'Complete message
            Dim cmessage As String = "Перестроение файлов паролей завершено. Новые файлы паролей можно найти в папке ""Resorted""."
            If pfilesrejected > 0 Then cmessage &= Convert.ToChar(10) & Convert.ToChar(13) & Convert.ToChar(10) & Convert.ToChar(13) & "Внимание! В ходе пересортировки были обнаружены файлы принадлежащие одним и тем же компьютерам. Проверьте результаты пересортировки. При необходимости удалите лишние файлы паролей и повторите перестроение."
            MessageBox.Show(cmessage, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

            Me.Enabled = True
        End If
    End Sub
End Class
