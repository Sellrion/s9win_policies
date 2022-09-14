Imports System
Imports System.Management
Imports System.Security.Permissions
Imports System.Runtime.InteropServices
Imports System.Xml
Imports System.Xml.Serialization
Imports Microsoft.Win32

Public Class gpe_policies

    Private Structure LUID
        Private lowPart As UInt32
        Private highPart As Int32
    End Structure

    Private Structure LUID_AND_ATTRIBUTES
        Public luid As LUID
        Public attributes As UInt32
    End Structure

    Private Structure TOKEN_PRIVILEGES
        Public PrivilegeCount As UInteger
        Public Privileges As LUID_AND_ATTRIBUTES
        Public Function Size() As Integer
            Return Marshal.SizeOf(Me)
        End Function
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Private Structure USER_INFO_1
        Dim usri1_name As String
        Dim usri1_password As String
        Dim usri1_password_age As Integer
        Dim usri1_priv As Integer
        Dim usri1_home_dir As String
        Dim usri1_comment As String
        Dim usri1_flags As Integer
        Dim usri1_script_path As String
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Private Structure USER_INFO_20
        Dim usri23_name As String
        Dim usri23_full_name As String
        Dim usri23_comment As String
        Dim usri23_flags As Integer
        Dim usri23_user_id As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Private Structure USER_INFO_1003
        Dim usri1003_password As String
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Private Structure USER_INFO_1008
        Dim usri1008_flags As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Private Structure LOCALGROUP_MEMBERS_INFO_0
        Dim lgrmi0_sid As IntPtr
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Private Structure LOCALGROUP_INFO_0
        Dim lgrpi0_name As String
    End Structure

    Private Enum SID_NAME_USE
        SidTypeUser = 1
        SidTypeGroup
        SidTypeDomain
        SidTypeAlias
        SidTypeWellKnownGroup
        SidTypeDeletedAccount
        SidTypeInvalid
        SidTypeUnknown
        SidTypeComputer
        SidTypeLabel
    End Enum

    Private Structure USER
        Dim username As String
        Dim SID As String
        Dim POLICIES() As Object
    End Structure

    Public Structure setting
        Dim hive As Integer
        Dim subkey As String
        Dim kind As RegistryValueKind
        Dim key As String
        Dim value As String
        Dim vnegative As String
        Dim vpositive As String
        Dim vnotset As String
        Dim controltype As Integer '0 - CheckBox, 1 - TextBox, 2 - ComboBox
        Dim controlname As String
        Dim minosversion As String
    End Structure

    Private Structure profile
        Dim p_name As String
        Dim p_settings As Dictionary(Of String, String)
    End Structure

    Private ReadOnly SETTINGS As New Dictionary(Of String, setting)     'Settings table that contain default registry scheme
    Private ReadOnly SYS_SETTINGS As New Dictionary(Of String, setting) 'System settings table that contain default registry scheme
    Public ReadOnly USERS(-1) As Object                                 'Array to store information about users
    Public PROFILES(-1) As Object                                       'Presets information
    Public SYS_POLICIES(-1) As Object                                   'Array to store system policies
    Public ReadOnly objectstatus As Integer                             'Status of the gpe_policies object
    Private ReadOnly osversion_major As Integer                         'Major version of OS running on that computer
    Private ReadOnly osversion_minor As Integer                         'Minor version of OS running on that computer
    Private ReadOnly osversion_build As Integer                         'Build number of OS running on that computer
    Public ReadOnly machine_name As String                              'NetBIOS machine name
    Private ReadOnly usersgroup_name As String                          'Name of the Users group on that computer
    Private USERS_LOADED(-1) As String                                  'List of users that have been loaded into registry
    Private Const HKEY_USERS As UInteger = &H80000003UI                 'USERS refistry hive digital representation
    Private Const TOKEN_ADJUST_PRIVILEGES As UInteger = &H20UI          'TOKEN_ADJUST_PRIVILEGES flag digital representation
    Private Const TOKEN_QUERY As UInteger = &H8UI                       'TOKEN_QUERY flag digital representation
    Private Const SE_PRIVILEGE_ENABLED As UInteger = &H2UI              'SE_PRIVILEGE_ENABLED flag digital representation
    Private Const SE_RESTORE_NAME As String = "SeRestorePrivilege"      'SE_RESTORE_NAME flag string
    Private Const SE_BACKUP_NAME As String = "SeBackupPrivilege"        'SE_BACKUP_NAME flag string
    Private Const UF_PASSWD_CANT_CHANGE As Integer = &H40               'UF_PASSWD_CANT_CHANGE flag digital representation
    Private Const UF_DONT_EXPIRE_PASSWD As Integer = &H10000            'UF_DONT_EXPIRE_PASSWD flag digital representation
    Private Const UF_SCRIPT As Integer = &H1                            'UF_SCRIPT flag digital representation
    Private Token As IntPtr                                             'Token pointer
    Private TP1, TP2 As TOKEN_PRIVILEGES                                'Token priviliges variables
    Public Exeption As String                                           'String that contain the information about last exeption
    Public AdministratorState As String                                 'Current state of an administrator account
    Public ReadOnly CurrentUser As String                               'Current thread starter user name
    Private ReadOnly SDRIVE As String                                   'System drive letter
    Public ReadOnly USE_INTERFACE As Integer = 0                        'Flag show what interface is used. 0 - WMI (default), 1 - WinAPI
    Private ReadOnly MAX_PASSFILE_SIZE As Integer = 100024              'Max size of passwords file to read in bytes
    Public ReadOnly MAX_PASSPHRASE_LENGTH As Integer = 14               'Max password phrase length
    Public ReadOnly MAX_USERNAME_LENGTH As Integer = 20                 'Max username string length


    <DllImport("kernel32")>
    Private Shared Function GetCurrentProcess() As IntPtr
    End Function

    <DllImport("kernel32")>
    Private Shared Function GetLastError() As Integer
    End Function

    <DllImport("advapi32", SetLastError:=True)>
    Private Shared Function OpenProcessToken(
                                            ByVal ProcessHandle As IntPtr,
                                            ByVal DesiredAccess As UInteger,
                                            ByRef TokenHandle As IntPtr
                                            ) As Integer
    End Function

    <DllImport("advapi32", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Private Shared Function LookupPrivilegeValue(
                                                ByVal lpSystemName As String,
                                                ByVal lpName As String,
                                                ByRef lpLuid As LUID
                                                ) As Integer
    End Function

    <DllImport("advapi32", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Private Shared Function AdjustTokenPrivileges(
                                                 ByVal TokenHandle As IntPtr,
                                                 ByVal DisableAllPrivileges As Integer,
                                                 ByRef NewState As TOKEN_PRIVILEGES,
                                                 ByVal BufferLength As Integer,
                                                 ByRef PreviousState As TOKEN_PRIVILEGES,
                                                 ByRef ReturnLength As Integer
                                                 ) As Integer
    End Function

    <DllImport("advapi32", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Private Shared Function RegLoadKey(
                                      ByVal hKey As UInteger,
                                      ByVal lpSubKey As String,
                                      ByVal lpFile As String
                                      ) As Integer
    End Function

    <DllImport("advapi32", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Private Shared Function RegUnLoadKey(
                                        ByVal hKey As UInteger,
                                        ByVal lpSubKey As String
                                        ) As Integer
    End Function

    <DllImport("advapi32", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Private Shared Function ConvertStringSidToSid(
                                                 ByVal StringSid As String,
                                                 ByRef Sid As IntPtr
                                                 ) As Boolean
    End Function

    <DllImport("advapi32", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Private Shared Function ConvertSidToStringSid(
                                                 ByVal Sid As IntPtr,
                                                 ByRef StringSid As IntPtr
                                                 ) As Boolean
    End Function

    <DllImport("advapi32", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Private Shared Function LookupAccountName(
                                             ByVal lpSystemName As String,
                                             ByVal lpAccountName As String,
                                             ByVal Sid As IntPtr,
                                             ByRef cbSid As UInteger,
                                             ByVal ReferencedDomainName As System.Text.StringBuilder,
                                             ByRef cchReferencedDomainName As UInteger,
                                             ByRef peUse As SID_NAME_USE
                                             ) As Boolean

    End Function

    <DllImport("netapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function NetUserAdd(
                                      <MarshalAs(UnmanagedType.LPWStr)> ByVal servername As String,
                                      ByVal level As Integer,
                                      ByRef buf As USER_INFO_1,
                                      ByRef parm_err As Integer
                                      ) As Integer
    End Function

    <DllImport("netapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function NetUserSetInfo(
                                          <MarshalAs(UnmanagedType.LPWStr)> ByVal servername As String,
                                          <MarshalAs(UnmanagedType.LPWStr)> ByVal username As String,
                                          ByVal level As Integer,
                                          ByRef buf As USER_INFO_1003,
                                          ByRef parm_err As Integer
                                          ) As Integer
    End Function

    <DllImport("netapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function NetUserSetInfo(
                                          <MarshalAs(UnmanagedType.LPWStr)> ByVal servername As String,
                                          <MarshalAs(UnmanagedType.LPWStr)> ByVal username As String,
                                          ByVal level As Integer,
                                          ByRef buf As USER_INFO_1008,
                                          ByRef parm_err As Integer
                                          ) As Integer
    End Function

    <DllImport("netapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function NetLocalGroupAddMembers(
                                                    <MarshalAs(UnmanagedType.LPWStr)> ByVal servername As String,
                                                    <MarshalAs(UnmanagedType.LPWStr)> ByVal groupname As String,
                                                    ByVal level As Integer,
                                                    ByRef buf As LOCALGROUP_MEMBERS_INFO_0,
                                                    ByVal num_entries As Integer
                                                    ) As Integer
    End Function

    <DllImport("netapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function NetLocalGroupEnum(
                                             <MarshalAs(UnmanagedType.LPWStr)> ByVal servername As String,
                                             ByVal level As Integer,
                                             ByRef bufptr As IntPtr,
                                             ByVal prefmaxlen As Integer,
                                             ByRef entriesread As Integer,
                                             ByRef totalentries As Integer,
                                             ByRef resume_handle As Integer
                                             ) As Integer
    End Function

    <DllImport("netapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function NetUserEnum(
                                        <MarshalAs(UnmanagedType.LPWStr)> ByVal servername As String,
                                        ByVal level As Integer,
                                        ByVal filter As Integer,
                                        ByRef bufptr As IntPtr,
                                        ByVal prefmaxlen As Integer,
                                        ByRef entriesread As Integer,
                                        ByRef totalentries As Integer,
                                        ByRef resume_handle As Integer
                                        ) As Integer
    End Function

    Sub New()
        'Get the nesessary privileges to be able to load registry hives
        If Not Me.AdjustPrivileges() Then
            If Me.Exeption = "" Then Me.Exeption = "Программе не удалось получить необходимые для работы привилегии."
            Me.objectstatus = 1
            Exit Sub
        End If

        'Important members
        Me.osversion_major = System.Environment.OSVersion.Version.Major
        Me.osversion_minor = System.Environment.OSVersion.Version.Minor
        Me.osversion_build = System.Environment.OSVersion.Version.Build
        Me.machine_name = System.Environment.MachineName

        'If WMI fails, let's try to find group by name
        If Me.isLocalGroupExists_WINAPI("Пользователи") Then
            Me.usersgroup_name = "Пользователи"
        ElseIf Me.isLocalGroupExists_WINAPI("Users") Then
            Me.usersgroup_name = "Users"

        Else
            Me.Exeption = "Не удалось получить информацию о локальных группах."
            Me.objectstatus = 1
            Exit Sub
        End If

        'Current thread starter
        Me.CurrentUser = Environment.UserName

        'Detect admin account
        Dim adminfo() As String = Me.detectAdminAccount_WINAPI()

        Me.AdministratorState = adminfo(0)

        'Init the settings list
        Me.SettingsInit(Me.SETTINGS, Me.SYS_SETTINGS, Me.PROFILES, adminfo(1))

        'Get system drive letter
        Dim sdrv() As String = Environment.SystemDirectory.Split("\")
        Me.SDRIVE = sdrv(0)

        'Get all information about users
        'WMI interface
        Dim p As Boolean = Me.getUsers_WINAPI(Me.SETTINGS, Me.USERS)

        If Not p Then
            Me.Exeption = "Не удалось получить информацию о локальных пользователях."
            Me.objectstatus = 1
            Exit Sub
        End If

        'Get system policies
        Dim i As Integer = 0
        Dim c_item As KeyValuePair(Of String, setting)
        Dim rs As New setting
        ReDim Me.SYS_POLICIES(Me.SYS_SETTINGS.Count - 1)
        For Each c_item In Me.SYS_SETTINGS
            rs = c_item.Value
            If isRegSubkey_Exists("", c_item.Value.hive, c_item.Value.subkey) Then
                rs.value = getRegPolicy("", c_item.Value.hive, c_item.Value.subkey, c_item.Value.key)
            Else
                rs.value = Nothing
            End If
            Me.SYS_POLICIES(i) = rs
            i = i + 1
        Next c_item

        'Main object built successfuly
        Me.objectstatus = 0
    End Sub

    Private Function AdjustPrivileges() As Boolean
        Try
            Dim pProc As IntPtr = GetCurrentProcess()
            Dim ret As Integer = OpenProcessToken(pProc, TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, Token)
            If ret = 0 Then Return False
            Dim RestoreLuid As LUID
            ret = LookupPrivilegeValue(Nothing, SE_RESTORE_NAME, RestoreLuid)
            If ret = 0 Then Return False
            Dim BackupLuid As LUID
            ret = LookupPrivilegeValue(Nothing, SE_BACKUP_NAME, BackupLuid)
            If ret = 0 Then Return False
            TP1.PrivilegeCount = 1
            TP1.Privileges.attributes = SE_PRIVILEGE_ENABLED
            TP1.Privileges.luid = RestoreLuid
            Dim returnLength As Integer = 0
            Dim oldPrivileges As TOKEN_PRIVILEGES
            ret = AdjustTokenPrivileges(Token, 0, TP1, TP1.Size(), oldPrivileges, returnLength)
            'Here we need more information about return value
            'So use the GetLastError core function
            If (ret = 0 Or GetLastError() = 1300) Then Return False
            TP2.PrivilegeCount = 1
            TP2.Privileges.attributes = SE_PRIVILEGE_ENABLED
            TP2.Privileges.luid = BackupLuid
            ret = AdjustTokenPrivileges(Token, 0, TP2, TP2.Size(), oldPrivileges, returnLength)
            If (ret = 0 Or GetLastError() = 1300) Then Return False
        Catch ex As Exception
            Me.Exeption = ex.Message()
            Return False
        End Try
        Return True
    End Function

    Public Function shutdown_Privileges() As Boolean
        Dim p As Boolean = True
        'Firstly we need to unload all the loaded hives
        If Me.USERS_LOADED.Length > 0 Then
            Dim i, j As Integer
            For i = 0 To Me.USERS_LOADED.Length - 1
                j = RegUnLoadKey(HKEY_USERS, Me.USERS_LOADED(i))
                If j <> 0 Then p = False
            Next
        End If

        'Lowering back privileges
        Dim returnLength As Integer = 0
        Dim ret As Integer = AdjustTokenPrivileges(Token, 0, TP1, TP1.Size(), Nothing, returnLength)
        If (ret = 0 Or GetLastError() = 1300) Then Return False
        ret = AdjustTokenPrivileges(Token, 0, TP2, TP2.Size(), Nothing, returnLength)
        If (ret = 0 Or GetLastError() = 1300) Then p = False
        Return p
    End Function

    Private Function getUsers_WINAPI(ByRef regtable As Dictionary(Of String, setting), ByRef usertable As Object) As Boolean
        Dim bufPtr As IntPtr
        Dim entrread As Integer
        Dim totalentr As Integer
        Dim reshandle As Integer
        Dim FILTER_NORMAL_ACCOUNT As Integer = &H2
        Dim ptrSid As IntPtr
        Dim cbSid As Integer
        Dim ptrSidString As IntPtr
        Dim refDomainName As New System.Text.StringBuilder
        Dim cbRefDomainName As Integer
        Dim peUse As SID_NAME_USE
        Dim StringSid As String = String.Empty
        Try
            Dim retval As Integer = NetUserEnum(String.Empty, 20, FILTER_NORMAL_ACCOUNT, bufPtr, -1, entrread, totalentr, reshandle)
            If Not retval = 0 Then Return False

            If entrread > 0 Then
                Dim iptr As IntPtr = bufPtr
                Dim i, j, k As Integer
                k = -1
                Dim uinf As USER_INFO_20
                Dim rs As New setting
                Dim c_item As KeyValuePair(Of String, setting)
                For i = 1 To entrread
                    uinf = New USER_INFO_20
                    uinf = CType(Marshal.PtrToStructure(iptr, GetType(USER_INFO_20)), USER_INFO_20)
                    'If user is not admin or guest
                    If Not uinf.usri23_user_id = 500 And Not uinf.usri23_user_id = 501 Then
                        'No we have to get a SID of that user
                        'First call
                        LookupAccountName(String.Empty, uinf.usri23_name, Nothing, cbSid, Nothing, cbRefDomainName, peUse)
                        'Adjust buffers
                        ptrSid = Marshal.AllocHGlobal(cbSid)
                        refDomainName.EnsureCapacity(cbRefDomainName)
                        If LookupAccountName(String.Empty, uinf.usri23_name, ptrSid, cbSid, refDomainName, cbRefDomainName, peUse) Then
                            If ConvertSidToStringSid(ptrSid, ptrSidString) Then StringSid = Marshal.PtrToStringAuto(ptrSidString)
                        Else
                            iptr = New IntPtr(iptr.ToInt32 + Marshal.SizeOf(GetType(USER_INFO_20)))
                            Continue For
                        End If

                        'Check if SID is loaded to the registry
                        If Not Me.isSID_Loaded(StringSid) Then
                            iptr = New IntPtr(iptr.ToInt32 + Marshal.SizeOf(GetType(USER_INFO_20)))
                            Continue For
                        End If

                        Dim us As USER
                        'Fill common information
                        us.username = uinf.usri23_name
                        us.SID = StringSid
                        ReDim us.POLICIES(regtable.Count - 1)
                        j = 0
                        'Registry settings table
                        For Each c_item In regtable
                            rs = c_item.Value
                            If Me.isRegSubkey_Exists(us.SID, c_item.Value.hive, c_item.Value.subkey) Then
                                rs.value = Me.getRegPolicy(us.SID, c_item.Value.hive, c_item.Value.subkey, c_item.Value.key)
                            Else
                                rs.value = Nothing
                            End If
                            us.POLICIES(j) = rs
                            j += 1
                        Next c_item
                        'Add user to the users table
                        ReDim Preserve usertable(k + 1)
                        usertable(k + 1) = us
                        k += 1
                    End If

                    'Increment pointer
                    iptr = New IntPtr(iptr.ToInt32 + Marshal.SizeOf(GetType(USER_INFO_20)))
                Next i
            End If
        Catch ex As Exception
            MessageBox.Show("Critical error: " & ex.Message(), gpe_main.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try

        'All were fine, so return success
        Return True
    End Function

    Private Function isLocalGroupExists_WINAPI(ByVal groupname As String) As Boolean
        Dim bufPtr As IntPtr
        Dim entrread As Integer
        Dim totalentr As Integer
        Dim reshandle As Integer

        Try
            Dim retval As Integer = NetLocalGroupEnum(String.Empty, 0, bufPtr, -1, entrread, totalentr, reshandle)
            If Not retval = 0 Then Return False

            If entrread > 0 Then
                Dim iptr As IntPtr = bufPtr
                Dim i As Integer
                Dim ginf As LOCALGROUP_INFO_0
                For i = 1 To entrread
                    ginf = New LOCALGROUP_INFO_0
                    ginf = CType(Marshal.PtrToStructure(iptr, GetType(LOCALGROUP_INFO_0)), LOCALGROUP_INFO_0)
                    If ginf.lgrpi0_name = groupname Then Return True
                    iptr = New IntPtr(iptr.ToInt32 + Marshal.SizeOf(GetType(LOCALGROUP_INFO_0)))
                Next i
            Else
                Return False
            End If
        Catch ex As Exception
            MessageBox.Show("Critical error: " & ex.Message(), gpe_main.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
        Return False
    End Function

    Private Function getRegPolicy(ByVal SID As String, ByVal hive As Integer, ByVal branch As String, ByVal key As String) As String
        Dim reg As RegistryKey
        If hive = 1 Then
            reg = Registry.Users.OpenSubKey(SID & "\" & branch, False)
        Else
            reg = Registry.LocalMachine.OpenSubKey(branch, False)
        End If
        Dim rv As String
        Try
            rv = reg.GetValue(key)
        Catch ex As Exception
            rv = Nothing
        End Try
        reg.Close()
        Return rv
    End Function

    Private Function setRegPolicy(ByVal SID As String,
                                  ByVal hive As Integer,
                                  ByVal branch As String,
                                  ByVal key As String,
                                  ByVal value As String,
                                  ByVal kind As RegistryValueKind
                                  ) As Boolean
        Dim reg As RegistryKey
        If hive = 1 Then
            reg = Registry.Users.OpenSubKey(SID & "\" & branch, True)
        Else
            reg = Registry.LocalMachine.OpenSubKey(branch, True)
        End If

        Try
            reg.SetValue(key, value, kind)
        Catch ex As Exception
            MessageBox.Show("Critical error: " & ex.Message(), gpe_main.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
        reg.Close()
        Return True
    End Function

    Private Function isRegSubkey_Exists(ByVal SID As String, ByVal hive As Integer, ByVal subkey As String) As Boolean
        Dim reg As RegistryKey
        If hive = 1 Then
            reg = Registry.Users.OpenSubKey(SID & "\" & subkey, False)
        Else
            reg = Registry.LocalMachine.OpenSubKey(subkey, False)
        End If
        If reg Is Nothing Then
            Return False
        Else
            reg.Close()
            Return True
        End If
    End Function

    Public Function vBool(ByRef needle As setting) As Boolean
        Select Case needle.controltype
            Case 0
                If needle.value = "0" Then
                    Return CBool(needle.vnegative)
                ElseIf needle.value = "1" Then
                    Return CBool(needle.vpositive)
                Else
                    Return CBool(needle.vnotset)
                End If
        End Select

        Return False
    End Function

    Private Function isSID_Loaded(ByVal SID As String) As Boolean
        Dim reg As RegistryKey
        reg = Registry.Users.OpenSubKey(SID, False)
        If Not reg Is Nothing Then Return True
        reg = Registry.LocalMachine.OpenSubKey("Software\Microsoft\Windows NT\CurrentVersion\ProfileList\" & SID, False)
        If reg Is Nothing Then Return False
        Dim impath As String = Me.getRegPolicy("", 2, "Software\Microsoft\Windows NT\CurrentVersion\ProfileList\" & SID, "ProfileImagePath")
        'Here we check if profile image path contains environment vars
        'Such as %HOMEDRIVE% or %SYSTEMDRIVE%
        Dim impath_lc As String = impath.ToLower()
        If impath_lc = "" Then
            Return False
        ElseIf impath_lc.StartsWith("%systemdrive%") Or impath_lc.StartsWith("%homedrive%") Then
            Dim impath_items() As String = impath.Split("\")
            impath_items(0) = Me.SDRIVE
            impath = String.Join("\", impath_items)
        End If
        reg.Close()

        Dim i As Integer = RegLoadKey(HKEY_USERS, SID, impath & "\NTUSER.DAT")

        If Not i <> 0 Then
            ReDim Preserve Me.USERS_LOADED(Me.USERS_LOADED.Length + 1)
            Me.USERS_LOADED(Me.USERS_LOADED.Length - 1) = SID
            Return True
        Else : Return False
        End If
    End Function

    Private Sub create_RegTree(ByVal SID As String, ByVal hive As Integer, ByVal branch As String)
        Dim reg As RegistryKey
        Dim leaves() As String = branch.Split("\")
        Dim i As Integer = 0
        Dim gtree As String = ""
        Dim gtree_temp As String = ""
        Do Until i = leaves.Length
            If i = 0 Then gtree_temp = gtree Else gtree_temp = gtree & "\" & leaves(i)
            If Not isRegSubkey_Exists(SID, hive, gtree_temp) Then
                If hive = 1 Then
                    reg = Registry.Users.OpenSubKey(SID & "\" & gtree, True)
                Else
                    reg = Registry.LocalMachine.OpenSubKey(gtree, True)
                End If
                reg.CreateSubKey(leaves(i))
                reg.Close()
            End If
            If i = 0 Then gtree = leaves(i) Else gtree = gtree & "\" & leaves(i)
            i = i + 1
        Loop
    End Sub

    Public Function apply_Policies(ByVal username As String, ByVal newpolicies As Dictionary(Of String, String)) As Boolean
        Dim i As Integer
        Dim SID As String = String.Empty
        For i = 0 To Me.USERS.Length - 1
            If Me.USERS(i).username = username Then
                SID = Me.USERS(i).SID
                Exit For
            End If
        Next i
        If SID = "" Then Return False

        Dim newpolicies_prepared As New Dictionary(Of String, String)
        newpolicies_prepared = Me.prepare_NewPolicies(newpolicies, 0)

        Try
            For j = 0 To Me.USERS(i).POLICIES.Length - 1
                If Me.isPolicy_Changed(Me.USERS(i).POLICIES(j).controlname, Me.USERS(i).POLICIES(j).value, newpolicies_prepared(Me.USERS(i).POLICIES(j).controlname), 0) And Me.osverison_Compare(Me.USERS(i).POLICIES(j).minosversion) Then
                    If Not Me.isRegSubkey_Exists(SID, Me.USERS(i).POLICIES(j).hive, Me.USERS(i).POLICIES(j).subkey) Then Me.create_RegTree(SID, Me.USERS(i).POLICIES(j).hive, Me.USERS(i).POLICIES(j).subkey)
                    If Not Me.setRegPolicy(SID, Me.USERS(i).POLICIES(j).hive, Me.USERS(i).POLICIES(j).subkey, Me.USERS(i).POLICIES(j).key, newpolicies_prepared(Me.USERS(i).POLICIES(j).controlname), Me.USERS(i).POLICIES(j).kind) Then Return False
                End If
                Dim newsetting As New setting
                newsetting = Me.SETTINGS(Me.USERS(i).POLICIES(j).controlname)
                newsetting.value = newpolicies_prepared(Me.USERS(i).POLICIES(j).controlname)
                Me.USERS(i).POLICIES(j) = newsetting
            Next j
        Catch ex As Exception
            Me.Exeption = ex.Message()
            Return False
        End Try

        Return True
    End Function

    Public Function apply_Policies(ByVal newpolicies As Dictionary(Of String, String)) As Boolean
        Dim newpolicies_prepared As New Dictionary(Of String, String)
        newpolicies_prepared = Me.prepare_NewPolicies(newpolicies, 1)

        Dim i As Integer
        Try
            For i = 0 To Me.SYS_POLICIES.Length - 1
                If Me.isPolicy_Changed(Me.SYS_POLICIES(i).controlname, Me.SYS_POLICIES(i).value, newpolicies_prepared(Me.SYS_POLICIES(i).controlname), 1) And Me.osverison_Compare(Me.SYS_POLICIES(i).minosversion) Then
                    If Not Me.isRegSubkey_Exists("", Me.SYS_POLICIES(i).hive, Me.SYS_POLICIES(i).subkey) Then
                        Me.create_RegTree("", Me.SYS_POLICIES(i).hive, Me.SYS_POLICIES(i).subkey)
                    End If
                    If Not Me.setRegPolicy("", Me.SYS_POLICIES(i).hive, Me.SYS_POLICIES(i).subkey, Me.SYS_POLICIES(i).key, newpolicies_prepared(Me.SYS_POLICIES(i).controlname), Me.SYS_POLICIES(i).kind) Then Return False
                End If
                Dim newsetting As New setting
                newsetting = Me.SYS_SETTINGS(Me.SYS_POLICIES(i).controlname)
                newsetting.value = newpolicies_prepared(Me.SYS_POLICIES(i).controlname)
                Me.SYS_POLICIES(i) = newsetting
            Next i
        Catch ex As Exception
            Me.Exeption = ex.Message()
            Return False
        End Try

        Return True
    End Function

    Private Function prepare_NewPolicies(ByVal newpolicies As Dictionary(Of String, String),
                                         ByVal ptype As Integer
                                         ) As Dictionary(Of String, String)
        Dim c_item As KeyValuePair(Of String, String)
        Dim newpolicies_prepared As New Dictionary(Of String, String)

        Select Case ptype
            Case 0
                For Each c_item In newpolicies
                    If Me.SETTINGS(c_item.Key).controltype = 0 Then
                        If c_item.Value = "0" Then
                            newpolicies_prepared.Add(c_item.Key, Me.SETTINGS(c_item.Key).vnegative)
                        ElseIf c_item.Value = "1" Then
                            newpolicies_prepared.Add(c_item.Key, Me.SETTINGS(c_item.Key).vpositive)
                        ElseIf c_item.Value = "2" Then
                            newpolicies_prepared.Add(c_item.Key, Me.SETTINGS(c_item.Key).vnotset)
                        End If
                    Else
                        newpolicies_prepared.Add(c_item.Key, c_item.Value)
                    End If
                Next c_item
            Case 1
                For Each c_item In newpolicies
                    If Me.SYS_SETTINGS(c_item.Key).controltype = 0 Then
                        If c_item.Value = "0" Then
                            newpolicies_prepared.Add(c_item.Key, Me.SYS_SETTINGS(c_item.Key).vnegative)
                        ElseIf c_item.Value = "1" Then
                            newpolicies_prepared.Add(c_item.Key, Me.SYS_SETTINGS(c_item.Key).vpositive)
                        ElseIf c_item.Value = "2" Then
                            newpolicies_prepared.Add(c_item.Key, Me.SYS_SETTINGS(c_item.Key).vnotset)
                        End If
                    Else
                        newpolicies_prepared.Add(c_item.Key, c_item.Value)
                    End If
                Next c_item
        End Select

        Return newpolicies_prepared
    End Function

    Public Function getSettings() As Object
        Return Me.SETTINGS
    End Function

    Public Function getSystemSettings() As Object
        Return Me.SYS_SETTINGS
    End Function

    Private Function isPolicy_Changed(ByVal PID As String,
                                      ByVal oldpolicy As String,
                                      ByVal newpolicy As String,
                                      ByVal ptype As Integer
                                      ) As Boolean
        Select Case ptype
            Case 0
                If oldpolicy = Nothing And (newpolicy = Me.SETTINGS(PID).vnegative) Then
                    Return False
                ElseIf oldpolicy = newpolicy Then
                    Return False
                End If
            Case 1
                If oldpolicy = Nothing And (newpolicy = Me.SYS_SETTINGS(PID).vnegative) Then
                    Return False
                ElseIf oldpolicy = newpolicy Then
                    Return False
                End If
        End Select

        Return True
    End Function

    Public Function activateAdminAccount() As Boolean
        Dim adminfo() As String = Me.detectAdminAccount_WINAPI()

        Dim sap As New SetAdminPassword
        Dim retval As Integer
        Select Case adminfo(0)
            Case "2"
                sap.ShowDialog()
                If sap.DialogResult = DialogResult.OK Then
                    If Not adminfo(1) = String.Empty Then
                        If Not sap.TextBox1.Text = "" Then
                            Dim usobj As New USER_INFO_1003
                            usobj.usri1003_password = sap.TextBox1.Text
                            retval = NetUserSetInfo(String.Empty, adminfo(1), 1003, usobj, Nothing)
                            If Not retval = 0 Then Return False
                        End If
                        Return True
                    End If
                End If
            Case "1"
                Dim usobj2 As New USER_INFO_1008
                usobj2.usri1008_flags = UF_DONT_EXPIRE_PASSWD + UF_SCRIPT
                retval = NetUserSetInfo(String.Empty, adminfo(1), 1008, usobj2, Nothing)
                If Not retval = 0 Then Return False

                sap.ShowDialog()
                If sap.DialogResult = DialogResult.OK Then
                    If Not sap.TextBox1.Text = "" Then
                        Dim usobj As New USER_INFO_1003
                        usobj.usri1003_password = sap.TextBox1.Text
                        retval = NetUserSetInfo(String.Empty, adminfo(1), 1003, usobj, Nothing)
                        If Not retval = 0 Then Return False
                    End If
                    Me.AdministratorState = "2"
                    Return True
                End If
        End Select
        Return False
    End Function

    Private Function detectAdminAccount_WINAPI() As Object
        Dim bufPtr As IntPtr
        Dim entrread As Integer
        Dim totalentr As Integer
        Dim reshandle As Integer
        Dim FILTER_NORMAL_ACCOUNT As Integer = &H2
        Dim admret(2) As String
        Dim p As Boolean = False
        Dim UF_ACCOUNTDISABLE As Integer = &H2

        Try
            Dim retval As Integer = NetUserEnum(String.Empty, 20, FILTER_NORMAL_ACCOUNT, bufPtr, -1, entrread, totalentr, reshandle)
            If Not retval = 0 Then
                admret = {"0", String.Empty}
                Return admret
            End If

            If entrread > 0 Then
                Dim iptr As IntPtr = bufPtr
                Dim i As Integer
                Dim uinf As USER_INFO_20
                For i = 1 To entrread
                    uinf = New USER_INFO_20
                    uinf = CType(Marshal.PtrToStructure(iptr, GetType(USER_INFO_20)), USER_INFO_20)
                    If uinf.usri23_user_id = 500 Then
                        If CBool(uinf.usri23_flags And UF_ACCOUNTDISABLE) Then admret(0) = "1" Else admret(0) = "2"
                        admret(1) = uinf.usri23_name
                        Return admret
                    End If

                    iptr = New IntPtr(iptr.ToInt32 + Marshal.SizeOf(GetType(USER_INFO_20)))
                Next i
            End If
        Catch ex As Exception
            admret = {"0", String.Empty}
            Return admret
        End Try

        admret = {"0", String.Empty}
        Return admret
    End Function

    Public Function osverison_Compare(ByVal minosversion As String) As Boolean 'True - less or equal, False - more
        Dim mv() = minosversion.Split(".")
        For i As Integer = 0 To mv.Length - 1
            mv(i) = Int32.Parse(mv(i))
        Next
        Return (mv(0) < Me.osversion_major) Or (mv(0) = Me.osversion_major And mv(1) < Me.osversion_minor) Or (mv(0) = Me.osversion_major And mv(1) = Me.osversion_minor And mv(2) <= Me.osversion_build)
    End Function

    Private Function getSIDByUsername_WINAPI(ByVal username As String) As IntPtr
        Dim ptrSid As IntPtr
        Dim cbSid As Integer
        Dim refDomainName As New System.Text.StringBuilder
        Dim cbRefDomainName As Integer
        Dim peUse As SID_NAME_USE
        Dim StringSid As String = String.Empty

        Try
            'First call
            LookupAccountName(String.Empty, username, Nothing, cbSid, Nothing, cbRefDomainName, peUse)
            'Adjust buffers
            ptrSid = Marshal.AllocHGlobal(cbSid)
            refDomainName.EnsureCapacity(cbRefDomainName)
            'Second call
            If LookupAccountName(String.Empty, username, ptrSid, cbSid, refDomainName, cbRefDomainName, peUse) Then Return ptrSid
        Catch ex As Exception
            MessageBox.Show("Critical error: " & ex.Message(), gpe_main.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try
        Return Nothing
    End Function

    Private Function fetchPassword(ByVal length As Integer, ByRef ent As Random) As String
        Dim vowels As String = "aeiouy"
        Dim consonants As String = "bcdfghkmnprstvwxz"
        Dim numerals As String = "123456789"
        Dim i As Integer = 0
        Dim password As String = ""
        Do Until i >= length
            If length - i = 1 Then
                password &= numerals(ent.Next(0, numerals.Length - 1))
                i = i + 1
            Else
                Select Case ent.Next(0, 3)
                    Case 0
                        password &= consonants(ent.Next(0, consonants.Length - 1)) & vowels(ent.Next(0, vowels.Length - 1))
                        i = i + 2
                    Case 1
                        password &= vowels(ent.Next(0, vowels.Length - 1)) & consonants(ent.Next(0, consonants.Length - 1))
                        i = i + 2
                    Case 2
                        password &= numerals(ent.Next(0, numerals.Length - 1))
                        i = i + 1
                End Select
            End If
        Loop
        Return password
    End Function

    Public Function createUser_List(ByVal usernames() As String, ByVal passlength As Integer) As Dictionary(Of String, String)
        Dim i As Integer
        Dim uobj As New USER_INFO_1
        Dim ugobj As New LOCALGROUP_MEMBERS_INFO_0
        Dim ucreated As New Dictionary(Of String, String)
        Dim ufailure As String = ""
        Dim npass As String = ""
        Dim tnow As New DateTime
        tnow = DateTime.Now
        Dim tx As New DateTime(1970, 1, 1, 0, 0, 0, 0)
        Dim tspan As TimeSpan = (tnow - tx.ToLocalTime())
        Dim rnd As New Random(tspan.TotalSeconds)
        For i = 0 To usernames.Length - 1
            uobj = New USER_INFO_1
            uobj.usri1_name = usernames(i)
            If passlength > 0 Then
                npass = Me.fetchPassword(passlength, rnd)
                uobj.usri1_password = npass
            Else
                uobj.usri1_password = ""
            End If
            uobj.usri1_password_age = 0
            uobj.usri1_priv = 1
            uobj.usri1_script_path = ""
            uobj.usri1_comment = "Created by " & Application.ProductName
            uobj.usri1_flags = UF_PASSWD_CANT_CHANGE + UF_DONT_EXPIRE_PASSWD + UF_SCRIPT
            uobj.usri1_home_dir = ""
            Dim retval As Integer = NetUserAdd(String.Empty, 1, uobj, Nothing)
            If Not retval = 0 Then
                If ufailure = "" Then ufailure = usernames(i) & " (" & New System.ComponentModel.Win32Exception(retval).Message & ")" Else ufailure &= ", " & usernames(i) & " (" & New System.ComponentModel.Win32Exception(retval).Message & ")"
            Else
                If passlength > 0 Then ucreated.Add(usernames(i), npass)
                ugobj = New LOCALGROUP_MEMBERS_INFO_0
                ugobj.lgrmi0_sid = Me.getSIDByUsername_WINAPI(usernames(i))
                retval = NetLocalGroupAddMembers(String.Empty, Me.usersgroup_name, 0, ugobj, 1)
                If Not retval = 0 Then
                    If ufailure = "" Then ufailure = usernames(i) & " (" & New System.ComponentModel.Win32Exception(retval).Message & ")" Else ufailure &= ", " & usernames(i) & " (" & New System.ComponentModel.Win32Exception(retval).Message & ")"
                End If
            End If
        Next i
        If Not ufailure = "" Then
            Me.Exeption = ufailure
            MessageBox.Show("Не удалось создать следующих пользователей: " & ufailure, gpe_main.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End If
        If passlength = 0 Then Return Nothing Else Return ucreated
    End Function

    Public Function changePassword_List(ByVal userlist() As String, ByVal passlength As Integer) As Dictionary(Of String, String)
        Dim i As Integer = 0
        Dim usobj As New USER_INFO_1003
        Dim passchanged As New Dictionary(Of String, String)
        Dim pfailure As String = ""
        Dim npass As String = ""
        Dim tnow As New DateTime
        tnow = DateTime.Now
        Dim tx As New DateTime(1970, 1, 1, 0, 0, 0, 0)
        Dim tspan As TimeSpan = (tnow - tx.ToLocalTime())
        Dim rnd As New Random(tspan.TotalSeconds)
        For i = 0 To userlist.Length - 1
            usobj = New USER_INFO_1003
            If passlength = 0 Then
                usobj.usri1003_password = ""
            Else
                npass = Me.fetchPassword(passlength, rnd)
                usobj.usri1003_password = npass
            End If
            Dim retval As Integer = NetUserSetInfo(String.Empty, userlist(i), 1003, usobj, Nothing)
            If Not retval = 0 Then
                If pfailure = "" Then pfailure = userlist(i) & " (" & New System.ComponentModel.Win32Exception(retval).Message & ")" Else pfailure &= ", " & userlist(i) & " (" & New System.ComponentModel.Win32Exception(retval).Message & ")"
            Else
                If passlength > 0 Then passchanged.Add(userlist(i), npass)
            End If
        Next i
        If Not pfailure = "" Then
            Me.Exeption = pfailure
            MessageBox.Show("Не удалось поменять пароли следующим пользователям: " & pfailure, gpe_main.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End If
        If passlength = 0 Then Return Nothing Else Return passchanged
    End Function

    Public Function setPassword_List(ByVal ulist As Dictionary(Of String, String)) As Boolean
        Dim usobj As New USER_INFO_1003
        Dim pfailure As String = ""
        Dim ulist_item As KeyValuePair(Of String, String)
        For Each ulist_item In ulist
            usobj = New USER_INFO_1003
            usobj.usri1003_password = ulist_item.Value
            Dim retval As Integer = NetUserSetInfo(String.Empty, ulist_item.Key, 1003, usobj, Nothing)
            If Not retval = 0 Then
                If pfailure = "" Then pfailure = ulist_item.Key & " (" & New System.ComponentModel.Win32Exception(retval).Message & ")" Else pfailure &= ", " & ulist_item.Key & " (" & New System.ComponentModel.Win32Exception(retval).Message & ")"
            End If
        Next ulist_item
        If Not pfailure = "" Then
            Me.Exeption = "Не удалось поменять пароли следующим пользователям: " & pfailure
            Return False
        End If
        Return True
    End Function

    Private Sub SettingsInit(ByRef settings As Dictionary(Of String, setting),
                             ByRef sys_settings As Dictionary(Of String, setting),
                             ByRef profiles() As Object,
                             ByVal adminName As String)

        'Profiles
        Dim p1 As New profile
        p1.p_name = "Без ограничений"
        p1.p_settings = New Dictionary(Of String, String) From {
            {"CheckBox1", "0"},
            {"CheckBox2", "0"},
            {"CheckBox5", "0"},
            {"CheckBox3", "0"},
            {"CheckBox7", "0"},
            {"CheckBox9", "0"},
            {"CheckBox4", "0"},
            {"ComboBox2", "0"},
            {"CheckBox10", "0"},
            {"CheckBox11", "0"},
            {"CheckBox8", "0"},
            {"CheckBox12", "0"},
            {"CheckBox14", "0"},
            {"CheckBox15", "0"},
            {"CheckBox16", "0"},
            {"CheckBox18", "0"},
            {"CheckBox17", "0"},
            {"CheckBox19", "0"},
            {"CheckBox20", "0"},
            {"CheckBox24", "0"},
            {"CheckBox25", "0"},
            {"CheckBox26", "0"},
            {"CheckBox27", "0"},
            {"CheckBox30", "0"},
            {"CheckBox28", "0"}
        }

        Dim p2 As New profile
        p2.p_name = "Ученик: один пользователь"
        p2.p_settings = New Dictionary(Of String, String) From {
            {"CheckBox1", "1"},
            {"CheckBox2", "1"},
            {"CheckBox5", "1"},
            {"CheckBox3", "0"},
            {"CheckBox7", "1"},
            {"CheckBox9", "1"},
            {"CheckBox4", "1"},
            {"ComboBox2", "1"},
            {"CheckBox10", "1"},
            {"CheckBox11", "1"},
            {"CheckBox8", "1"},
            {"CheckBox12", "1"},
            {"CheckBox14", "1"},
            {"CheckBox15", "1"},
            {"CheckBox16", "1"},
            {"CheckBox18", "1"},
            {"CheckBox17", "1"},
            {"CheckBox19", "1"},
            {"CheckBox20", "1"},
            {"CheckBox24", "1"},
            {"CheckBox25", "1"},
            {"CheckBox26", "1"},
            {"CheckBox27", "1"},
            {"CheckBox30", "1"},
            {"CheckBox28", "0"}
        }

        Dim p3 As New profile
        p3.p_name = "Ученик: несколько пользователей"
        p3.p_settings = New Dictionary(Of String, String) From {
            {"CheckBox1", "1"},
            {"CheckBox2", "1"},
            {"CheckBox5", "1"},
            {"CheckBox3", "0"},
            {"CheckBox7", "1"},
            {"CheckBox9", "1"},
            {"CheckBox4", "1"},
            {"ComboBox2", "1"},
            {"CheckBox10", "1"},
            {"CheckBox11", "1"},
            {"CheckBox8", "1"},
            {"CheckBox12", "1"},
            {"CheckBox14", "1"},
            {"CheckBox15", "1"},
            {"CheckBox16", "1"},
            {"CheckBox18", "1"},
            {"CheckBox17", "1"},
            {"CheckBox19", "1"},
            {"CheckBox20", "0"},
            {"CheckBox24", "1"},
            {"CheckBox25", "1"},
            {"CheckBox26", "1"},
            {"CheckBox27", "1"},
            {"CheckBox30", "1"},
            {"CheckBox28", "0"}
        }

        Dim p4 As New profile
        p4.p_name = "Учитель: один пользователь"
        p4.p_settings = New Dictionary(Of String, String) From {
            {"CheckBox1", "0"},
            {"CheckBox2", "0"},
            {"CheckBox5", "1"},
            {"CheckBox3", "0"},
            {"CheckBox7", "0"},
            {"CheckBox9", "1"},
            {"CheckBox4", "0"},
            {"ComboBox2", "1"},
            {"CheckBox10", "0"},
            {"CheckBox11", "0"},
            {"CheckBox8", "1"},
            {"CheckBox12", "0"},
            {"CheckBox14", "1"},
            {"CheckBox15", "1"},
            {"CheckBox16", "1"},
            {"CheckBox18", "0"},
            {"CheckBox17", "0"},
            {"CheckBox19", "0"},
            {"CheckBox20", "1"},
            {"CheckBox24", "0"},
            {"CheckBox25", "0"},
            {"CheckBox26", "1"},
            {"CheckBox27", "0"},
            {"CheckBox30", "1"},
            {"CheckBox28", "0"}
        }

        Dim p5 As New profile
        p5.p_name = "Учитель: несколько пользователей"
        p5.p_settings = New Dictionary(Of String, String) From {
            {"CheckBox1", "0"},
            {"CheckBox2", "0"},
            {"CheckBox5", "0"},
            {"CheckBox3", "0"},
            {"CheckBox7", "0"},
            {"CheckBox9", "1"},
            {"CheckBox4", "0"},
            {"ComboBox2", "1"},
            {"CheckBox10", "0"},
            {"CheckBox11", "0"},
            {"CheckBox8", "1"},
            {"CheckBox12", "1"},
            {"CheckBox14", "1"},
            {"CheckBox15", "1"},
            {"CheckBox16", "1"},
            {"CheckBox18", "0"},
            {"CheckBox17", "0"},
            {"CheckBox19", "0"},
            {"CheckBox20", "0"},
            {"CheckBox24", "0"},
            {"CheckBox25", "0"},
            {"CheckBox26", "1"},
            {"CheckBox27", "0"},
            {"CheckBox30", "1"},
            {"CheckBox28", "0"}
        }

        profiles = {p1, p2, p3, p4, p5}
    End Sub
End Class
