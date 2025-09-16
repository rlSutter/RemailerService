' Class for reading and writing the Windows Registry 
' overcoming the restrictions imposed by 
' GetSetting y SaveSetting, which only allow you to 
' read and write from  HKEY_CURRENT_USER\Software\VB and VBA
' Programming: Sinhué Báez
' Date: FEB/25/2003
Imports Microsoft.Win32

Public Class CRegistry

    '---------public properties of the class----------------------
    Public ReadOnly Property HKeyLocalMachine() As RegistryKey
        Get
            Return Registry.LocalMachine
        End Get
    End Property

    Public ReadOnly Property HkeyClassesRoot() As RegistryKey
        Get
            Return Registry.ClassesRoot
        End Get
    End Property

    Public ReadOnly Property HKeyCurrentUser() As RegistryKey
        Get
            Return Registry.CurrentUser
        End Get
    End Property

    Public ReadOnly Property HKeyUsers() As RegistryKey
        Get
            Return Registry.Users
        End Get
    End Property

    Public ReadOnly Property HKeyCurrentConfig() As RegistryKey
        Get
            Return Registry.CurrentConfig
        End Get
    End Property

    Public ReadOnly Property HKeyDynData() As RegistryKey
        Get
            Return Registry.PerformanceData
        End Get
    End Property


    '--------------------------------------------------------------------
    'Description: Writes a value in the Registry
    'Parameters: 
    '   ParentKey: a RegistryKey that represents any of the six partent keys
    '              where you want to write.
    '   SubKey: a String with the name of the subkey(or nested subkeys)
    '           where you want to write. This Subkey or subkeys may exist 
    '           or not. If a subkey(or subkeys) doesn't exist 
    '           this method will create it.
    '   ValueName: a String with the name of the value to be created.
    '   Value: an Object with the value to be stored.
    'Returns: True if the method succeded, otherwise False.
    'Date: FEB/25/2003
    'Programming: Sinhué Báez
    '--------------------------------------------------------------------
    Public Function WriteValue(ByVal ParentKey As RegistryKey, _
                               ByVal SubKey As String, _
                               ByVal ValueName As String, _
                               ByVal Value As Object) As Boolean

        Dim Key As RegistryKey

        Try
            'Opens the given subkey
            Key = ParentKey.OpenSubKey(SubKey, True)
            If Key Is Nothing Then 'if Key doesn't exist then create it
                Key = ParentKey.CreateSubKey(SubKey)
            End If

            'sets the value
            Key.SetValue(ValueName, Value)

            Return True
        Catch e As Exception
            Return False
        End Try
    End Function

    '--------------------------------------------------------------------
    'Description: Reads a value from the Registry
    'Parameters: 
    '   ParentKey: a RegistryKey that represents any of the six partent keys
    '              where you want to read.
    '   SubKey: a String with the name of the subkey(or nested subkeys)
    '           where you want to read.
    '   ValueName: a String with the name of the value to be read.
    '   Value: an Object with the value to be read.
    'Returns: True if the method succeded, otherwise False.
    'Date: FEB/25/2003
    'Programming: Sinhué Báez
    '--------------------------------------------------------------------
    Public Function ReadValue(ByVal ParentKey As RegistryKey, _
                              ByVal SubKey As String, _
                              ByVal ValueName As String, _
                              ByRef Value As Object) As Boolean
        Dim Key As RegistryKey

        Try
            'opens the given subkey
            Key = ParentKey.OpenSubKey(SubKey, True)
            If Key Is Nothing Then 'it Key doesn't exist then throw an exception
                Throw New Exception("La Key del registro no existe")
            End If

            'Gets the value
            Value = Key.GetValue(ValueName)

            Return True
        Catch e As Exception
            Return False
        End Try
    End Function

    '--------------------------------------------------------------------
    'Description: Deletes a value from the Registry
    'Parameters: 
    '   ParentKey: a RegistryKey that represents any of the six partent keys
    '              where you want to delete a value from.
    '   SubKey: a String with the name of the subkey(or nested subkeys)
    '           where you want to delete a value from.
    '   ValueName:  a String with the name of the value to be deleted.
    'Returns: True if the method succeded, otherwise False.
    'Date: MAR/3/2003
    'Programming: Sinhué Báez
    '--------------------------------------------------------------------
    Public Function DeleteValue(ByVal ParentKey As RegistryKey, _
                                ByVal SubKey As String, _
                                ByVal ValueName As String) As Boolean
        Dim Key As RegistryKey

        Try
            'opens the given subkey
            Key = ParentKey.OpenSubKey(SubKey, True)

            'deletes the value
            If Not Key Is Nothing Then
                Key.DeleteValue(ValueName)
                Return True
            Else
                Return False
            End If
        Catch e As Exception
            Return False
        End Try

    End Function

    '--------------------------------------------------------------------
    'Description: Deletes a key from the Registry
    'Parameters: 
    '   ParentKey: a RegistryKey that represents any of the six partent keys
    '              where you want to delete a subkey from.
    '   SubKey: a String with the name of the subkey to be deleted.
    'Returns: True if the method succeded, otherwise False.
    'Date: MAR/3/2003
    'Programming: Sinhué Báez
    '--------------------------------------------------------------------
    Public Function DeleteSubKey(ByVal ParentKey As RegistryKey, _
                                ByVal SubKey As String) As Boolean

        Try
            'deletes the subkey and returns a False if Subkey doesn't exist
            ParentKey.DeleteSubKey(SubKey, False)
            Return True
        Catch e As Exception
            Return False
        End Try
    End Function


    '--------------------------------------------------------------------
    'Description: Crea una Key del registro
    'Parameters: 
    '   ParentKey: a RegistryKey that represents any of the six partent keys
    '              where you want to create a subkey.
    '   SubKey:  a String with the name of the subkey to be created.
    'Returns: True if the method succeded, otherwise False.
    'Date: MAR/3/2003
    'Programming: Sinhué Báez
    '--------------------------------------------------------------------
    Public Function CreateSubKey(ByVal ParentKey As RegistryKey, _
                               ByVal SubKey As String) As Boolean

        Try
            'creates the given subkey
            ParentKey.CreateSubKey(SubKey)
            Return True
        Catch e As Exception
            Return False
        End Try
    End Function

    Public Function GetAllSubKeys(ByVal ParentKey As RegistryKey, _
                                     ByVal SubKey As String, _
                                     ByRef strSubKeys() As String) As Boolean
        Dim Key As RegistryKey

        Try
            'opens the given subkey
            Key = ParentKey.OpenSubKey(SubKey)
            If Not Key Is Nothing Then
                'get all the subKeys (child subkeys)
                strSubKeys = Key.GetSubKeyNames
                Return True
            Else
                Return False
            End If
        Catch e As Exception
            Return False
        End Try
    End Function
End Class
