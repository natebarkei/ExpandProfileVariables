Imports Microsoft.Win32

Module Program

    Private FLAG_VERBOSE As Boolean = False
    Private FLAG_QUIET As Boolean = False
    Private VAR_PREPEND As String = "USER"
    Private FAIL_PREPROCESS As Boolean = False
    Private FLAG_REMOVE As Boolean = False
    Private FLAG_LIST As Boolean = False

    Sub Main(ByVal args As String())
        For Each arg As String In args
            arg = arg.ToLower
            If arg.StartsWith("/?") Or arg.StartsWith("/help") Then
                DisplayHelp()
                FAIL_PREPROCESS = True
                Exit For
            End If
            If arg.StartsWith("/remove") Then FLAG_REMOVE = True
            If arg.StartsWith("/list") Then FLAG_LIST = True
            If arg.StartsWith("/verbose") Then FLAG_VERBOSE = True
            If arg.StartsWith("/quiet") Then FLAG_QUIET = True
            If arg.StartsWith("/prepend=") Then
                Dim name As String = arg.Split("=")(1).Trim.ToUpper
                If Not String.IsNullOrWhiteSpace(name) Then
                    If FLAG_VERBOSE Then
                        ConsoleWriteLine("Setting Prepend to '{0}'", name)
                        VAR_PREPEND = name
                    End If
                Else
                    Console.WriteLine("ERR: Invalid prepend value provided.")
                    FAIL_PREPROCESS = True
                    Exit For
                End If
            End If
        Next


        If Not FAIL_PREPROCESS Then
            If FLAG_REMOVE And FLAG_LIST Then
                ConsoleWriteLine("ERR: You can not use /list and /remove at the same time.")


            ElseIf FLAG_REMOVE Then
                RemoveVariables()
            ElseIf FLAG_LIST Then
                ListVariables()
            Else
                ExpandVariables()
            End If

        End If

        If Debugger.IsAttached Then
            ConsoleWriteLine("Press <enter> to continue...")
            Console.ReadLine()
        End If
    End Sub

    Public Sub DisplayHelp()
        ConsoleWriteLine("Create environment variables for special folders.")
        ConsoleWriteLine("")
        ConsoleWriteLine("expandprofilevariables [/prepend={USER}][/remove][/list][/quiet][/verbose]")
        ConsoleWriteLine("")
        ConsoleWriteLine("  /prepend={USER} Change the beginning of the variable name created.")
        ConsoleWriteLine("                  This will always be changed into uppercase and combined")
        ConsoleWriteLine("                  with the special folder name.")
        ConsoleWriteLine("")
        ConsoleWriteLine("  /remove         Remove mode, this removes all of the created variables.")
        ConsoleWriteLine("  /list           List mode, this outputs set commands to update cmd parent")
        ConsoleWriteLine("                  process environments.")
        ConsoleWriteLine("                  Example:")
        ConsoleWriteLine("                  FOR /F """"tokens=*"""" %i in ('ExpandProfileVariables.exe /list') DO %i")
        ConsoleWriteLine("")
        ConsoleWriteLine("  /quiet          Quiet mode, this does not show any output.")
        ConsoleWriteLine("  /verbose        Verbose mode, this shows all actions and their results.")
        ConsoleWriteLine("")
        ConsoleWriteLine("The environment variables will be stored in the users environment and")
        ConsoleWriteLine("will be accessable from all newly created processes.")
        ConsoleWriteLine("")
        ConsoleWriteLine("For More info:https://github.com/natebarkei/ExpandProfileVariables")
        ConsoleWriteLine("")
    End Sub


    Public Sub ListVariables()
        Try
            Dim Key As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders")
            Trace.WriteLine("Opened Key succesfully")
            For Each v In Key.GetValueNames
                If Not v.StartsWith("{") Then
                    Dim vv As String = Key.GetValue(v).ToString
                    Dim EnvName As String = VAR_PREPEND & v.ToUpper.Replace(" ", "")
                    Trace.WriteLine(" Variable:" & EnvName)
                    Dim Exists As Boolean = Not String.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable(EnvName, EnvironmentVariableTarget.User))
                    If Exists Then
                        ConsoleWriteLine("set {0}={1}", EnvName, vv)
                    End If
                End If
            Next
        Catch ex As Exception
        End Try
    End Sub

    Public Sub RemoveVariables()
        Try
            ConsoleWrite("Removing Profile Variables... ")
            Dim VariablesRemoved As New List(Of String)

            Dim Key As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders")
            Trace.WriteLine("Opened Key succesfully")
            For Each v In Key.GetValueNames
                If Not v.StartsWith("{") Then
                    Dim EnvName As String = VAR_PREPEND & v.ToUpper.Replace(" ", "")
                    Trace.WriteLine(" Variable:" & EnvName)
                    Dim Exists As Boolean = Not String.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable(EnvName, EnvironmentVariableTarget.User))
                    If Exists Then
                        Environment.SetEnvironmentVariable(EnvName, String.Empty, EnvironmentVariableTarget.User)
                        VariablesRemoved.Add(String.Format(" Removed {0}", EnvName))
                    Else
                        VariablesRemoved.Add(String.Format(" Variable {0} did not exist", EnvName))
                    End If
                End If
            Next

            ConsoleWriteLine("[√] Complete")
            If FLAG_VERBOSE Then
                VariablesRemoved.ForEach(Sub(a) ConsoleWriteLine(a))
            End If

        Catch ex As Exception
            ConsoleWriteLine("[X] --> Unable to remove variables: {0}", ex.Message)
            Trace.WriteLine(ex.ToString)
        End Try

    End Sub

    Public Sub ExpandVariables()
        Try
            ConsoleWrite("Expanding Profile Variables... ")
            Dim VariablesAdded As New List(Of String)

            Dim Key As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders")
            Trace.WriteLine("Opened Key succesfully")
            For Each v In Key.GetValueNames
                If Not v.StartsWith("{") Then
                    Dim vv As String = Key.GetValue(v).ToString
                    Dim EnvName As String = VAR_PREPEND & v.ToUpper.Replace(" ", "")
                    Trace.WriteLine(" Variable:" & EnvName & "=" & vv)
                    Dim Existing As String = Environment.GetEnvironmentVariable(EnvName, EnvironmentVariableTarget.User)
                    If String.IsNullOrWhiteSpace(Existing) Or Existing <> vv Then
                        Environment.SetEnvironmentVariable(EnvName, vv, EnvironmentVariableTarget.User)
                        'This is a fix
                        Environment.SetEnvironmentVariable(EnvName, vv, EnvironmentVariableTarget.Process)
                        VariablesAdded.Add(String.Format(" {0}={1}", EnvName, vv))
                    Else
                        VariablesAdded.Add(String.Format(" {0}={1} <-- Value already exists (not updated)", EnvName, vv))
                    End If
                End If
            Next


            ConsoleWriteLine("[√] Complete")
            If FLAG_VERBOSE Then
                VariablesAdded.ForEach(Sub(a) ConsoleWriteLine(a))
            End If

        Catch ex As Exception
            ConsoleWriteLine("[X] --> Unable to expand variables: {0}", ex.Message)
            Trace.WriteLine(ex.ToString)
        End Try
    End Sub


    ' These functions just allow us to control in one place how much we output
    ' to the screen.

    Private Sub ConsoleWriteLine(format As String, ParamArray args() As Object)
        ConsoleWrite(format & Environment.NewLine, args)
    End Sub

    Private Sub ConsoleWrite(format As String, ParamArray args() As Object)
        If FLAG_QUIET Then Exit Sub
        Try
            Console.Write(format, args)
        Catch ex As Exception
            Console.Write(format)
            Console.Write(String.Join(",", args))
        End Try
    End Sub


End Module

