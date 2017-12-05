Module Main

    Sub Main()
        'Check if program is not already running
        Dim appName As String = My.Application.Info.AssemblyName
        Dim appInstancesCounter As Integer = 0
        For Each process As Process In Process.GetProcesses()
            If process.ProcessName = appName Then
                appInstancesCounter += 1
            End If
        Next
        If appInstancesCounter > 1 Then
            Environment.Exit(0)
        End If

        'Reboot console application ans close
        Try
            For Each process As Process In Process.GetProcesses()
                If process.ProcessName = "SSD20172ConsoleApp" Then
                    process.Kill()
                End If
            Next

            Dim processStartInfo As New ProcessStartInfo
            processStartInfo.FileName = My.Application.Info.DirectoryPath + "/SSD20172ConsoleApp.exe"
            Dim consoleAppProcess As Process = Process.Start(processStartInfo)
            Environment.Exit(0)
        Catch ex As Exception
            Environment.Exit(0)
        End Try
    End Sub

End Module
