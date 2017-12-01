Module ModelExecutor

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

        'Test declares only to see the available functions
        Dim TestArenaApp As Arena.Application
        Dim TestArenaModel As Arena.Model
        'The real code begins

        '1) Initialize the program and the model
        Dim ArenaApp
        Dim ArenaModel
        ArenaApp = CType(CreateObject("Arena.Application"), Arena.Application)
        ArenaModel = ArenaApp.Models.Open("C:\Simulation\AirportServiceModel.doe")
        ArenaModel.Go()
        ArenaModel.SaveAs("C:\Simulation\Dump\Dump.doe")
        ArenaApp.Quit()
    End Sub

End Module
