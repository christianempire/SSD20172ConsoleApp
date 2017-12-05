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

        Dim failedConnectionAttempts As Integer
        Do
            If failedConnectionAttempts = 5 Then
                Try
                    Dim processStartInfo As New ProcessStartInfo
                    processStartInfo.FileName = My.Application.Info.DirectoryPath + "/SSD20172ConsoleAppRebooter.exe"
                    Dim consoleAppRebooterProcess As Process = Process.Start(processStartInfo)
                Catch ex As Exception

                Finally
                    failedConnectionAttempts = 0
                End Try
            End If

            Dim modelRequestString As String = Proxy.GetModelRequest()

            If modelRequestString = "" Then
                failedConnectionAttempts += 1
            Else
                failedConnectionAttempts = 0
            End If

            If modelRequestString <> "" And Interpreter.GetValueFromModelRequestString(modelRequestString, "status") = "Submitted" Then
                Dim inProgressRequest = New Gson()
                inProgressRequest.AddAttribute("Status", "InProgress")
                Proxy.PostModelRequest(inProgressRequest.GetString())

                Dim numExpertAgents As Integer = Integer.Parse(Interpreter.GetValueFromModelRequestString(modelRequestString, "numExpertAgents", False))
                Dim numNewAgents As Integer = Integer.Parse(Interpreter.GetValueFromModelRequestString(modelRequestString, "numNewAgents", False))
                Dim isAdvanced As Boolean = Boolean.Parse(Interpreter.GetValueFromModelRequestString(modelRequestString, "isAdvanced", False))

                Dim totalServiceDuration As Decimal = 8
                Dim agentLunchDuration As Decimal = 0.5
                Dim minAgentsDuringLunch As Integer = 4
                Dim meanArrivalTime As Decimal = 1.2
                Dim lowerTransferTime As Decimal = 2
                Dim upperTransferTime As Decimal = 3
                Dim expertAgentMeanServiceDuration As Decimal = 2.5
                Dim newAgentMeanServiceDuration As Decimal = 3.5
                If isAdvanced Then
                    totalServiceDuration = Decimal.Parse(Math.Round(Double.Parse(Interpreter.GetValueFromModelRequestString(modelRequestString, "totalServiceDuration", False)), 2))
                    agentLunchDuration = Decimal.Parse(Math.Round(Double.Parse(Interpreter.GetValueFromModelRequestString(modelRequestString, "agentLunchDuration", False)), 2))
                    minAgentsDuringLunch = Integer.Parse(Interpreter.GetValueFromModelRequestString(modelRequestString, "minAgentsDuringLunch", False))
                    meanArrivalTime = Decimal.Parse(Math.Round(Double.Parse(Interpreter.GetValueFromModelRequestString(modelRequestString, "meanArrivalTime", False)), 2))
                    lowerTransferTime = Decimal.Parse(Math.Round(Double.Parse(Interpreter.GetValueFromModelRequestString(modelRequestString, "lowerTransferTime", False)), 2))
                    upperTransferTime = Decimal.Parse(Math.Round(Double.Parse(Interpreter.GetValueFromModelRequestString(modelRequestString, "upperTransferTime", False)), 2))
                    expertAgentMeanServiceDuration = Decimal.Parse(Math.Round(Double.Parse(Interpreter.GetValueFromModelRequestString(modelRequestString, "expertAgentMeanServiceDuration", False)), 2))
                    newAgentMeanServiceDuration = Decimal.Parse(Math.Round(Double.Parse(Interpreter.GetValueFromModelRequestString(modelRequestString, "newAgentMeanServiceDuration", False)), 2))
                End If

                Dim model As New Model(numExpertAgents,
                                       numNewAgents,
                                       totalServiceDuration,
                                       agentLunchDuration,
                                       minAgentsDuringLunch,
                                       meanArrivalTime,
                                       lowerTransferTime,
                                       upperTransferTime,
                                       expertAgentMeanServiceDuration,
                                       newAgentMeanServiceDuration)
                model.Execute()

                Dim output As String = model.GetOutput()
                'Dim output As String = My.Computer.FileSystem.ReadAllText("C:\Simulation\AirportServiceModel.out")

                Dim avgTimeInSystem As Decimal = Interpreter.GetValueFromModelOutputString(output, "Cliente.TotalTime", "Average")
                Dim avgWaitingTime As Decimal = Interpreter.GetValueFromModelOutputString(output, "Atencion de un agentes.Queue.WaitingTime", "Average")
                Dim avgNumberInQueue As Decimal = Interpreter.GetValueFromModelOutputString(output, "Atencion de un agentes.Queue.NumberInQueue", "Average")
                Dim maxNumberInQueue As Decimal = Interpreter.GetValueFromModelOutputString(output, "Atencion de un agentes.Queue.NumberInQueue", "Maximum")

                Dim agents(numExpertAgents + numNewAgents - 1) As Gson
                Dim agentIndex = -1
                For i = 1 To numExpertAgents
                    agentIndex += 1
                    agents(agentIndex) = New Gson()
                    agents(agentIndex).AddAttribute("IsExpert", True)
                    agents(agentIndex).AddAttribute("Utilization", Interpreter.GetValueFromModelOutputString(output, "Agente Exp " & i & ".Utilization", "Average"))
                Next
                For i = 1 To numNewAgents
                    agentIndex += 1
                    agents(agentIndex) = New Gson()
                    agents(agentIndex).AddAttribute("IsExpert", False)
                    agents(agentIndex).AddAttribute("Utilization", Interpreter.GetValueFromModelOutputString(output, "Agente Nue " & i & ".Utilization", "Average"))
                Next

                Dim simulation As New Gson()
                'Inputs
                simulation.AddAttribute("IsAdvanced", isAdvanced)
                simulation.AddAttribute("TotalServiceDuration", totalServiceDuration)
                simulation.AddAttribute("AgentLunchDuration", agentLunchDuration)
                simulation.AddAttribute("MinAgentsDuringLunch", minAgentsDuringLunch)
                simulation.AddAttribute("MeanArrivalTime", meanArrivalTime)
                simulation.AddAttribute("LowerTransferTime", lowerTransferTime)
                simulation.AddAttribute("UpperTransferTime", upperTransferTime)
                simulation.AddAttribute("ExpertAgentMeanServiceDuration", expertAgentMeanServiceDuration)
                simulation.AddAttribute("NewAgentMeanServiceDuration", newAgentMeanServiceDuration)
                'Outputs
                simulation.AddAttribute("AvgTimeInSystem", avgTimeInSystem)
                simulation.AddAttribute("AvgWaitingTime", avgWaitingTime)
                simulation.AddAttribute("AvgNumberInQueue", avgNumberInQueue)
                simulation.AddAttribute("MaxNumberInQueue", maxNumberInQueue)
                simulation.AddAttribute("Agent", Gson.GetArrayString(agents))

                Dim middlewareRequest As New Gson()
                middlewareRequest.AddAttribute("Status", "Finished")
                middlewareRequest.AddAttribute("IsAdvanced", isAdvanced)
                middlewareRequest.AddAttribute("NumExpertAgents", numExpertAgents)
                middlewareRequest.AddAttribute("NumNewAgents", numNewAgents)
                middlewareRequest.AddAttribute("Simulation", simulation.GetString())

                Console.WriteLine(Proxy.PostModelRequest(middlewareRequest.GetString()))
            End If
            Threading.Thread.Sleep(1000)
        Loop
    End Sub

End Module
