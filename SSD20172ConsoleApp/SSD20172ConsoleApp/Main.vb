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

        Do
            Dim modelRequestString As String = Proxy.GetModelRequest()

            If Interpreter.GetValueFromModelRequestString(modelRequestString, "status") = "Submitted" Then
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

                Dim gson As New Gson()
                gson.AddAttribute("Status", "Finished")
                gson.AddAttribute("NumExpertAgents", numExpertAgents)
                gson.AddAttribute("NumNewAgents", numNewAgents)

                Console.WriteLine(Proxy.PostModelRequest(gson.GetString()))
            End If
            Threading.Thread.Sleep(1000)
        Loop
    End Sub

End Module
