Public Class Model
    'Global model variables
    Private ArenaApp
    Private ArenaModel

    Public Sub New(ByVal numExpertAgents As Integer,
                   ByVal numNewAgents As Integer,
                   ByVal totalServiceDuration As Decimal,
                   ByVal agentLunchDuration As Decimal,
                   ByVal minNumAgentsDuringLunch As Integer,
                   ByVal meanArrivalTime As Decimal,
                   ByVal lowerTransferTime As Decimal,
                   ByVal upperTransferTime As Decimal,
                   ByVal expertAgentMeanServiceDuration As Decimal,
                   ByVal newAgentMeanServiceDuration As Decimal)
        'START ARENA AND THE MODEL
        ArenaApp = CType(CreateObject("Arena.Application"), Arena.Application)
        ArenaModel = ArenaApp.Models.Open("C:\Simulation\Testing\AirportServiceModelInitial.doe")

        'SET NUMBER OF AGENTS AND AGENT LUNCH DURATION
        Dim resourceIndex As Integer
        Dim newModule
        resourceIndex = 1
        For i = 1 To numExpertAgents
            ArenaModel.Modules(ArenaModel.Modules.Find(Arena.smFindType.smFindTag, "object.157")).Data("Resource Name(" & resourceIndex & ")") = "Agente Exp " & i
            resourceIndex += 1
        Next
        For i = 1 To numNewAgents
            ArenaModel.Modules(ArenaModel.Modules.Find(Arena.smFindType.smFindTag, "object.157")).Data("Resource Name(" & resourceIndex & ")") = "Agente Nue " & i
            resourceIndex += 1
        Next
        'Set the agents in resources table
        For i = 1 To numExpertAgents
            newModule = ArenaModel.Modules.Create("BasicProcess", "Resource", 0, 0)
            newModule.Data("Name") = "Agente Exp " & i
        Next
        For i = 1 To numNewAgents
            newModule = ArenaModel.Modules.Create("BasicProcess", "Resource", 0, 0)
            newModule.Data("Name") = "Agente Nue " & i
        Next
        'Set the lunch processes for the agents
        Dim firstLunchModuleNumber As Integer = 0
        Dim lastLunchModuleNumber As Integer = 0
        Dim numAgents As Integer = numExpertAgents + numNewAgents
        Dim numLunchProcesses As Integer

        If numAgents >= (minNumAgentsDuringLunch * 2) Then
            numLunchProcesses = 2
        Else
            numLunchProcesses = Math.Ceiling((numExpertAgents + numNewAgents) / minNumAgentsDuringLunch)
        End If

        Dim numAgentsPerLunchProcess As Integer = Math.Ceiling(numAgents / numLunchProcesses)
        Dim lunchProcessIndex As Integer = 0
        Dim lunchProcessX As Integer = -650
        Dim agentIndex As Integer = numAgentsPerLunchProcess + 1
        Dim currentLunchProcessModule As Object = 0

        For i = 1 To numExpertAgents
            If agentIndex > numAgentsPerLunchProcess Then
                lunchProcessIndex += 1
                lunchProcessX += 650
                agentIndex = 1
                currentLunchProcessModule = ArenaModel.Modules.Create("BasicProcess", "Process", lunchProcessX, 3000)
                currentLunchProcessModule.Data("Name") = "Grupo Refrigerio " & lunchProcessIndex
                currentLunchProcessModule.Data("Action") = "Seize Delay Release"
                currentLunchProcessModule.Data("Priority") = "High(1)"
                currentLunchProcessModule.Data("Resource Type(1)") = "Resource"
                currentLunchProcessModule.Data("Resource Name(" & agentIndex & ")") = "Agente Exp " & i
                currentLunchProcessModule.Data("DelayType") = "Constant"
                currentLunchProcessModule.Data("Value") = agentLunchDuration

                lastLunchModuleNumber = ArenaModel.Shapes.Find(Arena.smFindType.smFindTag, "object." & currentLunchProcessModule.shape.SerialNumber)
                If lunchProcessIndex = 1 Then
                    firstLunchModuleNumber = ArenaModel.Shapes.Find(Arena.smFindType.smFindTag, "object." & currentLunchProcessModule.shape.SerialNumber)
                End If
            Else
                currentLunchProcessModule.Data("Resource Name(" & agentIndex & ")") = "Agente Exp " & i
            End If
            agentIndex += 1
        Next

        For i = 1 To numNewAgents
            If agentIndex > numAgentsPerLunchProcess Then
                lunchProcessIndex += 1
                lunchProcessX += 650
                agentIndex = 1
                currentLunchProcessModule = ArenaModel.Modules.Create("BasicProcess", "Process", lunchProcessX, 3000)
                currentLunchProcessModule.Data("Name") = "Grupo Refrigerio " & lunchProcessIndex
                currentLunchProcessModule.Data("Action") = "Seize Delay Release"
                currentLunchProcessModule.Data("Priority") = "High(1)"
                currentLunchProcessModule.Data("Resource Type(1)") = "Resource"
                currentLunchProcessModule.Data("Resource Name(" & agentIndex & ")") = "Agente Nue " & i
                currentLunchProcessModule.Data("DelayType") = "Constant"
                currentLunchProcessModule.Data("Value") = agentLunchDuration

                lastLunchModuleNumber = ArenaModel.Shapes.Find(Arena.smFindType.smFindTag, "object." & currentLunchProcessModule.shape.SerialNumber)
                If lunchProcessIndex = 1 And numExpertAgents = 0 Then
                    firstLunchModuleNumber = ArenaModel.Shapes.Find(Arena.smFindType.smFindTag, "object." & currentLunchProcessModule.shape.SerialNumber)
                End If
            Else
                currentLunchProcessModule.Data("Resource Name(" & agentIndex & ")") = "Agente Nue " & i
            End If
            agentIndex += 1
        Next

        'Connect lunch processes
        Dim startLunchProcessModuleNumber = ArenaModel.Shapes.Find(Arena.smFindType.smFindTag, "object.194")
        Dim finishLunchProcessModuleNumber = ArenaModel.Shapes.Find(Arena.smFindType.smFindTag, "object.292")
        ArenaModel.Connections.Create(ArenaModel.Shapes(startLunchProcessModuleNumber), ArenaModel.Shapes(firstLunchModuleNumber))
        ArenaModel.Connections.Create(ArenaModel.Shapes(lastLunchModuleNumber), ArenaModel.Shapes(finishLunchProcessModuleNumber))

        'Set the conditional value for expert agents
        ArenaModel.Modules(ArenaModel.Modules.Find(Arena.smFindType.smFindTag, "object.97")).Data("Value") = numExpertAgents

        'SET TOTAL SERVICE DURATION AND AGENT LUNCH DURATION
        Dim customerCreatorModel = ArenaModel.Modules(ArenaModel.Modules.Find(Arena.smFindType.smFindTag, "object.11"))
        customerCreatorModel.Data("Max Batches") = "(" & totalServiceDuration & "*60-TNOW)*99999"
        ArenaModel.Modules(ArenaModel.Modules.Find(Arena.smFindType.smFindTag, "object.49")).Data("Value") = "TNOW <= " & totalServiceDuration & "*60"
        Dim LunchCreateModule = ArenaModel.Modules(ArenaModel.Modules.Find(Arena.smFindType.smFindTag, "object.194"))
        LunchCreateModule.Data("Value") = agentLunchDuration
        LunchCreateModule.Data("Max Batches") = numLunchProcesses
        LunchCreateModule.Data("First Creation") = Math.Round(totalServiceDuration / 2, 2)

        'SET MEAN ARRIVAL TIME
        customerCreatorModel.Data("Value") = meanArrivalTime

        'SET LOWER TRANSFER TIME AND UPPER TRANSFER TIME
        Dim transferTimeProcessModule = ArenaModel.Modules(ArenaModel.Modules.Find(Arena.smFindType.smFindTag, "object.30"))
        transferTimeProcessModule.Data("Min") = lowerTransferTime
        transferTimeProcessModule.Data("Max") = upperTransferTime

        'SET EXPERT AGENT MEAN SERVICE DURATION AND NEW AGENT MEAN SERVICE DURATION
        ArenaModel.Modules(ArenaModel.Modules.Find(Arena.smFindType.smFindTag, "object.114")).Data("Expression") = "EXPO( " & expertAgentMeanServiceDuration & " )"
        ArenaModel.Modules(ArenaModel.Modules.Find(Arena.smFindType.smFindTag, "object.131")).Data("Expression") = "EXPO( " & newAgentMeanServiceDuration & " )"
    End Sub

    Public Sub Execute()
        'ArenaModel.Go()
        'ArenaModel.SaveAs("C:\Simulation\Testing\DumpTesting\DumpTesting.doe")
        'ArenaApp.Quit()
    End Sub

    Public Function GetOutput() As String
        Return My.Computer.FileSystem.ReadAllText("C:\Simulation\AirportServiceModel.out")
    End Function
End Class
