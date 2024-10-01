Imports System
Imports System.Diagnostics.Eventing
Imports System.IO
Imports System.Net.Http
Imports System.Text.Json

Module Program
    Public Class Completion
        Public Property name As String
        Public Property timestamp As String
        Public Property expires As String
    End Class

    Public Class Person
        Public Property name As String
        Public Property completions As List(Of Completion)
    End Class

    Public Class TrainingCount
        Public Property trainingName As String
        Public Property completionCount As Integer
    End Class

    Public Class TrainingFiscalYearResult
        Public Property trainingName As String
        Public Property completedBy As List(Of String)
    End Class

    Public Class ExpiryResult
        Public Property personName As String
        Public Property trainings As List(Of ExpiryTraining)
    End Class

    Public Class ExpiryTraining
        Public Property trainingName As String
        Public Property status As String
        Public Property expirationDate As String
    End Class

    ' Three parts: 1. List each completed training with a count of how many people have completed that training.
    '              2. Given a list of trainings and a fiscal year (defined as 7/1/n-1 – 6/30/n), for each specified training, list all people that completed that training in the specified fiscal year.
    '                 *	Use parameters: Trainings = "Electrical Safety for Labs", "X-Ray Safety", "Laboratory Safety Training"; Fiscal Year = 2024
    '              3. Given a date, find all people that have any completed trainings that have already expired, or will expire within one month of the specified date (A training is considered expired the day after its expiration date). For each person found, list each completed training that met the previous criteria, with an additional field to indicate expired vs expires soon.
    '                 * Use date: Oct 1st, 2023
    ' NOTE: The .json files are generated in the /Data directory

    Sub Main(args As String())

        ' Path to the JSON file
        Dim filePath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "trainings.json")
        Dim outputDirectory As String = "..\..\..\Data\"

        ' Check if the file exists
        If Not File.Exists(filePath) Then
            Console.WriteLine("File not found!")
            Return
        End If

        ' Read the JSON file content
        Dim jsonString As String = File.ReadAllText(filePath)

        ' Deserialize the JSON content into a list of Person objects
        Dim persons As List(Of Person) = JsonSerializer.Deserialize(Of List(Of Person))(jsonString)

        ' Part 1: Count the number of people who completed each training (only the most recent completion)
        Dim trainingCountData As New List(Of TrainingCount)
        Dim trainingCount As New Dictionary(Of String, Integer)

        For Each person In persons
            ' Dictionary to store the most recent completion for each training for this person
            Dim mostRecentCompletions As New Dictionary(Of String, DateTime)

            For Each completion In person.completions
                Dim completionDate As DateTime
                If DateTime.TryParse(completion.timestamp, completionDate) Then
                    ' If the training has already been completed by this person, check if the current completion is more recent
                    If mostRecentCompletions.ContainsKey(completion.name) Then
                        If completionDate > mostRecentCompletions(completion.name) Then
                            mostRecentCompletions(completion.name) = completionDate
                        End If
                    Else
                        mostRecentCompletions(completion.name) = completionDate
                    End If
                End If
            Next

            ' Update the global count based on the most recent completions
            For Each training In mostRecentCompletions.Keys
                If trainingCount.ContainsKey(training) Then
                    trainingCount(training) += 1
                Else
                    trainingCount(training) = 1
                End If
            Next
        Next
        Console.WriteLine()

        ' Display the training completion counts and store the counts into a list for JSON output
        Console.WriteLine("Training Completion Counts (most recent only):" & vbCrLf)
        For Each kvp In trainingCount
            Console.WriteLine($"{kvp.Key}: {kvp.Value} people have completed this training.")
            trainingCountData.Add(New TrainingCount With {
                .trainingName = kvp.Key,
                .completionCount = kvp.Value
            })
        Next

        Dim part1Json As String = JsonSerializer.Serialize(trainingCountData)
        File.WriteAllText(Path.Combine(outputDirectory, "Part1_TrainingCompletionCounts.json"), part1Json)

        Console.WriteLine("--------------")

        ' Part 2: Filter people who completed specified trainings in the given fiscal year
        ' "Electrical Safety for Labs", "X-Ray Safety", "Laboratory Safety Training"; Fiscal Year = 2024    (not sure if I was supposed to read this in as an external variable?) 
        Dim fiscalYear = 2024
        Dim startFiscalYear As New DateTime(fiscalYear - 1, 7, 1)
        Dim endFiscalYear As New DateTime(fiscalYear, 6, 30)

        ' List of specified trainings to check
        Dim trainingsToCheck As String() = {"Electrical Safety for Labs", "X-Ray Safety", "Laboratory Safety Training"}

        Dim fiscalYearData As New List(Of TrainingFiscalYearResult)

        ' Filter people who completed these trainings in the fiscal year (only most recent completion)
        Console.WriteLine(vbCrLf & $"People who completed specified trainings in Fiscal Year {fiscalYear}:")

        For Each training In trainingsToCheck
            Dim result As New TrainingFiscalYearResult With {
                .trainingName = training,
                .completedBy = New List(Of String)
            }

            Console.WriteLine(vbCrLf & $"Training: {training}")

            For Each person In persons
                ' Find the most recent completion for the current training
                Dim mostRecentCompletion As DateTime? = Nothing

                For Each completion In person.completions
                    Dim completionDate As DateTime
                    If DateTime.TryParse(completion.timestamp, completionDate) Then
                        If completion.name = training Then
                            If mostRecentCompletion Is Nothing OrElse completionDate > mostRecentCompletion Then
                                mostRecentCompletion = completionDate
                            End If
                        End If
                    End If
                Next

                ' Check if the most recent completion falls within the fiscal year
                If mostRecentCompletion.HasValue AndAlso mostRecentCompletion.Value >= startFiscalYear AndAlso mostRecentCompletion.Value <= endFiscalYear Then
                    result.completedBy.Add(person.name)
                    Console.WriteLine($"  {person.name} completed this training on {mostRecentCompletion.Value.ToShortDateString()}")
                End If
            Next

            fiscalYearData.Add(result)
        Next

        ' Write Part 2 JSON output to file
        Dim part2Json As String = JsonSerializer.Serialize(fiscalYearData)
        File.WriteAllText(Path.Combine(outputDirectory, "Part2_FiscalYearCompletions.json"), part2Json)

        Console.WriteLine("--------------")

        ' Part 3: Find trainings that have expired or will expire soon as of October 1st, 2023 (Oct 2nd, as the day after is when expiration occurs.) 
        ' Given a date, find all people that have any completed trainings that have already expired, or will expire within one month of the specified date
        '   (A training is considered expired the day after its expiration date).
        '   For each person found, list each completed training that met the previous criteria, with an additional field to indicate expired vs expires soon.
        ' Use date Oct 1St, 2023
        ' Assumptions: 1. null expiration date for "expires" means the training is good forever, or it never happened, so it can't expire either way. 
        '              2. The expire date is 10/2, the day after...

        Dim specifiedDate As New DateTime(2023, 10, 2)
        Dim expiresSoonThreshold As DateTime = specifiedDate.AddMonths(-1) ' 1 month before October 2nd, 2023

        Dim expiryData As New List(Of ExpiryResult)

        Console.WriteLine(vbCrLf & $"People with trainings that have expired or will expire soon (as of {specifiedDate.ToShortDateString()}):")

        For Each person In persons
            Dim expiryResult As New ExpiryResult With {
                .personName = person.name,
                .trainings = New List(Of ExpiryTraining)
            }
            Dim hasExpiredOrSoonTraining As Boolean = False

            ' Check each training for expiration
            For Each completion In person.completions
                Dim expirationDate As DateTime


                ' Check if the training has an expiration date
                If DateTime.TryParse(completion.expires, expirationDate) Then

                    ' Rule: Expires soon if it expires between September 2nd and October 2nd, 2023
                    If expirationDate >= expiresSoonThreshold AndAlso expirationDate <= specifiedDate Then
                        expiryResult.trainings.Add(New ExpiryTraining With {
                            .trainingName = completion.name,
                            .status = "Expires Soon",
                            .expirationDate = expirationDate.ToShortDateString()
                        })
                        If Not hasExpiredOrSoonTraining Then
                            Console.WriteLine(vbCrLf & $"Person: {person.name}")
                            hasExpiredOrSoonTraining = True
                        End If
                        Console.WriteLine($"  Training: {completion.name}, Status: Expires Soon (on {expirationDate.ToShortDateString()})")
                        ' Rule: Expired if the expiration date is before October 2nd, 2023
                    ElseIf expirationDate < specifiedDate Then
                        expiryResult.trainings.Add(New ExpiryTraining With {
                            .trainingName = completion.name,
                            .status = "Expired",
                            .expirationDate = expirationDate.ToShortDateString()
                        })
                        If Not hasExpiredOrSoonTraining Then
                            Console.WriteLine(vbCrLf & $"Person: {person.name}")
                            hasExpiredOrSoonTraining = True
                        End If
                        Console.WriteLine($"  Training: {completion.name}, Status: Expired (on {expirationDate.ToShortDateString()})")
                    End If
                End If
            Next

            If expiryResult.trainings.Count > 0 Then
                expiryData.Add(expiryResult)
            End If
        Next

        ' Write Part 3 JSON output to file
        Dim part3Json As String = JsonSerializer.Serialize(expiryData)
        File.WriteAllText(Path.Combine(outputDirectory, "Part3_ExpiredOrSoonToExpire.json"), part3Json)

    End Sub
End Module