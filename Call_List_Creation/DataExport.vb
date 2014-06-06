Imports System.Configuration
Imports System.IO
Imports Rebex.Net


'=================================================================================================
'Class Name:	DataExport
'Description:	Holds functions used to parse a report and create a formatted data file.
'Property of Archipelago Systems, LLC.
'=================================================================================================
Public Class DataExport
#Region "Enums"
    Friend Enum ProcessingStatus
        SuccessfulRow
        ErroredRow
        SkippedRow
        WrittenRow
    End Enum
#End Region
#Region "Variables"
    'Declare variable used as the return string for main
    Dim results As String
    Private Const _className As String = "DataExport"
    Private Const skipReasonMeetingType As String = "Meeting Type"
    Private Const skipReasonPreviousInvalid As String = "Previous Meeting or Invalid Home Phone"
    Private Const skipReasonPrivPhone As String = "PRIV Keyword"
    Private Const skipReasonInvalidHomePhone As String = "Invalid Home Phone"
    Private Const skipReasonNoProvider As String = "Missing Provider ID"
    Private Const skipReasonNoOK As String = "OK Keyword Not Found"
    Private Const invalidCallTime As String = "The call date/time cannot be in the past."
    Private Const skipReasonScheduledDaysPrior As String = "The appointment was created within the timeframe of the 'no reminder necessary' rule."
    Private Const xlsProblem As String = "There was a problem with the xls file.  Be sure it is in the correct directory (as specified in the xlsFile key of the config file).  Also verify that it is formatted and named properly.  Also, be sure it is not open."
    Private Const ProgramError As String = "Please close any open CallList or Report files and try again."
    Private Const OpenReportFile As String = "The Eclipse OPEN report is open.  Please close it and try again."
    Private Const providerIDNotFound As String = "A provider id/meeting type combination matching a combination in the Eclipse OPEN report was not found in the CSV file.  Please review the CSV file to be sure all meeting types are represented either by name or by the key word 'All-Else'."
    Private Const problemReadingXLS As String = "There is a problem with the columns in the CSV file.  Please check that they all exist."
    Dim ReportFile As FileInfo
    Dim CallListFile As FileInfo
    Dim inputReader As StreamReader
    Dim outputReader As StreamReader
    Dim outputWriter As StreamWriter
    Dim exceptionWriter As StreamWriter
    Dim callListFilePath2 As String
    Shared rowWritten As Boolean = False
#End Region
    '=================================================================================================
    'Method Name:	DataExport.Main
    'Description:	Entry Point for the executable
    'Property of Archipelago Systems, LLC.
    '=================================================================================================
    Public Function Main(ByVal custID As String, ByVal daysPrior As String, ByVal scheduledDaysPrior As String, _
            ByVal callHour As String, ByVal callMinute As String, ByVal meetingTypeArray() As String, _
            ByVal useCSV As Boolean, ByVal CSVFile As String) As String

        Dim reportFilePath, callListFilePath, exceptionFilePath As String 'Strings to store filenames

        Try
            'Get current settings from the configuration file
            reportFilePath = ConfigurationSettings.AppSettings("ReportFile").ToString
            callListFilePath = ConfigurationSettings.AppSettings("CallListFile").ToString
            exceptionFilePath = ConfigurationSettings.AppSettings("ExceptionFile").ToString
        Catch e As Exception
            'Problem with the config file
            UpdateResults("There was a problem with the configuration file:")
            UpdateResults(e.ToString)
            Return results
        End Try
        ProcessTransactions(reportFilePath, callListFilePath, exceptionFilePath, custID, daysPrior, scheduledDaysPrior, callHour, callMinute, meetingTypeArray, useCSV, CSVFile)
        Return results
    End Function
    '=================================================================================================
    'Method Name:	DataExport.ArchiveCallList
    'Description:	Writes and executes the FTP script
    'Property of Archipelago Systems, LLC.
    '=================================================================================================
    Public Function ArchiveCallList() As Boolean
        Dim callListFilePath As String
        Dim outputArchivePath As String
        Try
            callListFilePath = ConfigurationSettings.AppSettings("CallListFile").ToString
            outputArchivePath = ConfigurationSettings.AppSettings("OutputArchive").ToString
        Catch ex As Exception
            Return False
        End Try
        Try
            'Create the directories if they're not there already
            Dim archive As String()
            archive = outputArchivePath.Split(".")
            archive(0) = archive(0) & "_" & Today.Now.Ticks
            archive(0) = archive(0) & ".txt"
            CheckForMissingDirectory(outputArchivePath)
            'Move the file to the archive
            File.Move(callListFilePath, archive(0).ToString())
        Catch ex As Exception
            Throw ex
        End Try
        Return True
    End Function
    '=================================================================================================
    'Method Name:	DataExport.UpdateResults
    'Description:	Updates a string that holds the results of the execution of the integration file creation code
    'Property of Archipelago Systems, LLC.
    '=================================================================================================
    Private Sub UpdateResults(ByVal status As String)
        results = results & status
        CloseReaders()
    End Sub
    Private Sub UpdateResultsFinal(ByVal status As String)
        results = results & status
    End Sub
    Private Sub CloseReaders()
        Try
            If Not exceptionWriter Is Nothing Then
                exceptionWriter.Close()
                exceptionWriter = Nothing
            End If
            If Not outputWriter Is Nothing Then
                outputWriter.Close()
                outputWriter = Nothing
            End If
            If Not CallListFile Is Nothing Then
                CallListFile = Nothing
            End If
            If Not ReportFile Is Nothing Then
                ReportFile = Nothing
            End If
            If Not inputReader Is Nothing Then
                inputReader.Close()
                inputReader = Nothing
            End If
            If Not outputReader Is Nothing Then
                outputReader.Close()
                outputReader = Nothing
            End If
        Catch e As Exception
            'Problem with the config file
            UpdateResults(ProgramError)
            If Not CallListFile Is Nothing Then
                CallListFile = Nothing
            End If
        End Try
    End Sub
    Private Sub ErrorCloseReaders()
        'Close readers, writers and files
        CallListFile = New FileInfo(callListFilePath2)
        If CallListFile.Exists Then
            CallListFile.Delete()
        End If
        CloseReaders()
    End Sub
    '=================================================================================================
    'Method Name:	DataExport.ProcessTransactions
    'Description:	Processes all transactions in the Input file
    'Property of Archipelago Systems, LLC.
    '=================================================================================================
    Private Sub ProcessTransactions(ByVal reportFilePath As String, ByVal callListFilePath As String, _
                                     ByVal exceptionFilePath As String, ByVal custID As String, _
                                     ByVal daysPrior As String, ByVal scheduledDaysPrior As String, _
                                     ByVal callHour As String, ByVal callMinute As String, _
                                     ByVal meetingTypeArray() As String, ByVal useCSV As Boolean, ByVal CSVFile As String)

        Dim row As Trans
        Dim skipCounter As Integer
        Dim skipReason As String
        Dim processedCounter As Integer
        Dim rowCounter As Integer
        Dim line As String
        Dim splitout As Array
        Dim records As ArrayList
        Dim x As Integer
        Dim y As Integer
        Dim exists As Boolean
        callListFilePath2 = callListFilePath

        Try
            'Create the directories if they're not there already
            CheckForMissingDirectory(exceptionFilePath)
            'Create an instace of StreamWriter to interact with the log files
            exceptionWriter = New StreamWriter(exceptionFilePath, False)
            'Write out a header in the file to separate runs
            exceptionWriter.WriteLine(New String("*", 100))
            exceptionWriter.WriteLine("* Run Date: " & Date.Now.ToString("f"))
            exceptionWriter.WriteLine(New String("*", 100))
            exceptionWriter.WriteLine("")
        Catch e As Exception
            'Problem with the config file
            UpdateResults(ProgramError)
            Exit Sub
        End Try
        If File.Exists(reportFilePath) Then
            ReportFile = New FileInfo(reportFilePath)
            'Create the directories if they're not there already
            CheckForMissingDirectory(reportFilePath)
            CheckForMissingDirectory(callListFilePath)
            Try
                inputReader = New StreamReader(reportFilePath)
                If inputReader.EndOfStream Then
                    UpdateResults("No appointments in report")
                    Exit Sub
                End If
            Catch e As Exception
                UpdateResults(OpenReportFile)
                Exit Sub
            End Try
            

            line = inputReader.ReadLine
            'skip the first six lines, which are headers
            Do Until x = 10
                line = inputReader.ReadLine
                If Mid(line, 10, 10) = "----------" Then
                    Exit Do
                End If
            Loop
            line = inputReader.ReadLine

            'Delete any lines in the output file left from the last run of the program      
            Try
                CallListFile = New FileInfo(callListFilePath)
                If CallListFile.Exists Then
                    CallListFile.Delete()
                End If
                outputWriter = New StreamWriter(callListFilePath)
            Catch e As Exception
                'Problem with the config file
                UpdateResults(ProgramError)
                Exit Sub
            End Try
            Do While Not line Is Nothing
                If line.Length < 2 Then
                    x = 0
                    Do While line Is Nothing OrElse line.Length < 2
                        If x = 10 Then Exit Do
                        line = inputReader.ReadLine
                        x += 1
                    Loop
                End If
                If line Is Nothing Then Exit Do
                If line.Length > 0 Then
                    'Increment the row counter by 1
                    rowCounter += 1
                    'Read the line and create a Trans object
                    row = New Trans(line, inputReader)
                    Select Case row.ProcessRow(reportFilePath, outputWriter, exceptionFilePath, _
                                inputReader, processedCounter, skipCounter, skipReason, custID, _
                                daysPrior, scheduledDaysPrior, callHour, callMinute, exceptionWriter, _
                                meetingTypeArray, useCSV, CSVFile)
                        Case ProcessingStatus.SuccessfulRow
                            processedCounter += 1
                        Case ProcessingStatus.SkippedRow
                            skipCounter += 1
                        Case ProcessingStatus.ErroredRow
                            If row.ProcessError = invalidCallTime Then
                                UpdateResults(invalidCallTime)
                                Exit Sub
                            ElseIf row.ProcessError = xlsProblem Then
                                UpdateResults(xlsProblem)
                                Exit Sub
                            ElseIf row.ProcessError = providerIDNotFound Then
                                UpdateResults(providerIDNotFound)
                                Exit Sub
                            ElseIf row.ProcessError = problemReadingXLS Then
                                UpdateResults(problemReadingXLS)
                                Exit Sub
                            End If
                            exceptionWriter.WriteLine("")
                        Case ProcessingStatus.WrittenRow
                    End Select
                End If
                'Get the next line
                line = inputReader.ReadLine
            Loop
            outputWriter.Close()
            outputWriter = Nothing
            CallListFile = Nothing
            'Count the number of lines in the output file
            outputReader = New StreamReader(callListFilePath)
            'Add the header lines to the records array
            records = New ArrayList
            records.Add(outputReader.ReadLine)
            records.Add(outputReader.ReadLine)
            records.Add(outputReader.ReadLine)
            records.Add(outputReader.ReadLine)
            line = outputReader.ReadLine
            Dim recordCount As Integer
            Do While Not line Is Nothing
                If line.Length > 0 Then
                    'Count the number records printed to the call list
                    'Find any duplicate phone numbers and only print the first appointment of the day
                    splitout = Split(line, ",")
                    'Create a new array with no duplicates
                    'Look through records to see if the phone numbers have been added yet
                    Dim recordsSplitout As Array
                    Dim replace As Boolean
                    y = 0
                    Do Until y = records.Count
                        recordsSplitout = Split(records(y), ",")
                        Try
                            If Trim(recordsSplitout(0)) = Trim(splitout(0)) And Trim(recordsSplitout(4)) = Trim(splitout(4)) Then
                                exists = True
                                If Convert.ToDateTime(Trim(recordsSplitout(2))) > Convert.ToDateTime(Trim(splitout(2))) Then
                                    'Replace the existing record with this one because the appt time is earlier
                                    records.RemoveAt(y)
                                    replace = True
                                End If
                                Exit Do
                            End If
                        Catch ex As Exception

                            'Handle argument out-of-range exception
                        End Try

                        y += 1
                    Loop
                    If Not exists Or replace Then
                        records.Add(line)
                        recordCount += 1
                    End If

                    exists = False
                    line = outputReader.ReadLine
                End If
            Loop

            Try
                x = 0
                outputReader.Close()
                outputReader = Nothing
                outputWriter = New StreamWriter(callListFilePath)
                Do Until x = (records.Count)
                    outputWriter.WriteLine(records(x))
                    x += 1
                Loop
                outputWriter.WriteLine("*EOF*")
            Catch ex As Exception
                UpdateResults(ProgramError)
                Exit Sub
            End Try

            exceptionWriter.WriteLine("")
            exceptionWriter.WriteLine("Rows Written: " & recordCount)
            exceptionWriter.Close()
            exceptionWriter = Nothing
            'Need to subtract the last line (*EOF*) from the count
            'Write out the processing results to the screen
            UpdateResultsFinal(New String("-", 100))
            UpdateResultsFinal(vbCrLf & "Run Date: " & Date.Now.ToString("s"))
            UpdateResultsFinal(vbCrLf & "Processing Complete" & vbCrLf)
            'If the row count is less than 0, set it to 0
            UpdateResultsFinal(vbCrLf & "Rows Written: " & recordCount & vbCrLf & vbCrLf)
            UpdateResultsFinal("Call list created in ")
            UpdateResultsFinal(System.Configuration.ConfigurationSettings.AppSettings.Item("CallListFile") & vbCrLf)
            UpdateResultsFinal(New String("-", 100))
        Else
            'Write a line to the log file to indicate that the input file did not exist
            exceptionWriter.WriteLine("Input File: " & reportFilePath & " does not exist.")
            exceptionWriter.Close()
            exceptionWriter = Nothing
            'Let user know that input file does not exist
            UpdateResults("Input File: " & reportFilePath & " does not exist.")
        End If
        'Clean up objects
        CloseReaders()
    End Sub
    '=================================================================================================
    'Class Name:	Trans
    'Description:	Contains operations used for parsing and processing a single row of the input file
    'Property of Archipelago Systems, LLC.
    '=================================================================================================
    Private Class Trans
        Private _row As String 'Member level variable used to hold data passed to constructor
        Private _error As String 'Member level variable used to hold the error message
        '=================================================================================================
        'Method Name:	Trans.New
        'Description:	Constructor that takes in a row from the flat file as an input parameter.
        'Property of Archipelago Systems, LLC.
        '=================================================================================================
        Friend Sub New(ByRef row As String, ByRef inputReader As StreamReader)
            'Put the row data in a member level variable
            If Left(Trim(Right(row, 11)), 4) = "PAGE" Then
                'Skip four lines to get to the next record
                row = inputReader.ReadLine()
                row = inputReader.ReadLine()
                row = inputReader.ReadLine()
                row = inputReader.ReadLine()
                row = inputReader.ReadLine()
                row = inputReader.ReadLine()

                'Alternate report has extra lines so check for this
                If Not row.Trim.Length > 1 Then
                    Do Until row.Trim.Length > 1
                        row = inputReader.ReadLine()
                    Loop
                End If

            End If

            _row = row
        End Sub

        '=================================================================================================
        'Method Name:	Trans.ProcessRow
        'Description:	Processes the row of data from the flat file
        'Property of Archipelago Systems, LLC.
        '=================================================================================================
        Friend Function ProcessRow(ByVal reportFilePath As String, ByVal outputWriter As StreamWriter, _
                            ByVal exceptionFilePath As String, ByRef inputReader As StreamReader, _
                            ByVal processedCounter As Integer, _
                            ByRef skipCounter As Integer, ByRef skipReason As String, _
                            ByVal custID As String, ByVal daysPrior As String, ByVal scheduledDaysPrior As String, _
                            ByVal callHour As String, ByVal callMinute As String, _
                            ByRef exceptionWriter As StreamWriter, _
                            ByVal meetingTypeArray() As String, ByVal useCSV As Boolean, ByVal CSVFile As String) As ProcessingStatus

            Dim strPhone As String
            Dim strNewPhone As String
            Dim apptString As String
            Dim createString As String
            Dim apptDateCompare As Date
            Dim apptTime As String
            Dim apptDate As Date
            Dim createDate As Date
            Dim callDate As Date
            Dim msgID As String
            Dim name As String
            Dim meetingType As String
            Dim skipMeeting As Boolean
            Dim provider As String
            Dim providerID As String
            Dim xlsFile As String
            Dim xlsReader As StreamReader
            Dim xlsLine As String
            Dim splitOut As Array
            Dim lookupEclipseOpenProviderID As String
            Dim lookupMeetingType As String
            skipMeeting = False
            Dim header As String
            'If this is NoneBut logic do one thing but if it's AllBut logic do something else
            If Trim(ConfigurationSettings.AppSettings("CallLogic").ToString.ToUpper) = "ALLBUT" Then
                'If this is the first record to be processed, we need to get the phone number and the appt date 
                If processedCounter = 0 Then
                    strNewPhone = "0"
                    strPhone = Trim(Mid(_row, 47, 18))
                    name = FormatName(Trim(Left(_row, 28)))
                    strNewPhone = FormatPhone(strPhone, name, exceptionWriter)
                    _row = inputReader.ReadLine
                    'Alternate report has extra lines so check for this
                    If Not _row.Trim.Length > 1 Then
                        Do Until _row.Trim.Length > 1
                            _row = inputReader.ReadLine()
                        Loop
                    End If

                    header = Trim(Right(_row, 11))
                    If Left(header, 4) = "PAGE" Then
                        'Skip four lines to get to the next record
                        _row = inputReader.ReadLine()
                        _row = inputReader.ReadLine()
                        _row = inputReader.ReadLine()
                        _row = inputReader.ReadLine()
                        _row = inputReader.ReadLine()
                        _row = inputReader.ReadLine()

                        'Alternate report has extra lines so check for this
                        If Not _row.Trim.Length > 1 Then
                            Do Until _row.Trim.Length > 1
                                _row = inputReader.ReadLine()
                            Loop
                        End If


                    End If

                    apptString = Mid(_row, 65, 10)
                    callDate = Convert.ToDateTime(apptString).AddDays(-Convert.ToInt32(daysPrior))
                    If callDate < Today Then
                        _error = invalidCallTime
                        Return ProcessingStatus.ErroredRow
                    End If
                    meetingType = Trim(Mid(_row, 87, 11))
                    skipMeeting = DetermineMeeting(meetingType, meetingTypeArray)
                    WriteHeader(outputWriter, custID, callHour, callMinute, callDate)
                    provider = Trim(Mid(_row, 26, 4)).ToUpper()
                    If useCSV = True Then
                        '*****************************************************************************************************
                        '   Find the corresponding IVR provider id based on Eclipse OPEN provider 
                        '   id and meeting type by parsing the xls file
                        '*****************************************************************************************************
                        Try
                            xlsFile = CSVFile
                            xlsReader = New StreamReader(xlsFile)
                            xlsLine = xlsReader.ReadLine
                            xlsLine = xlsReader.ReadLine
                            splitOut = Split(xlsLine, ",")
                            lookupEclipseOpenProviderID = Trim(splitOut(3)).ToUpper()
                            lookupMeetingType = Trim(splitOut(2)).ToUpper.ToUpper
                        Catch ex As Exception
                            _error = xlsProblem
                            If Not xlsReader Is Nothing Then
                                xlsReader.Close()
                                xlsReader = Nothing
                            End If
                            Return ProcessingStatus.ErroredRow
                        End Try
                        Try
                            Do While Not xlsLine Is Nothing
                                If xlsLine.Length > 0 Then
                                    splitOut = Split(xlsLine, ",")
                                    lookupEclipseOpenProviderID = Trim(splitOut(3)).ToUpper()
                                    lookupMeetingType = Trim(splitOut(2)).ToUpper
                                    If provider = lookupEclipseOpenProviderID And meetingType.ToUpper = lookupMeetingType.ToUpper Then
                                        providerID = Trim(splitOut(1))
                                        msgID = Trim(splitOut(4))
                                        Exit Do
                                    End If
                                End If
                                xlsLine = xlsReader.ReadLine
                            Loop
                        Catch e As Exception
                            _error = problemReadingXLS
                            If Not xlsReader Is Nothing Then
                                xlsReader.Close()
                                xlsReader = Nothing
                            End If
                            Return ProcessingStatus.ErroredRow
                        End Try
                        If Not xlsReader Is Nothing Then
                            xlsReader.Close()
                            xlsReader = Nothing
                        End If
                        xlsReader = New StreamReader(xlsFile)
                        xlsLine = xlsReader.ReadLine
                        xlsLine = xlsReader.ReadLine
                        If providerID = Nothing Then
                            'loop through the excel spreadsheet again to see if an All-Else exists
                            Try
                                Do While Not xlsLine Is Nothing
                                    If xlsLine.Length > 0 Then
                                        splitOut = Split(xlsLine, ",")
                                        lookupEclipseOpenProviderID = Trim(splitOut(3)).ToUpper()
                                        lookupMeetingType = Trim(splitOut(2)).ToUpper
                                        If provider = lookupEclipseOpenProviderID And lookupMeetingType.ToUpper = "ALL-ELSE" Then
                                            providerID = Trim(splitOut(1))
                                            msgID = Trim(splitOut(4))
                                            Exit Do
                                        End If
                                    End If
                                    xlsLine = xlsReader.ReadLine
                                Loop
                            Catch e As Exception
                                _error = problemReadingXLS
                                If Not xlsReader Is Nothing Then
                                    xlsReader.Close()
                                    xlsReader = Nothing
                                End If
                                Return ProcessingStatus.ErroredRow
                            End Try
                            If providerID = Nothing Then
                                _error = providerIDNotFound
                                If Not xlsReader Is Nothing Then
                                    xlsReader.Close()
                                    xlsReader = Nothing
                                End If
                                Return ProcessingStatus.ErroredRow
                            End If
                        End If
                        If Not xlsReader Is Nothing Then
                            xlsReader.Close()
                            xlsReader = Nothing
                        End If
                    Else
                        GetProviderID_FromConfig(providerID, provider)
                    End If
                    apptString = Trim(Mid(_row, 65, 10))
                    apptDateCompare = Convert.ToDateTime(apptString)
                    apptTime = Trim(Mid(_row, 113, 10))
                    If apptTime = "12:00n" Then
                        apptTime = "12:00pm"
                    End If
                    apptDate = Convert.ToDateTime(apptString & " " & apptTime)
                    If scheduledDaysPrior <> "" Then
                        createString = Trim(Mid(_row, 77, 10))
                        createDate = Convert.ToDateTime(createString)
                        If createDate.AddDays(Trim(scheduledDaysPrior)).Date >= apptDateCompare Then
                            skipReason = skipReasonScheduledDaysPrior
                        Else
                            WriteAllButRecord(strPhone, strNewPhone, providerID, apptDate, skipMeeting, msgID, name, outputWriter)
                        End If
                    Else
                        WriteAllButRecord(strPhone, strNewPhone, providerID, apptDate, skipMeeting, msgID, name, outputWriter)
                    End If
                Else
                    If _row Is Nothing Then _row = inputReader.ReadLine()
                    If Not _row.Trim.Length > 1 Then
                        If Not _row.Trim.Length > 1 Then
                            Do Until _row.Trim.Length > 1
                                _row = inputReader.ReadLine()
                            Loop
                        End If
                    End If
                    If _row.Trim.Chars(0).IsLetter(_row.Chars(0)) Then
                        'It's a new record so get the phone number
                        strPhone = Trim(Mid(_row, 47, 18))
                        name = FormatName(Trim(Left(_row, 28)))
                        strNewPhone = FormatPhone(strPhone, name, exceptionWriter)
                        If strNewPhone.Length = 10 Then
                            _row = inputReader.ReadLine
                            'Alternate report has extra lines so check for this
                            If Not _row.Trim.Length > 1 Then
                                Do Until _row.Trim.Length > 1
                                    _row = inputReader.ReadLine()
                                Loop
                            End If

                            header = Trim(Right(_row, 11))
                            If Left(header, 4) = "PAGE" Then
                                'Skip four lines to get to the next record
                                _row = inputReader.ReadLine()
                                _row = inputReader.ReadLine()
                                _row = inputReader.ReadLine()
                                _row = inputReader.ReadLine()
                                _row = inputReader.ReadLine()
                                _row = inputReader.ReadLine()

                                'Alternate report has extra lines so check for this
                                If Not _row.Trim.Length > 1 Then
                                    Do Until _row.Trim.Length > 1
                                        _row = inputReader.ReadLine()
                                    Loop
                                End If

                            End If

                            meetingType = Trim(Mid(_row, 87, 11))
                            'Look through the meeting types that should be skipped
                            skipMeeting = DetermineMeeting(meetingType, meetingTypeArray)
                            If skipMeeting <> True Then
                                'Get the provider ID
                                provider = Trim(Mid(_row, 26, 4)).ToUpper()
                                If useCSV = True Then
                                    '*****************************************************************************************************
                                    '   Find the corresponding IVR provider id based on Eclipse OPEN provider 
                                    '   id and meeting type by parsing the xls file
                                    '*****************************************************************************************************
                                    Try
                                        xlsFile = CSVFile
                                        xlsReader = New StreamReader(xlsFile)
                                        xlsLine = xlsReader.ReadLine
                                        xlsLine = xlsReader.ReadLine
                                        splitOut = Split(xlsLine, ",")
                                        lookupEclipseOpenProviderID = Trim(splitOut(3)).ToUpper()
                                        lookupMeetingType = Trim(splitOut(2)).ToUpper
                                    Catch ex As Exception
                                        _error = xlsProblem
                                        If Not xlsReader Is Nothing Then
                                            xlsReader.Close()
                                            xlsReader = Nothing
                                        End If
                                        Return ProcessingStatus.ErroredRow
                                    End Try
                                    Try
                                        Do While Not xlsLine Is Nothing
                                            If xlsLine.Length > 0 Then
                                                splitOut = Split(xlsLine, ",")
                                                lookupEclipseOpenProviderID = Trim(splitOut(3)).ToUpper()
                                                lookupMeetingType = Trim(splitOut(2)).ToUpper
                                                If provider = lookupEclipseOpenProviderID And meetingType.ToUpper = lookupMeetingType.ToUpper Then
                                                    providerID = Trim(splitOut(1))
                                                    msgID = Trim(splitOut(4))
                                                    Exit Do
                                                End If
                                            End If
                                            xlsLine = xlsReader.ReadLine
                                        Loop
                                    Catch e As Exception
                                        _error = problemReadingXLS
                                        If Not xlsReader Is Nothing Then
                                            xlsReader.Close()
                                            xlsReader = Nothing
                                        End If
                                        Return ProcessingStatus.ErroredRow
                                    End Try
                                    If Not xlsReader Is Nothing Then
                                        xlsReader.Close()
                                    End If
                                    xlsReader = New StreamReader(xlsFile)
                                    xlsLine = xlsReader.ReadLine
                                    xlsLine = xlsReader.ReadLine
                                    If providerID = Nothing Then
                                        'loop through the excel spreadsheet again to see if an All-Else exists
                                        Try
                                            Do While Not xlsLine Is Nothing
                                                If xlsLine.Length > 0 Then
                                                    splitOut = Split(xlsLine, ",")
                                                    lookupEclipseOpenProviderID = Trim(splitOut(3)).ToUpper()
                                                    lookupMeetingType = Trim(splitOut(2)).ToUpper
                                                    If provider = lookupEclipseOpenProviderID And lookupMeetingType.ToUpper = "ALL-ELSE" Then
                                                        providerID = Trim(splitOut(1))
                                                        msgID = Trim(splitOut(4))
                                                        Exit Do
                                                    End If
                                                End If
                                                xlsLine = xlsReader.ReadLine
                                            Loop
                                        Catch e As Exception
                                            _error = problemReadingXLS
                                            If Not xlsReader Is Nothing Then
                                                xlsReader.Close()
                                                xlsReader = Nothing
                                            End If
                                            Return ProcessingStatus.ErroredRow
                                        End Try
                                        If providerID = Nothing Then
                                            _error = providerIDNotFound
                                            If Not xlsReader Is Nothing Then
                                                xlsReader.Close()
                                                xlsReader = Nothing
                                            End If
                                            Return ProcessingStatus.ErroredRow
                                        End If
                                    End If
                                    If Not xlsReader Is Nothing Then
                                        xlsReader.Close()
                                        xlsReader = Nothing
                                    End If
                                Else
                                    GetProviderID_FromConfig(providerID, provider)
                                End If
                                apptString = Trim(Mid(_row, 65, 10))
                                apptDateCompare = Convert.ToDateTime(apptString)
                                apptTime = Trim(Mid(_row, 113, 10))
                                'Must handle the case of 12:00n 
                                If apptTime = "12:00n" Then
                                    apptTime = "12:00pm"
                                End If
                                'Convert to date
                                apptDate = Convert.ToDateTime(apptString & " " & apptTime)
                                If Trim(scheduledDaysPrior) <> "" Then
                                    createString = Trim(Mid(_row, 77, 10))
                                    createDate = Convert.ToDateTime(createString)
                                    If createDate.AddDays(scheduledDaysPrior).Date >= apptDateCompare Then
                                        skipReason = skipReasonScheduledDaysPrior
                                        Return ProcessingStatus.SkippedRow
                                    Else
                                        If strNewPhone.Length <> 10 Then
                                            If Not rowWritten Then
                                                exceptionWriter.WriteLine("Invalid phone number: " & name & " " & strNewPhone)
                                            End If
                                            rowWritten = True
                                        End If
                                        WriteAllButRecord(strPhone, strNewPhone, providerID, apptDate, skipMeeting, msgID, name, outputWriter)
                                    End If
                                Else
                                    If strNewPhone.Length <> 10 Then
                                        If name <> "written" Then
                                            exceptionWriter.WriteLine("Invalid phone number: " & name & " " & strNewPhone)
                                        End If
                                        rowWritten = True
                                    End If
                                    WriteAllButRecord(strPhone, strNewPhone, providerID, apptDate, skipMeeting, msgID, name, outputWriter)
                                End If
                            Else
                                Return ProcessingStatus.SkippedRow
                            End If
                        Else
                            Return ProcessingStatus.SkippedRow
                        End If
                    End If
                End If
                Return ProcessingStatus.SuccessfulRow
            ElseIf Trim(ConfigurationSettings.AppSettings("CallLogic").ToString.ToUpper) = "NONEBUT" Then
                'If this is the first record to be processed, we need to get the phone number and the appt date 
                If processedCounter = 0 Then
                    strNewPhone = "0"
                    strPhone = Trim(Mid(_row, 47, 18))
                    name = FormatName(Trim(Left(_row, 28)))
                    strNewPhone = FormatPhone(strPhone, name, exceptionWriter)
                    _row = inputReader.ReadLine
                    'Alternate report has extra lines so check for this
                    If Not _row.Trim.Length > 1 Then
                        Do Until _row.Trim.Length > 1
                            _row = inputReader.ReadLine()
                        Loop
                    End If

                    header = Trim(Right(_row, 11))
                    If Left(header, 4) = "PAGE" Then
                        'Skip four lines to get to the next record
                        _row = inputReader.ReadLine()
                        _row = inputReader.ReadLine()
                        _row = inputReader.ReadLine()
                        _row = inputReader.ReadLine()
                        _row = inputReader.ReadLine()
                        _row = inputReader.ReadLine()

                        'Alternate report has extra lines so check for this
                        If Not _row.Trim.Length > 1 Then
                            Do Until _row.Trim.Length > 1
                                _row = inputReader.ReadLine()
                            Loop
                        End If

                    End If

                    apptString = Mid(_row, 65, 10)
                    callDate = Convert.ToDateTime(apptString).AddDays(-Convert.ToInt32(daysPrior))
                    If callDate < Today Then
                        _error = invalidCallTime
                        Return ProcessingStatus.ErroredRow
                    End If
                    meetingType = Trim(Mid(_row, 87, 11))
                    'Look through the meeting types that should be skipped
                    skipMeeting = DetermineMeeting(meetingType, meetingTypeArray)
                    provider = Trim(Mid(_row, 26, 4)).ToUpper()
                    If useCSV = True Then
                        '*****************************************************************************************************
                        '   Find the corresponding IVR provider id based on Eclipse OPEN provider 
                        '   id and meeting type by parsing the CSV file
                        '*****************************************************************************************************
                        Try
                            xlsFile = CSVFile
                            xlsReader = New StreamReader(xlsFile)
                            xlsLine = xlsReader.ReadLine
                            xlsLine = xlsReader.ReadLine
                            splitOut = Split(xlsLine, ",")
                            lookupEclipseOpenProviderID = Trim(splitOut(3)).ToUpper()
                            lookupMeetingType = Trim(splitOut(2)).ToUpper
                        Catch ex As Exception
                            _error = xlsProblem
                            If Not xlsReader Is Nothing Then
                                xlsReader.Close()
                                xlsReader = Nothing
                            End If
                            Return ProcessingStatus.ErroredRow
                        End Try
                        Try
                            Do While Not xlsLine Is Nothing
                                If xlsLine.Length > 0 Then
                                    splitOut = Split(xlsLine, ",")
                                    lookupEclipseOpenProviderID = Trim(splitOut(3)).ToUpper()
                                    lookupMeetingType = Trim(splitOut(2)).ToUpper
                                    If provider = lookupEclipseOpenProviderID And meetingType.ToUpper = lookupMeetingType.ToUpper Then
                                        providerID = Trim(splitOut(1))
                                        msgID = Trim(splitOut(4))
                                        Exit Do
                                    End If
                                End If
                                xlsLine = xlsReader.ReadLine
                            Loop
                        Catch e As Exception
                            _error = problemReadingXLS
                            If Not xlsReader Is Nothing Then
                                xlsReader.Close()
                                xlsReader = Nothing
                            End If
                            Return ProcessingStatus.ErroredRow
                        End Try
                        If Not xlsReader Is Nothing Then
                            xlsReader.Close()
                        End If
                        xlsReader = New StreamReader(xlsFile)
                        xlsLine = xlsReader.ReadLine
                        xlsLine = xlsReader.ReadLine
                        If providerID = Nothing Then
                            'loop through the excel spreadsheet again to see if an All-Else exists
                            Try
                                Do While Not xlsLine Is Nothing
                                    If xlsLine.Length > 0 Then
                                        splitOut = Split(xlsLine, ",")
                                        lookupEclipseOpenProviderID = Trim(splitOut(3)).ToUpper()
                                        lookupMeetingType = Trim(splitOut(2)).ToUpper
                                        If provider = lookupEclipseOpenProviderID And lookupMeetingType.ToUpper = "ALL-ELSE" Then
                                            providerID = Trim(splitOut(1))
                                            msgID = Trim(splitOut(4))
                                            Exit Do
                                        End If
                                    End If
                                    xlsLine = xlsReader.ReadLine
                                Loop
                            Catch e As Exception
                                _error = problemReadingXLS
                                If Not xlsReader Is Nothing Then
                                    xlsReader.Close()
                                    xlsReader = Nothing
                                End If
                                Return ProcessingStatus.ErroredRow
                            End Try
                        End If
                        If Not xlsReader Is Nothing Then
                            xlsReader.Close()
                            xlsReader = Nothing
                        End If
                    Else
                        GetProviderID_FromConfig(providerID, provider)
                    End If
                    WriteHeader(outputWriter, custID, callHour, callMinute, callDate)
                    apptString = Trim(Mid(_row, 65, 10))
                    apptTime = Trim(Mid(_row, 113, 10))
                    If apptTime = "12:00n" Then
                        apptTime = "12:00pm"
                    End If
                    apptDate = Convert.ToDateTime(apptString & " " & apptTime)
                    If scheduledDaysPrior <> "" Then
                        createString = Trim(Mid(_row, 77, 10))
                        createDate = Convert.ToDateTime(createString)
                        If createDate.AddDays(Trim(scheduledDaysPrior)).Date >= apptDateCompare Then
                            skipReason = skipReasonScheduledDaysPrior
                        Else
                            If strNewPhone.Length <> 10 Then
                                If name <> "written" Then
                                    exceptionWriter.WriteLine("Invalid phone number: " & name & " " & strNewPhone)
                                End If
                                rowWritten = True
                            End If
                            If useCSV = False Then
                                WriteNoneButRecord(strPhone, strNewPhone, providerID, apptDate, skipMeeting, msgID, name, outputWriter)
                            Else
                                WriteNoneButRecord(strPhone, strNewPhone, providerID, apptDate, skipMeeting, msgID, name, outputWriter)
                            End If
                        End If
                    Else
                        If strNewPhone.Length <> 10 Then
                            If name <> "written" Then
                                exceptionWriter.WriteLine("Invalid phone number: " & name & " " & strNewPhone)
                            End If
                            rowWritten = True
                        End If
                        WriteNoneButRecord(strPhone, strNewPhone, providerID, apptDate, skipMeeting, msgID, name, outputWriter)
                    End If
                End If
                If _row.Trim.Chars(0).IsLetter(_row.Chars(0)) Then
                    'It's a new record so get the phone number
                    strPhone = Trim(Mid(_row, 47, 18))
                    name = FormatName(Trim(Left(_row, 28)))

                    strNewPhone = FormatPhone(strPhone, name, exceptionWriter)
                    If Right(strPhone, 2).ToUpper = "OK" Then
                        If strNewPhone.Length = 10 Then
                            _row = inputReader.ReadLine
                            'Alternate report has extra lines so check for this
                            If Not _row.Trim.Length > 1 Then
                                Do Until _row.Trim.Length > 1
                                    _row = inputReader.ReadLine()
                                Loop
                            End If

                            header = Trim(Right(_row, 11))
                            If Left(header, 4) = "PAGE" Then
                                'Skip lines to get to the next record
                                _row = inputReader.ReadLine()
                                _row = inputReader.ReadLine()
                                _row = inputReader.ReadLine()
                                _row = inputReader.ReadLine()
                                _row = inputReader.ReadLine()
                                _row = inputReader.ReadLine()

                                'Alternate report has extra lines so check for this
                                If Not _row.Trim.Length > 1 Then
                                    Do Until _row.Trim.Length > 1
                                        _row = inputReader.ReadLine()
                                    Loop
                                End If

                            End If
                            meetingType = Trim(Mid(_row, 87, 11))
                            'Look through the meeting types that should be skipped
                            skipMeeting = DetermineMeeting(meetingType, meetingTypeArray)
                            If skipMeeting <> True Then
                                provider = Trim(Mid(_row, 26, 4)).ToUpper()
                                If useCSV = True Then
                                    '*****************************************************************************************************
                                    '   Find the corresponding IVR provider id based on Eclipse OPEN provider 
                                    '   id and meeting type by parsing the xls file
                                    '*****************************************************************************************************
                                    Try
                                        xlsFile = CSVFile
                                        xlsReader = New StreamReader(xlsFile)
                                        xlsLine = xlsReader.ReadLine
                                        xlsLine = xlsReader.ReadLine
                                        splitOut = Split(xlsLine, ",")
                                        lookupEclipseOpenProviderID = Trim(splitOut(3)).ToUpper()
                                        lookupMeetingType = Trim(splitOut(2)).ToUpper
                                    Catch ex As Exception
                                        _error = xlsProblem
                                        xlsReader.Close()
                                        xlsReader = Nothing
                                        Return ProcessingStatus.ErroredRow
                                    End Try
                                    Try
                                        Do While Not xlsLine Is Nothing
                                            If xlsLine.Length > 0 Then
                                                splitOut = Split(xlsLine, ",")
                                                lookupEclipseOpenProviderID = Trim(splitOut(3)).ToUpper()
                                                lookupMeetingType = Trim(splitOut(2)).ToUpper
                                                If provider = lookupEclipseOpenProviderID And meetingType.ToUpper = lookupMeetingType.ToUpper Then
                                                    providerID = Trim(splitOut(1))
                                                    msgID = Trim(splitOut(4))
                                                    Exit Do
                                                End If
                                            End If
                                            xlsLine = xlsReader.ReadLine
                                        Loop
                                    Catch e As Exception
                                        _error = problemReadingXLS
                                        If Not xlsReader Is Nothing Then
                                            xlsReader.Close()
                                            xlsReader = Nothing
                                        End If
                                        Return ProcessingStatus.ErroredRow
                                    End Try
                                    xlsReader.Close()
                                    xlsReader = New StreamReader(xlsFile)
                                    xlsLine = xlsReader.ReadLine
                                    xlsLine = xlsReader.ReadLine
                                    If providerID = Nothing Then
                                        Try
                                            'loop through the excel spreadsheet again to see if an All-Else exists
                                            Do While Not xlsLine Is Nothing
                                                If xlsLine.Length > 0 Then
                                                    splitOut = Split(xlsLine, ",")
                                                    lookupEclipseOpenProviderID = Trim(splitOut(3)).ToUpper()
                                                    lookupMeetingType = Trim(splitOut(2)).ToUpper
                                                    If provider = lookupEclipseOpenProviderID And lookupMeetingType.ToUpper = "ALL-ELSE" Then
                                                        providerID = Trim(splitOut(1))
                                                        msgID = Trim(splitOut(4))
                                                        Exit Do
                                                    End If
                                                End If
                                                xlsLine = xlsReader.ReadLine
                                            Loop
                                        Catch e As Exception
                                            _error = problemReadingXLS
                                            If Not xlsReader Is Nothing Then
                                                xlsReader.Close()
                                                xlsReader = Nothing
                                            End If
                                            Return ProcessingStatus.ErroredRow
                                        End Try
                                        If providerID = Nothing Then
                                            _error = providerIDNotFound
                                            If Not xlsReader Is Nothing Then
                                                xlsReader.Close()
                                                xlsReader = Nothing
                                            End If
                                            Return ProcessingStatus.ErroredRow
                                        End If
                                    End If
                                    If Not xlsReader Is Nothing Then
                                        xlsReader.Close()
                                        xlsReader = Nothing
                                    End If
                                Else
                                    GetProviderID_FromConfig(providerID, provider)
                                End If
                                apptString = Trim(Mid(_row, 65, 10))
                                apptTime = Trim(Mid(_row, 113, 10))
                                If apptTime = "12:00n" Then
                                    apptTime = "12:00pm"
                                End If
                                apptDate = Convert.ToDateTime(apptString & " " & apptTime)
                                If scheduledDaysPrior <> "" Then
                                    createString = Trim(Mid(_row, 77, 10))
                                    createDate = Convert.ToDateTime(createString)
                                    If createDate.AddDays(Trim(scheduledDaysPrior)).Date >= apptDateCompare Then
                                        skipReason = skipReasonScheduledDaysPrior
                                    Else
                                        If useCSV = False Then
                                            WriteNoneButRecord(strPhone, strNewPhone, providerID, apptDate, skipMeeting, msgID, name, outputWriter)
                                        Else
                                            WriteNoneButRecord(strPhone, strNewPhone, providerID, apptDate, skipMeeting, msgID, name, outputWriter)
                                        End If
                                    End If
                                Else
                                    WriteNoneButRecord(strPhone, strNewPhone, providerID, apptDate, skipMeeting, msgID, name, outputWriter)
                                End If
                            End If
                        End If
                    End If
                End If
                Return ProcessingStatus.SuccessfulRow
            End If
        End Function
        Private Function FormatName(ByVal name As String) As String
            Dim space As Integer
            space = name.IndexOf(" ")
            If name.Length > 0 Then
                If space > 0 Then
                    name = Trim(Left(name, space))
                Else
                    name = Trim(Left(name, 9))
                End If
                Return OnlyAlphaNumericChars(name)
            End If
        End Function

        Public Function OnlyAlphaNumericChars(ByVal OrigString As String) As String
            '***********************************************************
            'INPUT:  Any String
            'OUTPUT: The Input String with all non-alphanumeric characters 
            '        removed
            '***********************************************************
            Dim lLen As Long
            Dim sAns As String
            Dim lCtr As Long
            Dim sChar As String

            OrigString = Trim(OrigString)
            lLen = Len(OrigString)
            For lCtr = 1 To lLen
                sChar = Mid(OrigString, lCtr, 1)
                If IsAlphaNumeric(Mid(OrigString, lCtr, 1)) Then
                    sAns = sAns & sChar
                End If
            Next
            Return sAns
        End Function

        Private Function IsAlphaNumeric(ByVal sChr As String) As Boolean
            IsAlphaNumeric = sChr Like "[0-9A-Za-z]"
        End Function
        '=================================================================================================
        'Method Name:	Trans.FormatPhone
        'Description:	Takes all spaces out the phone and adds area code if relevant
        'Property of Archipelago Systems, LLC.
        '=================================================================================================
        Private Function FormatPhone(ByVal strPhone As String, ByRef name As String, ByRef exceptionWriter As StreamWriter) As String
            Dim x As Integer
            Dim cur As String
            Dim strNewPhone As String = ""
            If strPhone.Length = 0 Then
                strNewPhone = ""
            End If
            Do Until x = strPhone.Length
                cur = strPhone.Chars(x)
                Select Case cur
                    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                        'If this is the first time in the loop, take the zero out
                        If x = 0 Then
                            strNewPhone = ""
                        End If
                        'The character is a number so add it to the new string
                        strNewPhone += cur
                End Select
                x += 1
            Loop

            If strNewPhone.Length <> 10 And strNewPhone.Length > 6 Then
                strNewPhone = Mid(strNewPhone, 1, 10)
            End If
            If strNewPhone.Length <> 10 Then
                strNewPhone = ConfigurationSettings.AppSettings("DefaultAreaCode").ToString & Mid(strNewPhone, 1, 7)
            End If
            If strNewPhone.Length = 10 And name <> "written" Then rowWritten = True Else exceptionWriter.WriteLine("Invalid phone number: " & name & " " & strNewPhone)

            Return strNewPhone
        End Function
        '=================================================================================================
        'Method Name:	Trans.DetermineMeeting
        'Description:	Returns boolean value of whether the record's meeting type should be skipped
        'Property of Archipelago Systems, LLC.
        '=================================================================================================
        Private Function DetermineMeeting(ByVal meetingType As String, ByVal meetingTypeArray() As String)
            'Get the meeting type: If the type is listed in the config file as a type to skip, don't print the line
            Dim y = 0
            Dim skipMeeting As Boolean
            Do Until y = (CType(ConfigurationSettings.AppSettings("MeetingSkipTypeTotal"), Integer) + 1)
                If Trim(meetingTypeArray(y)) <> "" Then
                    If meetingTypeArray(y).ToUpper = meetingType.ToUpper Then
                        skipMeeting = True
                        y = (CType(ConfigurationSettings.AppSettings("MeetingSkipTypeTotal"), Integer))
                    End If
                End If
                y += 1
            Loop
            Return skipMeeting
        End Function
        '=================================================================================================
        'Method Name:	Trans.ProcessError
        'Description:	Exposes Error Message to consuming class
        'Property of Archipelago Systems, LLC.
        '=================================================================================================
        Public ReadOnly Property ProcessError() As String
            Get
                Return _error
            End Get
        End Property
        '=================================================================================================
        'Method Name:	Trans.ToString
        'Description:	Overrides default ToString Functionality
        'Property of Archipelago Systems, LLC.
        '=================================================================================================
        Public Overrides Function ToString() As String
            Try
                Return _row
            Catch
                Return Nothing
            End Try
        End Function
        '=================================================================================================
        'Method Name:	Trans.CheckForNullCharacters
        'Description:	Checks to see if a string contains null characters
        'Property of Archipelago Systems, LLC.
        '=================================================================================================
        Private Function CheckForNullCharacters(ByVal s As String) As Boolean

            Dim reader As StringReader
            Dim n As Integer

            Try
                reader = New StringReader(s)

                'Loop through each character looking for a null character
                Do
                    'Get the character code for the reader
                    n = reader.Read

                    'If its null, return true, exit do
                    If n = 0 Then
                        Return True
                        Exit Do

                        '-1 means the end of the string had been reached
                    ElseIf n = -1 Then
                        Return False
                        Exit Do
                    End If
                Loop

            Finally
                'Clean up reader
                reader = Nothing
            End Try
        End Function
        '=================================================================================================
        'Method Name:	DataExport.WriteHeader
        'Description:	Writes the four first lines needed in the Call List file
        'Property of Archipelago Systems, LLC.
        '=================================================================================================
        Private Sub WriteHeader(ByRef outputWriter As StreamWriter, ByVal custID As String, _
                                ByVal callHour As String, ByVal callMinute As String, ByVal callDate As Date)

            'Write the customer id to the output file
            outputWriter.WriteLine("CUSTOMER_ID:" & custID)

            'Write the calling time to the output file
            outputWriter.WriteLine("CALLING_TIME:" & callHour & ":" & callMinute & ":00")

            'Write the calling date to the output file
            outputWriter.WriteLine("CALLING_DATE:" & callDate.ToShortDateString)

            'Write the calling engine to the output file
            outputWriter.WriteLine("ENGINE:" & ConfigurationSettings.AppSettings("Engine").ToString())
        End Sub
        Private Sub GetProviderID_FromConfig(ByRef providerID As String, ByVal provider As String)
            Dim z As Integer = 1
            Dim done As Boolean = False
            Do Until done = True Or z = CType(ConfigurationSettings.AppSettings("EngineProviderTotal"), Integer)
                providerID = "EngineProvider" & z
                If ConfigurationSettings.AppSettings(providerID).ToString().ToUpper = provider.ToUpper Then
                    providerID = z
                    done = True
                Else
                    z += 1
                End If
            Loop
            done = False
        End Sub
        '=================================================================================================
        'Method Name:	DataExport.WriteAllButRecord
        'Description:	Writes a record in the 'AllBut' logic
        'Property of Archipelago Systems, LLC.
        '=================================================================================================
        Private Sub WriteAllButRecord(ByVal strPhone As String, ByVal strNewPhone As String, ByVal providerID As String, _
                        ByVal apptDate As Date, ByVal skipMeeting As Boolean, ByVal msgID As String, ByVal name As String, ByRef outputWriter As StreamWriter)
            Dim spanish As Boolean
            If strPhone.ToUpper.IndexOf("PRIV") = -1 Or strPhone.ToUpper.IndexOf("PRIV") = 0 Then
                spanish = strPhone.ToUpper.IndexOf("SP") > 0
                If Not providerID Is Nothing Then
                    If providerID.Length <> 0 Then
                        If Left(providerID, 14).ToUpper <> "ENGINEPROVIDER" Then
                            If strNewPhone.Length = 10 And providerID <> Nothing And apptDate <> Nothing Then
                                If skipMeeting = False Then
                                    outputWriter.Write(strNewPhone)
                                    outputWriter.Write(", " & providerID)
                                    outputWriter.Write(", " & apptDate & ", ")
                                    If Not msgID Is Nothing Then
                                        If msgID.Length <> 0 Then
                                            outputWriter.Write(msgID)
                                        End If
                                    End If
                                    outputWriter.Write(", " & name)
                                    If spanish Then outputWriter.Write(", " & "SP")
                                    outputWriter.WriteLine("")
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End Sub
        '=================================================================================================
        'Method Name:	DataExport.WriteNoneButRecord
        'Description:	Writes a record in the 'NoneBut' logic
        'Property of Archipelago Systems, LLC.
        '=================================================================================================
        Private Sub WriteNoneButRecord(ByVal strPhone As String, ByVal strNewPhone As String, ByVal providerID As String, _
                        ByVal apptDate As Date, ByVal skipMeeting As Boolean, ByVal msgID As String, ByVal name As String, ByRef outputWriter As StreamWriter)
            Dim spanish As Boolean
            If Right(strPhone, 2).ToUpper = "OK" Then
                spanish = strPhone.ToUpper.IndexOf("SP") > 0
                If Not providerID Is Nothing Then
                    If providerID.Length <> 0 Then
                        If Left(providerID, 14).ToUpper <> "ENGINEPROVIDER" Then
                            If strNewPhone.Length = 10 And providerID <> Nothing And apptDate <> Nothing Then
                                If skipMeeting = False Then
                                    outputWriter.Write(strNewPhone)
                                    outputWriter.Write(", " & providerID)
                                    outputWriter.Write(", " & apptDate & ", ")
                                    If Not msgID Is Nothing Then
                                        If msgID.Length <> 0 Then
                                            outputWriter.Write(msgID)
                                        End If
                                    End If
                                    outputWriter.Write(", " & name)
                                    If spanish Then outputWriter.Write(", " & "SP")
                                    outputWriter.WriteLine("")
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End Sub

    End Class
    '=================================================================================================
    'Method Name:	DataExport.CheckForMissingDirectory
    'Description:	Verifies that the directory exists and creates it if necessary
    'Property of Archipelago Systems, LLC.
    '=================================================================================================
    Private Sub CheckForMissingDirectory(ByVal filePath As String)
        Dim directoryEndPosition As Integer

        directoryEndPosition = filePath.LastIndexOf("\")

        'If the necessary directory is not found, it is created
        If Not Directory.Exists(filePath.Substring(0, directoryEndPosition + 1)) Then
            Directory.CreateDirectory(filePath.Substring(0, directoryEndPosition + 1))
        End If
    End Sub
End Class