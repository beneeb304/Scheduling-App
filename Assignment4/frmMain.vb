Imports System.Data.SqlClient
Imports System.IO

Public Class frmMain
    'Name of database
    Private Const strDBNAME As String = "Scheduling"

    'Name of the database server
    Private Const strSERVERNAME As String = "(localdb)\MSSQLLocalDB"

    'Number of days in a week
    Private Const intDAYSINWEEK As Integer = 7

    'Time increment for appointments
    Private Const intTIMEINCREMENT As Integer = 15

    'Path to database in executable
    Private ReadOnly strDBPATH As String = My.Application.Info.DirectoryPath & "\" & strDBNAME & ".mdf"

    'This is the full connection string
    Private ReadOnly strCONNECTION As String = "SERVER=" & strSERVERNAME & ";DATABASE=" &
                     strDBNAME & ";Integrated Security=SSPI;AttachDbFileName=" & strDBPATH

    Private DBConn As SqlConnection

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'If the database doesn't exist, create it
        If Not IO.File.Exists(strDBPATH) Then
            'Create database, build tables, and insert initial data
            CreateDatabase()
        End If
    End Sub

    Private Sub btnSelectFile_Click(sender As Object, e As EventArgs) Handles btnSelectFile.Click
        'Create an openfiledialog
        Dim ofd As OpenFileDialog = New OpenFileDialog

        'Filter to only allow txt files
        ofd.Filter = "txt files (*.txt)|*.txt"

        'Show to user
        ofd.ShowDialog()

        'Set name of txt file to textbox
        txtFileName.Text = ofd.FileName
    End Sub

    Private Sub btnEnterFile_Click(sender As Object, e As EventArgs) Handles btnEnterFile.Click
        'Read in file
        Try
            'String to hold text file
            Dim strText As String = ""

            'Read each line of text in the text file
            For Each line As String In File.ReadAllLines(txtFileName.Text)
                'Read line of text and append a CR
                strText &= (line & vbCr)
            Next

            'Send string off to be parsed (add CR as splitter)
            ParseTextFile(strText.Split(vbCr))
        Catch ex As Exception
            MessageBox.Show("Please read in a valid text file!")
        End Try

        'Empty controls so user is forced to realize new data (if applicable)
        txtFileName.Text = ""
        cmbFilter.SelectedItem = Nothing
        lstSelect.Items.Clear()
        lstDisplay.Items.Clear()
    End Sub

    Private Sub btnQuit_Click(sender As Object, e As EventArgs) Handles btnQuit.Click
        'Quit nicely
        Close()
    End Sub

    Private Sub frmMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        'Ask the user if they want to delete the current db
        Dim result As DialogResult = MessageBox.Show("Would you like to blast away the current database?", "Exiting", MessageBoxButtons.YesNoCancel)

        'If yes, delete it. Otherwise, keep exiting
        If result = DialogResult.Yes Then
            DeleteDatabase()
        ElseIf result = DialogResult.Cancel Then
            e.Cancel = True
        End If
    End Sub

    Private Sub ParseTextFile(strLines() As String)
        'Create String array for pieces of data in each line of text
        Dim strLine() As String
        Dim strPatients As String = ""
        Dim strAppointments As String = ""
        Dim lstPhone As New List(Of String)
        Dim intConsRows As Integer

        'Cycle through each line
        For Each line As String In strLines
            'Split data by vbTab
            strLine = line.Split(vbTab)

            'Patient
            If strLine(0) = "P" Then
                'If first time, add beginning of insert statement
                If strPatients.Length = 0 Then
                    strPatients = "INSERT INTO Patients VALUES "
                End If

                'Append to insert statement
                strPatients &= "('" & strLine(1).Substring(0, strLine(1).IndexOf(" ")) & "', " &
                    "'" & strLine(1).Substring(strLine(1).IndexOf(" ") + 1) & "', " &
                    "'" & strLine(2) & "', " &
                    "'" & strLine(3) & "')," & vbCrLf
            End If
        Next

        'Chop off last comma
        strPatients = strPatients.Substring(0, strPatients.LastIndexOf(","))

        'Connect to SQL Server instance
        ConnectToSQL(False)

        'Execute SQL command
        ExecuteSQLNonQ(strPatients)

        For Each line As String In strLines
            'Split data by vbTab
            strLine = line.Split(vbTab)

            'Appointment
            If strLine(0) = "A" Then

                'FIND ME figure this out to only run once...
                'If first time, add phone numbers to listbox
                'If strAppointments.Length = 0 Then
                '    Dim dataReader = ExecuteSQLReader("SELECT Phone FROM Patients")

                '    'While there's data, read it
                '    While dataReader.Read()
                '        'Add phone numbers to list
                '        lstPhone.Add(dataReader("Phone"))
                '    End While

                '    dataReader.Close()
                'End If

                ''Check if patient doesn't exist yet
                'If lstPhone.Contains(strLine(1)) = False Then
                '    'If not, add them
                '    AddPatient(strLine(1))
                'End If

                'Operation
                Select Case strLine(2)
                    'Adding appointment
                    Case "A"
                        'Get selector
                        Dim chrSelector As Char = strLine(3).Substring(0, 1)

                        'Get patient tuid
                        Dim dataReader = ExecuteSQLReader("SELECT TUID FROM Patients WHERE Phone = '" & strLine(1) & "'")
                        Dim intPatientTUID As Integer
                        While dataReader.Read()
                            'Get patient TUID
                            intPatientTUID = Integer.Parse(dataReader("TUID").ToString())
                        End While
                        dataReader.Close()

                        'Determine preference
                        Select Case chrSelector
                            'Next available preference
                            Case "N"
                                'Get number of consecutive rows we need for this appointment
                                intConsRows = Integer.Parse(strLine(4)) / 15

                                'Get all appointments that are available
                                dataReader = ExecuteSQLReader("SELECT StartTime, EndTime, Day FROM Appointments WHERE PatientTUID IS NULL")

                                Dim thisTime As Date = Nothing
                                Dim thisDate As Date = Nothing

                                Dim intCtr As Integer = 1

                                Dim blnFound As Boolean = False

                                'While there's data, read it
                                While dataReader.Read()

                                    Debug.WriteLine(thisTime.ToString())
                                    Debug.WriteLine(CDate(dataReader("StartTime").ToString()))

                                    'Initialize thisDate
                                    If thisTime = Nothing Then
                                        thisTime = CDate(dataReader("StartTime").ToString())
                                    Else
                                        'Check if this loop and last loop are consecutive (15 minutes apart)
                                        If DateDiff(DateInterval.Minute, thisTime, CDate(dataReader("StartTime").ToString())) = 15 Then 'Long.Parse(15) 
                                            'Increment counter
                                            intCtr += 1

                                            'Replace last time with this time
                                            thisTime = CDate(dataReader("StartTime").ToString())

                                            'Check if we have enough slots
                                            If intCtr = intConsRows Then
                                                'Set appointment end time
                                                thisTime = CDate(dataReader("EndTime").ToString())

                                                'Set date
                                                thisDate = CDate(dataReader("Day").ToString())

                                                'Set found flag
                                                blnFound = True

                                                'Exit loop
                                                Exit While
                                            End If
                                        Else
                                            'Reset counter
                                            intConsRows = 1

                                            'Reset time
                                            thisTime = Nothing
                                        End If
                                    End If
                                End While

                                dataReader.Close()

                                If blnFound Then
                                    'Alter statement
                                    Dim strSQL As String
                                    strSQL = "SET ROWCOUNT 1 " &
                                        "UPDATE Appointments " &
                                        "SET PatientTUID = " & intPatientTUID & ", EndTime = '" & thisTime & "', AppointmentLength = " & strLine(4) &
                                        " WHERE StartTime = '" & thisTime.AddMinutes(Integer.Parse(strLine(4) / -1)) & "' AND Day = '" & thisDate & "'"

                                    Debug.WriteLine(strSQL)

                                    ExecuteSQLNonQ(strSQL)

                                    'Delete following rows


                                Else
                                    MessageBox.Show("We could not find an appointment for " & strLine(1))
                                End If

                            'Doctor preference
                            Case "D"

                                'Dim s As String = ""
                                'For Each thign In strLine
                                '    s &= thign & " "
                                'Next
                                'MessageBox.Show(s)

                            'Time/day preference
                            Case "T"

                                'Dim s As String = ""
                                'For Each thign In strLine
                                '    s &= thign & " "
                                'Next
                                'MessageBox.Show(s)

                        End Select

                    'Deleting appointment
                    Case "D"

                    'Changing appointment
                    Case "C"

                End Select
            End If
        Next

        'Disconnect from the database
        DisconnectFromSQL()
    End Sub

    Private Sub AddPatient(strPhone As String)
        Dim strData(2) As String

        'Get patient's first name
        strData(0) = InputBox("Please enter the patient's first name.", "New patient! Please fill out their information.")

        'Get patient's last name
        strData(1) = InputBox("Please enter the patient's last name.", "New patient! Please fill out their information.")

        'Get patient's insurance provider
        strData(2) = InputBox("Please enter the patient's insurance provider.", "New patient! Please fill out their information.")

        'Make SQL command
        Dim strSQLCommand As String = "INSERT INTO Patients VALUES " &
            "('" & strData(0) & "', " &
            "'" & strData(1) & "', " &
            "'" & strPhone & "', " &
            "'" & strData(2) & "')"

        'Execute SQL command
        ExecuteSQLNonQ(strSQLCommand)
    End Sub

    Private Sub ExecuteSQLNonQ(strSQL As String)
        'Build a SQL Server database from scratch
        Dim DBCmd As New SqlCommand With {
            .CommandText = strSQL,
            .Connection = DBConn
        }

        Try
            DBCmd.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show(ex.ToString(), "Error executing SQL statement.")
        End Try
    End Sub

    Private Function ExecuteSQLReader(strSQL As String) As SqlDataReader
        'Build a SQL Server database from scratch
        Dim DBCmd As New SqlCommand()
        Dim dataReader As SqlDataReader = Nothing

        DBCmd.CommandText = strSQL
        DBCmd.Connection = DBConn

        Try
            dataReader = DBCmd.ExecuteReader()
        Catch ex As Exception
            MessageBox.Show(ex.ToString(), "Error executing SQL statement.")
        End Try

        Return dataReader
    End Function

    Private Sub ConnectToSQL(blnInitial As Boolean)
        If blnInitial Then
            'Initial connection
            DBConn = New SqlConnection("Server=" & strSERVERNAME)
        Else
            'Use the full connection string with the Integrated Security line
            DBConn = New SqlConnection(strCONNECTION)
        End If

        'Try to open the SQL Server connection
        Try
            DBConn.Open()
        Catch ex As Exception
            MessageBox.Show(ex.ToString(), "Failed to connect to SQL Server... Closing down!")
            Me.Close()
        End Try
    End Sub

    Private Sub DisconnectFromSQL()
        If (DBConn.State = ConnectionState.Open) Then
            DBConn.Close()
        End If
    End Sub

    Private Sub CreateDatabase()

        'String to hold SQL command text
        Dim strSQLCommand As String

        'Build the database command
        strSQLCommand = "CREATE DATABASE " & strDBNAME & " On " &
            "(NAME = '" & strDBNAME & "', " &
            "FILENAME = '" & strDBPATH & "')"

        'Connect to SQL Server instance
        ConnectToSQL(True)

        'Execute SQL command
        ExecuteSQLNonQ(strSQLCommand)

        'Close the connection and reopen it pointing at the scheduling database
        DisconnectFromSQL()

        'Connect to SQL with security
        ConnectToSQL(False)

        'Build database tables
        BuildTables()

        'Insert the initial datasets
        InsertInitialData()

        'We can check to see if we're open before trying to issue a connection close
        DisconnectFromSQL()
    End Sub

    Private Sub BuildTables()
        'String to hold SQL command text
        Dim strSQLCommand As String

        'Build the Doctors Table
        strSQLCommand = "CREATE TABLE Doctors
            (TUID int IDENTITY(1,1) PRIMARY KEY,
            FirstName varchar(50) NOT NULL,
            LastName varchar(50) NOT NULL)"

        'Execute SQL command
        ExecuteSQLNonQ(strSQLCommand)

        'Build the Availability Table
        strSQLCommand = "CREATE TABLE Availability
            (TUID int IDENTITY(1,1) PRIMARY KEY,
            DoctorTUID int FOREIGN KEY REFERENCES Doctors(TUID) NOT NULL,
            Day char(1) NOT NULL,
            StartTime Time(0) NOT NULL,
            EndTime Time(0) NOT NULL,
            PatientCount int NOT NULL)"

        'Execute SQL command
        ExecuteSQLNonQ(strSQLCommand)

        'Build the Patients Table
        strSQLCommand = "CREATE TABLE Patients
             (TUID int IDENTITY(1,1) PRIMARY KEY,
             FirstName varchar(50) NOT NULL,
             LastName varchar(50) NOT NULL,
             Phone varchar(8) NOT NULL UNIQUE,
             Insurance varchar(20) NOT NULL)"

        'Execute SQL command
        ExecuteSQLNonQ(strSQLCommand)

        'Build the Appointments Table
        strSQLCommand = "CREATE TABLE Appointments
            (TUID int IDENTITY(1,1) PRIMARY KEY, 
            PatientTUID int FOREIGN KEY REFERENCES Patients(TUID), 
            DoctorTUID int FOREIGN KEY REFERENCES Doctors(TUID) NOT NULL,
            Day Date NOT NULL,
            AppointmentLength int NOT NULL,
            StartTime TIME(0), 
            EndTime TIME(0))"

        'Execute SQL command
        ExecuteSQLNonQ(strSQLCommand)
    End Sub

    Private Sub InsertInitialData()

        'String to hold SQL command text
        Dim strSQLCommand As String
        Dim strDayofWeek As String

        strSQLCommand = "INSERT INTO Doctors VALUES 
            ('Ray', 'Strantz'),
            ('Henry', 'Jones, Jr.'),
            ('Emmett', 'Brown')"

        'Execute SQL command
        ExecuteSQLNonQ(strSQLCommand)

        strSQLCommand = "INSERT INTO Availability VALUES 
            (1, 'M', '10:00', '14:00', 7),
            (1, 'W', '10:00', '14:00', 7),
            (1, 'F', '10:00', '14:00', 7),
            (2, 'M', '8:00', '13:00', 8),
            (2, 'W', '8:00', '13:00', 8),
            (2, 'R', '8:00', '13:00', 8),
            (3, 'M', '11:00', '16:00', 9),
            (3, 'T', '11:00', '16:00', 9),
            (3, 'R', '11:00', '16:00', 9),
            (3, 'F', '11:00', '16:00', 9)"

        'Execute SQL command
        ExecuteSQLNonQ(strSQLCommand)

        'Create today date object
        Dim thisDate As Date = Date.Today

        'Cycle through each day of the week starting with tomorrow
        For i = 1 To intDAYSINWEEK
            'Get day of the week
            strDayofWeek = thisDate.AddDays(i).DayOfWeek.ToString()

            Select Case strDayofWeek
                Case "Monday"
                    'Select all doctors who have availability for M
                    PopulateAppointments(CStr(thisDate.AddDays(i)), "M")
                Case "Tuesday"
                    'Select all doctors who have availability for T
                    PopulateAppointments(CStr(thisDate.AddDays(i)), "T")
                Case "Wednesday"
                    'Select all doctors who have availability for W
                    PopulateAppointments(CStr(thisDate.AddDays(i)), "W")
                Case "Thursday"
                    'Select all doctors who have availability for R
                    PopulateAppointments(CStr(thisDate.AddDays(i)), "R")
                Case "Friday"
                    'Select all doctors who have availability for W
                    PopulateAppointments(CStr(thisDate.AddDays(i)), "F")
            End Select
        Next
    End Sub

    Private Sub PopulateAppointments(strDate As String, chrDay As Char)


        'For loop through each

        'Create variable for start time

        'While we are before end time
        'add available to appointment table
        'increment time var by 15 minutes



        'Insert initial available appointments for 1 week out
        Dim dataReader As SqlDataReader

        'Create time variables
        Dim startTime, endTime As Date

        'Doctor TUID
        Dim intDoc As Integer

        'Create SQL Command
        Dim strSQLCommand As String = "SELECT DoctorTUID, StartTime, EndTime FROM Availability WHERE Day = '" & chrDay & "'"

        'Execute SQL command
        dataReader = ExecuteSQLReader(strSQLCommand)

        'Empty SQL command
        strSQLCommand = ""

        If dataReader.HasRows() Then

            'While there's data, read it
            While dataReader.Read()
                'Capture data returned from query
                intDoc = dataReader("DoctorTUID")

                startTime = CDate(dataReader("StartTime").ToString())
                endTime = CDate(dataReader("EndTime").ToString())

                'Loop while the start time is before the end time
                While startTime < endTime

                    'If first time adding to SQL command, create it
                    If strSQLCommand.Length = 0 Then
                        strSQLCommand = "INSERT INTO Appointments VALUES" & vbCrLf
                    End If

                    'Append to SQL command
                    strSQLCommand &= "(NULL, " & intDoc & ", '" & strDate & "', " &
                        intTIMEINCREMENT & ", '" & startTime.ToString("HH:mm") &
                        "', '" & startTime.AddMinutes(intTIMEINCREMENT).ToString("HH:mm") & "')," & vbCrLf

                    'Increment time by 15 minutes
                    startTime = startTime.AddMinutes(intTIMEINCREMENT)
                End While
            End While

            'Close datareader
            dataReader.Close()

            'Chop off last comma
            strSQLCommand = strSQLCommand.Substring(0, strSQLCommand.LastIndexOf(","))

            ExecuteSQLNonQ(strSQLCommand)
        End If
    End Sub

    Private Sub DeleteDatabase()
        'This routine deletes a database completely from code. Credit to CIS 311 ch 15 notes.

        'String to hold SQL command text
        Dim strSQLCommand As String

        'Connect to SQL Server instance
        ConnectToSQL(True)

        'Try to force single ownership of the database so that we have the permissions to delete it
        strSQLCommand = "ALTER DATABASE [" & strDBNAME & "] SET " &
            "SINGLE_USER WITH ROLLBACK IMMEDIATE"

        'Execute SQL command
        ExecuteSQLNonQ(strSQLCommand)

        'Now, drop the database
        strSQLCommand = "DROP DATABASE " & strDBNAME

        'Execute SQL command
        ExecuteSQLNonQ(strSQLCommand)

        DisconnectFromSQL()
    End Sub

    Private Sub txtFileName_TextChanged(sender As Object, e As EventArgs) Handles txtFileName.TextChanged
        If txtFileName.Text.Length > 0 Then
            btnEnterFile.Enabled = True
        Else
            btnEnterFile.Enabled = False
        End If
    End Sub

    Private Sub cmbFilter_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbFilter.SelectedIndexChanged
        'String to hold SQL command text
        Dim strSQLCommand As String

        'Clear listboxes
        lstSelect.Items.Clear()
        lstDisplay.Items.Clear()

        ConnectToSQL(False)

        Dim dataReader As SqlDataReader

        If cmbFilter.SelectedItem = "Patients" Then
            strSQLCommand = "SELECT Phone, FirstName, LastName FROM Patients"

            dataReader = ExecuteSQLReader(strSQLCommand)

            'While there's data, read it
            While dataReader.Read()
                'Add to listbox
                lstSelect.Items.Add(dataReader("Phone") & " - " & dataReader("FirstName") & " " & dataReader("LastName"))
            End While

            'Tidy up
            dataReader.Close()
        ElseIf cmbFilter.SelectedItem = "Doctor Appointments" Or cmbFilter.SelectedItem = "Doctor Availability" Then
            strSQLCommand = "SELECT TUID, FirstName, LastName FROM Doctors"

            dataReader = ExecuteSQLReader(strSQLCommand)

            'While there's data, read it
            While dataReader.Read()
                'Add to listbox
                lstSelect.Items.Add(dataReader("TUID") & " - Dr. " & dataReader("FirstName") & " " & dataReader("LastName"))
            End While

            'Tidy up
            dataReader.Close()
        End If

        'Tidy up
        DisconnectFromSQL()
    End Sub

    Private Sub btnAddPatient_Click(sender As Object, e As EventArgs) Handles btnAddPatient.Click
        'Get patient's insurance provider
        Dim strPhone As String = InputBox("Please enter the patient's phone number.", "New patient! Please fill out their information.")

        AddPatient(strPhone)
    End Sub

    Private Sub lstSelect_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstSelect.SelectedIndexChanged
        'Clear listboxes
        lstDisplay.Items.Clear()

        ConnectToSQL(False)

        'Determine ff selecting Patients, Doctor Appointments, or Doctor Availability
        If cmbFilter.SelectedItem = "Patients" Then

        ElseIf cmbFilter.SelectedItem = "Doctor Appointments" Then

        ElseIf cmbFilter.SelectedItem = "Doctor Availability" Then
            Dim intTUID As Integer = lstSelect.SelectedItem.ToString().Substring(0, 1)
            Dim dataReader = ExecuteSQLReader("SELECT Day, StartTime, EndTime FROM Appointments WHERE DoctorTUID = " & intTUID & "AND PatientTUID IS NULL")
            Dim thisDate As Date

            'While there's data, read it
            While dataReader.Read()
                'Get day of week
                thisDate = CDate(dataReader("Day").ToString())

                'Add to listbox
                lstDisplay.Items.Add(thisDate.DayOfWeek.ToString() & ", " & CStr(dataReader("Day")) & " from " & dataReader("StartTime").ToString() & " to " & dataReader("EndTime").ToString())
            End While

            dataReader.Close()
        End If
    End Sub
End Class