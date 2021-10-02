Imports System.Data.SqlClient
Imports System.IO

Public Class frmMain
    'Name of database
    Const strDBNAME As String = "Scheduling"

    'Name of the database server
    Const strSERVERNAME As String = "(localdb)\MSSQLLocalDB"

    'Path to database in executable
    Dim strDBPATH As String = My.Application.Info.DirectoryPath & "\" & strDBNAME & ".mdf"

    'This is the full connection string
    Dim strCONNECTION As String = "SERVER=" & strSERVERNAME & ";DATABASE=" &
                     strDBNAME & ";Integrated Security=SSPI;AttachDbFileName=" & strDBPATH

    Dim DBConn As SqlConnection

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'If the database doesn't exist, create it
        If Not IO.File.Exists(strDBPATH) Then
            'Create database, build tables, and insert initial data
            CreateDatabase()
        End If
    End Sub

    Private Sub btnSelectFile_Click(sender As Object, e As EventArgs) Handles btnSelectFile.Click
        'Get rid of error provider
        ErrorProvider.Clear()

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
        'Get rid of error provider
        ErrorProvider.Clear()

        'Make sure the user has selected a text file before continuing
        If txtFileName.Text = Nothing Then
            ErrorProvider.SetError(btnEnterFile, "Please select a txt file before continuing!")
        Else
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
        End If

        'Empty textbox
        txtFileName.Text = ""
    End Sub

    Private Sub btnQuit_Click(sender As Object, e As EventArgs) Handles btnQuit.Click
        'Get rid of error provider
        ErrorProvider.Clear()

        'Quit nicely
        Me.Close()
    End Sub

    Private Sub frmMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        'Ask the user if they want to delete the current db
        Dim result As DialogResult = MessageBox.Show("Would you like to blast away the current database?", "Exiting", MessageBoxButtons.YesNoCancel)

        'If yes, delete it. Otherwise, keep exiting
        If result = DialogResult.Yes Then
            DeleteDatabase(strSERVERNAME, strDBNAME)
        ElseIf result = DialogResult.Cancel Then
            e.Cancel = True
        End If
    End Sub

    Private Sub ParseTextFile(strLines() As String)
        'Create String array for pieces of data in each line of text
        Dim strLine() As String
        Dim strAppointments As String = ""
        Dim strPatients As String = ""

        'Cycle through each line
        For Each line As String In strLines
            'Split data by vbTab
            strLine = line.Split(vbTab)

            'Parse with conditional statements
            If strLine(0) = "A" Then
                'Appointment entry

            ElseIf strLine(0) = "P" Then
                'Patient entry

                'If first time, add beginning of insert statement
                If strPatients.Length = 0 Then
                    strPatients = "INSERT INTO dbo.Patients VALUES "
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
        ExecuteSQL(strPatients)
    End Sub

    Private Sub ExecuteSQL(strSQL As String)
        'Build a SQL Server database from scratch
        Dim DBCmd As SqlCommand = New SqlCommand()

        DBCmd.CommandText = strSQL
        DBCmd.Connection = DBConn

        Try
            DBCmd.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show(ex.ToString(), "Error executing SQL statement.")
        End Try
    End Sub

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
        ExecuteSQL(strSQLCommand)

        'Close the connection and reopen it pointing at the scheduling database
        If (DBConn.State = ConnectionState.Open) Then
            DBConn.Close()
        End If

        'Connect to SQL with security
        ConnectToSQL(False)

        'Build database tables
        BuildTables()

        'Insert the initial datasets
        InsertInitialData()

        'We can check to see if we're open before trying to issue a connection close
        If DBConn.State = ConnectionState.Open Then
            DBConn.Close()
        End If
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
        ExecuteSQL(strSQLCommand)

        'Build the Availability Table
        strSQLCommand = "CREATE TABLE Availability
            (TUID int IDENTITY(1,1) PRIMARY KEY,
            DoctorTUID int FOREIGN KEY REFERENCES Doctors(TUID) NOT NULL,
            Day char(1) NOT NULL,
            StartTime Time(0) NOT NULL,
            EndTime Time(0) NOT NULL,
            PatientCount int NOT NULL)"

        'Execute SQL command
        ExecuteSQL(strSQLCommand)

        'Build the Patients Table
        strSQLCommand = "CREATE TABLE Patients
             (TUID int IDENTITY(1,1) PRIMARY KEY,
             FirstName varchar(50) NOT NULL,
             LastName varchar(50) NOT NULL,
             Phone varchar(8) NOT NULL,
             Insurance varchar(20) NOT NULL)"

        'Execute SQL command
        ExecuteSQL(strSQLCommand)

        'Build the Appointments Table
        strSQLCommand = "CREATE TABLE Appointments
            (TUID int IDENTITY(1,1) PRIMARY KEY,
            PatientTUID int FOREIGN KEY REFERENCES Patients(TUID) NOT NULL,
            DoctorTUID int FOREIGN KEY REFERENCES Doctors(TUID) NOT NULL,
            Day Date NOT NULL,
            AppointmentLength int NOT NULL,
            StartTime Time(0) NOT NULL,
            EndTime Time(0) NOT NULL)"

        'Execute SQL command
        ExecuteSQL(strSQLCommand)
    End Sub

    Private Sub InsertInitialData()

        'String to hold SQL command text
        Dim strSQLCommand As String

        strSQLCommand = "INSERT INTO Doctors VALUES 
            ('Ray', 'Strantz'),
            ('Henry', 'Jones, Jr.'),
            ('Emmett', 'Brown')"

        'Execute SQL command
        ExecuteSQL(strSQLCommand)

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
        ExecuteSQL(strSQLCommand)
    End Sub

    Private Sub DeleteDatabase(strSERVERNAME As String, strDBNAME As String)
        'This routine deletes a database completely from code. Credit to CIS 311 ch 15 notes.

        'String to hold SQL command text
        Dim strSQLCommand As String

        'Connect to SQL Server instance
        ConnectToSQL(True)

        'Try to force single ownership of the database so that we have the permissions to delete it
        strSQLCommand = "ALTER DATABASE [" & strDBNAME & "] SET " &
            "SINGLE_USER WITH ROLLBACK IMMEDIATE"

        'Execute SQL command
        ExecuteSQL(strSQLCommand)

        'Now, drop the database
        strSQLCommand = "DROP DATABASE " & strDBNAME

        'Execute SQL command
        ExecuteSQL(strSQLCommand)

        If (DBConn.State = ConnectionState.Open) Then
            DBConn.Close()
        End If
    End Sub
End Class