Imports System.Data.SqlClient

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

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'If the database doesn't exist, create it
        If Not IO.File.Exists(strDBPATH) Then
            'Create database
            CreateDatabase(strSERVERNAME, strDBNAME, strDBPATH, strCONNECTION)
        End If
    End Sub

    Private Sub CreateDatabase(ByVal strSERVERNAME As String, ByVal strDBNAME As String, ByVal strDBPATH As String, ByVal strCONNECTION As String)
        'Build a SQL Server database from scratch
        Dim DBConn As SqlConnection
        Dim strSQLCmd As String
        Dim DBCmd As SqlCommand = New SqlCommand()

        'Point at the server
        DBConn = New SqlConnection("Server=" & strSERVERNAME)

        'Build the database
        strSQLCmd = "CREATE DATABASE " & strDBNAME & " On " &
                        "(NAME = '" & strDBNAME & "', " &
                        "FILENAME = '" & strDBPATH & "')"

        DBCmd.CommandText = strSQLCmd
        DBCmd.Connection = DBConn

        Try
            'Open the connection and try running the command
            DBConn.Open()
            DBCmd.ExecuteNonQuery()
            MessageBox.Show("Database was successfully created", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
            MessageBox.Show("Cannot build database! Closing program down...")
        End Try

        'Close the connection and reopen it pointing at the scheduling database
        If (DBConn.State = ConnectionState.Open) Then
            DBConn.Close()
        End If

        'Use the full connection string with the Integrated Security line
        DBConn = New SqlConnection(strCONNECTION)
        DBConn.Open()

        BuildTables(DBConn, DBCmd)

        'Insert the initial datasets
        InsertInitialData(DBConn, DBCmd)

        'We can check to see if we're open before trying to issue a connection close
        If DBConn.State = ConnectionState.Open Then
            DBConn.Close()
        End If
    End Sub

    Private Sub BuildTables(DBConn As SqlConnection, DBCmd As SqlCommand)
        'Build the Doctors Table
        DBCmd.CommandText = "CREATE TABLE Doctors 
                            (TUID int IDENTITY(1,1) PRIMARY KEY,
                            FirstName varchar(50) NOT NULL,
                            LastName varchar(50) NOT NULL)"

        DBCmd.Connection = DBConn

        Try
            DBCmd.ExecuteNonQuery()
            MessageBox.Show("Created Doctors Table")
        Catch Ex As Exception
            MessageBox.Show("Doctors Table Already Exists")
        End Try

        'Build the Availability Table
        DBCmd.CommandText = "CREATE TABLE Availability
                            (TUID int IDENTITY(1,1) PRIMARY KEY,
                            DoctorTUID int FOREIGN KEY REFERENCES Doctors(TUID) NOT NULL,
                            Day char(1) NOT NULL,
                            StartTime Time(0) NOT NULL,
                            EndTime Time(0) NOT NULL,
                            PatientCount int NOT NULL)"

        DBCmd.Connection = DBConn

        Try
            DBCmd.ExecuteNonQuery()
            MessageBox.Show("Created Availability Table")
        Catch Ex As Exception
            MessageBox.Show("Availability Table Already Exists")
        End Try

        'Build the Patients Table
        DBCmd.CommandText = "CREATE TABLE Patients
                            (TUID int IDENTITY(1,1) PRIMARY KEY,
                            FirstName varchar(50) NOT NULL,
                            LastName varchar(50) NOT NULL,
                            Phone varchar(8) NOT NULL,
                            Insurance varchar(20) NOT NULL)"

        DBCmd.Connection = DBConn

        Try
            DBCmd.ExecuteNonQuery()
            MessageBox.Show("Created Patients Table")
        Catch Ex As Exception
            MessageBox.Show("Patients Table Already Exists")
        End Try

        'Build the Appointments Table
        DBCmd.CommandText = "CREATE TABLE Appointments
                            (TUID int IDENTITY(1,1) PRIMARY KEY,
                            PatientTUID int FOREIGN KEY REFERENCES Patients(TUID) NOT NULL,
                            DoctorTUID int FOREIGN KEY REFERENCES Doctors(TUID) NOT NULL,
                            Day Date NOT NULL,
                            AppointmentLength int NOT NULL,
                            StartTime Time(0) NOT NULL,
                            EndTime Time(0) NOT NULL)"

        DBCmd.Connection = DBConn

        Try
            DBCmd.ExecuteNonQuery()
            MessageBox.Show("Created AppointmentsTable")
        Catch Ex As Exception
            MessageBox.Show("Appointments Table Already Exists")
        End Try
    End Sub

    Private Sub InsertInitialData(DBConn As SqlConnection, DBCmd As SqlCommand)
        DBCmd.CommandText = "INSERT INTO Doctors VALUES 
                            ('Ray', 'Strantz'),
                            ('Henry', 'Jones, Jr.'),
                            ('Emmett', 'Brown')"

        DBCmd.Connection = DBConn

        Try
            DBCmd.ExecuteNonQuery()
            MessageBox.Show("Added to Doctors Table")
        Catch Ex As Exception
            MessageBox.Show("Error adding to Doctors Table")
        End Try

        DBCmd.CommandText = "INSERT INTO Availability VALUES 
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

        Try
            DBCmd.ExecuteNonQuery()
            MessageBox.Show("Added to Availability Table")
        Catch Ex As Exception
            MessageBox.Show("Error adding to Availability Table")
        End Try
    End Sub

    Private Sub DeleteDataFromTables()
        Dim DBConn As SqlConnection
        Dim DBCmd As SqlCommand = New SqlCommand()

        'Use the full connection string with the Integrated Security line
        DBConn = New SqlConnection(strCONNECTION)
        DBConn.Open()

        DBCmd.CommandText = "DELETE FROM Patients"

        DBCmd.Connection = DBConn

        Try
            DBCmd.ExecuteNonQuery()
            MessageBox.Show("Deleted from Patients Table")
        Catch Ex As Exception
            MessageBox.Show("Failed to delete patients table")
        End Try

        DBCmd.CommandText = "DELETE FROM Appointments"

        DBCmd.Connection = DBConn

        Try
            DBCmd.ExecuteNonQuery()
            MessageBox.Show("Deleted from Appointments Table")
        Catch Ex As Exception
            MessageBox.Show("Failed to delete appointments table")
        End Try
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
        'Make sure the user has selected a text file before continuing
        If txtFileName.Text = Nothing Then
            MessageBox.Show("Please select a txt file before continuing!")
        End If


    End Sub

    Private Sub btnQuit_Click(sender As Object, e As EventArgs) Handles btnQuit.Click
        'Quit nicely
        Me.Close()
    End Sub

    Private Sub frmMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        'Ask the user if they want to delete the current db
        Dim result As DialogResult = MessageBox.Show("Would you like to blast away the current data?", "Exiting", MessageBoxButtons.YesNo)

        'If yes, delete it. Otherwise, keep exiting
        If result = DialogResult.Yes Then
            DeleteDataFromTables()
        End If
    End Sub
End Class