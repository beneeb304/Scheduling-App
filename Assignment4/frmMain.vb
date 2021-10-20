Imports System.Data.SqlClient
Imports System.IO

Public Class frmMain
    '------------------------------------------------------------
    '-                File Name : frmMain.vb                    -
    '-                Part of Project: Assignment4              -
    '------------------------------------------------------------
    '-                Written By: Benjamin Neeb                 -
    '-                Written On: October 21, 2021              -
    '------------------------------------------------------------
    '- File Purpose:                                            -
    '-                                                          -
    '- This file contains the main Sub for the form             -
    '- application. This file performs all of the interaction   -
    '- with the user and database.                              -
    '------------------------------------------------------------
    '- Program Purpose:                                         -
    '-                                                          -
    '- This program simulates a doctor scheduling application.  -
    '------------------------------------------------------------
    '- Global Variable Dictionary (alphabetically):             -
    '- strDBPath        - String that holds the database path   -
    '- strConnection    - Connection String to database         -
    '- DBConn           - SQL Connection object                 -
    '------------------------------------------------------------

    'Path to database in executable
    Private ReadOnly strDBPath As String = My.Application.Info.DirectoryPath & "\" & strDBNAME & ".mdf"

    'This is the full connection string
    Private ReadOnly strConnection As String = "SERVER=" & strSERVERNAME & ";DATABASE=" &
                     strDBNAME & ";Integrated Security=SSPI;AttachDbFileName=" & strDBPath

    'SQL Connection object
    Private DBConn As SqlConnection

    '---------------------------------------------------------------------------------------
    '--- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS ---
    '--- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS ---
    '--- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS ---
    '---------------------------------------------------------------------------------------

    'Name of database
    Private Const strDBNAME As String = "Scheduling"

    'Name of the database server
    Private Const strSERVERNAME As String = "(localdb)\MSSQLLocalDB"

    'Number of days in a week
    Private Const intDAYSINWEEK As Integer = 7

    'Time increment for appointments
    Private Const intTIMEINCREMENT As Integer = 15

    '-----------------------------------------------------------------------------------
    '--- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS ---
    '--- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS ---
    '--- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS ---
    '-----------------------------------------------------------------------------------

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '------------------------------------------------------------
        '-              Subprogram Name: frmMain_Load               -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: October 21, 2021              -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine is called when the for loads. It checks  -
        '- to see if the database has been created. If it hasn't,   -
        '- it creates it.                                           -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- sender – Identifies which particular control that raised -
        '-          the click event                                 - 
        '- e – Holds the EventArgs object sent to the routine       -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- (None)                                                   -
        '------------------------------------------------------------

        'If the database doesn't exist, create it
        If Not File.Exists(strDBPath) Then
            'Create database, build tables, and insert initial data
            CreateDatabase()
        End If
    End Sub

    Private Sub btnSelectFile_Click(sender As Object, e As EventArgs) Handles btnSelectFile.Click
        '------------------------------------------------------------
        '-              Subprogram Name: btnSelectFile_Click        -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: October 21, 2021              -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine is called when the user clicks on the    -
        '- select file button. This opens an openfiledialog that    -
        '- allows the user to select a txt file.                    -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- sender – Identifies which particular control that raised -
        '-          the click event                                 - 
        '- e – Holds the EventArgs object sent to the routine       -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- ofd      - Open File Dialog instance                     -
        '------------------------------------------------------------

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
        '------------------------------------------------------------
        '-              Subprogram Name: btnEnterFile_Click         -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: October 21, 2021              -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine is called when the user clicks on the    -
        '- enter file button. This begins processing the file for   -
        '- input.                                                   -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- sender – Identifies which particular control that raised -
        '-          the click event                                 - 
        '- e – Holds the EventArgs object sent to the routine       -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- strText      - String to hold each line in the text file -
        '------------------------------------------------------------

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
            MessageBox.Show(ex.ToString(), "Please read in a valid text file!")
        End Try

        'Empty controls so user is forced to realize new data (if applicable)
        txtFileName.Text = ""
        cmbFilter.SelectedItem = Nothing
        lstSelect.Items.Clear()
        lstDisplay.Items.Clear()
    End Sub

    Private Sub btnQuit_Click(sender As Object, e As EventArgs) Handles btnQuit.Click
        '------------------------------------------------------------
        '-              Subprogram Name: btnQuit_Click              -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: October 21, 2021              -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine is called when the user clicks on the    -
        '- quit button. This begins processing of closing the form  -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- sender – Identifies which particular control that raised -
        '-          the click event                                 - 
        '- e – Holds the EventArgs object sent to the routine       -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- (None)                                                   -
        '------------------------------------------------------------

        'Quit nicely
        Close()
    End Sub

    Private Sub frmMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        '------------------------------------------------------------
        '-              Subprogram Name: frmMain_FormClosing        -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: October 21, 2021              -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine is called when the form starts to close. -
        '- The program asks the user if they want to keep the data- -
        '- base or blast it.                                        -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- sender – Identifies which particular control that raised -
        '-          the click event                                 - 
        '- e – Holds the EventArgs object sent to the routine       -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- result       - DialogResult that holds the answer to     -
        '-                keeping or blasting the database.         -
        '------------------------------------------------------------

        'Ask the user if they want to delete the current db
        Dim result As DialogResult = MessageBox.Show("Would you like to blast away the current database?", "Exiting", MessageBoxButtons.YesNoCancel)

        'If yes, delete it. Otherwise, keep exiting
        If result = DialogResult.Yes Then
            DeleteDatabase()
        ElseIf result = DialogResult.Cancel Then
            e.Cancel = True
        End If
    End Sub

    Private Sub txtFileName_TextChanged(sender As Object, e As EventArgs) Handles txtFileName.TextChanged
        '------------------------------------------------------------
        '-              Subprogram Name: txtFileName_TextChanged    -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: October 21, 2021              -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine is called when the text in txtFileName   -
        '- changes. It checks whether it is empty or not.           -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- sender – Identifies which particular control that raised -
        '-          the click event                                 - 
        '- e – Holds the EventArgs object sent to the routine       -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- (None)                                                   -
        '------------------------------------------------------------

        If txtFileName.Text.Length > 0 Then
            btnEnterFile.Enabled = True
        Else
            btnEnterFile.Enabled = False
        End If
    End Sub

    Private Sub cmbFilter_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbFilter.SelectedIndexChanged
        '------------------------------------------------------------
        '-     Subprogram Name: cmbFilter_SelectedIndexChanged      -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: October 21, 2021              -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine is called when the user selects an item  -
        '- from the combo-box.                                      -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- sender – Identifies which particular control that raised -
        '-          the click event                                 - 
        '- e – Holds the EventArgs object sent to the routine       -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- dataReader   - SqlDataReader that will return SQL result -
        '------------------------------------------------------------

        'Clear listboxes
        lstSelect.Items.Clear()
        lstDisplay.Items.Clear()

        Dim dataReader As SqlDataReader

        If cmbFilter.SelectedItem = "Patients" Then
            'Datareader and SQL command
            dataReader = ExecuteSQLReader("SELECT Phone, FirstName, LastName FROM Patients")

            'While there's data, read it
            While dataReader.Read()
                'Add to listbox
                lstSelect.Items.Add(dataReader("Phone") & " - " & dataReader("FirstName") & " " & dataReader("LastName"))
            End While

            'Tidy up
            dataReader.Close()
            DisconnectFromSQL()
        ElseIf cmbFilter.SelectedItem = "Doctor Appointments" Or cmbFilter.SelectedItem = "Doctor Availability" Then
            'Datareader and SQL command
            dataReader = ExecuteSQLReader("SELECT TUID, FirstName, LastName FROM Doctors")

            'While there's data, read it
            While dataReader.Read()
                'Add to listbox
                lstSelect.Items.Add(dataReader("TUID") & " - Dr. " & dataReader("FirstName") & " " & dataReader("LastName"))
            End While

            'Tidy up
            dataReader.Close()
            DisconnectFromSQL()
        End If
    End Sub

    Private Sub lstSelect_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstSelect.SelectedIndexChanged
        '------------------------------------------------------------
        '-     Subprogram Name: lstSelect_SelectedIndexChanged      -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: October 21, 2021              -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine is called when the user selects an item  -
        '- from the list-box.                                       -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- sender – Identifies which particular control that raised -
        '-          the click event                                 - 
        '- e – Holds the EventArgs object sent to the routine       -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- dataReader   - SqlDataReader that will return SQL result -
        '- intTUID      - Int to hold the TUID of patient or doctor -
        '- thisDate     - Date to hold current date                 -
        '------------------------------------------------------------

        'Clear listboxes
        lstDisplay.Items.Clear()

        Dim intTUID As Integer
        Dim thisDate As Date

        'Determine if selecting Patients, Doctor Appointments, or Doctor Availability
        If cmbFilter.SelectedItem = "Patients" And lstSelect.SelectedIndex > -1 Then
            Dim dataReader = ExecuteSQLReader("SELECT TUID FROM Patients WHERE Phone = '" & lstSelect.SelectedItem.ToString().Substring(0, 8) & "'")

            While dataReader.Read()
                intTUID = dataReader("TUID")
            End While

            dataReader.Close()
            DisconnectFromSQL()

            dataReader = ExecuteSQLReader("SELECT Day, StartTime, EndTime FROM Appointments WHERE PatientTUID = " & intTUID & " ORDER BY Day, StartTime")

            While dataReader.Read()
                'Get day of week
                thisDate = CDate(dataReader("Day").ToString())

                'Add to listbox
                lstDisplay.Items.Add(thisDate.DayOfWeek.ToString() & ", " & CStr(dataReader("Day")) & " from " & dataReader("StartTime").ToString() & " to " & dataReader("EndTime").ToString())
            End While

            dataReader.Close()
            DisconnectFromSQL()

        ElseIf cmbFilter.SelectedItem = "Doctor Appointments" And lstSelect.SelectedIndex > -1 Then
            intTUID = lstSelect.SelectedItem.ToString().Substring(0, 1)
            Dim dataReader = ExecuteSQLReader("SELECT Day, StartTime, EndTime FROM Appointments WHERE DoctorTUID = " & intTUID & "AND PatientTUID IS NOT NULL ORDER BY Day, StartTime")

            'While there's data, read it
            While dataReader.Read()
                'Get day of week
                thisDate = CDate(dataReader("Day").ToString())

                'Add to listbox
                lstDisplay.Items.Add(thisDate.DayOfWeek.ToString() & ", " & CStr(dataReader("Day")) & " from " & dataReader("StartTime").ToString() & " to " & dataReader("EndTime").ToString())
            End While

            dataReader.Close()
            DisconnectFromSQL()
        ElseIf cmbFilter.SelectedItem = "Doctor Availability" And lstSelect.SelectedIndex > -1 Then
            intTUID = lstSelect.SelectedItem.ToString().Substring(0, 1)
            Dim dataReader = ExecuteSQLReader("SELECT Day, StartTime, EndTime FROM Appointments WHERE DoctorTUID = " & intTUID & "AND PatientTUID IS NULL ORDER BY Day, StartTime")

            'While there's data, read it
            While dataReader.Read()
                'Get day of week
                thisDate = CDate(dataReader("Day").ToString())

                'Check availability for day
                If GetAvailability(thisDate, intTUID) Then
                    'Add to listbox
                    lstDisplay.Items.Add(thisDate.DayOfWeek.ToString() & ", " & CStr(dataReader("Day")) & " from " & dataReader("StartTime").ToString() & " to " & dataReader("EndTime").ToString())
                End If
            End While
            dataReader.Close()
            DisconnectFromSQL()
        End If
    End Sub

    Private Sub ParseTextFile(strLines() As String)
        '------------------------------------------------------------
        '-              ubprogram Name: ParseTextFile               -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: October 21, 2021              -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine is used to parse the text from the text  -
        '- file                                                     -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- strLines()   - String array to hold lines from text file -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- dataReader   - SqlDataReader that will return SQL result -
        '- strLine()    - String array to hold line split by tabs   -
        '- strSQL       - String to hold SQL command                -
        '- intTUID      - Int to hold patient TUID                  -
        '------------------------------------------------------------

        'Create String array for pieces of data in each line of text
        Dim strLine() As String
        Dim strSQL As String = ""
        Dim intTUID = 0

        'Cycle through each line
        For Each line As String In strLines
            'Split data by vbTab
            strLine = line.Split(vbTab)

            'Patient
            If strLine(0) = "P" Then
                'If first time, add beginning of insert statement
                If strSQL.Length = 0 Then
                    strSQL = "INSERT INTO Patients VALUES "
                End If

                'Append to insert statement
                strSQL &= "('" & strLine(1).Substring(0, strLine(1).IndexOf(" ")) & "', " &
                    "'" & strLine(1).Substring(strLine(1).IndexOf(" ") + 1) & "', " &
                    "'" & strLine(2) & "', " &
                    "'" & strLine(3) & "')," & vbCrLf
            End If
        Next

        'Chop off last comma
        strSQL = strSQL.Substring(0, strSQL.LastIndexOf(","))

        'Execute SQL command
        ExecuteSQLNonQ(strSQL, False)

        For Each line As String In strLines
            'Split data by vbTab
            strLine = line.Split(vbTab)

            'Appointment
            If strLine(0) = "A" Then
                'Operation
                Select Case strLine(2)
                    'Adding appointment
                    Case "A"
                        AddAppointment(strLine)

                    'Deleting appointment
                    Case "D"
                        DeleteAppointment(strLine)

                    'Changing appointment
                    Case "C"
                        'Delete first appointment for the patient with phone number

                        'Get patient TUID
                        strSQL = "SELECT TUID FROM Patients WHERE Phone = '" & strLine(1) & "'"
                        Dim dataReader As SqlDataReader = ExecuteSQLReader(strSQL)
                        While dataReader.Read()
                            intTUID = dataReader("TUID")
                        End While
                        dataReader.Close()
                        DisconnectFromSQL()

                        strSQL = "SET ROWCOUNT 1 DELETE FROM Appointments WHERE PatientTUID = " & intTUID

                        ExecuteSQLNonQ(strSQL, False)

                        'Add back in
                        AddAppointment(strLine)
                End Select
            End If
        Next
    End Sub

    Private Sub AddAppointment(strLine() As String)
        '------------------------------------------------------------
        '-              Subprogram Name: AddAppointment             -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: October 21, 2021              -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine is used to add an appointment to the     -
        '- data-base                                                -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- strLine()    - String array to hold line split by tabs   -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- dataReader   - SqlDataReader that will return SQL result -
        '- strSQL       - String to hold SQL command                -
        '- intPTUID     - Int to hold patient TUID                  -
        '- intDTUID     - Int to hold doctor TUID                   -
        '- intConsRows  - Int to hold the amount of consecutive     -
        '-                rows needed for db entry                  -
        '- chrSelector  - Char used to hold the selector for        -
        '-                appointment preference                    -
        '- thisDate     - Date var to hold current date             -
        '- thisTime     - Date var to hold current time             -
        '- intCtr       - Int to act as a loop counter              -
        '- blnFound     - Boolean flag to indicate found data       -
        '- dctDocs      - Dictionary of doctor names and their TUID -
        '- intDelta     - Int to determine positive or negative day -
        '------------------------------------------------------------

        'SQL String
        Dim strSQL As String

        'Get consecutive rows needed
        Dim intConsRows As Integer

        'Vars to hold date and time
        Dim thisTime As Date = Nothing
        Dim thisDate As Date = Nothing

        'Vars to help with looping
        Dim intCtr As Integer = 1
        Dim blnFound As Boolean = False

        'TUIDs
        Dim intDTUID = 0
        Dim intPTUID As Integer

        'Get patient TUID
        Dim dataReader = ExecuteSQLReader("SELECT TUID FROM Patients WHERE Phone = '" & strLine(1) & "'")
        While dataReader.Read()
            'Get patient TUID
            intPTUID = Integer.Parse(dataReader("TUID").ToString())
        End While
        dataReader.Close()
        DisconnectFromSQL()

        'Get selector
        Dim chrSelector As Char = strLine(3).Substring(0, 1)

        'Determine preference
        Select Case chrSelector
            'Next available preference
            Case "N"
                'Get number of consecutive rows we need for this appointment
                intConsRows = Integer.Parse(strLine(4)) / 15

                'Get all appointments that are available
                dataReader = ExecuteSQLReader("SELECT * FROM Appointments WHERE PatientTUID IS NULL ORDER BY Day, StartTime")

                'While there's data, read it
                While dataReader.Read()
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
                DisconnectFromSQL()

                If blnFound Then
                    'Alter statement
                    strSQL = "SET ROWCOUNT 1 " &
                        "UPDATE Appointments " &
                        "SET PatientTUID = " & intPTUID & ", EndTime = '" & thisTime & "', AppointmentLength = " & strLine(4) &
                        " OUTPUT inserted.[DoctorTUID]" &
                        " WHERE StartTime = '" & thisTime.AddMinutes(Integer.Parse(strLine(4) / -1)) & "' AND Day = '" & thisDate & "'"

                    'Get doctorTUID
                    dataReader = ExecuteSQLReader(strSQL)

                    While dataReader.Read()
                        intDTUID = dataReader("DoctorTUID")
                    End While

                    'Close dataReader
                    dataReader.Close()
                    DisconnectFromSQL()

                    'Make new SQL command to delete those empty appointment slots
                    strSQL = "DELETE FROM Appointments WHERE " & "PatientTUID IS NULL AND DoctorTUID = " & intDTUID &
                        " AND Day = '" & thisDate & "' AND EndTime BETWEEN '" & thisTime.AddMinutes(Integer.Parse(strLine(4) / -1)) & "' AND '" & thisTime & "'"

                    ExecuteSQLNonQ(strSQL, False)
                Else
                    MessageBox.Show("We could not find an appointment for " & strLine(1))
                End If
                            'Doctor preference
            Case "D"
                'Get doctor TUID matched with doc last name
                dataReader = ExecuteSQLReader("SELECT TUID, LastName FROM Doctors")

                'Make collection to hold results
                Dim dctDocs As New Dictionary(Of String, String)

                While dataReader.Read()
                    dctDocs.Add(dataReader("LastName").ToString(), dataReader("TUID").ToString())
                End While

                'Close dataReader
                dataReader.Close()
                DisconnectFromSQL()

                'Get number of consecutive rows we need for this appointment
                intConsRows = Integer.Parse(strLine(4)) / 15

                'Get the doctor TUID
                intDTUID = dctDocs.Item(strLine(3).Substring(2))

                'Get all appointments that are available
                dataReader = ExecuteSQLReader("SELECT * FROM Appointments WHERE PatientTUID IS NULL AND DoctorTUID = " & intDTUID & " ORDER BY Day, StartTime")

                'While there's data, read it
                While dataReader.Read()
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
                DisconnectFromSQL()

                If blnFound Then
                    'Alter statement
                    strSQL = "SET ROWCOUNT 1 " &
                        "UPDATE Appointments " &
                        "SET PatientTUID = " & intPTUID & ", EndTime = '" & thisTime & "', AppointmentLength = " & strLine(4) &
                        " WHERE StartTime = '" & thisTime.AddMinutes(Integer.Parse(strLine(4) / -1)) & "' AND Day = '" & thisDate & "' AND DoctorTUID = " & intDTUID

                    ExecuteSQLNonQ(strSQL, False)

                    'Make new SQL command to delete those empty appointment slots
                    strSQL = "DELETE FROM Appointments WHERE " & "PatientTUID IS NULL AND DoctorTUID = " & intDTUID &
                        " AND Day = '" & thisDate & "' AND EndTime BETWEEN '" & thisTime.AddMinutes(Integer.Parse(strLine(4) / -1)) & "' AND '" & thisTime & "'"

                    ExecuteSQLNonQ(strSQL, False)

                Else
                    MessageBox.Show("We could not find an appointment for " & strLine(1))
                End If
                            'Time/day preference
            Case "T"
                'Get day data
                Dim intToday As Integer = Today.DayOfWeek
                Dim dayOfWeek As DayOfWeek
                Select Case strLine(3).Substring(2, 1)
                    Case "M"
                        dayOfWeek = DayOfWeek.Monday
                    Case "T"
                        dayOfWeek = DayOfWeek.Tuesday
                    Case "W"
                        dayOfWeek = DayOfWeek.Wednesday
                    Case "R"
                        dayOfWeek = DayOfWeek.Thursday
                    Case "F"
                        dayOfWeek = DayOfWeek.Friday
                End Select

                'Find next day of the week
                Dim intDelta As Integer = dayOfWeek - intToday
                If intDelta > 0 Then
                    thisDate = Today.AddDays(intDelta)
                Else
                    thisDate = Today.AddDays(7 + intDelta)
                End If

                'Get start and end times
                Dim startTime As Date = strLine(3).Substring(4)
                Dim endTime As Date = startTime.AddMinutes(strLine(4))

                strSQL = "SET ROWCOUNT 1 " &
                "UPDATE Appointments " &
                "SET PatientTUID = " & intPTUID & ", EndTime = '" & endTime & "', AppointmentLength = " & strLine(4) &
                " OUTPUT inserted.[DoctorTUID]" &
                " WHERE StartTime = '" & strLine(3).Substring(4) & "' AND Day = '" & thisDate & "'"

                'Get doctorTUID
                dataReader = ExecuteSQLReader(strSQL)

                Dim intDoctorTUID = 0

                While dataReader.Read()
                    intDoctorTUID = dataReader("DoctorTUID")
                End While

                'Close dataReader
                dataReader.Close()
                DisconnectFromSQL()

                'Make new SQL command to delete those empty appointment slots
                strSQL = "DELETE FROM Appointments WHERE " & "PatientTUID IS NULL AND DoctorTUID = " & intDoctorTUID &
                        " AND Day = '" & thisDate & "' AND EndTime BETWEEN '" & startTime.AddMinutes(15) & "' AND '" & endTime & "'"

                ExecuteSQLNonQ(strSQL, False)
        End Select
    End Sub

    Private Sub DeleteAppointment(strLine() As String)
        '------------------------------------------------------------
        '-          Subprogram Name: DeleteAppointment              -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: October 21, 2021              -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine is used to delete an appointment from    -
        '- the data-base                                            -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- strLine()    - String array to hold line split by tabs   -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- dataReader   - SqlDataReader that will return SQL result -
        '- strSQL       - String to hold SQL command                -
        '- intPTUID     - Int to hold patient TUID                  -
        '- intDTUID     - Int to hold doctor TUID                   -
        '- aptDate      - Date var to hold current date             -
        '- startTime    - Date var to hold current time             -
        '- intToday     - Int to hold weekday number                -
        '- dayOfWeek    - DayOfWeek holds the current weekday       -
        '- intDelta     - Int to determine positive or negative day -
        '------------------------------------------------------------

        'SQL String
        Dim strSQL As String
        Dim intDuration As Integer
        Dim intPTUID = 0
        Dim intDTUID As Integer
        Dim aptDay As Date

        'Get day data
        Dim startTime As Date = strLine(3).Substring(2)
        Dim intToday As Integer = Today.DayOfWeek
        Dim dayOfWeek As DayOfWeek
        Select Case strLine(3).Substring(0, 1)
            Case "M"
                dayOfWeek = DayOfWeek.Monday
            Case "T"
                dayOfWeek = DayOfWeek.Tuesday
            Case "W"
                dayOfWeek = DayOfWeek.Wednesday
            Case "R"
                dayOfWeek = DayOfWeek.Thursday
            Case "F"
                dayOfWeek = DayOfWeek.Friday
        End Select

        'Find next day of the week
        Dim intDelta As Integer = dayOfWeek - intToday
        If intDelta > 0 Then
            aptDay = Today.AddDays(intDelta)
        Else
            aptDay = Today.AddDays(7 + intDelta)
        End If

        'Get patient TUID
        strSQL = "SELECT TUID FROM Patients WHERE Phone = '" & strLine(1) & "'"
        Dim dataReader As SqlDataReader = ExecuteSQLReader(strSQL)
        While dataReader.Read()
            intPTUID = dataReader("TUID")
        End While
        dataReader.Close()
        DisconnectFromSQL()

        strSQL = "DELETE FROM Appointments OUTPUT deleted.[AppointmentLength], deleted.[DoctorTUID] " &
            "WHERE PatientTUID = " & intPTUID & " AND Day = '" & aptDay & "' AND StartTime = '" & startTime & "'"

        'Get EndTime
        dataReader = ExecuteSQLReader(strSQL)

        While dataReader.Read()
            intDuration = dataReader("AppointmentLength")
            intDTUID = dataReader("DoctorTUID")
        End While

        'Close dataReader
        dataReader.Close()
        DisconnectFromSQL()

        strSQL = "INSERT INTO Appointments VALUES"

        'Loop for apt duration /15 which will give us how many slots to add back to availability
        For i = 1 To intDuration / 15
            strSQL &= " (NULL, " & intDTUID & ", '" & aptDay & "', 15, '" & startTime & "', '" & startTime.AddMinutes(intTIMEINCREMENT) & "'),"

            'increment start time
            startTime = startTime.AddMinutes(15)
        Next

        'Chop off last comma
        strSQL = strSQL.Substring(0, strSQL.LastIndexOf(","))

        ExecuteSQLNonQ(strSQL, False)
    End Sub

    Private Sub AddPatient(strPhone As String)
        '------------------------------------------------------------
        '-              Subprogram Name: AddPatient                 -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: October 21, 2021              -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine is used to add a patient to the database -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- strPhone     - String to hold patients phone number      -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- strData      - String array to hold patient information  -
        '- strSQL       - String to hold SQL command                -
        '------------------------------------------------------------

        Dim strData(2) As String

        'Get patient's first name
        strData(0) = InputBox("Please enter the patient's first name.", "New patient! Please fill out their information.")

        'Get patient's last name
        strData(1) = InputBox("Please enter the patient's last name.", "New patient! Please fill out their information.")

        'Get patient's insurance provider
        strData(2) = InputBox("Please enter the patient's insurance provider.", "New patient! Please fill out their information.")

        'Make SQL command
        Dim strSQL As String = "INSERT INTO Patients VALUES " &
            "('" & strData(0) & "', " &
            "'" & strData(1) & "', " &
            "'" & strPhone & "', " &
            "'" & strData(2) & "')"

        'Execute SQL command
        ExecuteSQLNonQ(strSQL, False)
    End Sub

    Private Sub ExecuteSQLNonQ(strSQL As String, blnSecurity As Boolean)
        '------------------------------------------------------------
        '-              Subprogram Name: ExecuteSQLNonQ             -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: October 21, 2021              -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine is used to execute SQL commands          -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- strSQL       - String to hold SQL command                -
        '- blnSecurity  - Boolean to determine if we need to add    -
        '-                security to the SQL connection string     -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- DBCmd        - SQLCommand to be sent to the database     -
        '------------------------------------------------------------

        'Connect to SQL
        ConnectToSQL(blnSecurity)

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

        'Disconnect from Database
        DisconnectFromSQL()
    End Sub

    Private Function ExecuteSQLReader(strSQL As String) As SqlDataReader
        '------------------------------------------------------------
        '-              Function Name: ExecuteSQLReader             -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: October 21, 2021              -
        '------------------------------------------------------------
        '- Function Purpose:                                        -
        '-                                                          -
        '- This Function is used execute SQL queries and return the -
        '- result                                                   -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- strSQL       - String to hold SQL command                -
        '------------------------------------------------------------
        '- Return Data                                              -
        '- dataReader   - SqlDataReader that will return SQL result -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- DBCmd        - SQLCommand to be sent to the database     -
        '- dataReader   - SqlDataReader that will return SQL result -
        '------------------------------------------------------------

        'Connect to SQL
        ConnectToSQL(False)

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
        '------------------------------------------------------------
        '-              Subprogram Name: ExecuteSQLNonQ             -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: October 21, 2021              -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine is used to connect to the database       -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- blnInitial   - Boolean to determine if this is the first -
        '-                time connecting                           -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- (None)                                                   -
        '------------------------------------------------------------

        If blnInitial Then
            'Initial connection
            DBConn = New SqlConnection("Server=" & strSERVERNAME)
        Else
            'Use the full connection string with the Integrated Security line
            DBConn = New SqlConnection(strConnection)
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
        '------------------------------------------------------------
        '-              Subprogram Name: DisconnectFromSQL          -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: October 21, 2021              -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine is used to disconnect from the database  -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- (None)                                                   -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- (None)                                                   -
        '------------------------------------------------------------

        If (DBConn.State = ConnectionState.Open) Then
            DBConn.Close()
        End If
    End Sub

    Private Sub CreateDatabase()
        '------------------------------------------------------------
        '-              Subprogram Name: CreateDatabase             -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: October 21, 2021              -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine is used to create the database           -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- (None)                                                   -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- strSQL       - String to hold SQL command                -
        '------------------------------------------------------------

        'String to hold SQL command text
        Dim strSQL As String

        'Build the database command
        strSQL = "CREATE DATABASE " & strDBNAME & " On " &
            "(NAME = '" & strDBNAME & "', " &
            "FILENAME = '" & strDBPath & "')"

        'Execute SQL command
        ExecuteSQLNonQ(strSQL, True)

        'Build database tables
        BuildTables()

        'Insert the initial datasets
        InsertInitialData()
    End Sub

    Private Sub BuildTables()
        '------------------------------------------------------------
        '-              Subprogram Name: BuildTables                -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: October 21, 2021              -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine is used to create the database tables    -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- (None)                                                   -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- strSQL       - String to hold SQL command                -
        '------------------------------------------------------------

        'String to hold SQL command text
        Dim strSQL As String

        'Build the Doctors Table
        strSQL = "CREATE TABLE Doctors
            (TUID int IDENTITY(1,1) PRIMARY KEY,
            FirstName varchar(50) NOT NULL,
            LastName varchar(50) NOT NULL)"

        'Execute SQL command
        ExecuteSQLNonQ(strSQL, False)

        'Build the Availability Table
        strSQL = "CREATE TABLE Availability
            (TUID int IDENTITY(1,1) PRIMARY KEY,
            DoctorTUID int FOREIGN KEY REFERENCES Doctors(TUID) NOT NULL,
            Day char(1) NOT NULL,
            StartTime Time(0) NOT NULL,
            EndTime Time(0) NOT NULL,
            PatientCount int NOT NULL)"

        'Execute SQL command
        ExecuteSQLNonQ(strSQL, False)

        'Build the Patients Table
        strSQL = "CREATE TABLE Patients
             (TUID int IDENTITY(1,1) PRIMARY KEY,
             FirstName varchar(50) NOT NULL,
             LastName varchar(50) NOT NULL,
             Phone varchar(8) NOT NULL UNIQUE,
             Insurance varchar(20) NOT NULL)"

        'Execute SQL command
        ExecuteSQLNonQ(strSQL, False)

        'Build the Appointments Table
        strSQL = "CREATE TABLE Appointments
            (TUID int IDENTITY(1,1) PRIMARY KEY, 
            PatientTUID int FOREIGN KEY REFERENCES Patients(TUID), 
            DoctorTUID int FOREIGN KEY REFERENCES Doctors(TUID) NOT NULL,
            Day Date NOT NULL,
            AppointmentLength int NOT NULL,
            StartTime TIME(0), 
            EndTime TIME(0))"

        'Execute SQL command
        ExecuteSQLNonQ(strSQL, False)
    End Sub

    Private Sub InsertInitialData()
        '------------------------------------------------------------
        '-              Subprogram Name: InsertInitialData          -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: October 21, 2021              -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine is used to insert the initial data into  -
        '- the database                                             -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- (None)                                                   -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- strSQL       - String to hold SQL command                -
        '- thisDate     - Date to hold current date                 -
        '- strDayofWeek - String to hold the day of the week        -
        '------------------------------------------------------------

        'String to hold SQL command text
        Dim strSQL As String

        'Create today date object
        Dim thisDate As Date = Date.Today

        'String for date
        Dim strDayofWeek As String

        strSQL = "INSERT INTO Doctors VALUES 
            ('Ray', 'Stantz'),
            ('Henry', 'Jones'),
            ('Emmett', 'Brown')"

        'Execute SQL command
        ExecuteSQLNonQ(strSQL, False)

        strSQL = "INSERT INTO Availability VALUES 
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
        ExecuteSQLNonQ(strSQL, False)

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
        '------------------------------------------------------------
        '-              Subprogram Name: PopulateAppointments       -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: October 21, 2021              -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine is used to populate the initially blank  -
        '- appointments into the database                           -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- strDate      - String to hold the start date             -
        '- chrDay       - Char to hold the weekday prefix           -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- dataReader   - SqlDataReader that will return SQL result -
        '- strSQL       - String to hold SQL command                -
        '- startTime    - Date to hold appointment start time       -
        '- endTime      - Date to hold appointment end time         -
        '- intDTUID     - Int to hold the doctor TUID               -
        '------------------------------------------------------------

        'Create time variables
        Dim startTime, endTime As Date

        'Create SQL Command
        Dim strSQL As String = "SELECT DoctorTUID, StartTime, EndTime FROM Availability WHERE Day = '" & chrDay & "'"

        'Doctor TUID
        Dim intDTUID As Integer

        'Insert initial available appointments for 1 week out
        Dim dataReader = ExecuteSQLReader(strSQL)

        'Empty SQL command
        strSQL = ""

        If dataReader.HasRows() Then

            'While there's data, read it
            While dataReader.Read()
                'Capture data returned from query
                intDTUID = dataReader("DoctorTUID")

                startTime = CDate(dataReader("StartTime").ToString())
                endTime = CDate(dataReader("EndTime").ToString())

                'Loop while the start time is before the end time
                While startTime < endTime

                    'If first time adding to SQL command, create it
                    If strSQL.Length = 0 Then
                        strSQL = "INSERT INTO Appointments VALUES" & vbCrLf
                    End If

                    'Append to SQL command
                    strSQL &= "(NULL, " & intDTUID & ", '" & strDate & "', " &
                        intTIMEINCREMENT & ", '" & startTime.ToString("HH:mm") &
                        "', '" & startTime.AddMinutes(intTIMEINCREMENT).ToString("HH:mm") & "')," & vbCrLf

                    'Increment time by 15 minutes
                    startTime = startTime.AddMinutes(intTIMEINCREMENT)
                End While
            End While

            'Close datareader
            dataReader.Close()
            DisconnectFromSQL()

            'Chop off last comma
            strSQL = strSQL.Substring(0, strSQL.LastIndexOf(","))

            ExecuteSQLNonQ(strSQL, False)
        End If
    End Sub

    Private Sub DeleteDatabase()
        '------------------------------------------------------------
        '-              Subprogram Name: DeleteDatabase             -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: October 21, 2021              -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine is used to delete a database (credit to  -
        '- CIS 311 chapter 15 notes)                                -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- (None)                                                   -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- strSQL       - String to hold SQL command                -
        '------------------------------------------------------------

        'Try to force single ownership of the database so that we have the permissions to delete it
        Dim strSQL As String = "ALTER DATABASE [" & strDBNAME & "] SET " &
            "SINGLE_USER WITH ROLLBACK IMMEDIATE"

        'Execute SQL command
        ExecuteSQLNonQ(strSQL, True)

        'Now, drop the database
        strSQL = "DROP DATABASE " & strDBNAME

        'Execute SQL command
        ExecuteSQLNonQ(strSQL, True)
    End Sub

    Private Function GetAvailability(thisDate As Date, intDTUID As Integer) As Boolean
        '------------------------------------------------------------
        '-              Function Name: GetAvailability            -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: October 21, 2021              -
        '------------------------------------------------------------
        '- Function Purpose:                                        -
        '-                                                          -
        '- This Function is used to get a doctor's availability     -
        '- by checking if the appointment limit has been reached    -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- thisDate     - Date to check on particular day           -
        '- intDTUID     - Doctor TUID                               -
        '------------------------------------------------------------
        '- Return Data                                              -
        '- True or False                                            -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- strSQL       - String to hold SQL command                -
        '- chrDay       - Char to hold weekday prefix               -
        '- intLimit     - Int to hold the doctor's patient limit    -
        '- intCount     - Int to hold the doctor's current patients -
        '------------------------------------------------------------

        Dim chrDay As Char = ""
        Dim intLimit As Integer
        Dim intCount As Integer

        'First, get day letter
        Select Case thisDate.DayOfWeek.ToString()
            Case "Monday"
                chrDay = "M"
            Case "Tuesday"
                chrDay = "T"
            Case "Wednesday"
                chrDay = "W"
            Case "Thursday"
                chrDay = "R"
            Case "Friday"
                chrDay = "F"
        End Select

        Dim strSQL As String = "SELECT PatientCount FROM Availability WHERE Day = '" & chrDay & "' AND DoctorTUID = " & intDTUID

        'Get patient limit for our doc
        Dim dataReader = ExecuteSQLReader(strSQL)
        While dataReader.Read()
            intLimit = dataReader("PatientCount")
        End While
        dataReader.Close()
        DisconnectFromSQL()

        'Get current patients assigned to our doc on this day
        strSQL = "SELECT COUNT(*) FROM Appointments WHERE Day = '" & thisDate & "' AND DoctorTUID = " & intDTUID & " AND PatientTUID IS NOT NULL"

        dataReader = ExecuteSQLReader(strSQL)
        While dataReader.Read()
            intCount = dataReader("")
        End While
        dataReader.Close()
        DisconnectFromSQL()

        'Check if the doc has reached his patient limit
        If intCount >= intLimit Then
            'If yes
            Return False
        Else
            'If no
            Return True
        End If
    End Function
End Class