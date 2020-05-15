Imports System.Data.SqlClient


Public Class Main

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load

        Using Synergy As New SqlConnection(
                "Data Source=SYNERGY;" _
                & "Initial Catalog=SynergyOne;" _
                & "User ID = <USER_ID>;Password = <PASSWORD>")

            Using CommandSTU As New SqlCommand(
                String.Format("SELECT DISTINCT 
                        Community.ID, 
                        Community.Surname, 
                        Community.Preferred 
                    FROM SubjectClasses 
                    INNER JOIN StudentClasses 
                        ON SubjectClasses.FileYear = StudentClasses.FileYear 
                        AND SubjectClasses.FileSemester = StudentClasses.FileSemester 
                        AND SubjectClasses.ClassCode = StudentClasses.ClassCode 
                    INNER JOIN Community 
                        ON StudentClasses.ID = Community.ID 
                    WHERE StudentClasses.FileYear = {0} 
                        AND StudentClasses.FileSemester = {1}
                    ORDER BY Community.Surname",
                   2020, 1),
                Synergy)


                Synergy.Open()

                Using ReaderSTU As SqlDataReader = CommandSTU.ExecuteReader()
                    If ReaderSTU.HasRows Then
                        Do While ReaderSTU.Read()
                            ' Add students to the comboBox as Objects, not a list of strings. 
                            ' I had to modify the MatchingComboBox code to allow this. 
                            MatchingComboBox1.Items.Add(New Student(
                                ReaderSTU("Preferred").ToString,
                                ReaderSTU("Surname").ToString,
                                ReaderSTU("ID").ToString))
                        Loop
                    End If

                    ReaderSTU.Close()
                    Synergy.Close()

                End Using 'ReaderSTU

            End Using 'CommandSTU

        End Using 'Synergy

    End Sub

End Class
