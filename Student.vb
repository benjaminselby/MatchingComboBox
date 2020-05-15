Public Class Student
    Public FirstName As String
    Public Surname As String
    Public ID As String
    Public Sub New(fname As String, sname As String, id As String)
        Me.FirstName = fname
        Me.Surname = sname
        Me.ID = id
    End Sub
    Public Overrides Function ToString() As String
        Return "Student with ID = " & ID
    End Function
End Class
