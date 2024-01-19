Imports Newtonsoft.Json
Imports System.IO
Imports System.Diagnostics

Public Class SubjectData
    Public Property Subject As String

    Public Property subjectallowwildcard As Boolean
    Public Property SubjectReplaceTo As Integer
    Public Property MatchBodyText As String
End Class


Public Class SubjectList


    Public Subjects As List(Of SubjectData)


    Public Sub New()

    End Sub


    Public Sub New(ByVal Location)
        Dim jsonLocation As String = Location
       Write("LiveLocation:" & jsonLocation)


        Subjects = New List(Of SubjectData)

        ' Deserialize JSON data and add to the Subjects list
        Dim json As String = File.ReadAllText(jsonLocation)
        Dim subjectListData As List(Of SubjectData) = JsonConvert.DeserializeObject(Of List(Of SubjectData))(json)

        If subjectListData IsNot Nothing Then
            Subjects.AddRange(subjectListData)
        End If


    End Sub


    Function swapBodyTexttoSubject(ByVal emailSubject As String, ByVal emailBodyText As String)

        '' search bo

    End Function



End Class

' MY json file example
'[
'  {
'    "subject": "1010199",
'    "subjectallowwildcard" "false",
'    "subjectReplaceTo": 10199,
'    "matchbodytext": ""
'  },
'  {
'    "subject": "2010199 Doe",
'    "subjectallowwildcard" "false",
'    "subjectReplaceTo": 10199,
'    "matchbodytext": ""
'  }
']