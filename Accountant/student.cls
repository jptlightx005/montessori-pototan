VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "student"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public studentID As Integer
Public queueID As Integer
Public isNew As Boolean
Public grade As String
Public firstName As String
Public middleName As String
Public lastName As String
Public gender As String
Public dateofBirth As Date
Public placeOfBirth As String
Public fathersName As String
Public fatherOccupation As String
Public mothersName As String
Public motherOccupation As String
Public homeAddress As String
Public homeNumber As String
Public guardianName As String
Public guardianRelation As String
Public guardianAddress As String
Public guardianNumnber As String
Public lastSchoolAttended As String
Public religion As String
Public isBaptized As Boolean
Public firstCommunion As Boolean


Public balancePaid As Double
Public datePaid As Date


Public Function fullName() As String
    fullName = firstName & " " & Left$(CStr(middleName), 1) & " " & lastName
End Function

Public Function GradeText() As String

End Function

