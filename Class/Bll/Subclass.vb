Imports HsinYi.EIP.JobDescription.DataAccessLayer


Namespace BusinessLogicLayer

    Public Class FollowMeSubClass
        Private _classID As Integer
        Private _subClassID As Integer
        Private _subClassTime As String
        Private _subClassTopic As String
        Private _subClassLecturer As String
        Private _subClassProductID As String
        Private _subClassNumberOfStudents As Short
        Private _subClassRemainingNumber As Short
        Private _subClassDeadLine As Date


#Region "  ÄÝ©Ê¦s¨ú "

        Public Property ClassID() As Integer
            Get
                Return _classID
            End Get
            Set(ByVal Value As Integer)
                _classID = Value
            End Set
        End Property

        Public Property SubClassID() As Integer
            Get
                Return _subClassID
            End Get
            Set(ByVal Value As Integer)
                _subClassID = Value
            End Set
        End Property

        Public Property SubClassTime() As String
            Get
                Return _subClassTime
            End Get
            Set(ByVal Value As String)
                _subClassTime = Value
            End Set
        End Property

        Public Property SubClassTopic() As String
            Get
                Return _subClassTopic
            End Get
            Set(ByVal Value As String)
                _subClassTopic = Value
            End Set
        End Property

        Public Property SubClassLecturer() As String
            Get
                Return _subClassLecturer
            End Get
            Set(ByVal Value As String)
                _subClassLecturer = Value
            End Set
        End Property

        Public Property SubClassProductID() As String
            Get
                Return _subClassProductID
            End Get
            Set(ByVal Value As String)
                _subClassProductID = Value
            End Set
        End Property

        Public Property SubClassNumberOfStudents() As Integer
            Get
                Return _subClassNumberOfStudents
            End Get
            Set(ByVal Value As Integer)
                _subClassNumberOfStudents = Value
            End Set
        End Property

        Public Property SubClassRemainingNumber() As Integer
            Get
                Return _subClassRemainingNumber
            End Get
            Set(ByVal Value As Integer)
                _subClassRemainingNumber = Value
            End Set
        End Property

        Public Property SubClassDeadLine() As Date
            Get
                Return _subClassDeadLine
            End Get
            Set(ByVal Value As Date)
                _subClassDeadLine = Value
            End Set
        End Property
#End Region


        Public Sub New()
        End Sub 'New

        Public Sub New(ByVal SubClassID As Integer)
            _subClassID = SubClassID
        End Sub 'New



        Public Function Load() As Boolean
            ' The Get Projects stored procedure returns a dataset with 3 tables: Projects, Members, and Categories.
            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyHsinYiJDConnString), "v2004_ParentClass_GetSubClasses", _subClassID)

            If ds.Tables(0).Rows.Count < 1 Then
                Return False
            End If

            _classID = Convert.ToInt32(ds.Tables(0).Rows(0)("ClassID"))
            _subClassID = Convert.ToInt32(ds.Tables(0).Rows(0)("SubClassID"))
            _subClassTime = Convert.ToString(ds.Tables(0).Rows(0)("SubClassTime"))
            _subClassTopic = Convert.ToString(ds.Tables(0).Rows(0)("SubClassTopic"))
            _subClassLecturer = Convert.ToString(ds.Tables(0).Rows(0)("SubClassLecturer"))
            _subClassProductID = Convert.ToString(ds.Tables(0).Rows(0)("SubClassProductID"))
            _subClassNumberOfStudents = Convert.ToInt16(ds.Tables(0).Rows(0)("SubClassNumberOfStudents"))
            _subClassRemainingNumber = Convert.ToInt16(ds.Tables(0).Rows(0)("SubClassRemainingNumber"))
            _subClassDeadLine = Convert.ToDateTime(ds.Tables(0).Rows(0)("SubClassDeadLine"))


            Return True
        End Function 'Load

        Public Overloads Shared Function GetSubClasss(ByVal ClassID As Integer) As FollowMeSubClassCollection
            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyHsinYiJDConnString), "v2004_ParentClass_GetSubClassesByClassID", ClassID)

            ' Separate Data into a Collection of projects
            Dim SubClasses As New FollowMeSubClassCollection
            If ds.Tables.Count > 0 Then
                Dim r As DataRow
                For Each r In ds.Tables(0).Rows
                    Dim SubClass As New FollowMeSubClass
                    SubClass.ClassID = Convert.ToInt32(r("ClassID"))
                    SubClass.SubClassID = Convert.ToInt32(r("SubClassID"))
                    SubClass.SubClassDeadLine = Convert.ToDateTime(r("SubClassDeadLine"))
                    SubClass.SubClassLecturer = Convert.ToString(r("SubClassLecturer"))
                    SubClass.SubClassNumberOfStudents = Convert.ToInt16(r("SubClassNumberOfStudents"))
                    SubClass.SubClassRemainingNumber = Convert.ToInt16(r("SubClassRemainingNumber"))
                    SubClass.SubClassTime = Convert.ToString(r("SubClassTime"))
                    SubClass.SubClassTopic = Convert.ToString(r("SubClassTopic"))
                    SubClass.SubClassProductID = Convert.ToString(r("SubClassProductID"))


                    SubClasses.Add(SubClass)
                Next r

            End If
            Return SubClasses
        End Function 'GetExtendReadBooks

        '*********************************************************************
        '
        ' Remove static method
        '
        '*********************************************************************

        Public Shared Sub Remove(ByVal subClassID As Integer)
            SqlHelper.ExecuteNonQuery(ConfigurationSettings.AppSettings(Web.Global.CfgKeyHsinYiJDConnString), "v2004_ParentClass_DeleteSubClass", subClassID)


        End Sub 'Remove
        '*********************************************************************
        '
        ' Calls Insert or Save based on the _subClassID
        '
        '*********************************************************************

        Public Function Save() As Boolean
            If _subClassID = 0 Then
                Return Insert()
            Else
                If _subClassID > 0 Then
                    Return Update()
                Else
                    _subClassID = 0
                    Return False
                End If
            End If
        End Function 'Save
        '*********************************************************************
        '
        ' Inserts a SubClass along with its members and categories into the databse
        '
        '*********************************************************************

        Private Function Insert() As Boolean

            Dim blnResult As Boolean

            Try
                _subClassID = Convert.ToInt32(SqlHelper.ExecuteScalar(ConfigurationSettings.AppSettings(Web.Global.CfgKeyHsinYiJDConnString), "v2004_ParentClass_AddSubClass", _classID, _subClassTime, _subClassTopic, _subClassLecturer, _subClassProductID, _subClassNumberOfStudents, _subClassRemainingNumber, _subClassDeadLine))

                blnResult = True

            Catch ex As Exception
                Throw ex
                blnResult = False
            End Try


            Return blnResult

        End Function 'Insert
        '*********************************************************************
        '
        ' Updates a SubClass along with its members and categories.
        '
        '*********************************************************************

        Private Function Update() As Boolean
            Dim blnResult As Boolean

            Try
                SqlHelper.ExecuteNonQuery(ConfigurationSettings.AppSettings(Web.Global.CfgKeyHsinYiJDConnString), "v2004_ParentClass_UpdateSubClass", _subClassID, _classID, _subClassTime, _subClassTopic, _subClassLecturer, _subClassProductID, _subClassNumberOfStudents, _subClassRemainingNumber, _subClassDeadLine)
                blnResult = True
            Catch ex As Exception
                Throw ex
                blnResult = False
            End Try

            Return blnResult
        End Function 'Update

        'Private Function getExtendReadBooks() As String
        '    Dim selectedBooks As New StringBuilder(_ExtendReadBooks.Count)
        '    Dim index As Integer = 1
        '    Dim ebook As ExtendReadBook
        '    For Each ebook In _ExtendReadBooks
        '        selectedBooks.Append(ebook.ProductID)

        '        ' Don't append separator if this is the last item
        '        If index <> _ExtendReadBooks.Count Then
        '            selectedBooks.Append(",")
        '        End If
        '        index += 1
        '    Next ebook
        '    Return selectedBooks.ToString()
        'End Function
    End Class

End Namespace

