Imports System
Imports System.Data
Imports System.Text
Imports System.Configuration
Imports HsinYi.EIP.JobDescription.DataAccessLayer

Namespace BusinessLogicLayer

    '****************************************************************************
    '
    ' FollowMeClass Class
    '
    ' The FollowMeClass class represents a list of FollowMe Class.
    '
    '****************************************************************************

    Public Class FollowMeClass
        Private _classID As Integer
        Private _Subject As String
        Private _intro As String
        Private _descTime As String
        Private _descTarget As String
        Private _descCharge As String
        Private _descLocation As String
        Private _descNotice As String
        Private _groupProductID As String

        Private _mainImage As String
        Private _subImage As String

        Private _createdUserID As Integer
        Private _createdDate As Date
        Private _modifiedUserID As Integer
        Private _modifiedDate As Date
        Private _isDel As Boolean

        Private _displayStarting As Date
        Private _displayEnding As Date
        Private _isWebDisplay As Boolean

        'Private _ExtendReadBooks As ExtendReadBookCollection
        Private _SubClassess As FollowMeSubClassCollection

#Region "  ÄÝ©Ê¦s¨ú "
        Public Property ClassID() As Integer
            Get
                Return _classID
            End Get
            Set(ByVal Value As Integer)
                _classID = Value
            End Set
        End Property
        Public Property MainImage() As String
            Get
                Return _mainImage
            End Get
            Set(ByVal Value As String)
                _mainImage = Value
            End Set
        End Property
        Public Property SubImage() As String
            Get
                Return _subImage
            End Get
            Set(ByVal Value As String)
                _subImage = Value
            End Set
        End Property
        Public Property Subject() As String
            Get
                Return _Subject
            End Get
            Set(ByVal Value As String)
                _Subject = Value
            End Set
        End Property
        Public Property Intro() As String
            Get
                Return _intro

            End Get
            Set(ByVal Value As String)
                _intro = Value
            End Set
        End Property

        Public Property DescTime() As String
            Get
                Return _descTime

            End Get
            Set(ByVal Value As String)
                _descTime = Value
            End Set
        End Property
        Public Property DescTarget() As String
            Get
                Return _descTarget

            End Get
            Set(ByVal Value As String)
                _descTarget = Value
            End Set
        End Property
        Public Property DescCharge() As String
            Get
                Return _descCharge

            End Get
            Set(ByVal Value As String)
                _descCharge = Value
            End Set
        End Property
        Public Property DescLocation() As String
            Get
                Return _descLocation

            End Get
            Set(ByVal Value As String)
                _descLocation = Value
            End Set
        End Property
        Public Property DescNotice() As String
            Get
                Return _descNotice
            End Get

            Set(ByVal Value As String)
                _descNotice = Value
            End Set
        End Property

        Public Property GroupProductID() As String
            Get
                Return _groupProductID
            End Get

            Set(ByVal Value As String)
                _groupProductID = Value
            End Set
        End Property

        Public Property DisplayEnding() As Date
            Get
                Return _displayEnding
            End Get

            Set(ByVal Value As Date)
                _displayEnding = Value
            End Set
        End Property
        Public Property DisplayStarting() As Date
            Get
                Return _displayStarting
            End Get

            Set(ByVal Value As Date)
                _displayStarting = Value
            End Set
        End Property
        Public Property IsWebDisplay() As Boolean
            Get
                Return _isWebDisplay
            End Get

            Set(ByVal Value As Boolean)
                _isWebDisplay = Value
            End Set
        End Property
        'Public Property ExtendReadBooks() As ExtendReadBookCollection
        '    Get
        '        Return _ExtendReadBooks
        '    End Get
        '    Set(ByVal Value As ExtendReadBookCollection)
        '        _ExtendReadBooks = Value
        '    End Set
        'End Property
        Public Property FollowMeSubClasses() As FollowMeSubClassCollection
            Get
                Return _SubClassess
            End Get
            Set(ByVal Value As FollowMeSubClassCollection)
                _SubClassess = Value
            End Set
        End Property
        Public Property CreatedUserID() As Integer
            Get
                Return _createdUserID
            End Get
            Set(ByVal Value As Integer)
                _createdUserID = Value
            End Set
        End Property


        Public Property ModifiedUserID() As Integer
            Get
                Return _modifiedUserID
            End Get
            Set(ByVal Value As Integer)
                _modifiedUserID = Value
            End Set
        End Property

        Public Property CreatedDate() As DateTime
            Get
                Return _createdDate
            End Get

            Set(ByVal Value As DateTime)
                _createdDate = Value
            End Set
        End Property
        Public Property ModifiedDate() As DateTime
            Get
                Return _modifiedDate
            End Get
            Set(ByVal Value As DateTime)
                _modifiedDate = Value
            End Set
        End Property
        Public Property isDel() As Boolean
            Get
                Return _isDel
            End Get
            Set(ByVal Value As Boolean)
                _isDel = Value
            End Set
        End Property

#End Region

        Public Sub New()
        End Sub 'New

        Public Sub New(ByVal ClassID As Integer)
            _classID = ClassID
        End Sub 'New

        Public Sub New(ByVal ClassID As Integer, ByVal Subject As String, ByVal Intro As String, ByVal GroupProductID As String)
            _classID = ClassID
            _Subject = Subject
            _intro = Intro
            _groupProductID = GroupProductID
        End Sub 'New

        '*********************************************************************
        '
        ' Retrieves the list of all the projects
        '
        '*********************************************************************
        Public Overloads Shared Function GetClasses() As FollowMeClassesCollection
            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyHsinYiJDConnString), CommandType.StoredProcedure, "v2004_ParentClass_ListAllClasses")

            ' Separate Data into a Collection of projects
            Dim Classes As New FollowMeClassesCollection
            If ds.Tables.Count > 0 Then
                Dim r As DataRow
                For Each r In ds.Tables(0).Rows
                    Dim PClass As New FollowMeClass
                    PClass.ClassID = Convert.ToInt32(r("ClassID"))
                    PClass.DescCharge = Convert.ToString(r("DescCharge"))
                    PClass.DescLocation = Convert.ToString(r("DescLocation"))
                    PClass.DescNotice = Convert.ToString(r("DescNotice"))
                    PClass.DescTarget = Convert.ToString(r("DescTarget"))
                    PClass.DescTime = Convert.ToString(r("DescTime"))
                    'PClass.ExtendReadBooks = 
                    '
                    PClass.GroupProductID = Convert.ToString(r("GroupProductID"))
                    PClass.Intro = Convert.ToString(r("Intro"))
                    PClass.Subject = Convert.ToString(r("Subject"))

                    PClass.MainImage = Convert.ToString(r("MainImage"))
                    PClass.SubImage = Convert.ToString(r("SubImage"))


                    PClass.DisplayStarting = Convert.ToDateTime(r("DisplayStarting"))
                    PClass.DisplayEnding = Convert.ToDateTime(r("DisplayEnding"))
                    PClass.IsWebDisplay = Convert.ToBoolean(r("IsWebDisplay"))

                    PClass.CreatedDate = Convert.ToDateTime(r("CreatedDate"))
                    PClass.ModifiedDate = Convert.ToDateTime(r("ModifiedDate"))

                    Dim dv As DataView = New DataView(ds.Tables(1), "ClassID = " & Convert.ToString(r("ClassID")), "SubClassID", DataViewRowState.None)

                    Dim row As DataRowView
                    Dim i As Integer

                    Dim SubClasses = New FollowMeSubClassCollection
                    For Each row In dv
                        Dim SubClass As New FollowMeSubClass
                        SubClass.ClassID = Convert.ToInt32(row("ClassID"))
                        SubClass.SubClassID = Convert.ToInt32(row("SubClassID"))
                        SubClass.SubClassDeadLine = Convert.ToDateTime(row("SubClassDeadLine"))
                        SubClass.SubClassLecturer = Convert.ToString(row("SubClassLecturer"))
                        SubClass.SubClassNumberOfStudents = Convert.ToString(row("SubClassNumberOfStudents"))
                        SubClass.SubClassProductID = Convert.ToString(row("SubClassProductID"))
                        SubClass.SubClassRemainingNumber = Convert.ToString(row("SubClassRemainingNumber"))
                        SubClass.SubClassTime = Convert.ToString(row("SubClassTime"))
                        SubClass.SubClassTopic = Convert.ToString(row("SubClassTopic"))
                        SubClasses.Add(SubClass)
                    Next
                    PClass.FollowMeSubClasses = SubClasses

                    dv = New DataView(ds.Tables(2), "ClassID = " & Convert.ToString(r("ClassID")), "ProductName", DataViewRowState.None)
                    dv.RowFilter = "ClassID = " & Convert.ToString(r("ClassID"))
                    'Dim ExtendReadBooks = New ExtendReadBookCollection
                    'For Each row In dv
                    '    Dim ExtendReadBook As New ExtendReadBook
                    '    ExtendReadBook.ClassID = Convert.ToInt32(row("ClassID"))
                    '    ExtendReadBook.ProductID = Convert.ToInt32(row("ProductID"))
                    '    ExtendReadBook.ProductName = Convert.ToString(row("ProductName"))
                    '    ExtendReadBook.ProductSellPrice = Convert.ToInt32(row("ProductSellPrice"))
                    '    ExtendReadBooks.Add(ExtendReadBook)
                    'Next
                    'PClass.ExtendReadBooks = ExtendReadBooks




                    Classes.Add(PClass)
                Next r

            End If
            Return Classes
        End Function 'GetEvents


        '*********************************************************************
        '
        ' Populates the object with the Event info from the database.
        '
        '*********************************************************************

        Public Function Load() As Boolean
            ' The Get Projects stored procedure returns a dataset with 3 tables: Projects, Members, and Categories.
            Dim ds As DataSet = SqlHelper.ExecuteDataset(ConfigurationSettings.AppSettings(Web.Global.CfgKeyHsinYiJDConnString), "v2004_ParentClass_GetClass", _classID)

            If ds.Tables(0).Rows.Count < 1 Then
                Return False
            End If



            _classID = Convert.ToInt32(ds.Tables(0).Rows(0)("ClassID"))
            _Subject = Convert.ToString(ds.Tables(0).Rows(0)("Subject"))
            _intro = Convert.ToString(ds.Tables(0).Rows(0)("Intro"))
            _descTime = Convert.ToString(ds.Tables(0).Rows(0)("descTime"))
            _descTarget = Convert.ToString(ds.Tables(0).Rows(0)("descTarget"))
            _descCharge = Convert.ToString(ds.Tables(0).Rows(0)("descCharge"))
            _descLocation = Convert.ToString(ds.Tables(0).Rows(0)("descLocation"))
            _descNotice = Convert.ToString(ds.Tables(0).Rows(0)("descNotice"))

            _groupProductID = Convert.ToString(ds.Tables(0).Rows(0)("groupProductID"))

            _displayStarting = Convert.ToDateTime(ds.Tables(0).Rows(0)("DisplayStarting"))
            _displayEnding = Convert.ToDateTime(ds.Tables(0).Rows(0)("DisplayEnding"))
            _isWebDisplay = Convert.ToBoolean(ds.Tables(0).Rows(0)("IsWebDisplay"))

            _mainImage = Convert.ToString(ds.Tables(0).Rows(0)("MainImage"))
            _subImage = Convert.ToString(ds.Tables(0).Rows(0)("SubImage"))


            _createdUserID = Convert.ToInt32(ds.Tables(0).Rows(0)("CreatedUserID"))
            _createdDate = Convert.ToDateTime(ds.Tables(0).Rows(0)("CreatedDate"))
            _modifiedUserID = Convert.ToInt32(ds.Tables(0).Rows(0)("ModifiedUserID"))
            _modifiedDate = Convert.ToDateTime(ds.Tables(0).Rows(0)("ModifiedDate"))


            Dim row As DataRow

            _SubClassess = New FollowMeSubClassCollection
            For Each row In ds.Tables(1).Rows
                Dim subClass As New FollowMeSubClass
                subClass.ClassID = Convert.ToInt32(row("ClassID"))
                subClass.SubClassID = Convert.ToInt32(row("SubClassID"))
                subClass.SubClassDeadLine = Convert.ToDateTime(row("SubClassDeadLine"))
                subClass.SubClassLecturer = Convert.ToString(row("SubClassLecturer"))
                subClass.SubClassNumberOfStudents = Convert.ToString(row("SubClassNumberOfStudents"))
                subClass.SubClassProductID = Convert.ToString(row("SubClassProductID"))
                subClass.SubClassRemainingNumber = Convert.ToString(row("SubClassRemainingNumber"))
                subClass.SubClassTime = Convert.ToString(row("SubClassTime"))
                subClass.SubClassTopic = Convert.ToString(row("SubClassTopic"))
                _SubClassess.Add(subClass)
            Next row

            '_ExtendReadBooks = New ExtendReadBookCollection

            'For Each row In ds.Tables(2).Rows
            '    Dim ExBook As New ExtendReadBook
            '    ExBook.ClassID = Convert.ToInt32(row("ClassID"))
            '    ExBook.ProductID = Convert.ToString(row("ProductID"))
            '    ExBook.ProductName = Convert.ToString(row("ProductName"))
            '    ExBook.ProductSellPrice = Convert.ToInt32(row("ProductSellPrice"))
            '    _ExtendReadBooks.Add(ExBook)
            'Next row

            Return True
        End Function 'Load

        '*********************************************************************
        '
        ' Remove static method
        '
        '*********************************************************************

        Public Shared Sub Remove(ByVal classID As Integer)
            SqlHelper.ExecuteNonQuery(ConfigurationSettings.AppSettings(Web.Global.CfgKeyHsinYiJDConnString), "v2004_ParentClass_DeleteClass", classID)

        End Sub 'Remove
        '*********************************************************************
        '
        ' Calls Insert or Save based on the _classID
        '
        '*********************************************************************

        Public Function Save() As Boolean
            If _classID = 0 Then
                Return Insert()
            Else
                If _classID > 0 Then
                    Return Update()
                Else
                    _classID = 0
                    Return False
                End If
            End If
        End Function 'Save
        '*********************************************************************
        '
        ' Inserts a Class along with its members and categories into the databse
        '
        '*********************************************************************

        Private Function Insert() As Boolean

            Dim blnResult As Boolean

            Try
                _classID = Convert.ToInt32(SqlHelper.ExecuteScalar(ConfigurationSettings.AppSettings(Web.Global.CfgKeyHsinYiJDConnString), "v2004_ParentClass_AddClass", _Subject, _intro, _descTime, _descTarget, _descCharge, _descLocation, _descNotice, _groupProductID, _displayStarting, _displayEnding, _isWebDisplay, _mainImage, _subImage, _createdUserID, _createdDate, _modifiedUserID, _modifiedDate))

                blnResult = True

            Catch ex As Exception
                Throw ex
                blnResult = False
            End Try


            Return blnResult

        End Function 'Insert
        '*********************************************************************
        '
        ' Updates a Class along with its members and categories.
        '
        '*********************************************************************

        Private Function Update() As Boolean
            Dim blnResult As Boolean

            Try
                SqlHelper.ExecuteNonQuery(ConfigurationSettings.AppSettings(Web.Global.CfgKeyHsinYiJDConnString), "v2004_ParentClass_UpdateClass", _classID, _Subject, _intro, _descTime, _descTarget, _descCharge, _descLocation, _descNotice, _groupProductID, _displayStarting, _displayEnding, _isWebDisplay, _mainImage, _subImage, _createdUserID, _createdDate, _modifiedUserID, _modifiedDate)
                blnResult = True
            Catch ex As Exception
                Throw ex
                blnResult = False
            End Try

            Return blnResult
        End Function 'Update

        Private Function getExtendReadBooks() As String
            'If _ExtendReadBooks Is Nothing Then
            '    Return String.Empty
            'End If
            'Dim selectedBooks As New StringBuilder(_ExtendReadBooks.Count)
            'Dim index As Integer = 1
            'Dim ebook As ExtendReadBook
            'For Each ebook In _ExtendReadBooks
            '    selectedBooks.Append(ebook.ProductID)

            '    ' Don't append separator if this is the last item
            '    If index <> _ExtendReadBooks.Count Then
            '        selectedBooks.Append(",")
            '    End If
            '    index += 1
            'Next ebook
            'Return selectedBooks.ToString()
        End Function

    End Class

End Namespace


