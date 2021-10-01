Imports System
Imports System.Collections

Namespace BusinessLogicLayer

    Public Class FollowMeSubClassCollection
        Inherits ArrayList
        Public Enum SubClassFields
            InitValue
            ClassID
            SubClassID
            SubClassTime
            NumberOfStudents
            RemainingNumber
            DeadLine

        End Enum


        Public Overloads Sub Sort(ByVal sortField As SubClassFields, ByVal isAscending As Boolean)
            Select Case sortField
                Case SubClassFields.ClassID
                    MyBase.Sort(New ClassIDComparer)
                Case SubClassFields.DeadLine
                    MyBase.Sort(New DeadLineComparer)
                Case SubClassFields.SubClassTime
                    MyBase.Sort(New SubClassTimeeComparer)

            End Select

            If Not isAscending Then
                MyBase.Reverse() ' 反向整個 System.Collections.ArrayList 中元素的順序。  
            End If
        End Sub 'Sort

        Private NotInheritable Class ClassIDComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As FollowMeSubClass = CType(x, FollowMeSubClass)
                Dim second As FollowMeSubClass = CType(y, FollowMeSubClass)
                Return first.ClassID.CompareTo(second.ClassID)
            End Function 'Compare
        End Class 'ClassIDComparer

        Private NotInheritable Class DeadLineComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As FollowMeSubClass = CType(x, FollowMeSubClass)
                Dim second As FollowMeSubClass = CType(y, FollowMeSubClass)
                Return first.SubClassDeadLine.CompareTo(second.SubClassDeadLine)
            End Function 'Compare
        End Class 'ClassIDComparer

        Private NotInheritable Class SubClassTimeeComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As FollowMeSubClass = CType(x, FollowMeSubClass)
                Dim second As FollowMeSubClass = CType(y, FollowMeSubClass)
                Return first.SubClassTime.CompareTo(second.SubClassTime)
            End Function 'Compare
        End Class 'ClassIDComparer


    End Class

End Namespace

