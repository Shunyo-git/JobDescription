Imports System
Imports System.Collections

Namespace BusinessLogicLayer

    '*********************************************************************
    '
    ' EventsCollection Class 
    '
    '  This is a skeleton class for creating sortable collection.
    '
    '*********************************************************************

    Public Class FollowMeClassesCollection
        Inherits ArrayList
        Public Enum ClassFields
            InitValue
            ClassID
            isWebDisplay
            DisplayStarting
            DisplayEnding
            CreatedDate
            ModifiedDate
        End Enum 'EventFields

        Public Overloads Sub Sort(ByVal sortField As ClassFields, ByVal isAscending As Boolean)
            Select Case sortField
                Case ClassFields.ClassID
                    MyBase.Sort(New ClassIDComparer)
                Case ClassFields.isWebDisplay
                    MyBase.Sort(New isWebDisplayComparer)
                Case ClassFields.DisplayStarting
                    MyBase.Sort(New DisplayStartingComparer)
                Case ClassFields.DisplayEnding
                    MyBase.Sort(New DisplayEndingComparer)
                Case ClassFields.CreatedDate
                    MyBase.Sort(New CreatedDateComparer)
                Case ClassFields.ModifiedDate
                    MyBase.Sort(New ModifiedDateComparer)

            End Select

            If Not isAscending Then
                MyBase.Reverse() ' 反向整個 System.Collections.ArrayList 中元素的順序。  
            End If
        End Sub 'Sort

        Private NotInheritable Class ClassIDComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As FollowMeClass = CType(x, FollowMeClass)
                Dim second As FollowMeClass = CType(y, FollowMeClass)
                Return first.ClassID.CompareTo(second.ClassID)
            End Function 'Compare
        End Class 'ClassIDComparer

        Private NotInheritable Class isWebDisplayComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As FollowMeClass = CType(x, FollowMeClass)
                Dim second As FollowMeClass = CType(y, FollowMeClass)
                Return first.IsWebDisplay.CompareTo(second.IsWebDisplay)
            End Function 'Compare
        End Class 'isWebDisplayComparer

        Private NotInheritable Class DisplayStartingComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As FollowMeClass = CType(x, FollowMeClass)
                Dim second As FollowMeClass = CType(y, FollowMeClass)
                Return first.DisplayStarting.CompareTo(second.DisplayStarting)
            End Function 'Compare
        End Class 'DisplayStartingComparer

        Private NotInheritable Class DisplayEndingComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As FollowMeClass = CType(x, FollowMeClass)
                Dim second As FollowMeClass = CType(y, FollowMeClass)
                Return first.DisplayEnding.CompareTo(second.DisplayEnding)
            End Function 'Compare
        End Class 'DisplayEndingComparer

        Private NotInheritable Class CreatedDateComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As FollowMeClass = CType(x, FollowMeClass)
                Dim second As FollowMeClass = CType(y, FollowMeClass)
                Return first.CreatedDate.CompareTo(second.CreatedDate)
            End Function 'Compare
        End Class 'CreatedDateComparer



        Private NotInheritable Class ModifiedDateComparer
            Implements IComparer '公開比較二個物件的方法。  

            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
                Dim first As FollowMeClass = CType(x, FollowMeClass)
                Dim second As FollowMeClass = CType(y, FollowMeClass)
                Return first.ModifiedDate.CompareTo(second.ModifiedDate)
            End Function 'Compare
        End Class 'ModifiedDateComparer
    End Class
End Namespace


