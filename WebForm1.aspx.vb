

Imports System.DirectoryServices


Public Class WebForm1
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    '�`�N: �U�C�w�d��m�ŧi�O Web Form �]�p�u��ݭn�����ءC
    '�ФŧR���β��ʥ��C
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: ���� Web Form �]�p�u��һݪ���k�I�s
        '�ФŨϥε{���X�s�边�i��ק�C
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '�b�o�̩�m�ϥΪ̵{���X�H��l�ƺ���

        Dim myADSPath As String = "WinNT://HsinYI"
        Dim de As DirectoryEntry
        de = New DirectoryEntry(myADSPath)
        de.Username = "sean"
        de.Password = "2idulggr"
        Dim c As DirectoryEntry
        For Each c In de.Children
            Response.Write(c.Path & " :  ")
            Response.Write(c.Username & " :  ")
            Response.Write(c.Name & " <BR>  ")

        Next


    End Sub

End Class