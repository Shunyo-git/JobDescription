Imports System.Web
Imports System.Web.SessionState

Namespace Web

    Public Class Global
        Inherits System.Web.HttpApplication

#Region " ����]�p�u�㲣�ͪ��{���X "

        Public Sub New()
            MyBase.New()

            '��������]�p�u��һݪ��I�s�C
            InitializeComponent()

            '�b InitializeComponent() �I�s����[�J�Ҧ�����l�]�w

        End Sub

        '������]�p�u�㪺���n��
        Private components As System.ComponentModel.IContainer

        '�`�N: �H�U������]�p�u��һݪ��{��
        '�z�i�H�ϥΤ���]�p�u��i��ק�A
        '�ФŨϥε{���X�s�边�i��ק�C
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
            components = New System.ComponentModel.Container
        End Sub

#End Region

        Public Const CfgKeyHsinYiJDConnString As String = "HsinYiJDConnectionString"

        Public Const CfgKeySMTPServer As String = "SMTPServer"

        Public Const CfgKeyAttachmentWebPath As String = "AttachmentWebPath"
        Public Const CfgKeyAttachmentSrvPath As String = "AttachmentSrvPath"

        Public Const CfgKeyUserAcctSource As String = "UserAccountSource"
        Public Const CfgKeyDefaultRole As String = "DefaultRoleForNewUser"
        Public Const CfgKeyFirstDayOfWeek As String = "FirstDayOfWeek"

        Public UserRoles As String = "userroles"
        Public UserRolesMobile As String = "userrolesmobile"

        Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)
            ' �Ұ����ε{���ɤ޵o
        End Sub

        Sub Session_Start(ByVal sender As Object, ByVal e As EventArgs)
            ' �Ұʤu�@���q�ɤ޵o
        End Sub

        Sub Application_BeginRequest(ByVal sender As Object, ByVal e As EventArgs)
            ' ��C�@�ӭn�D�}�l�ɤ޵o
        End Sub

        Sub Application_AuthenticateRequest(ByVal sender As Object, ByVal e As EventArgs)
            ' �������Ҩϥήɤ޵o
        End Sub

        Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)
            ' �o�Ϳ��~�ɤ޵o
        End Sub

        Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)
            ' ��u�@���q�����ɤ޵o
        End Sub

        Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)
            ' �����ε{�������ɤ޵o
        End Sub

    End Class

End Namespace


