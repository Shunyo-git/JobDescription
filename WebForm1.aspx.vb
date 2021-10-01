

Imports System.DirectoryServices


Public Class WebForm1
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁

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
