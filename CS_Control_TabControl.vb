Public Class CS_Control_TabControl
    Inherits System.Windows.Forms.TabControl

    Private mTabPagesHidden As New Dictionary(Of String, System.Windows.Forms.TabPage)
    Private mTabPagesOrder As List(Of String)

    Public Sub HideTabPageByName(ByVal TabPageName As String)
        If mTabPagesOrder Is Nothing Then
            ' The first time the Hide method is called, save the original order of the TabPages
            mTabPagesOrder = New List(Of String)
            For Each TabPageCurrent As TabPage In Me.TabPages
                mTabPagesOrder.Add(TabPageCurrent.Name)
            Next
        End If

        If Me.TabPages.ContainsKey(TabPageName) Then
            Dim TabPageToHide As TabPage

            ' Get the TabPage object
            TabPageToHide = TabPages(TabPageName)
            ' Add the TabPage to the internal List
            mTabPagesHidden.Add(TabPageName, TabPageToHide)
            ' Remove the TabPage from the TabPages collection of the TabControl
            Me.TabPages.Remove(TabPageToHide)
        End If
    End Sub

    Public Sub ShowTabPageByName(ByVal TabPageName As String)
        If mTabPagesHidden.ContainsKey(TabPageName) Then
            Dim TabPageToShow As TabPage

            ' Get the TabPage object
            TabPageToShow = mTabPagesHidden(TabPageName)
            ' Add the TabPage to the TabPages collection of the TabControl
            Me.TabPages.Insert(GetTabPageInsertionPoint(TabPageName), TabPageToShow)
            ' Remove the TabPage from the internal List
            mTabPagesHidden.Remove(TabPageName)
        End If
    End Sub

    Private Function GetTabPageInsertionPoint(ByVal TabPageName As String) As Integer
        Dim TabPageIndex As Integer
        Dim TabPageCurrent As TabPage
        Dim TabNameIndex As Integer
        Dim TabNameCurrent As String

        For TabPageIndex = 0 To Me.TabPages.Count - 1
            TabPageCurrent = Me.TabPages(TabPageIndex)
            For TabNameIndex = TabPageIndex To mTabPagesOrder.Count - 1
                TabNameCurrent = mTabPagesOrder(TabNameIndex)
                If TabNameCurrent = TabPageCurrent.Name Then
                    Exit For
                End If
                If TabNameCurrent = TabPageName Then
                    Return TabPageIndex
                End If
            Next
        Next
        Return TabPageIndex
    End Function

    Protected Overrides Sub Dispose(Disposing As Boolean)
        If Not Disposing Then
            mTabPagesHidden = Nothing
            mTabPagesOrder = Nothing
            Dispose(False)
        End If
    End Sub

End Class