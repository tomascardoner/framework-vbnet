Namespace CardonerSistemas.Controls

	Friend Class TabControlExtension

#Region "Declarations"

		Private ReadOnly _TabPagesOrder As List(Of String) = New List(Of String)
		Private ReadOnly _TabPagesHidden As List(Of TabPage) = New List(Of TabPage)

#End Region

		Public Sub New(tabControl As System.Windows.Forms.TabControl)
			' This Is necesary because the TabControl have a bug
			Debug.Print(tabControl.Handle.ToString)

			For Each tabPageCurrent As TabPage In tabControl.TabPages
				_TabPagesOrder.Add(tabPageCurrent.Name)
			Next
		End Sub

		Public Sub HidePage(tabControl As TabControl, tabPage As TabPage)
			If tabControl.TabPages.Contains(tabPage) Then
				_TabPagesHidden.Add(tabPage)
				tabControl.TabPages.Remove(tabPage)
			End If
		End Sub

		Public Sub ShowPage(tabControl As TabControl, tabPage As TabPage)
			If _TabPagesHidden.Contains(tabPage) Then
				tabControl.TabPages.Insert(GetTabPageInsertionPoint(tabControl, tabPage.Name), tabPage)
				_TabPagesHidden.Remove(tabPage)
			End If
		End Sub

		Public Sub PageVisible(tabControl As TabControl, tabPage As TabPage, value As Boolean)
			If value Then
				ShowPage(tabControl, tabPage)
			Else
				HidePage(tabControl, tabPage)
			End If
		End Sub

		Private Function GetTabPageInsertionPoint(tabControl As TabControl, tabPageName As String) As Integer
			Dim tabPageIndex As Integer
			Dim tabPageCurrent As TabPage
			Dim tabNameIndex As Integer
			Dim tabNameCurrent As String

			For tabPageIndex = 0 To tabControl.TabPages.Count - 1
				tabPageCurrent = tabControl.TabPages(tabPageIndex)
				For tabNameIndex = tabPageIndex To _TabPagesOrder.Count - 1
					tabNameCurrent = _TabPagesOrder(tabNameIndex)
					If tabNameCurrent = tabPageCurrent.Name Then
						Exit For
					End If
					If tabNameCurrent = tabPageName Then
						Return tabPageIndex
					End If
				Next
			Next
			Return TabPageIndex
		End Function
	End Class
End Namespace