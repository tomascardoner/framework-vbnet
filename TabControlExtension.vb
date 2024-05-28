Namespace CardonerSistemas
    Friend Class TabControlExtension

#Region "Declarations"

        Private ReadOnly tabPagesOrder As List(Of String)
        Private ReadOnly tabPagesHidden As List(Of TabPage)
        Private ReadOnly tabControl As TabControl

#End Region

        Public Sub New(theControl As TabControl)
            ' Esto es necesario porque el TabControl tiene un bug que no inserta las TabPages si no se creó el handle
#Disable Warning S1481 ' Unused local variables should be removed
            Dim handle As IntPtr = theControl.Handle
#Enable Warning S1481 ' Unused local variables should be removed
            tabControl = theControl

            tabPagesOrder = New List(Of String)
            tabPagesHidden = New List(Of TabPage)
            For Each tabPageCurrent As TabPage In theControl.TabPages
                tabPagesOrder.Add(tabPageCurrent.Name)
            Next
        End Sub

        Friend Sub HidePage(thePage As TabPage)
            If tabControl.TabPages.Contains(thePage) Then
                tabPagesHidden.Add(thePage)
                tabControl.TabPages.Remove(thePage)
            End If
        End Sub

        Friend Sub ShowPage(thePage As TabPage)
            If tabPagesHidden.Contains(thePage) Then
                tabControl.TabPages.Insert(GetTabPageInsertionPoint(tabControl, thePage.Name), thePage)
                tabPagesHidden.Remove(thePage)
            End If
        End Sub

        Friend Sub PageVisible(thePage As TabPage, value As Boolean)
            If value Then
                ShowPage(thePage)
            Else
                HidePage(thePage)
            End If
        End Sub

        Private Function GetTabPageInsertionPoint(theControl As TabControl, tabPageName As String) As Integer
            Dim tabPageIndex As Integer
            Dim tabPageCurrent As TabPage
            Dim tabNameIndex As Integer
            Dim tabNameCurrent As String

            For tabPageIndex = 0 To theControl.TabPages.Count - 1
                tabPageCurrent = theControl.TabPages(tabPageIndex)
                For tabNameIndex = tabPageIndex To tabPagesOrder.Count - 1
                    tabNameCurrent = tabPagesOrder(tabNameIndex)
                    If tabNameCurrent = tabPageCurrent.Name Then
                        Return tabPageIndex
                    End If
                    If tabNameCurrent = tabPageName Then
                        Return tabPageIndex
                    End If
                Next tabNameIndex
            Next tabPageIndex
            Return tabPageIndex
        End Function

    End Class
End Namespace