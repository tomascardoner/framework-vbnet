Public Class CS_Control_CheckedComboBox
    Inherits System.Windows.Forms.ComboBox

    Private WithEvents mCheckedListBox As New CheckedListBox

    Public Overloads Property DataSource As Object
        Get
            Return MyBase.DataSource
        End Get
        Set(value As Object)
            MyBase.DataSource = value
        End Set
    End Property

    Protected Overrides Sub OnDropDown(e As EventArgs)
        MyBase.OnDropDown(e)
    End Sub

    Protected Overrides Sub Finalize()
        mCheckedListBox.Dispose()
        MyBase.Finalize()
    End Sub

    Public Property CheckedListBox() As CheckedListBox
        Get
            Return mCheckedListBox
        End Get
        Set(value As CheckedListBox)
            mCheckedListBox = value
        End Set
    End Property
End Class