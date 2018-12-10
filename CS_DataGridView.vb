Module CS_DataGridView
    Private Sub Column_SetStandardValues(ByRef col As DataGridViewColumn, ByVal Name As String, ByVal Title As String, ByVal DataField As String, ByVal Alignment As DataGridViewContentAlignment)
        With col
            .Name = Name
            .HeaderText = Title
            .DataPropertyName = DataField
            .DefaultCellStyle.Alignment = Alignment
            .HeaderCell.Style.Alignment = Alignment
        End With
    End Sub

    Friend Function CreateColumn_TextBox(ByVal Name As String, ByVal Title As String, ByVal DataField As String, ByVal Alignment As DataGridViewContentAlignment) As DataGridViewTextBoxColumn
        Dim col As New DataGridViewTextBoxColumn
        Dim cell As New DataGridViewTextBoxCell

        Column_SetStandardValues(CType(col, DataGridViewColumn), Name, Title, DataField, Alignment)

        col.CellTemplate = cell

        Return col
    End Function

    Friend Function CreateColumn_CheckBox(ByVal Name As String, ByVal Title As String, ByVal DataField As String, ByVal Alignment As DataGridViewContentAlignment, ByVal TriState As Boolean, ByVal TrueValue As Object, ByVal FalseValue As Object, ByVal IndeterminateValue As Object) As DataGridViewCheckBoxColumn
        Dim col As New DataGridViewCheckBoxColumn
        Dim cell As New DataGridViewCheckBoxCell(TriState)

        Column_SetStandardValues(CType(col, DataGridViewColumn), Name, Title, DataField, Alignment)

        cell.TrueValue = TrueValue
        cell.FalseValue = FalseValue
        cell.IndeterminateValue = IndeterminateValue

        col.CellTemplate = cell

        Return col
    End Function

    Friend Function CreateColumn_ComboBox(ByVal Name As String, ByVal Title As String, ByVal DataField As String, ByVal Alignment As DataGridViewContentAlignment, ByVal DataSource As IList, ByVal ValueMember As String, ByVal DisplayMember As String) As DataGridViewComboBoxColumn
        Dim col As New DataGridViewComboBoxColumn
        Dim cell As New DataGridViewComboBoxCell()

        Column_SetStandardValues(CType(col, DataGridViewColumn), Name, Title, DataField, Alignment)

        col.DataSource = DataSource
        col.ValueMember = ValueMember
        col.DisplayMember = DisplayMember

        cell.DataSource = DataSource
        cell.ValueMember = ValueMember
        cell.DisplayMember = DisplayMember

        col.CellTemplate = cell

        Return col
    End Function
End Module
