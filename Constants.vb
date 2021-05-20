Namespace CardonerSistemas

    Module Constants
        Friend Const PublicEncryptionPassword As String = "CmcaTlMdmA,aTmP,am2CyalhdEb"

        Friend Const ComboBoxAllYesNo_AllListindex As Byte = 0
        Friend Const ComboBoxAllYesNo_YesListindex As Byte = 1
        Friend Const ComboBoxAllYesNo_NoListindex As Byte = 2

        Friend Const DateTimePickerMinimumValue As Date = #1753/01/01#

        '////////////////////
        '    FIELD VALUES
        '////////////////////
        Friend Const FIELD_VALUE_NOTSPECIFIED_BYTE As Byte = 0

        Friend Const FIELD_VALUE_NOTSPECIFIED_SHORT As Short = 0
        Friend Const FIELD_VALUE_NOTSPECIFIED_INTEGER As Integer = 0
        Friend Const FIELD_VALUE_NOTSPECIFIED_DATE As Date = #12:00:00 AM#
        Friend Const FIELD_VALUE_NOTSPECIFIED_STRING As String = ""

        Friend Const FIELD_VALUE_ALL_BYTE As Byte = Byte.MaxValue
        Friend Const FIELD_VALUE_ALL_SHORT As Short = Short.MaxValue
        Friend Const FIELD_VALUE_ALL_INTEGER As Integer = Integer.MaxValue
        Friend Const FIELD_VALUE_ALL_STRING As String = "#"

        Friend Const FIELD_VALUE_OTHER_BYTE As Byte = Byte.MaxValue - 1
        Friend Const FIELD_VALUE_OTHER_SHORT As Short = Short.MaxValue - 1
        Friend Const FIELD_VALUE_OTHER_INTEGER As Integer = Integer.MaxValue - 1
        Friend Const FIELD_VALUE_OTHER_STRING As String = "@"

        '//////////////////
        '    STRINGS
        '//////////////////
        Friend Const KEY_STRINGER As String = "@"

        Friend Const KEY_DELIMITER As String = "|@|"

        Friend Const STRING_LIST_SEPARATOR As Char = ";"c
        Friend Const STRING_LIST_DELIMITER As String = "¬"
    End Module

End Namespace