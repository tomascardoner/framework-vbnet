Module CS_Constants
    Friend Const PUBLIC_ENCRYPTION_PASSWORD As String = "CmcaTlMdmA,aTmP,am2CyalhdEb"

    Friend Const COMBOBOX_ALLYESNO_ALL_LISTINDEX As Byte = 0
    Friend Const COMBOBOX_ALLYESNO_YES_LISTINDEX As Byte = 1
    Friend Const COMBOBOX_ALLYESNO_NO_LISTINDEX As Byte = 2

    Friend Const DATETIMEPICKER_MINIMUM_VALUE As Date = #1753/01/01#

    '////////////////////
    '    FIELD VALUES
    '////////////////////
    Friend Const FIELD_VALUE_NOTSPECIFIED_BYTE As Byte = 0

    Friend Const FIELD_VALUE_NOTSPECIFIED_SHORT As Short = 0
    Friend Const FIELD_VALUE_NOTSPECIFIED_INTEGER As Integer = 0
    Friend Const FIELD_VALUE_NOTSPECIFIED_DATE As Date = #12:00:00 AM#

    Friend Const FIELD_VALUE_ALL_BYTE As Byte = Byte.MaxValue
    Friend Const FIELD_VALUE_ALL_SHORT As Short = Short.MaxValue
    Friend Const FIELD_VALUE_ALL_INTEGER As Integer = Integer.MaxValue

    Friend Const FIELD_VALUE_OTHER_BYTE As Byte = 254
    Friend Const FIELD_VALUE_OTHER_SHORT As Short = 32766
    Friend Const FIELD_VALUE_OTHER_INTEGER As Integer = 2147483646

    '//////////////////
    '    STRINGS
    '//////////////////
    Friend Const KEY_STRINGER As String = "@"

    Friend Const KEY_DELIMITER As String = "|@|"

    Friend Const STRING_LIST_SEPARATOR As Char = ";"c
    'Friend Const STRING_LIST_DELIMITER As String = "¬"

    '//////////////////
    '    E-MAIL
    '//////////////////
    Friend Const EMAIL_CLIENT_NETDLL As String = "NETCLIENT"

    Friend Const EMAIL_CLIENT_MSOUTLOOK As String = "MSOUTLOOK"
    Friend Const EMAIL_CLIENT_CRYSTALREPORTSMAPI As String = "CRYSTALMAPI"
End Module