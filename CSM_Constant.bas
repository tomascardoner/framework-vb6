Attribute VB_Name = "CSM_Constant"
Option Explicit

Public Const ITEM_START_CHAR As String = "«"
Public Const ITEM_END_CHAR As String = "»"

Public Const ITEM_ALL_MALE As String = ITEM_START_CHAR & "Todos" & ITEM_END_CHAR
Public Const ITEM_ALL_FEMALE As String = ITEM_START_CHAR & "Todas" & ITEM_END_CHAR

Public Const ITEM_EMPTY_MALE As String = ITEM_START_CHAR & "Vacío" & ITEM_END_CHAR
Public Const ITEM_EMPTY_FEMALE As String = ITEM_START_CHAR & "Vacía" & ITEM_END_CHAR

Public Const ITEM_COMPLETE_MALE As String = ITEM_START_CHAR & "Completo" & ITEM_END_CHAR
Public Const ITEM_COMPLETE_FEMALE As String = ITEM_START_CHAR & "Completa" & ITEM_END_CHAR

Public Const ITEM_POSITIVE_MALE As String = "Positivo"
Public Const ITEM_POSITIVE_FEMALE As String = "Positiva"
Public Const ITEM_NEGATIVE_MALE As String = "Negativo"
Public Const ITEM_NEGATIVE_FEMALE As String = "Negativa"

Public Const ITEM_NONE_MALE As String = ITEM_START_CHAR & "Ninguno" & ITEM_END_CHAR
Public Const ITEM_NONE_FEMALE As String = ITEM_START_CHAR & "Ninguna" & ITEM_END_CHAR
Public Const ITEM_NONE_CHARS2 As String = "--"
Public Const ITEM_NONE_CHARS5 As String = "-----"
Public Const ITEM_NONE_CHARS10 As String = "----------"
Public Const ITEM_NONE_CHARS20 As String = "--------------------"

Public Const ITEM_NOTSPECIFIED As String = ITEM_START_CHAR & "No especifica" & ITEM_END_CHAR

Public Const ITEM_DEFAULT_MALE As String = ITEM_START_CHAR & "Predeterminado" & ITEM_END_CHAR
Public Const ITEM_DEFAULT_FEMALE As String = ITEM_START_CHAR & "Predeterminada" & ITEM_END_CHAR

Public Const BOOLEAN_STRING_YES As String = "Sí"
Public Const BOOLEAN_STRING_NO As String = "No"

Public Const KEY_STRINGER As String = "@"
Public Const KEY_DELIMITER As String = "|@|"

Public Const STRING_LIST_SEPARATOR As String = "|"
Public Const STRING_LIST_DELIMITER As String = "¬"

Public Const FILTER_ACTIVO_LIST_INDEX As Long = 1

Public Const DATE_TIME_FIELD_NULL_VALUE As Date = #1/1/1900#

Public Const DATATYPE_LONG_VALUE_MAX As Long = 2147483647
Public Const DATATYPE_LONG_VALUE_MIN As Long = -2147483648#

Public Const DATATYPE_INTEGER_VALUE_MAX As Integer = 32767
Public Const DATATYPE_INTEGER_VALUE_MIN As Integer = -32768

Public Const LOCATION_LATITUDE_NULL_VALUE As Double = 99
Public Const LOCATION_LONGITUDE_NULL_VALUE As Double = 999
