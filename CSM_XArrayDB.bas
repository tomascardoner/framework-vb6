Attribute VB_Name = "CSM_XArrayDB"
Option Explicit

Public Function ConvertADOTypeToXTYPE(ByVal ADODBDataType As ADODB.DataTypeEnum) As XArrayDBObject.XTYPE
    Select Case ADODBDataType
        Case ADODB.DataTypeEnum.adBoolean
            ConvertADOTypeToXTYPE = XTYPE_BOOLEAN
        Case ADODB.DataTypeEnum.adChar, ADODB.DataTypeEnum.adLongVarChar, ADODB.DataTypeEnum.adLongVarWChar, ADODB.DataTypeEnum.adVarChar, ADODB.DataTypeEnum.adVarWChar
            ConvertADOTypeToXTYPE = XTYPE_STRING
        Case ADODB.DataTypeEnum.adCurrency
            ConvertADOTypeToXTYPE = XTYPE_CURRENCY
        Case ADODB.DataTypeEnum.adDate, ADODB.DataTypeEnum.adDBDate, ADODB.DataTypeEnum.adDBTime, ADODB.DataTypeEnum.adDBTimeStamp
            ConvertADOTypeToXTYPE = XTYPE_DATE
        Case ADODB.DataTypeEnum.adDecimal, ADODB.DataTypeEnum.adDouble, ADODB.DataTypeEnum.adNumeric, ADODB.DataTypeEnum.adSingle, ADODB.DataTypeEnum.adVarNumeric
            ConvertADOTypeToXTYPE = XTYPE_DOUBLE
        Case ADODB.DataTypeEnum.adInteger, ADODB.DataTypeEnum.adSmallInt, ADODB.DataTypeEnum.adTinyInt, ADODB.DataTypeEnum.adUnsignedBigInt, ADODB.DataTypeEnum.adUnsignedInt, ADODB.DataTypeEnum.adUnsignedSmallInt, ADODB.DataTypeEnum.adUnsignedTinyInt
            ConvertADOTypeToXTYPE = XTYPE_INTEGER
        Case Else
            Stop
    End Select
End Function
