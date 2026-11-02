<%
' ============================================================================
'  config/dbConnection.asp  (PLANTILLA - NO contiene contraseñas reales)
'  Conexion a SQL Server (BDD_PoliMusic) + utilidades anti-inyeccion SQL.
'  Regla de oro: NINGUN valor del usuario se concatena en el SQL. Todo entra
'  por parametros ADO (ADODB.Command / CreateParameter).
'
'  Para usar esta plantilla:
'    1. Copia este archivo con el nombre: dbConnection.asp
'    2. Reemplaza USUARIO_BD y CONTRASENA_BD con tus credenciales reales.
'  (El archivo dbConnection.asp real está excluido del repositorio por seguridad.)
' ============================================================================
Option Explicit

' --- Constantes ADO (sustituyen a adovbs.inc) ---
Const adCmdText = 1
Const adParamInput = 1
Const adInteger = 3
Const adVarChar = 200

' ============================================================================
'  CADENA DE CONEXION
'  Si tu SQL Server usa nombre de instancia, cambia "localhost" por
'  "localhost\NOMBRE_INSTANCIA". Alternativas de provider comentadas abajo.
' ============================================================================
Const DB_CONN_STRING = "Provider=SQLNCLI11;Server=localhost;Database=BDD_PoliMusic;Uid=USUARIO_BD;Pwd=CONTRASENA_BD;"

' Alternativa B: si "SQLNCLI11" no esta instalado (SQL 2019+):
'Const DB_CONN_STRING = "Provider=MSOLEDBSQL;Server=localhost;Database=BDD_PoliMusic;Uid=USUARIO_BD;Pwd=CONTRASENA_BD;"
' Alternativa C: via ODBC:
'Const DB_CONN_STRING = "Driver={ODBC Driver 17 for SQL Server};Server=localhost;Database=BDD_PoliMusic;Uid=USUARIO_BD;Pwd=CONTRASENA_BD;"

' ============================================================================
'  ABRIR CONEXION / CREAR COMANDOS
' ============================================================================

Function OpenConnection()
    Dim oConn
    Set oConn = Server.CreateObject("ADODB.Connection")
    oConn.ConnectionString = DB_CONN_STRING
    oConn.ConnectionTimeout = 15
    oConn.CommandTimeout = 30
    oConn.Open
    Set OpenConnection = oConn
End Function

Function NewCommand(oConn, sSql)
    Dim oCmd
    Set oCmd = Server.CreateObject("ADODB.Command")
    Set oCmd.ActiveConnection = oConn   ' NOTE: "Set" es obligatorio
    oCmd.CommandType = adCmdText
    oCmd.CommandText = sSql
    Set NewCommand = oCmd
End Function

' ============================================================================
'  PARAMETROS TIPIADOS (impiden el SQL-Injection)
' ============================================================================

Sub AddVarCharParam(oCmd, sName, sValue, iLength)
    Dim sVal
    sVal = "" & sValue
    If IsNull(sValue) Then sVal = ""
    oCmd.Parameters.Append oCmd.CreateParameter(sName, adVarChar, adParamInput, iLength, Left(sVal, iLength))
End Sub

Sub AddIntParam(oCmd, sName, vValue)
    oCmd.Parameters.Append oCmd.CreateParameter(sName, adInteger, adParamInput, 4, ToLong(vValue))
End Sub

' Admite valores vacios -> guarda NULL en la BD (ej. RELEASE_YEAR / PLAYS)
Sub AddNullableIntParam(oCmd, sName, vValue)
    If Len(Trim("" & vValue)) = 0 Then
        oCmd.Parameters.Append oCmd.CreateParameter(sName, adInteger, adParamInput, 4, Null)
    Else
        oCmd.Parameters.Append oCmd.CreateParameter(sName, adInteger, adParamInput, 4, CLng(Trim("" & vValue)))
    End If
End Sub

' ============================================================================
'  CONVERSIONES / LIMPIEZA
' ============================================================================

Function ToLong(vValue)
    ToLong = 0
    If IsNull(vValue) Or IsEmpty(vValue) Then Exit Function
    If IsNumeric(vValue) Then ToLong = CLng(vValue)
End Function

' Recorta y limpia texto de formularios al tamano de la columna
Function DbText(vValue, iMaxLen)
    Dim sVal
    sVal = Trim("" & vValue)
    If IsNull(vValue) Then sVal = ""
    If Len(sVal) > iMaxLen Then sVal = Left(sVal, iMaxLen)
    DbText = sVal
End Function

' ============================================================================
'  HELPERS DE EJECUCION (devuelven 0 = OK, o el numero de error ADO/SQL)
' ============================================================================

Function ExecNonQuery(oConn, sSql)
    Dim oCmd
    On Error Resume Next
    Err.Clear
    Set oCmd = NewCommand(oConn, sSql)
    oCmd.Execute
    ExecNonQuery = Err.Number
    Set oCmd = Nothing
End Function

Function ExecNonQueryId(oConn, sSql, vId)
    Dim oCmd
    On Error Resume Next
    Err.Clear
    Set oCmd = NewCommand(oConn, sSql)
    AddIntParam oCmd, "@id", vId
    oCmd.Execute
    ExecNonQueryId = Err.Number
    Set oCmd = Nothing
End Function

Function ExecNonQuery2Id(oConn, sSql, vId1, vId2)
    Dim oCmd
    On Error Resume Next
    Err.Clear
    Set oCmd = NewCommand(oConn, sSql)
    AddIntParam oCmd, "@id1", vId1
    AddIntParam oCmd, "@id2", vId2
    oCmd.Execute
    ExecNonQuery2Id = Err.Number
    Set oCmd = Nothing
End Function

Function ExecNonQueryStr(oConn, sSql, sValue, iLength)
    Dim oCmd
    On Error Resume Next
    Err.Clear
    Set oCmd = NewCommand(oConn, sSql)
    AddVarCharParam oCmd, "@txt", sValue, iLength
    oCmd.Execute
    ExecNonQueryStr = Err.Number
    Set oCmd = Nothing
End Function

' ============================================================================
'  HELPERS DE CONSULTA (devuelven ADODB.Recordset abierto; cerrar con CloseRs)
' ============================================================================

Function OpenRs(oConn, sSql)
    Dim oCmd, oRs
    On Error Resume Next
    Err.Clear
    Set oCmd = NewCommand(oConn, sSql)
    Set oRs = oCmd.Execute
    If Err.Number <> 0 Then
        Set oRs = Nothing
    End If
    Set OpenRs = oRs
End Function

Function OpenRsId(oConn, sSql, vId)
    Dim oCmd, oRs
    On Error Resume Next
    Err.Clear
    Set oCmd = NewCommand(oConn, sSql)
    AddIntParam oCmd, "@id", vId
    Set oRs = oCmd.Execute
    If Err.Number <> 0 Then
        Set oRs = Nothing
    End If
    Set OpenRsId = oRs
End Function

Function OpenRsStr(oConn, sSql, sValue, iLength)
    Dim oCmd, oRs
    On Error Resume Next
    Err.Clear
    Set oCmd = NewCommand(oConn, sSql)
    AddVarCharParam oCmd, "@txt", sValue, iLength
    Set oRs = oCmd.Execute
    If Err.Number <> 0 Then
        Set oRs = Nothing
    End If
    Set OpenRsStr = oRs
End Function

Sub CloseRs(oRs)
    On Error Resume Next
    If IsObject(oRs) Then
        If oRs.State <> 0 Then oRs.Close
        Set oRs = Nothing
    End If
End Sub

Sub CloseConn(oConn)
    On Error Resume Next
    If IsObject(oConn) Then
        If oConn.State <> 0 Then oConn.Close
        Set oConn = Nothing
    End If
End Sub
%>
