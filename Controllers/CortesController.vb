Imports System.Net
Imports System.Net.Http
Imports System.Web.Http
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Converters
Imports conexion

<RoutePrefix("api/v1/cortes")>
Public Class CortesController
    Inherits ApiController

    <AcceptVerbs("GET")>
    <HttpGet()>
    <Route("dia")>
    Public Function GetCortesDelDia(<FromUri()> ByVal Optional fechaInicio As String = Nothing,
                                    <FromUri()> ByVal Optional fechaFin As String = Nothing,
                                    <FromUri()> ByVal Optional tiposSalida As String = Nothing,
                                    <FromUri()> ByVal Optional pageNumber As Integer = 1,
                                    <FromUri()> ByVal Optional pageSize As Integer = 50,
                                    <FromUri()> ByVal Optional searchTerm As String = Nothing) As HttpResponseMessage

        Try
            If String.IsNullOrEmpty(fechaInicio) Then fechaInicio = Date.Now().ToString("yyyy-MM-dd")
            If String.IsNullOrEmpty(fechaFin) Then fechaFin = Date.Now().ToString("yyyy-MM-dd")
            If String.IsNullOrEmpty(tiposSalida) Then tiposSalida = "CA,SM,ED,CR,NC,SX,S,X,DX"

            If pageNumber < 1 Then pageNumber = 1
            If pageSize < 1 OrElse pageSize > 200 Then pageSize = 50

            Dim lsError As String = ""
            Dim oConsultas As New consultas()
            Dim dt As DataTable = oConsultas.getCortesDelDia(fechaInicio, fechaFin, tiposSalida, pageNumber, pageSize, searchTerm, lsError)

            If dt Is Nothing Then
                If String.IsNullOrEmpty(lsError) Then lsError = "Error al obtener los cortes del día."
                Return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, lsError)
            End If

            Dim totalRecords As Integer = 0
            If dt.Rows.Count > 0 AndAlso dt.Columns.Contains("TotalRecords") Then
                totalRecords = Convert.ToInt32(dt.Rows(0)("TotalRecords"))
            End If

            Dim cortes As New List(Of CorteDTO)
            For Each row As DataRow In dt.Rows
                Dim dto As New CorteDTO()
                dto.Id = Convert.ToInt64(row("RowNum"))
                dto.IdLote = If(row("IdLote") IsNot DBNull.Value, row("IdLote").ToString(), String.Empty)
                dto.Folio = If(row("Folio") IsNot DBNull.Value, row("Folio").ToString(), String.Empty)
                dto.FechaHora = If(row("FechaHora") IsNot DBNull.Value, row("FechaHora").ToString(), String.Empty)
                dto.TipoSalida = If(row("TipoSalida") IsNot DBNull.Value, row("TipoSalida").ToString(), String.Empty)
                dto.TipoDescripcion = If(row("TipoDescripcion") IsNot DBNull.Value, row("TipoDescripcion").ToString(), String.Empty)
                dto.Usuario = If(row("Usuario") IsNot DBNull.Value, row("Usuario").ToString(), String.Empty)
                dto.Cantidad = If(row("Cantidad") IsNot DBNull.Value, Convert.ToDecimal(row("Cantidad")), 0)
                dto.Estatus = If(row("Estatus") IsNot DBNull.Value, row("Estatus").ToString(), String.Empty)
                cortes.Add(dto)
            Next

            Dim totalPages As Integer = CInt(Math.Ceiling(Convert.ToDouble(totalRecords) / Convert.ToDouble(pageSize)))

            Dim result As New PagedResult(Of CorteDTO)()
            result.TotalRecords = totalRecords
            result.PageNumber = pageNumber
            result.PageSize = pageSize
            result.TotalPages = totalPages
            result.Data = cortes

            Dim response As New ApiResponse(Of PagedResult(Of CorteDTO))()
            response.Success = True
            response.Message = "OK"
            response.Data = result

            Return Request.CreateResponse(HttpStatusCode.OK, response)

        Catch ex As Exception
            Return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message)
        End Try
    End Function

End Class
