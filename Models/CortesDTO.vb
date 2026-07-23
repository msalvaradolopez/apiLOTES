Public Class CortesFilterRequest
    Public Property FechaInicio As String
    Public Property FechaFin As String
    Public Property TiposSalida As List(Of String)
    Public Property PageNumber As Integer
    Public Property PageSize As Integer
    Public Property SearchTerm As String
End Class

Public Class CorteDTO
    Public Property Id As Long
    Public Property IdLote As String
    Public Property Folio As String
    Public Property FechaHora As String
    Public Property TipoSalida As String
    Public Property Cantidad As Decimal
    Public Property Existencia As Decimal
    Public Property Ubicacion As String
    Public Property Usuario As String
    Public Property Almacen As String
    Public Property Socio As String
End Class

Public Class PagedResult(Of T)
    Public Property TotalRecords As Integer
    Public Property PageNumber As Integer
    Public Property PageSize As Integer
    Public Property TotalPages As Integer
    Public Property Data As List(Of T)
End Class

Public Class ApiResponse(Of T)
    Public Property Success As Boolean
    Public Property Message As String
    Public Property Data As T
End Class
