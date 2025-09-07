Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.IO
Imports System.Drawing
Imports System.Drawing.Imaging
Imports iTextSharp
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports iTextSharp.text.pdf.draw
Imports System.Globalization

Public Class consultas

    ' conexión string
    Dim conn As String = System.Configuration.ConfigurationManager.ConnectionStrings("connDB").ConnectionString
    Dim DBConn As New SqlConnection(conn)
    Dim DBAdaptador As SqlDataAdapter
    Dim DBcomando As SqlCommand

    ' http://localhost/ctrlapi/api/consultas
    Public Function catArticulos(valorBuscar As String, ExistenciaSiNo As String) As DataTable
        Dim odtProveedores As New DataTable
        Try
            DBConn.Open()
            DBcomando = New SqlCommand("select Articulo = t1.ItemCode, Nombre = t2.ItemName, Almacen = t1.WhsCode, Existencia =  t1.OnHand, Unidad = t2.SalUnitMsr,
                                          Lotes = isnull((select sum(t10.EXISTENCIA)
                                                                from LOTESMOV t10
                                                              where t10.U_SO1_NUMEROARTICULO = t1.ItemCode
                                                                      AND t10.ALMACEN = t1.WhsCode
                                                                      AND t10.TIPOMOV in ('10', 'DE', 'EP', 'EX', 'FP', 'EM', 'RP', 'E')),0)
                                    from OITW T1 inner join OITM t2 on t2.ItemCode = t1.ItemCode
			                                      INNER JOIN LOTESPARAM T3 on t3.VALOR = T1.WhsCode and t3.IDPARAM = 'almacenDefault'
                                     where T2.U_Lote = 'Y'
											AND (
													(isnull(t1.OnHand, 0) > 0 and @ExistenciaSiNo = '1') 
													 or @ExistenciaSiNo = '0'
												)
                                          and ((T2.ItemName like '%' + @valorBuscar + '%'  or T2.ItemCode like '%' + @valorBuscar + '%')  or @valorBuscar = 'TODOS') ", DBConn)
            DBcomando.Parameters.AddWithValue("@ExistenciaSiNo", ExistenciaSiNo)
            DBcomando.Parameters.AddWithValue("@valorBuscar", valorBuscar)

            DBAdaptador = New SqlDataAdapter(DBcomando)

            DBAdaptador.Fill(odtProveedores)

            Return odtProveedores
        Catch ex As Exception
            Return Nothing
        Finally
            DBConn.Close()
        End Try
    End Function

    Public Function catArticuloPorID(itemcode As String, ExistenciaSiNo As String) As DataTable
        Dim odtProveedores As New DataTable
        Try
            DBConn.Open()
            DBcomando = New SqlCommand(" select Articulo = t1.ItemCode, Nombre = t2.ItemName, Almacen = t1.WhsCode, Existencia =  t1.OnHand, Unidad = t2.SalUnitMsr,
                                          Lotes = isnull((select sum(t10.EXISTENCIA)
                                                                from LOTESMOV t10
                                                              where t10.U_SO1_NUMEROARTICULO = t1.ItemCode
                                                                      AND t10.ALMACEN = t1.WhsCode
                                                                      AND t10.TIPOMOV in ('10', 'DE', 'EP', 'EX', 'FP', 'EM', 'RP', 'E')),0)
                                    from OITW T1 inner join OITM t2 on t2.ItemCode = t1.ItemCode
			                                      INNER JOIN LOTESPARAM T3 on t3.VALOR = T1.WhsCode and t3.IDPARAM = 'almacenDefault'
                                     where T2.U_Lote = 'Y'
											AND (
													(isnull(t1.OnHand, 0) > 0 and @ExistenciaSiNo = '1') 
													 or @ExistenciaSiNo = '0'
												)
                                                  AND T1.ItemCode = @itemcode 
                                     order by t2.ItemName ", DBConn)
            DBcomando.Parameters.AddWithValue("@itemcode", itemcode)
            DBcomando.Parameters.AddWithValue("@ExistenciaSiNo", ExistenciaSiNo)

            DBAdaptador = New SqlDataAdapter(DBcomando)

            DBAdaptador.Fill(odtProveedores)

            Return odtProveedores
        Catch ex As Exception
            Return Nothing
        Finally
            DBConn.Close()
        End Try
    End Function

    Public Function getLotePorID(idmov As String) As DataTable
        Dim odtConsulta As New DataTable
        Try
            DBConn.Open()
            DBcomando = New SqlCommand("select IDMOV, T1.ITEMCODE, NOMBRE = T2.ItemName, IDLOTE, ALMACEN, TIPOMOV, DOCTO, FECHADOC, VENDEDOR, CANTIDAD, EXISTENCIA,
                                         UBICACION, USUARIO, FECHAMOV, CONS, UNIDAD = t1.SalUnitMsr
                                  from LOTESMOV T1 INNER JOIN OITM T2 ON T2.ITEMCODE = T1.ITEMCODE
	                                 where T1.IDMOV = @IDMOV ", DBConn)
            DBcomando.Parameters.AddWithValue("@IDMOV", idmov)

            DBAdaptador = New SqlDataAdapter(DBcomando)

            DBAdaptador.Fill(odtConsulta)

            Return odtConsulta
        Catch ex As Exception
            Return Nothing
        Finally
            DBConn.Close()
        End Try
    End Function

    Public Sub setLoteEntrada(oLote As tbLote, ByRef asError As String)
        Dim odtProveedores As New DataTable
        Dim fechaActual = Date.Now()
        Dim cons As Integer = 0
        Try
            If DBConn.State = ConnectionState.Closed Then DBConn.Open()

            If oLote.IDMOV = 0 Then

                If String.IsNullOrEmpty(oLote.IDLOTE) Then

                    If oLote.CANTIDAD <= 0 Then
                        asError = "Deb de capturar valor mayor a 0 en el campo CANTIDAD."
                        Return
                    End If
                    DBcomando = New SqlCommand("select isnull(max(cons), 0) from LOTESMOV where convert(varchar, fechamov, 112) = convert(varchar, @fecha, 112) ", DBConn)
                    DBcomando.Parameters.AddWithValue("@fecha", fechaActual)
                    cons = Convert.ToInt32(DBcomando.ExecuteScalar())
                    cons = cons + 1
                    oLote.CONS = cons

                    DBcomando = New SqlCommand("INSERT INTO LOTESMOV(U_SO1_NUMEROARTICULO,U_SO1_FOLIO,U_SO1_NUMPARTIDA,IDLOTE,ALMACEN,TIPOMOV,DOCTO,FECHADOC,SOCIO,
                                                            VENDEDOR,CANTIDAD,EXISTENCIA,UBICACION,USUARIO_R1,USERID,FECHAMOV,CONS, IMPRIMIR)
                                    values(@U_SO1_NUMEROARTICULO, @U_SO1_FOLIO, @U_SO1_NUMPARTIDA, convert(varchar, @FECHAMOV, 112) + '-' + RIGHT('000' + Ltrim(Rtrim(cast(@CONS as varchar))),3), @ALMACEN, @TIPOMOV, @DOCTO,  @FECHAMOV, @SOCIO,
                                                    @VENDEDOR, @CANTIDAD, @EXISTENCIA, @UBICACION, @USUARIO_R1, @USERID, @FECHAMOV, @CONS, 'S') ", DBConn)
                Else
                    DBcomando = New SqlCommand("INSERT INTO LOTESMOV(U_SO1_NUMEROARTICULO,U_SO1_FOLIO,U_SO1_NUMPARTIDA,IDLOTE,ALMACEN,TIPOMOV,DOCTO,FECHADOC,SOCIO,
                                                            VENDEDOR,CANTIDAD,EXISTENCIA,UBICACION,USUARIO_R1,USERID,FECHAMOV,CONS)
                                    values(@U_SO1_NUMEROARTICULO, @U_SO1_FOLIO, @U_SO1_NUMPARTIDA, @IDLOTE, @ALMACEN, @TIPOMOV, @DOCTO,  @FECHAMOV, @SOCIO,
                                                    @VENDEDOR, @CANTIDAD, @EXISTENCIA, @UBICACION, @USUARIO_R1, @USERID, @FECHAMOV, @CONS, 'S') ", DBConn)
                End If


                If oLote.TIPOMOV = "10" Then
                    oLote.DOCTO = "*NA"
                    oLote.VENDEDOR = "*NA"
                    oLote.U_SO1_FOLIO = "*NA"
                    oLote.U_SO1_NUMPARTIDA = -1
                    oLote.SOCIO = "*NA"
                    oLote.USUARIO_R1 = "*NA"
                End If

            Else
                DBcomando = New SqlCommand("UPDATE LOTESMOV
                                    SET U_SO1_NUMEROARTICULO = @U_SO1_NUMEROARTICULO,
                                        U_SO1_FOLIO = @U_SO1_FOLIO,
                                        U_SO1_NUMPARTIDA = @U_SO1_NUMPARTIDA,
                                        IDLOTE = @IDLOTE,
                                        ALMACEN = @ALMACEN,
                                        TIPOMOV = @TIPOMOV,
                                        DOCTO = @DOCTO,
                                        FECHADOC = @FECHADOC,
                                        SOCIO = @SOCIO,
                                        VENDEDOR = @VENDEDOR,
                                        CANTIDAD = @CANTIDAD,
                                        EXISTENCIA = @EXISTENCIA,
                                        UBICACION = @UBICACION,
                                        USUARIO_R1 = @USUARIO_R1,
                                        USERID = @USERID,
                                        FECHAMOV = @FECHAMOV,
                                        CONS =  @CONS
                                    WHERE IDMOV = @IDMOV", DBConn)
                DBcomando.Parameters.AddWithValue("@IDMOV", oLote.IDMOV)

            End If

            If String.IsNullOrEmpty(oLote.DOCTO) Then oLote.DOCTO = "*NA"
            If String.IsNullOrEmpty(oLote.U_SO1_FOLIO) Then oLote.U_SO1_FOLIO = "*NA"
            If String.IsNullOrEmpty(oLote.U_SO1_NUMPARTIDA) Then oLote.U_SO1_NUMPARTIDA = "*NA"
            If String.IsNullOrEmpty(oLote.SOCIO) Then oLote.SOCIO = "*NA"
            If String.IsNullOrEmpty(oLote.USUARIO_R1) Then oLote.USUARIO_R1 = "*NA"

            DBcomando.Parameters.AddWithValue("@U_SO1_NUMEROARTICULO", oLote.U_SO1_NUMEROARTICULO)
            DBcomando.Parameters.AddWithValue("@U_SO1_FOLIO", oLote.U_SO1_FOLIO)
            DBcomando.Parameters.AddWithValue("@U_SO1_NUMPARTIDA", oLote.U_SO1_NUMPARTIDA)
            DBcomando.Parameters.AddWithValue("@IDLOTE", oLote.IDLOTE)
            DBcomando.Parameters.AddWithValue("@ALMACEN", oLote.ALMACEN)
            DBcomando.Parameters.AddWithValue("@TIPOMOV", oLote.TIPOMOV)
            DBcomando.Parameters.AddWithValue("@DOCTO", oLote.DOCTO)
            If String.IsNullOrEmpty(oLote.FECHADOC) Then
                DBcomando.Parameters.AddWithValue("@FECHADOC", fechaActual)
            Else
                DBcomando.Parameters.AddWithValue("@FECHADOC", DateTime.ParseExact(oLote.FECHADOC, "dd/MM/yyyy", CultureInfo.InvariantCulture))
            End If

            DBcomando.Parameters.AddWithValue("@SOCIO", oLote.SOCIO)
            DBcomando.Parameters.AddWithValue("@VENDEDOR", oLote.VENDEDOR)
            DBcomando.Parameters.AddWithValue("@CANTIDAD", oLote.CANTIDAD)

            oLote.EXISTENCIA = oLote.CANTIDAD
            DBcomando.Parameters.AddWithValue("@EXISTENCIA", oLote.EXISTENCIA)
            DBcomando.Parameters.AddWithValue("@UBICACION", oLote.UBICACION)
            DBcomando.Parameters.AddWithValue("@FECHAMOV", fechaActual)
            DBcomando.Parameters.AddWithValue("@USUARIO_R1", oLote.USUARIO_R1)
            DBcomando.Parameters.AddWithValue("@USERID", oLote.USERID)
            DBcomando.Parameters.AddWithValue("@CONS", oLote.CONS)
            DBcomando.ExecuteNonQuery()

        Catch ex As Exception
            asError = ex.Message.ToString()
            Return
        Finally
            If DBConn.State = ConnectionState.Open Then DBConn.Close()
        End Try
    End Sub

    Public Sub delLoteEntrada(ByRef oLote As tbLote, ByRef _Error As String)

        Dim ctdMov As Integer
        Try
            DBConn.Open()
            DBcomando = New SqlCommand("select ctdMov = count(1)
                                  from LOTESMOV t1
	                                 where T1.IDMOV = @IDMOV
                                          AND T1.CANTIDAD <> T1.EXISTENCIA ", DBConn)
            DBcomando.Parameters.AddWithValue("@IDMOV", oLote.IDMOV)
            ctdMov = Convert.ToInt32(DBcomando.ExecuteScalar())

            If ctdMov >= 1 Then
                _Error = "Existen movimientos asociados al lote : " + oLote.IDLOTE + " (Salidas)"
            Else
                oLote.TIPOMOV = "99" ' MOVTOS ENTRADAS EN BAJA.
                setLoteEntrada(oLote, _Error)
            End If

        Catch ex As Exception
            _Error = ex.Message.ToString
        Finally
            If DBConn.State = ConnectionState.Open Then DBConn.Close()
        End Try
    End Sub

    ' PARA USO DE ENTRADAS DE INVENTARIO INICIAL.
    Public Function getLotesPorArticulo(itemcode As String) As DataTable
        Dim odtConsulta As New DataTable
        Try
            DBConn.Open()
            DBcomando = New SqlCommand("select IDMOV, U_SO1_NUMEROARTICULO, U_SO1_FOLIO, U_SO1_NUMPARTIDA, IDLOTE, ALMACEN, TIPOMOV, DOCTO, FECHADOC = convert(varchar, FECHADOC, 103), SOCIO,
                                    VENDEDOR, CANTIDAD, EXISTENCIA, UBICACION, USUARIO_R1, USERID, FECHAMOV = convert(varchar, FECHAMOV, 103), CONS, NOMBRE = t2.ItemName, T2.SalUnitMsr
                                  from LOTESMOV t1 inner join OITM t2 on t2.ItemCode = t1.U_SO1_NUMEROARTICULO
                                  where TIPOMOV = '10' 
                                        AND T1.U_SO1_NUMEROARTICULO = @itemcode ", DBConn)
            DBcomando.Parameters.AddWithValue("@itemcode", itemcode)

            DBAdaptador = New SqlDataAdapter(DBcomando)

            DBAdaptador.Fill(odtConsulta)

            Return odtConsulta
        Catch ex As Exception
            Return Nothing
        Finally
            DBConn.Close()
        End Try
    End Function

    Public Function getMovtosEntSal(valorBuscar As String, fechaFin As String, asError As String) As DataTable
        Dim odtConsulta As New DataTable
        Dim fechaActual = Date.Now()

        Dim aFecha As String() = fechaFin.Split("/")
        Dim FechaFinAux As New DateTime(aFecha(2), aFecha(1), aFecha(0))

        Try
            DBConn.Open()

            DBcomando = New SqlCommand("select valor
                                  from LOTESPARAM
	                                 where IDPARAM = 'sysFechaInicio' ", DBConn)

            fechaActual = Convert.ToDateTime(DBcomando.ExecuteScalar())

            DBcomando = New SqlCommand("SELECT V.U_SO1_FECHA, 'VENTA' [DOCUMENTO], V.U_SO1_TIPO, IIF(V.U_SO1_FACTURA ='Y', 'FISCAL', 'NO FISCAL') [TIPO],
	                                  V.Name [U_SO1_FOLIO], VD.U_SO1_NUMPARTIDA, V.U_SO1_FOLIOCORTEX, V.U_SO1_FOLIOCONSOLID, 
	                                  convert(varchar, CAST(V.U_SO1_FECHA AS DATE), 103) [FECHA], V.U_SO1_HORACADENA [HORA], 
	                                  CAST(V.U_SO1_USUARIO AS nvarchar(10)) + ' - ' + OHEM.firstName + ' ' + ISNULL(OHEM.middleName,'') + ' ' + ISNULL(OHEM.lastName,'')  [USUARIO], 
	                                  CAST(V.U_SO1_VENDEDOR AS nvarchar(10)) + ' - ' + OSLP.SlpName [VENDEDOR], 
	                                  V.U_SO1_CLIENTE + ' - ' + OCRD.CardName [SOCIO],
	                                  VD.U_SO1_NUMEROARTICULO, VD.U_SO1_DESCRIPCION, VD.U_SO1_CANTIDAD [CANTIDADVENTA], U_SO1_NUMEROLOTE = '', [CANTIDADLOTE] = 0,
                                    VD.U_SO1_ALMACEN, UNIDAD = T2.SalUnitMsr
                                  FROM [@SO1_01VENTA] V WITH (NOLOCK)
                                  JOIN [@SO1_01VENTADETALLE] VD WITH (NOLOCK) ON VD.U_SO1_FOLIO = V.Name
                                  -- JOIN [@SO1_01NUMEROLOTE] NL ON NL.U_SO1_FOLIO = VD.U_SO1_FOLIO 
                                  JOIN OSLP ON OSLP.SlpCode = V.U_SO1_VENDEDOR
	                                 -- AND NL.U_SO1_NUMEROARTICULO = VD.U_SO1_NUMEROARTICULO
                                  JOIN OHEM WITH (NOLOCK) ON OHEM.empID = V.U_SO1_USUARIO
                                  JOIN OCRD WITH (NOLOCK) ON OCRD.CardCode = V.U_SO1_CLIENTE
                                  INNER JOIN OITM t2 WITH (NOLOCK) on t2.ItemCode = VD.U_SO1_NUMEROARTICULO AND T2.U_Lote = 'Y'
                                  WHERE V.U_SO1_TIPO IN ('CR', 'CA')
                                       AND EXISTS (SELECT 1 
				                                        FROM LOTESMOV T1 WITH (NOLOCK) 
				                                        WHERE T1.U_SO1_NUMEROARTICULO = VD.U_SO1_NUMEROARTICULO 
						                                        AND T1.ALMACEN = VD.U_SO1_ALMACEN 
						                                        AND T1.TIPOMOV in ('10', 'DE', 'EP', 'EX', 'FP', 'EM', 'RP', 'E')
						                                        AND T1.EXISTENCIA > 0
						                                        AND (CONVERT(VARCHAR, T1.FECHAMOV, 112) <= CONVERT(VARCHAR, V.U_SO1_FECHA, 112)))
                                      AND VD.U_SO1_CANTIDAD > ISNULL((SELECT SUM (T10.CANTIDAD)
                                                                      FROM LOTESMOV T10 WITH (NOLOCK)
                                                                      WHERE T10.TIPOMOV <> '99'
                                                                            AND T10.U_SO1_NUMEROARTICULO = VD.U_SO1_NUMEROARTICULO
                                                                            AND t10.U_SO1_NUMPARTIDA = VD.U_SO1_NUMPARTIDA
                                                                            AND T10.U_SO1_FOLIO = VD.U_SO1_FOLIO ), 0)
	                                  AND ((VD.U_SO1_DESCRIPCION like '%' + @valorBuscar + '%'  or VD.U_SO1_NUMEROARTICULO like '%' + @valorBuscar + '%')  or @valorBuscar = 'TODOS') 
                                      AND CONVERT(VARCHAR, V.U_SO1_FECHA, 112) <= CONVERT(VARCHAR, @fechaFin, 112)
                                  UNION ALL
                                  SELECT V.U_SO1_FECHA, 'DEVOLUCION' [DOCUMENTO], V.U_SO1_TIPO, IIF(V.U_SO1_FACTURA ='Y', 'FISCAL', 'NO FISCAL') [TIPO],
	                                  V.Name [U_SO1_FOLIO], VD.U_SO1_NUMPARTIDA, V.U_SO1_FOLIOCORTEX, V.U_SO1_FOLIOCONSOLID, 
	                                  convert(varchar, CAST(V.U_SO1_FECHA AS DATE), 103) [FECHA], V.U_SO1_HORACADENA [HORA], 
	                                  CAST(V.U_SO1_USUARIO AS nvarchar(10)) + ' - ' + OHEM.firstName + ' ' + ISNULL(OHEM.middleName,'') + ' ' + ISNULL(OHEM.lastName,'')  [USUARIO], 
	                                  CAST(V.U_SO1_VENDEDOR AS nvarchar(10)) + ' - ' + OSLP.SlpName [VENDEDOR], 
	                                  V.U_SO1_CLIENTE + ' - ' + OCRD.CardName [SOCIO],
	                                  VD.U_SO1_NUMEROARTICULO, VD.U_SO1_DESCRIPCION, VD.U_SO1_CANTIDAD [CANTIDADVENTA], U_SO1_NUMEROLOTE = '', [CANTIDADLOTE] = 0,
                                    VD.U_SO1_ALMACEN, UNIDAD = T2.SalUnitMsr
                                  FROM [@SO1_01DEVOLUCION] V WITH (NOLOCK)
                                  JOIN [@SO1_01DEVOLUCIONDET] VD WITH (NOLOCK) ON VD.U_SO1_FOLIO = V.Name
                                  -- JOIN [@SO1_01NUMEROLOTE] NL ON NL.U_SO1_FOLIO = VD.U_SO1_FOLIO 
                                  JOIN OSLP WITH (NOLOCK) ON OSLP.SlpCode = V.U_SO1_VENDEDOR
	                                 -- AND NL.U_SO1_NUMEROARTICULO = VD.U_SO1_NUMEROARTICULO
                                  JOIN OHEM WITH (NOLOCK) ON OHEM.empID = V.U_SO1_USUARIO
                                  JOIN OCRD WITH (NOLOCK) ON OCRD.CardCode = V.U_SO1_CLIENTE
                                  INNER JOIN OITM t2 WITH (NOLOCK) on t2.ItemCode = VD.U_SO1_NUMEROARTICULO AND T2.U_Lote = 'Y'
                                  WHERE CONVERT(VARCHAR, V.U_SO1_FECHA, 112) >= CONVERT(VARCHAR, @fechaActual, 112)
                                         AND VD.U_SO1_CANTIDAD > ISNULL((SELECT SUM (T10.CANTIDAD)
                                                                          FROM LOTESMOV T10 WITH (NOLOCK)
                                                                          WHERE T10.TIPOMOV <> '99'
                                                                                AND T10.U_SO1_NUMEROARTICULO = VD.U_SO1_NUMEROARTICULO
                                                                                AND t10.U_SO1_NUMPARTIDA = VD.U_SO1_NUMPARTIDA
                                                                                AND T10.U_SO1_FOLIO = VD.U_SO1_FOLIO ), 0)
                                        AND ((VD.U_SO1_DESCRIPCION like '%' + @valorBuscar + '%'  or VD.U_SO1_NUMEROARTICULO like '%' + @valorBuscar + '%')  or @valorBuscar = 'TODOS') 
                                        AND CONVERT(VARCHAR, V.U_SO1_FECHA, 112) <= CONVERT(VARCHAR, @fechaFin, 112)
                                  UNION ALL
                                  SELECT V.U_SO1_FECHA, 'COMPRA' [DOCUMENTO], V.U_SO1_TIPO,
	                                  CASE V.U_SO1_TIPO WHEN 'FP' THEN 'FACTURA PROVEEDOR'
					                                    WHEN 'EP' THEN 'ENTRADA PROVEEDOR'
					                                    ELSE 'OTROS'
	                                  END [TIPO],
	                                  V.Name [U_SO1_FOLIO], VD.U_SO1_NUMPARTIDA, V.U_SO1_FOLIOCORTEX, '', 
	                                  convert(varchar, CAST(V.U_SO1_FECHA AS DATE), 103) [FECHA], V.U_SO1_HORACADENA [HORA], 
	                                  CAST(V.U_SO1_USUARIO AS nvarchar(10)) + ' - ' + OHEM.firstName + ' ' + ISNULL(OHEM.middleName,'') + ' ' + ISNULL(OHEM.lastName,'')  [USUARIO], 
	                                  '' [VENDEDOR], 
	                                  V.U_SO1_PROVEEDOR + ' - ' + OCRD.CardName [SOCIO],
	                                  VD.U_SO1_NUMEROARTICULO, VD.U_SO1_DESCRIPCION, VD.U_SO1_CANTIDAD [CANTIDADVENTA], U_SO1_NUMEROLOTE = '', [CANTIDADLOTE] = 0,
                                    VD.U_SO1_ALMACEN, UNIDAD = T2.SalUnitMsr 
                                  FROM [@SO1_01COMPRA] V WITH (NOLOCK)
                                  JOIN [@SO1_01COMPRADETALLE] VD WITH (NOLOCK) ON VD.U_SO1_FOLIO = V.Name
                                  -- JOIN [@SO1_01NUMEROLOTE] NL ON NL.U_SO1_FOLIO = VD.U_SO1_FOLIO 
                                  JOIN OHEM WITH (NOLOCK) ON OHEM.empID = V.U_SO1_USUARIO
                                  JOIN OCRD WITH (NOLOCK) ON OCRD.CardCode = V.U_SO1_PROVEEDOR
                                  INNER JOIN OITM t2 WITH (NOLOCK) on t2.ItemCode = VD.U_SO1_NUMEROARTICULO AND T2.U_Lote = 'Y'
                                  WHERE CONVERT(VARCHAR, V.U_SO1_FECHA, 112) >= CONVERT(VARCHAR, @fechaActual, 112)
                                        AND VD.U_SO1_CANTIDAD > ISNULL((SELECT SUM (T10.CANTIDAD)
                                                                          FROM LOTESMOV T10 WITH (NOLOCK)
                                                                          WHERE T10.TIPOMOV <> '99'
                                                                                AND T10.U_SO1_NUMEROARTICULO = VD.U_SO1_NUMEROARTICULO
                                                                                AND t10.U_SO1_NUMPARTIDA = VD.U_SO1_NUMPARTIDA
                                                                                AND T10.U_SO1_FOLIO = VD.U_SO1_FOLIO ), 0)
                                        AND ((VD.U_SO1_DESCRIPCION like '%' + @valorBuscar + '%'  or VD.U_SO1_NUMEROARTICULO like '%' + @valorBuscar + '%')  or @valorBuscar = 'TODOS') 
                                        AND CONVERT(VARCHAR, V.U_SO1_FECHA, 112) <= CONVERT(VARCHAR, @fechaFin, 112)
                                  UNION ALL
                                  SELECT V.U_SO1_FECHA, 'DEVOLUCION COMPRA' [DOCUMENTO], U_SO1_TIPO = 'DX',
	                                  CASE V.U_SO1_TIPODOCDESTINO WHEN 'N' THEN 'NOTA CREDITO PROVEEDOR'
					                                    WHEN 'D' THEN 'DEVOLUCION PROVEEDOR'
					                                    ELSE 'OTROS'
	                                  END [TIPO],
	                                  V.Name [U_SO1_FOLIO], VD.U_SO1_NUMPARTIDA, V.U_SO1_FOLIOCORTEX, '', 
	                                  convert(varchar, CAST(V.U_SO1_FECHA AS DATE), 103) [FECHA], V.U_SO1_HORACADENA [HORA], 
	                                  CAST(V.U_SO1_USUARIO AS nvarchar(10)) + ' - ' + OHEM.firstName + ' ' + ISNULL(OHEM.middleName,'') + ' ' + ISNULL(OHEM.lastName,'')  [USUARIO], 
	                                  '' [VENDEDOR], 
	                                  V.U_SO1_PROVEEDOR + ' - ' + OCRD.CardName [SOCIO],
	                                  VD.U_SO1_NUMEROARTICULO, OITM.ItemName, VD.U_SO1_CANTIDAD [CANTIDADVENTA], U_SO1_NUMEROLOTE = '', [CANTIDADLOTE] = 0,
                                    VD.U_SO1_ALMACEN, UNIDAD = T2.SalUnitMsr
                                  FROM [@SO1_01DEVOLCOMPRA] V WITH (NOLOCK)
                                  JOIN [@SO1_01DEVOLCOMPRADE] VD WITH (NOLOCK) ON VD.U_SO1_FOLIO = V.Name
                                  -- JOIN [@SO1_01NUMEROLOTE] NL ON NL.U_SO1_FOLIO = VD.U_SO1_FOLIO 
                                  JOIN OHEM WITH (NOLOCK) ON OHEM.empID = V.U_SO1_USUARIO
                                  JOIN OCRD WITH (NOLOCK) ON OCRD.CardCode = V.U_SO1_PROVEEDOR
                                  JOIN OITM WITH (NOLOCK) ON OITM.ItemCode = VD.U_SO1_NUMEROARTICULO AND OITM.U_Lote = 'Y'
                                  INNER JOIN OITM t2 WITH (NOLOCK) on t2.ItemCode = VD.U_SO1_NUMEROARTICULO
                                  WHERE EXISTS (SELECT 1 
				                                  FROM LOTESMOV T1 WITH (NOLOCK)
				                                  WHERE T1.U_SO1_NUMEROARTICULO = VD.U_SO1_NUMEROARTICULO 
						                                  AND T1.ALMACEN = VD.U_SO1_ALMACEN 
						                                  AND T1.TIPOMOV in ('10', 'DE', 'EP', 'EX', 'FP', 'EM', 'RP', 'E')
						                                  AND T1.EXISTENCIA > 0
						                                  AND (CONVERT(VARCHAR, T1.FECHAMOV, 112) <= CONVERT(VARCHAR, V.U_SO1_FECHA, 112)))
                                              AND VD.U_SO1_CANTIDAD > ISNULL((SELECT SUM (T10.CANTIDAD)
                                                                              FROM LOTESMOV T10 WITH (NOLOCK)
                                                                              WHERE T10.TIPOMOV <> '99'
                                                                                    AND T10.U_SO1_NUMEROARTICULO = VD.U_SO1_NUMEROARTICULO
                                                                                    AND t10.U_SO1_NUMPARTIDA = VD.U_SO1_NUMPARTIDA
                                                                                    AND T10.U_SO1_FOLIO = VD.U_SO1_FOLIO ), 0)
                                        AND ((OITM.ItemName like '%' + @valorBuscar + '%'  or VD.U_SO1_NUMEROARTICULO like '%' + @valorBuscar + '%')  or @valorBuscar = 'TODOS') 
                                        AND CONVERT(VARCHAR, V.U_SO1_FECHA, 112) <= CONVERT(VARCHAR, @fechaFin, 112)
                                  UNION ALL
                                  SELECT V.U_SO1_FECHA, 'TRASPASO SALIDA' [DOCUMENTO], V.U_SO1_TIPO,
	                                  CASE V.U_SO1_TIPO WHEN 'E' THEN 'ENTRADA POR TRASPASO'
					                                    WHEN 'S' THEN 'SALIDA POR TRASPASO'
					                                    ELSE 'OTROS'
	                                  END [TIPO],
	                                  V.U_SO1_FOLIO [U_SO1_FOLIO], VD.U_SO1_NUMPARTIDA, V.U_SO1_FOLIOCORTEX, '', 
	                                  convert(varchar, CAST(V.U_SO1_FECHA AS DATE), 103) [FECHA], V.U_SO1_HORACADENA [HORA], 
	                                  CAST(V.U_SO1_USUARIO AS nvarchar(10)) + ' - ' + OHEM.firstName + ' ' + ISNULL(OHEM.middleName,'') + ' ' + ISNULL(OHEM.lastName,'')  [USUARIO], 
	                                  '' [VENDEDOR], 
	                                  '' [SOCIO],
	                                  VD.U_SO1_NUMEROARTICULO, VD.U_SO1_DESCRIPCION, VD.U_SO1_CANTIDAD [CANTIDADVENTA], U_SO1_NUMEROLOTE = '', [CANTIDADLOTE] = 0,
                                    V.U_SO1_ALMACEN, UNIDAD = T2.SalUnitMsr
                                  FROM [@SO1_01TRASPASO] V WITH (NOLOCK)
                                  JOIN [@SO1_01TRASPASODET] VD WITH (NOLOCK) ON VD.U_SO1_FOLIO = V.U_SO1_FOLIO AND VD.U_SO1_INSTANCIA = V.U_SO1_INSTANCIA
                                  -- JOIN [@SO1_01NUMEROLOTE] NL ON NL.U_SO1_FOLIO = VD.U_SO1_FOLIO 
                                  JOIN OHEM ON OHEM.empID = V.U_SO1_USUARIO
                                  INNER JOIN OITM t2 on t2.ItemCode = VD.U_SO1_NUMEROARTICULO AND t2.U_Lote = 'Y'
                                  WHERE V.U_SO1_TIPO = 'S'
                                        AND EXISTS (SELECT 1 
				                                            FROM LOTESMOV T1 WITH (NOLOCK) 
				                                            WHERE T1.U_SO1_NUMEROARTICULO = VD.U_SO1_NUMEROARTICULO 
						                                            AND T1.ALMACEN = V.U_SO1_ALMACEN 
						                                            AND T1.TIPOMOV in ('10', 'DE', 'EP', 'EX', 'FP', 'EM', 'RP', 'E')
						                                            AND T1.EXISTENCIA > 0
						                                            AND (CONVERT(VARCHAR, T1.FECHAMOV, 112) <= CONVERT(VARCHAR, V.U_SO1_FECHA, 112)))
                                        AND VD.U_SO1_CANTIDAD > ISNULL((SELECT SUM (T10.CANTIDAD)
                                                                        FROM LOTESMOV T10 WITH (NOLOCK)
                                                                        WHERE T10.TIPOMOV <> '99'
                                                                              AND T10.U_SO1_NUMEROARTICULO = VD.U_SO1_NUMEROARTICULO
                                                                              AND t10.U_SO1_NUMPARTIDA = VD.U_SO1_NUMPARTIDA
                                                                              AND T10.U_SO1_FOLIO = VD.U_SO1_FOLIO ), 0)
		                                  AND ((VD.U_SO1_DESCRIPCION like '%' + @valorBuscar + '%'  or VD.U_SO1_NUMEROARTICULO like '%' + @valorBuscar + '%')  or @valorBuscar = 'TODOS')
                                          AND CONVERT(VARCHAR, V.U_SO1_FECHA, 112) <= CONVERT(VARCHAR, @fechaFin, 112)
                                  UNION ALL
                                  SELECT V.U_SO1_FECHA, 'TRASPASO SALIDA SAP' [DOCUMENTO], V.U_SO1_TIPO,
	                                  CASE V.U_SO1_TIPO WHEN 'E' THEN 'ENTRADA POR TRASPASO'
					                                    WHEN 'S' THEN 'SALIDA POR TRASPASO'
					                                    ELSE 'OTROS'
	                                  END [TIPO],
	                                  V.U_SO1_FOLIO [U_SO1_FOLIO], VD.U_SO1_NUMPARTIDA, V.U_SO1_FOLIOCORTEX, '', 
	                                  convert(varchar, CAST(V.U_SO1_FECHA AS DATE), 103) [FECHA], V.U_SO1_HORACADENA [HORA], 
	                                  CAST(V.U_SO1_USUARIO AS nvarchar(10)) + ' - ' + OHEM.firstName + ' ' + ISNULL(OHEM.middleName,'') + ' ' + ISNULL(OHEM.lastName,'')  [USUARIO], 
	                                  '' [VENDEDOR], 
	                                  '' [SOCIO],
	                                  VD.U_SO1_NUMEROARTICULO, VD.U_SO1_DESCRIPCION, VD.U_SO1_CANTIDAD [CANTIDADVENTA], U_SO1_NUMEROLOTE = '', [CANTIDADLOTE] = 0,
                                    V.U_SO1_ALMACEN, UNIDAD = T2.SalUnitMsr
                                  FROM [@SO1_01TRASPASO] V WITH (NOLOCK)
                                  JOIN [@SO1_01TRASPASODET] VD WITH (NOLOCK) ON VD.U_SO1_FOLIO = V.U_SO1_FOLIO
                                  -- JOIN [@SO1_01NUMEROLOTE] NL ON NL.U_SO1_FOLIO = VD.U_SO1_FOLIO 
                                  JOIN OHEM ON OHEM.empID = V.U_SO1_USUARIO
                                  INNER JOIN OITM t2 on t2.ItemCode = VD.U_SO1_NUMEROARTICULO AND t2.U_Lote = 'Y'
                                  WHERE V.U_SO1_TIPO = 'X'
                                        AND EXISTS (SELECT 1 
				                                            FROM LOTESMOV T1 WITH (NOLOCK) 
				                                            WHERE T1.U_SO1_NUMEROARTICULO = VD.U_SO1_NUMEROARTICULO 
						                                            AND T1.ALMACEN = V.U_SO1_ALMACEN 
						                                            AND T1.TIPOMOV in ('10', 'DE', 'EP', 'EX', 'FP', 'EM', 'RP', 'E')
						                                            AND T1.EXISTENCIA > 0
						                                            AND (CONVERT(VARCHAR, T1.FECHAMOV, 112) <= CONVERT(VARCHAR, V.U_SO1_FECHA, 112)))
                                        AND VD.U_SO1_CANTIDAD > ISNULL((SELECT SUM (T10.CANTIDAD)
                                                                        FROM LOTESMOV T10 WITH (NOLOCK)
                                                                        WHERE T10.TIPOMOV <> '99'
                                                                              AND T10.U_SO1_NUMEROARTICULO = VD.U_SO1_NUMEROARTICULO
                                                                              AND t10.U_SO1_NUMPARTIDA = VD.U_SO1_NUMPARTIDA
                                                                              AND T10.U_SO1_FOLIO = VD.U_SO1_FOLIO ), 0)
		                                  AND ((VD.U_SO1_DESCRIPCION like '%' + @valorBuscar + '%'  or VD.U_SO1_NUMEROARTICULO like '%' + @valorBuscar + '%')  or @valorBuscar = 'TODOS')
                                          AND CONVERT(VARCHAR, V.U_SO1_FECHA, 112) <= CONVERT(VARCHAR, @fechaFin, 112)
                                      UNION ALL
                                      SELECT V.U_SO1_FECHA, 'TRASPASO ENTRADA' [DOCUMENTO], V.U_SO1_TIPO,
	                                      CASE V.U_SO1_TIPO WHEN 'E' THEN 'ENTRADA POR TRASPASO'
					                                        WHEN 'S' THEN 'SALIDA POR TRASPASO'
					                                        ELSE 'OTROS'
	                                      END [TIPO],
	                                      V.U_SO1_FOLIO [U_SO1_FOLIO], VD.U_SO1_NUMPARTIDA, V.U_SO1_FOLIOCORTEX, '', 
	                                      convert(varchar, CAST(V.U_SO1_FECHA AS DATE), 103) [FECHA], V.U_SO1_HORACADENA [HORA], 
	                                      CAST(V.U_SO1_USUARIO AS nvarchar(10)) + ' - ' + OHEM.firstName + ' ' + ISNULL(OHEM.middleName,'') + ' ' + ISNULL(OHEM.lastName,'')  [USUARIO], 
	                                      '' [VENDEDOR], 
	                                      '' [SOCIO],
	                                      VD.U_SO1_NUMEROARTICULO, VD.U_SO1_DESCRIPCION, VD.U_SO1_CANTIDAD [CANTIDADVENTA], U_SO1_NUMEROLOTE = '', [CANTIDADLOTE] = 0,
                                        V.U_SO1_ALMACEN, UNIDAD = T2.SalUnitMsr
                                      FROM [@SO1_01TRASPASO] V WITH (NOLOCK)
                                      JOIN [@SO1_01TRASPASODET] VD WITH (NOLOCK) ON VD.U_SO1_FOLIO = V.U_SO1_FOLIO AND VD.U_SO1_INSTANCIA = V.U_SO1_INSTANCIA
                                      -- JOIN [@SO1_01NUMEROLOTE] NL ON NL.U_SO1_FOLIO = VD.U_SO1_FOLIO 
                                      JOIN OHEM ON OHEM.empID = V.U_SO1_USUARIO
                                      INNER JOIN OITM t2 on t2.ItemCode = VD.U_SO1_NUMEROARTICULO  AND t2.U_Lote = 'Y'
                                      WHERE V.U_SO1_TIPO = 'E'
                                          AND CONVERT(VARCHAR, V.U_SO1_FECHA, 112) >= CONVERT(VARCHAR, @fechaActual, 112)
                                          AND VD.U_SO1_CANTIDAD > ISNULL((SELECT SUM (T10.CANTIDAD)
                                                                          FROM LOTESMOV T10 WITH (NOLOCK)
                                                                          WHERE T10.TIPOMOV <> '99'
                                                                                AND T10.U_SO1_NUMEROARTICULO = VD.U_SO1_NUMEROARTICULO
                                                                                AND t10.U_SO1_NUMPARTIDA = VD.U_SO1_NUMPARTIDA
                                                                                AND T10.U_SO1_FOLIO = VD.U_SO1_FOLIO ), 0)
		                                      AND ((VD.U_SO1_DESCRIPCION like '%' + @valorBuscar + '%'  or VD.U_SO1_NUMEROARTICULO like '%' + @valorBuscar + '%')  or @valorBuscar = 'TODOS') 
                                              AND CONVERT(VARCHAR, V.U_SO1_FECHA, 112) <= CONVERT(VARCHAR, @fechaFin, 112)
                                  UNION ALL
                                  SELECT V.U_SO1_FECHA, 'SALIDA DE MERCANCIA' [DOCUMENTO], U_SO1_TIPO = 'SX',
	                                  '' [TIPO],
	                                  V.Name [U_SO1_FOLIO], VD.U_SO1_NUMPARTIDA, V.U_SO1_FOLIOCORTEX, '', 
	                                  convert(varchar, CAST(V.U_SO1_FECHA AS DATE), 103) [FECHA], V.U_SO1_HORACADENA [HORA], 
	                                  CAST(V.U_SO1_USUARIO AS nvarchar(10)) + ' - ' + OHEM.firstName + ' ' + ISNULL(OHEM.middleName,'') + ' ' + ISNULL(OHEM.lastName,'')  [USUARIO], 
	                                  '' [VENDEDOR], 
	                                  '' [SOCIO],
	                                  VD.U_SO1_NUMEROARTICULO, VD.U_SO1_DESCRIPCION, VD.U_SO1_CANTIDAD [CANTIDADVENTA], U_SO1_NUMEROLOTE = '', [CANTIDADLOTE] = 0,
                                    VD.U_SO1_ALMACEN, UNIDAD = T2.SalUnitMsr
                                  FROM [@SO1_01SALIDAMERCAN] V WITH (NOLOCK)
                                  JOIN [@SO1_01SALIDAMERDET] VD WITH (NOLOCK) ON VD.U_SO1_FOLIO = V.Name
                                  -- JOIN [@SO1_01NUMEROLOTE] NL ON NL.U_SO1_FOLIO = VD.U_SO1_FOLIO 
                                  JOIN OHEM WITH (NOLOCK) ON OHEM.empID = V.U_SO1_USUARIO
                                  INNER JOIN OITM t2 WITH (NOLOCK) on t2.ItemCode = VD.U_SO1_NUMEROARTICULO  AND t2.U_Lote = 'Y'
                                  WHERE EXISTS (SELECT 1 
				                                        FROM LOTESMOV T1 WITH (NOLOCK)
				                                        WHERE T1.U_SO1_NUMEROARTICULO = VD.U_SO1_NUMEROARTICULO 
						                                        AND T1.ALMACEN = VD.U_SO1_ALMACEN 
						                                        AND T1.TIPOMOV in ('10', 'DE', 'EP', 'EX', 'FP', 'EM', 'RP', 'E')
						                                        AND T1.EXISTENCIA > 0
						                                        AND (CONVERT(VARCHAR, T1.FECHAMOV, 112) <= CONVERT(VARCHAR, V.U_SO1_FECHA, 112)))
                                        AND VD.U_SO1_CANTIDAD > ISNULL((SELECT SUM (T10.CANTIDAD)
                                                                          FROM LOTESMOV T10 WITH (NOLOCK)
                                                                          WHERE T10.TIPOMOV <> '99'
                                                                                AND T10.U_SO1_NUMEROARTICULO = VD.U_SO1_NUMEROARTICULO
                                                                                AND t10.U_SO1_NUMPARTIDA = VD.U_SO1_NUMPARTIDA
                                                                                AND T10.U_SO1_FOLIO = VD.U_SO1_FOLIO ), 0)
		                                  AND ((VD.U_SO1_DESCRIPCION like '%' + @valorBuscar + '%'  or VD.U_SO1_NUMEROARTICULO like '%' + @valorBuscar + '%')  or @valorBuscar = 'TODOS') 
                                          AND CONVERT(VARCHAR, V.U_SO1_FECHA, 112) <= CONVERT(VARCHAR, @fechaFin, 112)    
                                  UNION ALL
                                  SELECT V.U_SO1_FECHA, 'ENTRADA DE MERCANCIA' [DOCUMENTO], U_SO1_TIPO = 'EX',
	                                  '' [TIPO],
	                                  V.Name [U_SO1_FOLIO], VD.U_SO1_NUMPARTIDA, V.U_SO1_FOLIOCORTEX, '', 
	                                  convert(varchar, CAST(V.U_SO1_FECHA AS DATE), 103) [FECHA], V.U_SO1_HORACADENA [HORA], 
	                                  CAST(V.U_SO1_USUARIO AS nvarchar(10)) + ' - ' + OHEM.firstName + ' ' + ISNULL(OHEM.middleName,'') + ' ' + ISNULL(OHEM.lastName,'')  [USUARIO], 
	                                  '' [VENDEDOR], 
	                                  '' [SOCIO],
	                                  VD.U_SO1_NUMEROARTICULO, VD.U_SO1_DESCRIPCION, VD.U_SO1_CANTIDAD [CANTIDADVENTA], U_SO1_NUMEROLOTE = '', [CANTIDADLOTE] = 0,
                                    VD.U_SO1_ALMACEN, UNIDAD = T2.SalUnitMsr
                                  FROM [@SO1_01ENTRADAMERCAN] V WITH (NOLOCK)
                                  JOIN [@SO1_01ENTRADAMERDET] VD WITH (NOLOCK) ON VD.U_SO1_FOLIO = V.Name
                                  -- JOIN [@SO1_01NUMEROLOTE] NL ON NL.U_SO1_FOLIO = VD.U_SO1_FOLIO 
                                  JOIN OHEM WITH (NOLOCK) ON OHEM.empID = V.U_SO1_USUARIO
                                  INNER JOIN OITM t2 WITH (NOLOCK) on t2.ItemCode = VD.U_SO1_NUMEROARTICULO  AND t2.U_Lote = 'Y'
                                  WHERE CONVERT(VARCHAR, V.U_SO1_FECHA, 112) >= CONVERT(VARCHAR, @fechaActual, 112)
                                        AND VD.U_SO1_CANTIDAD > ISNULL((SELECT SUM (T10.CANTIDAD)
                                                                          FROM LOTESMOV T10 WITH (NOLOCK)
                                                                          WHERE T10.TIPOMOV <> '99'
                                                                                AND T10.U_SO1_NUMEROARTICULO = VD.U_SO1_NUMEROARTICULO
                                                                                AND t10.U_SO1_NUMPARTIDA = VD.U_SO1_NUMPARTIDA
                                                                                AND T10.U_SO1_FOLIO = VD.U_SO1_FOLIO ), 0)
                                          AND ((VD.U_SO1_DESCRIPCION like '%' + @valorBuscar + '%'  or VD.U_SO1_NUMEROARTICULO like '%' + @valorBuscar + '%')  or @valorBuscar = 'TODOS') 
                                          AND CONVERT(VARCHAR, V.U_SO1_FECHA, 112) <= CONVERT(VARCHAR, @fechaFin, 112)
                                   UNION ALL
                                   SELECT V.U_SO1_FECHA, 'SALIDA PRODUCCION' [DOCUMENTO], U_SO1_TIPO = 'RP',
	                                  '' [TIPO],
	                                  V.Name [U_SO1_FOLIO], VD.U_SO1_NUMPARTIDA, V.U_SO1_FOLIOCORTEX, '', 
	                                  convert(varchar, CAST(V.U_SO1_FECHA AS DATE), 103) [FECHA], V.U_SO1_HORACADENA [HORA], 
	                                  CAST(V.U_SO1_USUARIO AS nvarchar(10)) + ' - ' + OHEM.firstName + ' ' + ISNULL(OHEM.middleName,'') + ' ' + ISNULL(OHEM.lastName,'')  [USUARIO], 
	                                  '' [VENDEDOR], 
	                                  '' [SOCIO],
	                                  VD.U_SO1_NUMEROARTICULO,  T2.ItemName [U_SO1_DESCRIPCION], VD.U_SO1_CANTIDADREQUE [CANTIDADVENTA], U_SO1_NUMEROLOTE = '', [CANTIDADLOTE] = 0,
                                    VD.U_SO1_ALMACEN, UNIDAD = T2.SalUnitMsr
                                  FROM [@SO1_01PRODUCCION] V WITH (NOLOCK)
                                  JOIN [@SO1_01PRODUCCIONDET] VD WITH (NOLOCK) ON VD.U_SO1_FOLIO = V.Name
                                  -- JOIN [@SO1_01NUMEROLOTE] NL ON NL.U_SO1_FOLIO = VD.U_SO1_FOLIO 
                                  JOIN OHEM WITH (NOLOCK) ON OHEM.empID = V.U_SO1_USUARIO
                                  INNER JOIN OITM t2 WITH (NOLOCK) on t2.ItemCode = VD.U_SO1_NUMEROARTICULO  AND t2.U_Lote = 'Y'
                                  WHERE EXISTS (SELECT 1 
				                                        FROM LOTESMOV T1 WITH (NOLOCK)
				                                        WHERE T1.U_SO1_NUMEROARTICULO = VD.U_SO1_NUMEROARTICULO 
						                                        AND T1.ALMACEN = VD.U_SO1_ALMACEN 
						                                        AND T1.TIPOMOV in ('10', 'DE', 'EP', 'EX', 'FP', 'EM', 'RP', 'E')
						                                        AND T1.EXISTENCIA > 0
						                                        AND (CONVERT(VARCHAR, T1.FECHAMOV, 112) <= CONVERT(VARCHAR, V.U_SO1_FECHA, 112)))
                                        AND VD.U_SO1_CANTIDADREQUE > ISNULL((SELECT SUM (T10.CANTIDAD)
                                                                          FROM LOTESMOV T10 WITH (NOLOCK)
                                                                          WHERE T10.TIPOMOV <> '99'
                                                                                AND T10.U_SO1_NUMEROARTICULO = VD.U_SO1_NUMEROARTICULO
                                                                                AND t10.U_SO1_NUMPARTIDA = VD.U_SO1_NUMPARTIDA
                                                                                AND T10.U_SO1_FOLIO = VD.U_SO1_FOLIO ), 0)
										AND ((t2.ItemName like '%' + @valorBuscar + '%'  or VD.U_SO1_NUMEROARTICULO like '%' + @valorBuscar + '%')  or @valorBuscar = 'TODOS') 
                                        AND CONVERT(VARCHAR, V.U_SO1_FECHA, 112) <= CONVERT(VARCHAR, @fechaFin, 112)
                                    ORDER BY 1 desc
                                    ", DBConn)
            DBcomando.Parameters.AddWithValue("@valorBuscar", valorBuscar)
            DBcomando.Parameters.AddWithValue("@fechaActual", fechaActual)
            DBcomando.Parameters.AddWithValue("@fechaFin", FechaFinAux)

            DBAdaptador = New SqlDataAdapter(DBcomando)

            DBAdaptador.Fill(odtConsulta)

            Return odtConsulta
        Catch ex As Exception
            asError = ex.Message.ToString()
            Return Nothing
        Finally
            DBConn.Close()
        End Try
    End Function

    'SE OBTIENE LOS LOTES TIPO ENTRADAS CON EXISTENCIAS PARA DISPONERLOS COMO SALIDAS.
    Public Function getMovLotesSalidas(_lote As tbLote) As DataTable
        Dim odtConsulta As New DataTable
        Try
            DBConn.Open()
            DBcomando = New SqlCommand("select IDMOV, U_SO1_NUMEROARTICULO, U_SO1_FOLIO, U_SO1_NUMPARTIDA, IDLOTE, ALMACEN, TIPOMOV, DOCTO, FECHADOC = convert(varchar, FECHADOC, 103), SOCIO,
                                    VENDEDOR, CANTIDAD, EXISTENCIA, UBICACION, USUARIO_R1, USERID, FECHAMOV = convert(varchar, FECHAMOV, 103), CONS, NOMBRE = t2.ItemName, UNIDAD = T2.SalUnitMsr
                                  from LOTESMOV t1 inner join OITM t2 on t2.ItemCode = t1.U_SO1_NUMEROARTICULO
                                  where TIPOMOV IN ('10', 'DE', 'EP', 'EX', 'FP', 'EM', 'RP', 'E')
                                        AND T1.EXISTENCIA > 0
                                        AND T1.U_SO1_NUMEROARTICULO = @U_SO1_NUMEROARTICULO
                                        AND not exists (select 1
                                                        from LOTESMOV t10
                                                        where T10.TIPOMOV NOT IN ('10', 'DE', 'EP', 'EX', 'FP', 'EM', 'RP', 'E', '99')
                                                              AND t10.U_SO1_NUMEROARTICULO = t1.U_SO1_NUMEROARTICULO
                                                              and t10.IDLOTE = t1.IDLOTE
                                                              and t10.U_SO1_NUMEROARTICULO = @U_SO1_NUMEROARTICULO
                                                              and t10.U_SO1_NUMPARTIDA = @U_SO1_NUMPARTIDA
                                                              AND T10.U_SO1_FOLIO = @U_SO1_FOLIO)", DBConn)
            DBcomando.Parameters.AddWithValue("@U_SO1_NUMEROARTICULO", _lote.U_SO1_NUMEROARTICULO)
            DBcomando.Parameters.AddWithValue("@U_SO1_NUMPARTIDA", _lote.U_SO1_NUMPARTIDA)
            DBcomando.Parameters.AddWithValue("@U_SO1_FOLIO", _lote.U_SO1_FOLIO)

            DBAdaptador = New SqlDataAdapter(DBcomando)

            DBAdaptador.Fill(odtConsulta)

            Return odtConsulta
        Catch ex As Exception
            Return Nothing
        Finally
            DBConn.Close()
        End Try
    End Function

    Public Sub setMovtoSalida(oLote As tbLote, ByRef asError As String)
        Dim odtProveedores As New DataTable
        Dim fechaActual = Date.Now()
        Dim yaExiste As Integer = 0

        If DBConn.State = ConnectionState.Closed Then DBConn.Open()

        Dim myTrans As SqlTransaction = DBConn.BeginTransaction()

        Try


            ' AFECTA INVENTARIO.
            DBcomando = New SqlCommand("UPDATE LOTESMOV
                                  set EXISTENCIA = EXISTENCIA - @ctdSalida,
                                       IMPRIMIR = 'S'
                                  where idmov = @idmov ", DBConn)
            DBcomando.Parameters.AddWithValue("@idmov", oLote.IDMOVORIGEN)
            DBcomando.Parameters.AddWithValue("@ctdSalida", oLote.CANTIDAD)
            DBcomando.Transaction = myTrans
            DBcomando.ExecuteNonQuery()

            oLote.CONS = 0

            If oLote.IDMOV = 0 Then


                If oLote.CANTIDAD <= 0 Then
                    asError = "Deb de capturar valor mayor a 0 en el campo CANTIDAD."
                    Return
                End If


                DBcomando = New SqlCommand("select yaExiste = isnull(count(1), 0)
                                      from LOTESMOV
                                      where U_SO1_NUMEROARTICULO = @itemcode
                                            AND U_SO1_FOLIO = @folio
                                            AND U_SO1_NUMPARTIDA = @numPartida
                                            AND IDLOTE = @idlote
                                            AND ALMACEN = @almacen
                                            AND TIPOMOV = @tipomov ", DBConn)
                DBcomando.Parameters.AddWithValue("@itemcode", oLote.U_SO1_NUMEROARTICULO)
                DBcomando.Parameters.AddWithValue("@folio", oLote.U_SO1_FOLIO)
                DBcomando.Parameters.AddWithValue("@numPartida", oLote.U_SO1_NUMPARTIDA)
                DBcomando.Parameters.AddWithValue("@idlote", oLote.IDLOTE)
                DBcomando.Parameters.AddWithValue("@almacen", oLote.ALMACEN)
                DBcomando.Parameters.AddWithValue("@tipomov", oLote.TIPOMOV)
                DBcomando.Transaction = myTrans
                yaExiste = Convert.ToInt32(DBcomando.ExecuteScalar())

                If yaExiste > 0 Then
                    asError = "ya existe un movimiento de salida."
                    Return
                End If

                DBcomando = New SqlCommand("INSERT INTO LOTESMOV(U_SO1_NUMEROARTICULO,U_SO1_FOLIO,U_SO1_NUMPARTIDA,IDLOTE,ALMACEN,TIPOMOV,DOCTO,FECHADOC,SOCIO,
                                                            VENDEDOR,CANTIDAD,EXISTENCIA,UBICACION,USUARIO_R1,USERID,FECHAMOV,CONS, IDMOVORIGEN)
                                    values(@U_SO1_NUMEROARTICULO, @U_SO1_FOLIO, @U_SO1_NUMPARTIDA, @IDLOTE, @ALMACEN, @TIPOMOV, @DOCTO,  @FECHAMOV, @SOCIO,
                                                    @VENDEDOR, @CANTIDAD, @EXISTENCIA, @UBICACION, @USUARIO_R1, @USERID, @FECHAMOV, @CONS, @IDMOVORIGEN) ", DBConn)


            Else
                DBcomando = New SqlCommand("UPDATE LOTESMOV
                                    SET U_SO1_NUMEROARTICULO = @U_SO1_NUMEROARTICULO,
                                        U_SO1_FOLIO = @U_SO1_FOLIO,
                                        U_SO1_NUMPARTIDA = @U_SO1_NUMPARTIDA,
                                        IDLOTE = @IDLOTE,
                                        ALMACEN = @ALMACEN,
                                        TIPOMOV = @TIPOMOV,
                                        DOCTO = @DOCTO,
                                        FECHADOC = @FECHADOC,
                                        SOCIO = @SOCIO,
                                        VENDEDOR = @VENDEDOR,
                                        CANTIDAD = @CANTIDAD,
                                        EXISTENCIA = @EXISTENCIA,
                                        UBICACION = @UBICACION,
                                        USUARIO_R1 = @USUARIO_R1,
                                        USERID = @USERID,
                                        FECHAMOV = @FECHAMOV,
                                        CONS =  @CONS,
                                        IDMOVORIGEN = @IDMOVORIGEN
                                    WHERE IDMOV = @IDMOV", DBConn)
                DBcomando.Parameters.AddWithValue("@IDMOV", oLote.IDMOV)

            End If

            DBcomando.Parameters.AddWithValue("@U_SO1_NUMEROARTICULO", oLote.U_SO1_NUMEROARTICULO)
            DBcomando.Parameters.AddWithValue("@U_SO1_FOLIO", oLote.U_SO1_FOLIO)
            DBcomando.Parameters.AddWithValue("@U_SO1_NUMPARTIDA", oLote.U_SO1_NUMPARTIDA)
            DBcomando.Parameters.AddWithValue("@IDLOTE", oLote.IDLOTE)
            DBcomando.Parameters.AddWithValue("@ALMACEN", oLote.ALMACEN)
            DBcomando.Parameters.AddWithValue("@TIPOMOV", oLote.TIPOMOV)
            DBcomando.Parameters.AddWithValue("@DOCTO", oLote.DOCTO)
            If String.IsNullOrEmpty(oLote.FECHADOC) Then
                DBcomando.Parameters.AddWithValue("@FECHADOC", fechaActual)
            Else
                DBcomando.Parameters.AddWithValue("@FECHADOC", oLote.FECHADOC)
            End If

            DBcomando.Parameters.AddWithValue("@SOCIO", oLote.SOCIO)
            DBcomando.Parameters.AddWithValue("@VENDEDOR", oLote.VENDEDOR)
            DBcomando.Parameters.AddWithValue("@CANTIDAD", oLote.CANTIDAD)

            ' LAS SALIDAS NO CONTROLAN EXISTENCIAS.
            oLote.EXISTENCIA = 0
            DBcomando.Parameters.AddWithValue("@EXISTENCIA", oLote.EXISTENCIA)
            DBcomando.Parameters.AddWithValue("@UBICACION", oLote.UBICACION)
            DBcomando.Parameters.AddWithValue("@FECHAMOV", fechaActual)
            DBcomando.Parameters.AddWithValue("@USUARIO_R1", oLote.USUARIO_R1)
            DBcomando.Parameters.AddWithValue("@USERID", oLote.USERID)
            DBcomando.Parameters.AddWithValue("@CONS", oLote.CONS)
            DBcomando.Parameters.AddWithValue("@IDMOVORIGEN", oLote.IDMOVORIGEN)
            DBcomando.Transaction = myTrans
            DBcomando.ExecuteNonQuery()

            myTrans.Commit()

        Catch ex As Exception
            asError = ex.Message.ToString()
            myTrans.Rollback()
            Return
        Finally
            If DBConn.State = ConnectionState.Open Then DBConn.Close()
        End Try
    End Sub

    Public Sub delMovtoSalida(oLote As tbLote, ByRef asError As String)
        Dim odtProveedores As New DataTable
        Dim fechaActual = Date.Now()
        Dim yaExiste As Integer = 0

        If DBConn.State = ConnectionState.Closed Then DBConn.Open()

        Dim myTrans As SqlTransaction = DBConn.BeginTransaction()

        Try


            ' AFECTA INVENTARIO.
            DBcomando = New SqlCommand("UPDATE LOTESMOV
                                  set EXISTENCIA = EXISTENCIA + @ctdSalida,
                                       IMPRIMIR = 'S'
                                  where idmov = @idmov ", DBConn)
            DBcomando.Parameters.AddWithValue("@idmov", oLote.IDMOVORIGEN)
            DBcomando.Parameters.AddWithValue("@ctdSalida", oLote.CANTIDAD)
            DBcomando.Transaction = myTrans
            DBcomando.ExecuteNonQuery()

            oLote.CONS = 0
            oLote.TIPOMOV = "99"

            DBcomando = New SqlCommand("UPDATE LOTESMOV
                                    SET TIPOMOV = @TIPOMOV
                                    WHERE IDMOV = @IDMOV", DBConn)
            DBcomando.Parameters.AddWithValue("@IDMOV", oLote.IDMOV)
            DBcomando.Parameters.AddWithValue("@TIPOMOV", oLote.TIPOMOV)

            DBcomando.Transaction = myTrans
            DBcomando.ExecuteNonQuery()

            myTrans.Commit()

        Catch ex As Exception
            asError = ex.Message.ToString()
            myTrans.Rollback()
            Return
        Finally
            If DBConn.State = ConnectionState.Open Then DBConn.Close()
        End Try
    End Sub


    ' OBTENER EXISTENCIA ACTUAL DEL MOVIMIENTO EN PROCESO.
    Public Function getMovExitencia(oLote As tbLote, ByRef asError As String) As Decimal
        Dim odtConsulta As New DataTable
        Dim ctdLotes As Decimal = 0
        Try
            DBConn.Open()
            DBcomando = New SqlCommand("select ISNULL(SUM(T1.CANTIDAD), 0)
                                  from LOTESMOV T1
                                  where TIPOMOV <> '99' 
                                        AND T1.U_SO1_NUMEROARTICULO = @U_SO1_NUMEROARTICULO
                                        AND T1.U_SO1_FOLIO = @U_SO1_FOLIO
                                        AND T1.U_SO1_NUMPARTIDA = @U_SO1_NUMPARTIDA ", DBConn)
            DBcomando.Parameters.AddWithValue("@U_SO1_NUMEROARTICULO", oLote.U_SO1_NUMEROARTICULO)
            DBcomando.Parameters.AddWithValue("@U_SO1_FOLIO", oLote.U_SO1_FOLIO)
            DBcomando.Parameters.AddWithValue("@U_SO1_NUMPARTIDA", oLote.U_SO1_NUMPARTIDA)

            ctdLotes = Convert.ToDecimal(DBcomando.ExecuteScalar())

            Return ctdLotes
        Catch ex As Exception
            Return 0
        Finally
            DBConn.Close()
        End Try
    End Function

    Public Function getMovLotesEntradas(_lote As tbLote) As DataTable
        Dim odtConsulta As New DataTable
        Try
            DBConn.Open()
            DBcomando = New SqlCommand("select IDMOV, U_SO1_NUMEROARTICULO, U_SO1_FOLIO, U_SO1_NUMPARTIDA, IDLOTE, ALMACEN, TIPOMOV, DOCTO, FECHADOC = convert(varchar, FECHADOC, 103), SOCIO,
                                    VENDEDOR, CANTIDAD, EXISTENCIA, UBICACION, USUARIO_R1, USERID, FECHAMOV = convert(varchar, FECHAMOV, 103), CONS, NOMBRE = t2.ItemName, UNIDAD = T2.SalUnitMsr
                                  from LOTESMOV t1 inner join OITM t2 on t2.ItemCode = t1.U_SO1_NUMEROARTICULO
                                  where T1.TIPOMOV <> '99'
                                       AND T1.U_SO1_NUMEROARTICULO = @U_SO1_NUMEROARTICULO
                                       AND T1.U_SO1_NUMPARTIDA = @U_SO1_NUMPARTIDA
                                        AND T1.U_SO1_FOLIO = @U_SO1_FOLIO ", DBConn)
            DBcomando.Parameters.AddWithValue("@U_SO1_NUMEROARTICULO", _lote.U_SO1_NUMEROARTICULO)
            DBcomando.Parameters.AddWithValue("@U_SO1_NUMPARTIDA", _lote.U_SO1_NUMPARTIDA)
            DBcomando.Parameters.AddWithValue("@U_SO1_FOLIO", _lote.U_SO1_FOLIO)

            DBAdaptador = New SqlDataAdapter(DBcomando)

            DBAdaptador.Fill(odtConsulta)

            Return odtConsulta
        Catch ex As Exception
            Return Nothing
        Finally
            DBConn.Close()
        End Try
    End Function

    ' PARA EL CASO DE LOTES DE NUEVO INGRESO POR MOVIMIENTOS DE ENTRADAS/SALIDAS.
    Public Sub setMovLoteEntrada(oLote As tbLote, ByRef asError As String)
        Dim odtProveedores As New DataTable
        Dim fechaActual = Date.Now()
        Dim cons As Integer = 0
        Try
            If DBConn.State = ConnectionState.Closed Then DBConn.Open()

            If oLote.IDMOV = 0 Then

                If String.IsNullOrEmpty(oLote.IDLOTE) Then

                    If oLote.CANTIDAD <= 0 Then
                        asError = "Deb de capturar valor mayor a 0 en el campo CANTIDAD."
                        Return
                    End If
                    DBcomando = New SqlCommand("select isnull(max(cons), 0) from LOTESMOV where convert(varchar, fechamov, 112) = convert(varchar, @fecha, 112) ", DBConn)
                    DBcomando.Parameters.AddWithValue("@fecha", fechaActual)
                    cons = Convert.ToInt32(DBcomando.ExecuteScalar())
                    cons = cons + 1
                    oLote.CONS = cons

                    DBcomando = New SqlCommand("INSERT INTO LOTESMOV(U_SO1_NUMEROARTICULO,U_SO1_FOLIO,U_SO1_NUMPARTIDA,IDLOTE,ALMACEN,TIPOMOV,DOCTO,FECHADOC,SOCIO,
                                                            VENDEDOR,CANTIDAD,EXISTENCIA,UBICACION,USUARIO_R1,USERID,FECHAMOV,CONS, IMPRIMIR)
                                    values(@U_SO1_NUMEROARTICULO, @U_SO1_FOLIO, @U_SO1_NUMPARTIDA, convert(varchar, @FECHAMOV, 112) + '-' + RIGHT('000' + Ltrim(Rtrim(cast(@CONS as varchar))),3), @ALMACEN, @TIPOMOV, @DOCTO,  @FECHAMOV, @SOCIO,
                                                    @VENDEDOR, @CANTIDAD, @EXISTENCIA, @UBICACION, @USUARIO_R1, @USERID, @FECHAMOV, @CONS, 'S') ", DBConn)
                Else
                    DBcomando = New SqlCommand("INSERT INTO LOTESMOV(U_SO1_NUMEROARTICULO,U_SO1_FOLIO,U_SO1_NUMPARTIDA,IDLOTE,ALMACEN,TIPOMOV,DOCTO,FECHADOC,SOCIO,
                                                            VENDEDOR,CANTIDAD,EXISTENCIA,UBICACION,USUARIO_R1,USERID,FECHAMOV,CONS)
                                    values(@U_SO1_NUMEROARTICULO, @U_SO1_FOLIO, @U_SO1_NUMPARTIDA, @IDLOTE, @ALMACEN, @TIPOMOV, @DOCTO,  @FECHAMOV, @SOCIO,
                                                    @VENDEDOR, @CANTIDAD, @EXISTENCIA, @UBICACION, @USUARIO_R1, @USERID, @FECHAMOV, @CONS, 'S') ", DBConn)
                End If

            Else
                DBcomando = New SqlCommand("UPDATE LOTESMOV
                                    SET U_SO1_NUMEROARTICULO = @U_SO1_NUMEROARTICULO,
                                        U_SO1_FOLIO = @U_SO1_FOLIO,
                                        U_SO1_NUMPARTIDA = @U_SO1_NUMPARTIDA,
                                        IDLOTE = @IDLOTE,
                                        ALMACEN = @ALMACEN,
                                        TIPOMOV = @TIPOMOV,
                                        DOCTO = @DOCTO,
                                        FECHADOC = @FECHADOC,
                                        SOCIO = @SOCIO,
                                        VENDEDOR = @VENDEDOR,
                                        CANTIDAD = @CANTIDAD,
                                        EXISTENCIA = @EXISTENCIA,
                                        UBICACION = @UBICACION,
                                        USUARIO_R1 = @USUARIO_R1,
                                        USERID = @USERID,
                                        FECHAMOV = @FECHAMOV,
                                        CONS =  @CONS
                                    WHERE IDMOV = @IDMOV", DBConn)
                DBcomando.Parameters.AddWithValue("@IDMOV", oLote.IDMOV)

            End If

            If String.IsNullOrEmpty(oLote.DOCTO) Then oLote.DOCTO = "*NA"
            If String.IsNullOrEmpty(oLote.U_SO1_FOLIO) Then oLote.U_SO1_FOLIO = "*NA"
            If String.IsNullOrEmpty(oLote.U_SO1_NUMPARTIDA) Then oLote.U_SO1_NUMPARTIDA = "*NA"
            If String.IsNullOrEmpty(oLote.SOCIO) Then oLote.SOCIO = "*NA"
            If String.IsNullOrEmpty(oLote.USUARIO_R1) Then oLote.USUARIO_R1 = "*NA"

            DBcomando.Parameters.AddWithValue("@U_SO1_NUMEROARTICULO", oLote.U_SO1_NUMEROARTICULO)
            DBcomando.Parameters.AddWithValue("@U_SO1_FOLIO", oLote.U_SO1_FOLIO)
            DBcomando.Parameters.AddWithValue("@U_SO1_NUMPARTIDA", oLote.U_SO1_NUMPARTIDA)
            DBcomando.Parameters.AddWithValue("@IDLOTE", oLote.IDLOTE)
            DBcomando.Parameters.AddWithValue("@ALMACEN", oLote.ALMACEN)
            DBcomando.Parameters.AddWithValue("@TIPOMOV", oLote.TIPOMOV)
            DBcomando.Parameters.AddWithValue("@DOCTO", oLote.DOCTO)
            If String.IsNullOrEmpty(oLote.FECHADOC) Then
                DBcomando.Parameters.AddWithValue("@FECHADOC", fechaActual)
            Else
                DBcomando.Parameters.AddWithValue("@FECHADOC", oLote.FECHADOC)
            End If

            DBcomando.Parameters.AddWithValue("@SOCIO", oLote.SOCIO)
            DBcomando.Parameters.AddWithValue("@VENDEDOR", oLote.VENDEDOR)
            DBcomando.Parameters.AddWithValue("@CANTIDAD", oLote.CANTIDAD)

            oLote.EXISTENCIA = oLote.CANTIDAD
            DBcomando.Parameters.AddWithValue("@EXISTENCIA", oLote.EXISTENCIA)
            DBcomando.Parameters.AddWithValue("@UBICACION", oLote.UBICACION)
            DBcomando.Parameters.AddWithValue("@FECHAMOV", fechaActual)
            DBcomando.Parameters.AddWithValue("@USUARIO_R1", oLote.USUARIO_R1)
            DBcomando.Parameters.AddWithValue("@USERID", oLote.USERID)
            DBcomando.Parameters.AddWithValue("@CONS", oLote.CONS)
            DBcomando.ExecuteNonQuery()

        Catch ex As Exception
            asError = ex.Message.ToString()
            Return
        Finally
            If DBConn.State = ConnectionState.Open Then DBConn.Close()
        End Try
    End Sub

    ' VALIDA LAS CANTIDADES A CAMBIAR DE LA ENTRADA EDITADA.
    Public Function validaExistenciaEntrada(oLote As tbLote) As Decimal
        Dim ctdMovEntSal As Decimal
        Dim ctdEntrada As Decimal
        Dim ctdActual As Decimal

        Try
            DBConn.Open()
            DBcomando = New SqlCommand(" select ctdMovEntSal = IsNull(sum(t1.CANTIDAD * CASE WHEN T1.TIPOMOV IN ('10', 'DE', 'EP', 'EX', 'FP', 'EM', 'RP', 'E') THEN 1 ELSE -1 END), 0.00 )
                                            from LOTESMOV t1 
                                            where t1.IDLOTE = @IDLOTE 
                                                    AND t1.TIPOMOV != '99'
                                                    AND t1.IDMOV != @IDMOV ", DBConn)
            DBcomando.Parameters.AddWithValue("@IDMOV", oLote.IDMOV)
            DBcomando.Parameters.AddWithValue("@IDLOTE", oLote.IDLOTE)

            ctdMovEntSal = Convert.ToInt32(DBcomando.ExecuteScalar())
            ctdEntrada = oLote.CANTIDAD

            If (ctdMovEntSal + ctdEntrada) >= 0 Then
                ctdActual = ctdEntrada + ctdMovEntSal  ' ctdMovEntSal = viene en negativo porque siempre acumula todas las salidas.
            End If

        Catch ex As Exception
            ctdActual = 0
        Finally
            DBConn.Close()
        End Try

        Return ctdActual
    End Function

    Public Function getKardex(valorBuscar As String, fecha As String) As DataTable
        Dim odtConsulta As New DataTable
        Try
            DBConn.Open()
            DBcomando = New SqlCommand("select t1.IDMOV, T1.IDLOTE, T1.U_SO1_NUMEROARTICULO, T1.U_SO1_FOLIO, T1.U_SO1_NUMPARTIDA, T1.ALMACEN, 
			                                    FECHADOC = CONVERT(VARCHAR, T1.FECHADOC, 103), FECHAMOV = CONVERT(VARCHAR, T1.FECHAMOV, 103), T1.SOCIO, T1.VENDEDOR, 
			                                    ENTRADA = CASE WHEN T1.TIPOMOV IN ('10', 'DE', 'EP', 'EX', 'FP', 'EM', 'RP', 'E') THEN T1.CANTIDAD ELSE 0 END,
			                                    SALIDA = CASE WHEN T1.TIPOMOV IN ('CA', 'SM', 'ED', 'CR', 'NC', 'SX', 'S', 'X', 'DX') THEN T1.CANTIDAD ELSE 0 END,
			                                    T1.USUARIO_R1, T1.USERID, T2.ItemName, T1.TIPOMOV, T1.UBICACION, T1.EXISTENCIA, UNIDAD = t2.SalUnitMsr,
			                                    TIPOMOVNOMBRE = CASE WHEN T1.TIPOMOV =  '10' THEN  'INVENTARIO LOTES.'
								                                    WHEN T1.TIPOMOV =  '99' THEN  'CANCELACION DE LOTES.'
								                                    WHEN T1.TIPOMOV =  'EP' THEN  'COMPRA - ENTRDA PROVEEDOR'
								                                    WHEN T1.TIPOMOV = 'EX' THEN 'ENTRADA DE MERCANCIA'
								                                    WHEN T1.TIPOMOV = 'DE' THEN 'DEVOLUCION DE VENTAS'
								                                    WHEN T1.TIPOMOV = 'FP' THEN 'COMPRA - FACTURA PROVEEDOR'
								                                    WHEN T1.TIPOMOV = 'EM' THEN 'ENTRADA MERCANCIA'
								                                    WHEN T1.TIPOMOV = 'RP' THEN 'RECIBO PRODUCCION'

								                                    WHEN T1.TIPOMOV = 'CA' THEN 'VENTA'
								                                    WHEN T1.TIPOMOV = 'CR' THEN 'VENTA FACTURA CREDITO'
								                                    WHEN T1.TIPOMOV = 'SM' THEN 'SALIDA MERCANCIA'
								                                    WHEN T1.TIPOMOV = 'ED' THEN 'VENTA'
								                                    WHEN T1.TIPOMOV = 'NC' THEN 'NOTA DE CREDITO'
								                                    WHEN T1.TIPOMOV = 'SX' THEN 'SALIDA DE MERCANCIA' 
								                                    ELSE 'OTRO TIPO DE MOV' END
                                    from LOTESMOV T1 INNER JOIN OITM T2 ON T2.ItemCode = T1.U_SO1_NUMEROARTICULO
                                    where CONVERT(VARCHAR, T1.FECHAMOV, 112) <= CONVERT(VARCHAR, CAST(@FECHA AS DATE), 112) 
                                          AND (
                                                T1.IDLOTE LIKE '%' + @valorBuscar + '%' 
                                                OR T1.U_SO1_NUMEROARTICULO LIKE  '%' + @valorBuscar + '%'
                                                OR T2.ItemName LIKE  '%' + @valorBuscar + '%'
                                                OR T1.USUARIO_R1 LIKE  '%' + @valorBuscar + '%'
                                                OR T1.SOCIO LIKE '%' + @valorBuscar + '%'
                                                OR T1.VENDEDOR LIKE '%' + @valorBuscar + '%'
                                                OR @valorBuscar = 'TODOS'
                                              )
                                    ORDER BY IDLOTE, IDMOV ", DBConn)
            DBcomando.Parameters.AddWithValue("@valorBuscar", valorBuscar)
            DBcomando.Parameters.AddWithValue("@FECHA", fecha)

            DBAdaptador = New SqlDataAdapter(DBcomando)

            DBAdaptador.Fill(odtConsulta)

            Return odtConsulta
        Catch ex As Exception
            Return Nothing
        Finally
            DBConn.Close()
        End Try
    End Function

    ' CONSULTA PARA USUARIOS DE VENTAS - SOLO SE MUESTRA ENTRADAS DE LOTES CON EXISTENCIAS.
    Public Function getLotesVentas(valorBuscar As String) As DataTable
        Dim odtConsulta As New DataTable
        Try
            DBConn.Open()
            DBcomando = New SqlCommand("select T1.IDLOTE, T1.U_SO1_NUMEROARTICULO, T1.U_SO1_FOLIO, T1.U_SO1_NUMPARTIDA, T1.ALMACEN, 
			                                    FECHADOC = CONVERT(VARCHAR, T1.FECHADOC, 103), FECHAMOV = CONVERT(VARCHAR, T1.FECHAMOV, 103), T1.SOCIO, T1.VENDEDOR, 
			                                    ENTRADA = CASE WHEN T1.TIPOMOV IN ('10', 'DE', 'EP', 'EX', 'FP', 'EM', 'RP', 'E') THEN T1.CANTIDAD ELSE 0 END,
			                                    SALIDA = CASE WHEN T1.TIPOMOV IN ('CA', 'SM', 'ED', 'CR', 'NC', 'SX', 'S', 'X', 'DX' ) THEN T1.CANTIDAD ELSE 0 END,
			                                    T1.USUARIO_R1, T1.USERID, T2.ItemName, T1.TIPOMOV, T1.EXISTENCIA, T1.UBICACION,
			                                    TIPOMOVNOMBRE = CASE WHEN T1.TIPOMOV =  '10' THEN  'INVENTARIO LOTES.'
								                                    WHEN T1.TIPOMOV =  '99' THEN  'CANCELACION DE LOTES.'
								                                    WHEN T1.TIPOMOV =  'EP' THEN  'COMPRA - ENTRDA PROVEEDOR'
								                                    WHEN T1.TIPOMOV = 'EX' THEN 'ENTRADA DE MERCANCIA'
								                                    WHEN T1.TIPOMOV = 'DE' THEN 'DEVOLUCION DE VENTAS'
								                                    WHEN T1.TIPOMOV = 'FP' THEN 'COMPRA - FACTURA PROVEEDOR'
								                                    WHEN T1.TIPOMOV = 'EM' THEN 'ENTRADA MERCANCIA'
								                                    WHEN T1.TIPOMOV = 'RP' THEN 'RECIBO PRODUCCION'

								                                    WHEN T1.TIPOMOV = 'CA' THEN 'VENTA'
								                                    WHEN T1.TIPOMOV = 'CR' THEN 'VENTA FACTURA CREDITO'
								                                    WHEN T1.TIPOMOV = 'SM' THEN 'SALIDA MERCANCIA'
								                                    WHEN T1.TIPOMOV = 'ED' THEN 'VENTA'
								                                    WHEN T1.TIPOMOV = 'NC' THEN 'NOTA DE CREDITO'
								                                    WHEN T1.TIPOMOV = 'SX' THEN 'SALIDA DE MERCANCIA' 
								                                    ELSE 'OTRO TIPO DE MOV' END
                                    from LOTESMOV T1 INNER JOIN OITM T2 ON T2.ItemCode = T1.U_SO1_NUMEROARTICULO
                                    where T1.TIPOMOV <> '99'
                                          AND T1.EXISTENCIA > 0
                                          AND (
                                                T1.IDLOTE LIKE '%' + @valorBuscar + '%' 
                                                OR T1.U_SO1_NUMEROARTICULO LIKE  '%' + @valorBuscar + '%'
                                                OR T2.ItemName LIKE  '%' + @valorBuscar + '%'
                                                OR @valorBuscar = 'TODOS'
                                              )
                                    ORDER BY IDLOTE, IDMOV ", DBConn)
            DBcomando.Parameters.AddWithValue("@valorBuscar", valorBuscar)

            DBAdaptador = New SqlDataAdapter(DBcomando)

            DBAdaptador.Fill(odtConsulta)

            Return odtConsulta
        Catch ex As Exception
            Return Nothing
        Finally
            DBConn.Close()
        End Try
    End Function


    Public Function getLotesImprimir(valorBuscar As String) As DataTable
        Dim odtConsulta As New DataTable
        Try
            DBConn.Open()
            DBcomando = New SqlCommand("select T1.IDMOV, T1.IDLOTE, T1.U_SO1_NUMEROARTICULO, T1.U_SO1_FOLIO, T1.U_SO1_NUMPARTIDA, T1.ALMACEN, UNIDAD = t2.SalUnitMsr,
			                                    FECHADOC = CONVERT(VARCHAR, T1.FECHADOC, 103), FECHAMOV = CONVERT(VARCHAR, T1.FECHAMOV, 103), T1.SOCIO, T1.VENDEDOR, 
			                                    ENTRADA = CASE WHEN T1.TIPOMOV IN ('10', 'DE', 'EP', 'EX', 'FP', 'EM', 'RP', 'E') THEN T1.CANTIDAD ELSE 0 END,
			                                    SALIDA = CASE WHEN T1.TIPOMOV IN ('CA', 'SM', 'ED', 'CR', 'NC', 'SX', 'S', 'X', 'DX') THEN T1.CANTIDAD ELSE 0 END,
			                                    T1.USUARIO_R1, T1.USERID, T2.ItemName, T1.TIPOMOV, T1.EXISTENCIA, T1.UBICACION, IMPRIMIR = ISNULL(T1.IMPRIMIR, 'N'),
			                                    TIPOMOVNOMBRE = CASE WHEN T1.TIPOMOV =  '10' THEN  'INVENTARIO LOTES.'
								                                    WHEN T1.TIPOMOV =  '99' THEN  'CANCELACION DE LOTES.'
								                                    WHEN T1.TIPOMOV =  'EP' THEN  'COMPRA - ENTRDA PROVEEDOR'
								                                    WHEN T1.TIPOMOV = 'EX' THEN 'ENTRADA DE MERCANCIA'
								                                    WHEN T1.TIPOMOV = 'DE' THEN 'DEVOLUCION DE VENTAS'
								                                    WHEN T1.TIPOMOV = 'FP' THEN 'COMPRA - FACTURA PROVEEDOR'
								                                    WHEN T1.TIPOMOV = 'EM' THEN 'ENTRADA MERCANCIA'
								                                    WHEN T1.TIPOMOV = 'RP' THEN 'RECIBO PRODUCCION'

								                                    WHEN T1.TIPOMOV = 'CA' THEN 'VENTA'
								                                    WHEN T1.TIPOMOV = 'CR' THEN 'VENTA FACTURA CREDITO'
								                                    WHEN T1.TIPOMOV = 'SM' THEN 'SALIDA MERCANCIA'
								                                    WHEN T1.TIPOMOV = 'ED' THEN 'VENTA'
								                                    WHEN T1.TIPOMOV = 'NC' THEN 'NOTA DE CREDITO'
								                                    WHEN T1.TIPOMOV = 'SX' THEN 'SALIDA DE MERCANCIA' 
								                                    ELSE 'OTRO TIPO DE MOV' END
                                    from LOTESMOV T1 INNER JOIN OITM T2 ON T2.ItemCode = T1.U_SO1_NUMEROARTICULO
                                    where T1.TIPOMOV <> '99'
                                          AND T1.EXISTENCIA > 0
                                          AND (
                                                (
                                                  ISNULL(T1.IMPRIMIR, 'N') = 'S' AND @valorBuscar = 'TODOS'
                                                )
                                                OR
                                                (
                                                  T1.IDLOTE = @valorBuscar
                                                )
                                              )
                                    ORDER BY IDLOTE, IDMOV ", DBConn)
            DBcomando.Parameters.AddWithValue("@valorBuscar", valorBuscar)

            DBAdaptador = New SqlDataAdapter(DBcomando)

            DBAdaptador.Fill(odtConsulta)

            Return odtConsulta
        Catch ex As Exception
            Return Nothing
        Finally
            DBConn.Close()
        End Try
    End Function

    ' CAMBIA EL VALOR DE IMPRESIÓN.
    Public Function setEstatusImprimir(oLote As tbLote) As Boolean
        Dim odtConsulta As New DataTable
        Dim resultado As Boolean = True

        Try
            DBConn.Open()
            DBcomando = New SqlCommand("UPDATE LOTESMOV
                                     SET IMPRIMIR = 'N'
                                    WHERE IDMOV = @IDMOV ", DBConn)
            DBcomando.Parameters.AddWithValue("@IDMOV", oLote.IDMOV)

            DBcomando.ExecuteNonQuery()

        Catch ex As Exception
            resultado = False
        Finally
            DBConn.Close()
        End Try

        Return resultado
    End Function


    Public Function GenerarPDF(oLotes() As tbLote) As String
        ' Dim oDoc As New iTextSharp.text.Document(PageSize.A4, 0, 0, 0, 0) 

        Dim pgSize As New iTextSharp.text.Rectangle(283.46, 141.732, RotateFlipType.Rotate90FlipNone)
        Dim oDoc As New iTextSharp.text.Document(pgSize, 0, 0, 0, 0)
        Dim pdfw As iTextSharp.text.pdf.PdfWriter
        Dim cb As PdfContentByte
        Dim fuente As iTextSharp.text.pdf.BaseFont
        Dim NombreArchivo As String = "C:\ejemplo.pdf"
        Dim pdfBase64 As String = ""
        Dim ctdLotesImpresos As Integer = 0
        Dim fechaActual = Date.Now()

        Dim dtLotes As New DataTable

        Try

            Dim ms As New MemoryStream()

            pdfw = PdfWriter.GetInstance(oDoc, ms)
            'Apertura del documento.
            oDoc.Open()
            cb = pdfw.DirectContent

            ' SE OBTIENEN LAS ETIQUETAS POR IMPRIMIR
            ' dtLotes = getLotesVentas("TODOS")

            For Each lote As tbLote In oLotes

                If lote.IMPRIMIR = "N" Then Continue For

                If Not setEstatusImprimir(lote) Then Continue For

                'Agregamos una pagina.
                ' If ctdLotesImpresos = 4 Or ctdLotesImpresos = 0 Then
                oDoc.NewPage()
                ctdLotesImpresos = 0
                ' End If

                ctdLotesImpresos = ctdLotesImpresos + 1

                'Iniciamos el flujo de bytes.
                cb.BeginText()
                'Instanciamos el objeto para la tipo de letra.
                fuente = FontFactory.GetFont(FontFactory.HELVETICA, iTextSharp.text.Font.DEFAULTSIZE, iTextSharp.text.Font.NORMAL).BaseFont
                'Seteamos el tipo de letra y el tamaño.
                cb.SetFontAndSize(fuente, 14)
                'Seteamos el color del texto a escribir.
                ' cb.SetColorFill(iTextSharp.text.Color.BLACK)
                'Aqui es donde se escribe el texto.
                'Aclaracion: Por alguna razon la coordenada vertical siempre es tomada desde el borde inferior (de ahi que se calcule como "PageSize.A4.Height - 50")
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, lote.IDLOTE, 10, 130, 0)

                ' GENERA BAR CODE,

                Dim cbx As New PdfContentByte(pdfw)
                Dim BarcodeLineas As New Barcode128()

                BarcodeLineas.BarHeight = 30
                BarcodeLineas.Code = lote.IDLOTE
                BarcodeLineas.GenerateChecksum = True
                BarcodeLineas.CodeType = Barcode128.CODE128

                ' Dim imganX As iTextSharp.text.Image = BarcodeLineas.CreateImageWithBarcode(cb, BaseColor.BLACK, BaseColor.BLACK)

                Dim bm As New Bitmap(BarcodeLineas.CreateDrawingImage(Color.Black, Color.White))
                Dim smBMP As New MemoryStream
                bm.Save(smBMP, ImageFormat.Bmp)
                ' bm.Save(smBMP, bm.RawFormat)
                Dim bmBytes As Byte() = smBMP.ToArray
                Dim imagen As iTextSharp.text.Image = iTextSharp.text.Image.GetInstance(bmBytes)
                ' Dim imagen As iTextSharp.text.Image = New Bitmap(BarcodeLineas.CreateDrawingImage(Color.Black, Color.White), 50, 50)
                ' Dim imagen As iTextSharp.text.Image = BarcodeLineas.CreateDrawingImage(Color.Black, Color.White)
                '  imagen = New Bitmap(BarcodeLineas.CreateDrawingImage(Color.Black, Color.White), 50, 50)
                imagen.ScaleToFit(150.0F, 40.0F)
                imagen.SetAbsolutePosition(10, 70)
                'imganX.SetAbsolutePosition(5.5F, 320)
                cb.AddImage(imagen)
                ' oDoc.Add(imagen)


                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, lote.EXISTENCIA, 50, 50, 0)
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, lote.UNIDAD, 85, 50, 0)
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, lote.UBICACION, 150, 50, 0)

                cb.SetFontAndSize(fuente, 8)
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, lote.U_SO1_NUMEROARTICULO, 10, 40, 0)
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, lote.ItemName, 10, 30, 0)


                cb.SetFontAndSize(fuente, 6)
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Imp: " + lote.USERID, 10, 10, 0)
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Fecha Impresión: " + fechaActual.ToString("dd/MM/yyyy HH:mm"), 150, 10, 0)
                'Fin del flujo de bytes.
                cb.EndText()
            Next

            'Forzamos vaciamiento del buffer.
            pdfw.Flush()
            'Cerramos el documento.
            oDoc.Close()

            ' Dim streamLength As Integer = Convert.ToInt32(ms.Length)

            ' Dim fileData As Byte() = New Byte(streamLength) {}
            Dim fileData As Byte() = ms.ToArray()

            ' Read the file into a byte array
            ' ms.Read(fileData, 0, streamLength)
            ms.Flush()
            ms.Close()

            pdfBase64 = Convert.ToBase64String(fileData)

        Catch ex As Exception

            Return Nothing
        Finally
            cb = Nothing
            pdfw = Nothing
            oDoc = Nothing
        End Try

        Return pdfBase64

    End Function


    Public Function setLoteUbicacion(oLote As tbLote, ByRef _error As String) As Boolean
        Dim resutado As Boolean = True
        Dim ctdExistencia As Decimal

        Try

            ' VALIDAR EXISTENCIA EN CASO DE CAMBIAR LA EXISTENCIA DE MOVIMIENTO DE ENTRADA.
            ctdExistencia = validaExistenciaEntrada(oLote)
            If ctdExistencia < 0 Then
                _error = "La cantidad capturada no cubre las salidas correspndintes al lote editado."
                Return False
            End If

            DBConn.Open()

            DBcomando = New SqlCommand("update LOTESMOV
                                  set UBICACION = @UBICACION,
                                      CANTIDAD = @CANTIDAD,
                                      EXISTENCIA = @EXISTENCIA,
                                      IMPRIMIR = 'S'
	                                 where IDMOV = @IDMOV ", DBConn)
            DBcomando.Parameters.AddWithValue("@UBICACION", oLote.UBICACION)
            DBcomando.Parameters.AddWithValue("@EXISTENCIA", ctdExistencia)
            DBcomando.Parameters.AddWithValue("@CANTIDAD", oLote.CANTIDAD)
            DBcomando.Parameters.AddWithValue("@IDMOV", oLote.IDMOV)

            DBcomando.ExecuteNonQuery()

        Catch ex As Exception
            resutado = False
            _error = ex.Message.ToString()
        Finally
            DBConn.Close()
        End Try

        Return resutado
    End Function

    Public Function getLogin(asUsuario As String, asPassw As String, ByRef _error As String) As Boolean

        Dim resutado As Boolean = True
        Dim pass As String = ""
        Try
            DBConn.Open()
            DBcomando = New SqlCommand("select passw = CONVERT(NVARCHAR(MAX), DECRYPTBYPASSPHRASE('Electrico2012', T2.PASSW))
                                  from  [@SO1_01AUTEMPROL] T1 INNER JOIN LOTESUSUARIOS T2 ON T2.CODE = T1.CODE
	                                 where t1.CODE = @USUARIO ", DBConn)
            DBcomando.Parameters.AddWithValue("@USUARIO", asUsuario)

            pass = DBcomando.ExecuteScalar()

            If pass <> asPassw Then

                pass = ""

                DBcomando = New SqlCommand("select passw = CONVERT(VARCHAR(MAX), DECRYPTBYPASSPHRASE('Electrico2012', T2.PASSW))
                                  from  [@SO1_01AUTEMPROL] T1 INNER JOIN LOTESUSUARIOS T2 ON T2.CODE = T1.CODE
	                                 where t1.CODE = @USUARIO ", DBConn)
                DBcomando.Parameters.AddWithValue("@USUARIO", asUsuario)

                pass = DBcomando.ExecuteScalar()

                If pass <> asPassw Then
                    resutado = False
                End If

            End If

        Catch ex As Exception
            resutado = False
            _error = ex.Message.ToString()
        Finally
            DBConn.Close()
        End Try

        Return resutado
    End Function

    Public Function getMenuUsuario(asUsuario As String, ByRef _error As String) As DataTable

        Dim odtConsulta As New DataTable
        Try
            DBConn.Open()
            DBcomando = New SqlCommand("select t1.IDMENU, t1.NOMBRE
                                  from  LOTESROLDET T1 INNER JOIN LOTESUSUARIOS T2 ON T2.IDROL = T1.IDROL
	                                 where T2.CODE = @USUARIO
                                    ORDER BY T1.ORDEN ", DBConn)
            DBcomando.Parameters.AddWithValue("@USUARIO", asUsuario)

            DBAdaptador = New SqlDataAdapter(DBcomando)

            DBAdaptador.Fill(odtConsulta)

        Catch ex As Exception
            odtConsulta = Nothing
            _error = ex.Message.ToString()
        Finally
            DBConn.Close()
        End Try

        Return odtConsulta
    End Function

    Public Function getPerfiles(ByRef _error As String) As DataTable

        Dim odtConsulta As New DataTable
        Try
            DBConn.Open()
            DBcomando = New SqlCommand("select t1.IDROL, t1.NOMBRE
                                  from  LOTESROL T1 ", DBConn)
            ' DBcomando.Parameters.AddWithValue("@USUARIO", asUsuario)

            DBAdaptador = New SqlDataAdapter(DBcomando)

            DBAdaptador.Fill(odtConsulta)

        Catch ex As Exception
            odtConsulta = Nothing
            _error = ex.Message.ToString()
        Finally
            DBConn.Close()
        End Try

        Return odtConsulta
    End Function

    Public Function setPerfil(otbPerfil As tbPerfil, ByRef _error As String) As Boolean
        Dim existeSINO As Integer = 0
        Dim odtConsulta As New DataTable
        Dim resultado As Boolean = True

        Try
            DBConn.Open()
            DBcomando = New SqlCommand("select existe = isnull(count(1), 0)
                                  from  LOTESROL
                                  where idrol = @IDROL", DBConn)
            DBcomando.Parameters.AddWithValue("@IDROL", otbPerfil.IDROL)

            existeSINO = Convert.ToInt32(DBcomando.ExecuteScalar())

            If existeSINO > 0 Then
                DBcomando = New SqlCommand(" update LOTESROL
                                        set nombre = @NOMBRE
                                     where idrol = @IDROL", DBConn)
                DBcomando.Parameters.AddWithValue("@IDROL", otbPerfil.IDROL)
                DBcomando.Parameters.AddWithValue("@NOMBRE", otbPerfil.NOMBRE)

            Else

                DBcomando = New SqlCommand(" INSERT INTO LOTESROL (IDROL, NOMBRE)
                                        VALUES (@IDROL, @NOMBRE) ", DBConn)
                DBcomando.Parameters.AddWithValue("@IDROL", otbPerfil.IDROL)
                DBcomando.Parameters.AddWithValue("@NOMBRE", otbPerfil.NOMBRE)

            End If

            DBcomando.ExecuteNonQuery()

        Catch ex As Exception
            resultado = False
            _error = ex.Message.ToString()
        Finally
            DBConn.Close()
        End Try

        Return resultado
    End Function

    Public Function delPerfil(otbPerfil As tbPerfil, ByRef _error As String) As Boolean
        Dim existeSINO As Integer = 0
        Dim odtConsulta As New DataTable
        Dim resultado As Boolean = True

        ' ESTA FUNCION SOLO ELIMINA EL REGISTRO ,

        Try
            DBConn.Open()

            DBcomando = New SqlCommand(" DELETE FROM LOTESROL
                                    WHERE IDROL = @IDROL ", DBConn)
            DBcomando.Parameters.AddWithValue("@IDROL", otbPerfil.IDROL)

            DBcomando.ExecuteNonQuery()

        Catch ex As Exception
            resultado = False
            _error = ex.Message.ToString()
        Finally
            DBConn.Close()
        End Try

        Return resultado
    End Function


    Public Function getPerfilMenu(idrol As String, ByRef _error As String) As DataTable

        Dim odtConsulta As New DataTable
        Try
            DBConn.Open()
            DBcomando = New SqlCommand("select IDROL, ORDEN, IDMENU, NOMBRE
                                  from  LOTESROLDET
                                  WHERE IDROL = @IDROL
                                  ORDER BY ORDEN ", DBConn)
            DBcomando.Parameters.AddWithValue("@IDROL", idrol)

            DBAdaptador = New SqlDataAdapter(DBcomando)

            DBAdaptador.Fill(odtConsulta)

        Catch ex As Exception
            odtConsulta = Nothing
            _error = ex.Message.ToString()
        Finally
            DBConn.Close()
        End Try

        Return odtConsulta
    End Function


    Public Function setPerfilMenu(otbPerfilMenu As tbPerfilMenu, ByRef _error As String) As Boolean
        Dim existeSINO As Integer = 0
        Dim odtConsulta As New DataTable
        Dim resultado As Boolean = True

        ' ESTA FUNCION SOLO INSERTA EL REGISTRO ,

        Try
            DBConn.Open()

            DBcomando = New SqlCommand(" INSERT INTO LOTESROLDET (IDROL, ORDEN, IDMENU, NOMBRE)
                                        VALUES (@IDROL, @ORDEN, @IDMENU, @NOMBRE) ", DBConn)
            DBcomando.Parameters.AddWithValue("@IDROL", otbPerfilMenu.IDROL)
            DBcomando.Parameters.AddWithValue("@ORDEN", otbPerfilMenu.ORDEN)
            DBcomando.Parameters.AddWithValue("@IDMENU", otbPerfilMenu.IDMENU)
            DBcomando.Parameters.AddWithValue("@NOMBRE", otbPerfilMenu.NOMBRE)

            DBcomando.ExecuteNonQuery()

        Catch ex As Exception
            resultado = False
            _error = ex.Message.ToString()
        Finally
            DBConn.Close()
        End Try

        Return resultado
    End Function

    Public Function delPerfilMenu(otbPerfilMenu As tbPerfilMenu, ByRef _error As String) As Boolean
        Dim existeSINO As Integer = 0
        Dim odtConsulta As New DataTable
        Dim resultado As Boolean = True

        ' ESTA FUNCION SOLO ELIMINA EL REGISTRO ,

        Try
            DBConn.Open()

            DBcomando = New SqlCommand(" DELETE FROM LOTESROLDET
                                    WHERE IDROL = @IDROL
                                          AND IDMENU = @IDMENU ", DBConn)
            DBcomando.Parameters.AddWithValue("@IDROL", otbPerfilMenu.IDROL)
            DBcomando.Parameters.AddWithValue("@IDMENU", otbPerfilMenu.IDMENU)

            DBcomando.ExecuteNonQuery()

        Catch ex As Exception
            resultado = False
            _error = ex.Message.ToString()
        Finally
            DBConn.Close()
        End Try

        Return resultado
    End Function

    Public Function getUsuarios(asUsuario As String, ByRef _error As String) As DataTable

        Dim odtConsulta As New DataTable
        Try
            DBConn.Open()
            DBcomando = New SqlCommand("select T1.CODE, T1.NAME, T2.IDROL, PASSW = CONVERT(NVARCHAR(MAX), DECRYPTBYPASSPHRASE('Electrico2012', T2.PASSW))
                                  from [@SO1_01AUTEMPROL] T1 LEFT JOIN  LOTESUSUARIOS T2 ON T2.CODE = T1.CODE
                                  WHERE T1.NAME LIKE '%' + @USUARIO + '%'
                                        OR T1.CODE LIKE '%' + @USUARIO + '%'
                                        OR @USUARIO = '0' ", DBConn)
            DBcomando.Parameters.AddWithValue("@USUARIO", asUsuario)

            DBAdaptador = New SqlDataAdapter(DBcomando)

            DBAdaptador.Fill(odtConsulta)

        Catch ex As Exception
            odtConsulta = Nothing
            _error = ex.Message.ToString()
        Finally
            DBConn.Close()
        End Try

        Return odtConsulta
    End Function

    Public Function setUsuario(otbUsuario As tbUsuario, ByRef _error As String) As Boolean
        Dim existeSINO As Integer = 0
        Dim odtConsulta As New DataTable
        Dim resultado As Boolean = True

        Try

            DBConn.Open()
            DBcomando = New SqlCommand("select existe = isnull(count(1), 0)
                                  from  LOTESUSUARIOS
                                  where CODE = @CODE", DBConn)
            DBcomando.Parameters.AddWithValue("@CODE", otbUsuario.CODE)

            existeSINO = Convert.ToInt32(DBcomando.ExecuteScalar())

            If existeSINO > 0 Then
                DBcomando = New SqlCommand(" update LOTESUSUARIOS
                                        set IDROL = @IDROL,
                                            PASSW  = ENCRYPTBYPASSPHRASE('Electrico2012', @PASSW) 
                                     where CODE = @CODE", DBConn)
                DBcomando.Parameters.AddWithValue("@CODE", otbUsuario.CODE)
                DBcomando.Parameters.AddWithValue("@IDROL", otbUsuario.IDROL)
                DBcomando.Parameters.AddWithValue("@PASSW", otbUsuario.PASSW)

            Else

                DBcomando = New SqlCommand(" INSERT INTO LOTESUSUARIOS (CODE, IDROL, PASSW)
                                        VALUES (@CODE, @IDROL, ENCRYPTBYPASSPHRASE('Electrico2012', @PASSW) ) ", DBConn)
                DBcomando.Parameters.AddWithValue("@CODE", otbUsuario.CODE)
                DBcomando.Parameters.AddWithValue("@IDROL", otbUsuario.IDROL)
                DBcomando.Parameters.AddWithValue("@PASSW", otbUsuario.PASSW)

            End If

            DBcomando.ExecuteNonQuery()

        Catch ex As Exception
            resultado = False
            _error = ex.Message.ToString()
        Finally
            DBConn.Close()
        End Try

        Return resultado
    End Function

    Public Function delUsuario(odtUsuario As tbUsuario, ByRef _error As String) As Boolean
        Dim existeSINO As Integer = 0
        Dim odtConsulta As New DataTable
        Dim resultado As Boolean = True

        ' ESTA FUNCION SOLO ELIMINA EL REGISTRO ,

        Try
            DBConn.Open()

            DBcomando = New SqlCommand(" DELETE FROM LOTESUSUARIOS
                                    WHERE CODE = @CODE ", DBConn)
            DBcomando.Parameters.AddWithValue("@CODE", odtUsuario.CODE)

            DBcomando.ExecuteNonQuery()

        Catch ex As Exception
            resultado = False
            _error = ex.Message.ToString()
        Finally
            DBConn.Close()
        End Try

        Return resultado
    End Function

End Class
