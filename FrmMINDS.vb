Option Explicit On

Imports System.Data.SqlClient
Imports System.Math
Imports System.IO
Imports System.Text.ASCIIEncoding
Public Class FrmMINDS
    Inherits System.Windows.Forms.Form
    Dim dtReporte As New DataTable("Reporte")
    Dim dtDetalle As New DataTable("Final")
    Dim strConn As String = "Server=SERVER-RAID; DataBase=production; User ID=User_pro; pwd=User_PRO2015"
    Dim fecha As Date = "19/09/2017" 'Date.Now.ToShortDateString julio fue el ULTIMO
    Dim fechaLim As Date = "19/09/2017"
    Dim Contador As Integer = 0
    Dim Pagos As New Minds2DSTableAdapters.layoutsCreditoTableAdapter


    Private Sub btnLCtos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLCtos.Click
        Dim Con1 As New ProductionDataSetTableAdapters.AnexosTableAdapter
        Dim Con2 As New ProductionDataSetTableAdapters.AviosTableAdapter
        Dim dsAgil As New DataSet()
        Dim Pagos As New Minds2DSTableAdapters.layoutsCreditoTableAdapter
        'Pagos.DeleteAll()
        Dim Cuentas As New Minds2DSTableAdapters.layoutsCuentaTableAdapter
        'Cuentas.DeleteAll()
        Dim cnAgil As New SqlConnection(strConn)
        Dim cm1 As New SqlCommand()
        Dim cm3 As New SqlCommand()
        Dim cm4 As New SqlCommand()
        Dim drAnexo As DataRow
        Dim drEdoctav As DataRow()
        Dim drDato As DataRow

        Dim cDia As String
        Dim i As Integer
        Dim cRenglon As String
        Dim cImporte As String
        Dim cAnexo As String
        Dim cCiclo As String
        Dim cCliente As String
        Dim cFecha As String
        Dim cFechafin As String
        Dim cPago As String
        Dim nCount As Integer
        Dim nPago As Decimal
        Dim cProduct As String
        Dim cSubProduct As String
        Dim cSucursal As String


        Dim cm2 As New SqlCommand()
        Dim dsReporte As New DataSet()
        Dim daAnexos As New SqlDataAdapter(cm1)
        Dim daEdoctav As New SqlDataAdapter(cm2)
        Dim daAvios As New SqlDataAdapter(cm3)
        Dim daCuentasConcetradoras As New SqlDataAdapter(cm4)
        Dim relAnexoEdoctav As DataRelation


        cDia = Mid(DTOC(Today), 7, 2) & Mid(DTOC(Today), 5, 2)
        cFecha = DTOC(Today)
        cnAgil.Open()

        With cm1
            .CommandType = CommandType.Text
            .CommandText = "SELECT cliente, Anexo, Fechacon, Mensu, MtoFin, Tipar, Sucursal FROM Minds_Cuentas "
            .Connection = cnAgil
        End With

        With cm3
            .CommandType = CommandType.Text
            .CommandText = "SELECT * FROM Minds_CuentasAvio"
            '.CommandText = "SELECT Clientes.cliente, Anexo + '-' + ciclo as Anexo, Fechaautorizacion, LineaActual, FechaTerminacion, Tipar FROM Clientes " & _
            '               "INNER JOIN Avios On Avios.Cliente = Clientes.Cliente WHERE Flcan = 'A' and fechaTerminacion >= '20130101' and (minds = 0 or minds is null)"
            .Connection = cnAgil
        End With

        With cm4
            .CommandType = CommandType.Text
            .CommandText = "SELECT * FROM Minds_CuentasConcetradoras"
            .Connection = cnAgil
        End With

        ' Este Stored Procedure trae la tabla de amortización del equipo de todos los contratos activos
        ' con fecha de contratación menor o igual a la de proceso

        With cm2
            .CommandType = CommandType.Text
            .CommandText = "SELECT * FROM Edoctav Order By Anexo, Letra Desc"
            .Connection = cnAgil
        End With
        daAnexos.Fill(dsAgil, "Anexos")
        daEdoctav.Fill(dsAgil, "Edoctav")
        daAvios.Fill(dsAgil, "Avios")
        daCuentasConcetradoras.Fill(dsAgil, "Cuentas")

        'CONCETRADORAS+++++++++++++++++++++++++++++++++++++
        For Each drAnexo In dsAgil.Tables("Cuentas").Rows
            cAnexo = drAnexo("Anexo")
            cCliente = drAnexo("Cliente")
            cSucursal = drAnexo("Mensu").ToString
            cImporte = drAnexo("MtoFin").ToString
            cFecha = CTOD(drAnexo("Fechacon")).ToShortDateString

            nCount = 0
            ' cProduct = "CREDITO"
            ' cSubProduct = "SIMPLE"
            cProduct = "3"

            cFechafin = "01/01/2030"
            cPago = 1

            If Cuentas.Existe(cAnexo).Value = 0 Then
                Cuentas.Insert(cAnexo, cCliente, 7, cProduct, cImporte, cFecha, cFechafin, 1, cPago)
            End If

            Label2.Text = "Procesando Contrato " & cAnexo
            Label2.Update()

        Next
        'CONCETRADORAS+++++++++++++++++++++++++++++++++++++

        ' Establecer la relación entre Anexos y Edoctav

        relAnexoEdoctav = New DataRelation("AnexoEdoctav", dsAgil.Tables("Anexos").Columns("Anexo"), dsAgil.Tables("Edoctav").Columns("Anexo"))
        dsAgil.EnforceConstraints = False
        dsAgil.Relations.Add(relAnexoEdoctav)

        For Each drAnexo In dsAgil.Tables("Anexos").Rows
            cAnexo = drAnexo("Anexo")
            cCliente = drAnexo("Cliente")
            cSucursal = drAnexo("Mensu").ToString
            cImporte = drAnexo("MtoFin").ToString
            cFecha = CTOD(drAnexo("Fechacon")).ToShortDateString
            drEdoctav = drAnexo.GetChildRows("AnexoEdoctav")
            Select Case drAnexo("Tipar")
                Case "F"
                    '    cProduct = "ARRENDAMIENTO"
                    '   cSubProduct = "FINANCIERO"
                    cProduct = "1"
                Case "P"
                    ' cProduct = "ARRENDAMIENTO"
                    ' cSubProduct = "PURO"
                    cProduct = "2"
                Case "R"
                    'cProduct = "CREDITO"
                    'cSubProduct = "REFACCIONARIO"
                    cProduct = "8"
                Case "S"
                    ' cProduct = "CREDITO"
                    ' cSubProduct = "SIMPLE"
                    cProduct = "3"
            End Select
            nCount = 0


            For Each drDato In drEdoctav
                If nCount = 0 Then
                    cFechafin = CTOD(drDato("Feven")).ToShortDateString
                    nPago = drDato("Abcap") + drDato("Inter")
                    cPago = nPago.ToString
                    nCount += 1
                End If
            Next

            cRenglon = cAnexo & "|" & cCliente & "|" & cProduct & "|" & cImporte & "|" & cFecha & "|" & cFechafin & "|1|" & cPago & "|" & cSucursal & "|"
            If Cuentas.Existe(cAnexo).Value = 0 Then
                Cuentas.Insert(cAnexo, cCliente, 7, cProduct, cImporte, cFecha, cFechafin, 1, cPago)
            Else
                'Cuentas.UpdateCuenta(cCliente, 7, cProduct, cImporte, cFecha, cFechafin, 1, cPago, cAnexo)
            End If

            'Con1.UpdateMinds(cAnexo)
            'stmWriter.WriteLine(cRenglon)
            Label2.Text = "Procesando Contrato " & cAnexo
            Label2.Update()

        Next



        For Each drAnexo In dsAgil.Tables("Avios").Rows
            If "081640011" = drAnexo("Anexo") Then
                cAnexo = drAnexo("Anexo")
            End If
            cAnexo = drAnexo("Anexo")
            cCliente = drAnexo("Cliente")
            cImporte = drAnexo("LineaActual").ToString
            cFecha = CTOD(drAnexo("FechaAutorizacion")).ToShortDateString
            Select Case drAnexo("Tipar")
                Case "A"
                    ' cProduct = "CREDITO"
                    'cSubProduct = "ANTICIPO DE AVIO"
                    cProduct = "114" ' como simple

                Case "C"
                    '   cProduct = "CREDITO"
                    '   cSubProduct = "CUENTA CORRIENTE"
                    cProduct = "4"
                Case "H"
                    ' cProduct = "CREDITO"
                    ' cSubProduct = "AVIO"
                    cProduct = "9"
            End Select
            cFechafin = CTOD(drAnexo("FechaTerminacion")).ToShortDateString
            nPago = drAnexo("LineaActual")
            cPago = nPago.ToString
            Label2.Text = "Procesando Contrato " & cAnexo & " de AVIO"
            Label2.Update()
            If drAnexo("Tipar") <> "A" Then
                If Cuentas.Existe(cAnexo).Value = 0 Then
                    Cuentas.Insert(cAnexo, cCliente, 7, cProduct, cImporte, cFecha, cFechafin, 1, cPago)
                Else
                    Cuentas.UpdateCuenta(cCliente, 7, cProduct, cImporte, cFecha, cFechafin, 1, cPago, cAnexo)
                End If

                cAnexo = Mid(cAnexo, 1, 9)
                cCiclo = Mid(cAnexo, 11, 2)
                Con2.UpdateMinds(cCiclo, cAnexo)
            End If
        Next
        cnAgil.Close()
        MsgBox("Proceso Terminado", MsgBoxStyle.Information, "Mensaje")

    End Sub

    Private Sub btnPagos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPagos.Click
        Cursor.Current = Cursors.WaitCursor
        fecha = dtpProcesar1.Value.ToShortDateString
        fechaLim = dtpProcesar2.Value.ToShortDateString
        Contador = 0

        Dim x As Integer = Pagos.DeleteFecha(fecha)
        Call Pagos_Tradicionales()
        Call Pagos_Avio()
        Cursor.Current = Cursors.Default
        MsgBox("Proceso Terminado" & vbCrLf & "Se cargaron :" & Contador & " transacciones", MsgBoxStyle.Information, "Mensaje")

    End Sub

    Private Sub btnCliente_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCliente.Click
        Cursor.Current = Cursors.WaitCursor
        Dim dsAgil As New DataSet()
        Dim Clientes As New Minds2DSTableAdapters.layoutsKYCTableAdapter
        Dim ClientesORG As New ProductionDataSetTableAdapters.ClientesTableAdapter
        Dim Municipio As New Minds2DSTableAdapters.Cat_MunicipioTableAdapter
        Dim Estado As New Minds2DSTableAdapters.Cat_EstadoTableAdapter
        Dim TMunicipio As New Minds2DS.Cat_MunicipioDataTable
        Dim TEstado As New Minds2DS.Cat_EstadoDataTable
        Dim cMuni As Double = 1
        Dim cEstado As Double = 0
        Dim xEstado As String = ""

        'Clientes.DeleteAll()

        Dim cnAgil As New SqlConnection(strConn)
        Dim cm1 As New SqlCommand()
        Dim cm2 As New SqlCommand()
        Dim cm3 As New SqlCommand()
        Dim drCliente As DataRow
        Dim drDato As DataRow
        Dim drAnexos As DataRow()
        Dim drAnexo As DataRow
        Dim drPlaza As DataRow

        Dim cDia As String
        Dim i As Integer
        Dim cRenglon As String
        Dim cCliente As String
        Dim cDescr As String
        Dim cPromo As String
        Dim cFecha As String
        Dim cGiro As String
        Dim cIdGiro As String
        Dim cProfGiro As String
        Dim cTipo As String
        Dim nCount As Integer
        Dim nDato As Integer
        Dim nIDEstado As Integer

        Dim aName As New ArrayList()
        Dim cDato As String
        Dim cNombre As String = ""
        Dim cApePaterno As String
        Dim cApeMaterno As String
        Dim cActivo As String = "2"
        Dim cDelegacion As String
        Dim nIdPlazam As Integer

        Dim dsReporte As New DataSet()
        Dim daCliente As New SqlDataAdapter(cm1)
        Dim daAnexos As New SqlDataAdapter(cm2)
        Dim daPlazas As New SqlDataAdapter(cm3)
        Dim relAnexoCliente As DataRelation

        cDia = Mid(DTOC(Today), 7, 2) & Mid(DTOC(Today), 5, 2)
        cFecha = dtpProcesar1.Value
        cnAgil.Open()

        With cm1
            .CommandType = CommandType.Text
            .CommandText = "SELECT Descr,Cliente,Promo,Tipo,Calle, Colonia,Delegacion,Copos,Telef1,Giro,Clientes.Plaza, DescPlaza, RFC,Curp,Email1, Fecha1, NombreCliente, ApellidoPaterno, ApellidoMaterno FROM Clientes " &
            "Inner Join Plazas ON Clientes.Plaza = Plazas.Plaza ORDER BY Cliente"
            '.CommandText = "SELECT Descr,Cliente,Promo,Tipo,Calle, Colonia,Delegacion,Copos,Telef1,Giro,Clientes.Plaza, DescPlaza, RFC,Curp,Email1, Fecha1, NombreCliente, ApellidoPaterno, ApellidoMaterno FROM Clientes " & _
            '"Inner Join Plazas ON Clientes.Plaza = Plazas.Plaza where siebel = 0 or siebel is null ORDER BY Cliente"
            .Connection = cnAgil
        End With
        daCliente.Fill(dsAgil, "Clientes")

        With cm2
            .CommandType = CommandType.Text
            .CommandText = "select Anexo, Fechacon, Flcan, Cliente FROM Anexos " &
            "UNION Select Anexo, FechaAutorizacion as Fechacon, Flcan, Cliente FROM Avios ORDER BY Cliente"
            .Connection = cnAgil
        End With
        daAnexos.Fill(dsAgil, "Anexos")

        ' Establecer la relación entre Anexos y Clientes

        relAnexoCliente = New DataRelation("AnexoCliente", dsAgil.Tables("Clientes").Columns("Cliente"), dsAgil.Tables("Anexos").Columns("Cliente"))
        dsAgil.EnforceConstraints = False
        dsAgil.Relations.Add(relAnexoCliente)

        nCount = 1
        For Each drCliente In dsAgil.Tables("Clientes").Rows
            cApePaterno = ""
            cApeMaterno = ""
            cNombre = ""
            cTipo = drCliente("Tipo")
            cGiro = drCliente("Giro")
            cDelegacion = Trim(drCliente("Delegacion"))
            nIDEstado = drCliente("Plaza")
            Select Case drCliente("Plaza")
                Case Is = "07"
                    nIDEstado = 5
                Case Is = "08"
                    nIDEstado = 6
                Case Is = "05"
                    nIDEstado = 7
                Case Is = "06"
                    nIDEstado = 8
                Case Is = "12"
                    nIDEstado = 11
                Case Is = "13"
                    nIDEstado = 12
                Case Is = "14"
                    nIDEstado = 13
                Case Is = "15"
                    nIDEstado = 14
                Case Is = "11"
                    nIDEstado = 15
            End Select

            With cm3
                .CommandType = CommandType.Text
                .CommandText = "SELECT IdPlazam, NombrePlaza FROM PlazasMinds WHERE IdEstado = " & nIDEstado
                .Connection = cnAgil
            End With
            daPlazas.Fill(dsAgil, "Plazas")

            nIdPlazam = 0
            For Each drPlaza In dsAgil.Tables("Plazas").Rows
                If Trim(cDelegacion) = Trim(drPlaza("NombrePlaza")) Then
                    nIdPlazam = drPlaza("IdPlazam")
                End If
            Next
            dsAgil.Tables.Remove("Plazas")

            If nIdPlazam = 0 Then
                Select Case nIDEstado
                    Case 1
                        nIdPlazam = 199009
                    Case 2
                        nIdPlazam = 299009
                    Case 3
                        nIdPlazam = 399009
                    Case 4
                        nIdPlazam = 499009
                    Case 5
                        nIdPlazam = 599009
                    Case 6
                        nIdPlazam = 699009
                    Case 7
                        nIdPlazam = 899009
                    Case 8
                        nIdPlazam = 999009
                    Case 9
                        nIdPlazam = 1001002
                    Case 10
                        nIdPlazam = 1199008
                    Case 11
                        nIdPlazam = 1299003
                    Case 12
                        nIdPlazam = 1399007
                    Case 13
                        nIdPlazam = 1499002
                    Case 14
                        nIdPlazam = 1699001
                    Case 15
                        nIdPlazam = 1899009
                    Case 16
                        nIdPlazam = 2099007
                    Case 17
                        nIdPlazam = 2199006
                    Case 18
                        nIdPlazam = 2299005
                    Case 19
                        nIdPlazam = 2399004
                    Case 20
                        nIdPlazam = 2999007
                    Case 21
                        nIdPlazam = 3299006
                    Case 22
                        nIdPlazam = 3399009
                    Case 23
                        nIdPlazam = 3499003
                    Case 24
                        nIdPlazam = 3599006
                    Case 25
                        nIdPlazam = 3699009
                    Case 26
                        nIdPlazam = 3799003
                    Case 27
                        nIdPlazam = 3899006
                    Case 28
                        nIdPlazam = 3999009
                    Case 29
                        nIdPlazam = 4099001
                    Case 30
                        nIdPlazam = 4399004
                    Case 31
                        nIdPlazam = 4599009
                    Case 32
                        nIdPlazam = 4699007
                End Select
            End If


            drAnexos = drCliente.GetChildRows("AnexoCliente")

            cPromo = drCliente("Promo")
            If drCliente("Tipo") = "F" Or drCliente("Tipo") = "E" Then
                cDescr = Trim(drCliente("Descr"))
                Dim texto() As String = Split(cDescr, " ")

                nCount = 0
                aName.Clear()
                For i = 0 To UBound(texto)
                    aName.Add(texto(i))
                    nCount += 1
                Next

                i = 1
                For Each cDato In aName
                    If i <= nCount - 2 Then
                        If cNombre = "" Then
                            cNombre = cDato
                        Else
                            cNombre = cNombre & " " & cDato
                        End If
                    ElseIf i = nCount - 1 Then
                        cApePaterno = cDato
                    ElseIf i = nCount Then
                        cApeMaterno = cDato
                    End If
                    i += 1
                Next

                cCliente = drCliente("Cliente")
                cPromo = drCliente("Promo")
            Else
                cNombre = Trim(drCliente("Descr"))

                cCliente = drCliente("Cliente")
                cPromo = drCliente("Promo")

            End If

            nDato = 0
            cActivo = "2"
            For Each drAnexo In drAnexos
                If nDato = 0 Then
                    cFecha = drAnexo("Fechacon")
                End If
                If drAnexo("Flcan") = "A" Then
                    cActivo = "1"
                End If
                nDato += 1
            Next

            If Trim(cGiro) = "" Then
                cGiro = "18"
            End If

            With cm3
                .CommandType = CommandType.Text
                .CommandText = "SELECT IdGiro, ActividadEconomica FROM GirosMinds WHERE Giro = " & cGiro
                .Connection = cnAgil
            End With
            daPlazas.Fill(dsAgil, "Giros")
            drPlaza = dsAgil.Tables("Giros").Rows(0)

            cIdGiro = drPlaza("IdGiro")
            cProfGiro = drPlaza("ActividadEconomica")
            dsAgil.Tables.Remove("Giros")

            If cTipo = "E" Then
                cTipo = 3
            ElseIf cTipo = "F" Then
                cTipo = 1
            Else
                cTipo = 2
            End If

            If Trim(cPromo) <> "" Then
                If Trim(cNombre) <> "" Then
                    If Len(cNombre) > 100 Then
                        cNombre = Mid(cNombre, 1, 100)
                    End If
                    cRenglon = cActivo & "|" & cPromo & "|0|0|0|0|0|0|" & cIdGiro & "|0||0|0|N|N|" & cNombre & "|" & cApePaterno & "|"
                    cRenglon = cRenglon & cApeMaterno & "|" & cProfGiro & "|" & drCliente("RFC") & "|" & cTipo & "|1|" & CTOD(cFecha).ToShortDateString & "|" & Trim(drCliente("Calle")) & "|0|0|" & Trim(drCliente("Colonia"))
                    cRenglon = cRenglon & "|" & drCliente("Copos") & "|" & Trim(drCliente("Delegacion")) & "|" & nIdPlazam & "|" & nIDEstado & "|237|0|0|0|0|0|0|0|0|" & (dtpProcesar1.Value).ToShortDateString & "|1|" & Trim(drCliente("CURP")) & "|"
                    cRenglon = cRenglon & Trim(drCliente("Telef1")) & "|1|" & cCliente & "|" & Val(cCliente) & "|" & (dtpProcesar1.Value).ToShortDateString & "|0|" & CTOD(drCliente("Fecha1")).ToShortDateString & "|" & Trim(drCliente("DescPlaza")) & "||" & Trim(drCliente("EMail1")) & "|237|"
                    'stmWriter.WriteLine(cRenglon)

                    Municipio.FillByMunicipio(TMunicipio, "%" & Trim(drCliente("delegacion")) & "%")
                    If TMunicipio.Rows.Count > 0 Then
                        cMuni = TMunicipio.Rows(0).Item(0)
                    Else
                        cMuni = 0
                    End If
                    xEstado = "%" & Trim(drCliente("DescPlaza")) & "%"
                    Select Case xEstado
                        Case "%ESTADO DE MEXICO%"
                            xEstado = "%MEXICO%"
                    End Select
                    Estado.FillByEstado(TEstado, xEstado)
                    If TEstado.Rows.Count > 0 Then
                        cEstado = TEstado.Rows(0).Item(0)
                    Else
                        cEstado = 0
                    End If
                    Try
                        If Clientes.Exsiste(Trim(drCliente("Cliente"))).Value = 0 Then
                            Clientes.Insert(Trim(drCliente("Cliente")), cActivo, 0, 0, 0, 0, 0, 0, cPromo, "Credito", "", cIdGiro, 0, 0, 0, cNombre, cApePaterno, cApeMaterno, cProfGiro, drCliente("RFC"), cTipo, 1, CTOD(cFecha).ToShortDateString, Trim(drCliente("Calle")), 0, 0, Trim(drCliente("Colonia")), drCliente("Copos") _
                            , cMuni, nIdPlazam, nIDEstado, 236, 0, 0, 0, 0, 0, 0, 0, 0, dtpProcesar1.Value.ToShortDateString, 1, Trim(drCliente("CURP")), Trim(drCliente("Telef1")), 1, Val(cCliente), dtpProcesar1.Value.ToShortDateString, 1, CTOD(drCliente("Fecha1")).ToShortDateString, cEstado, "", Trim(drCliente("EMail1")), 236, 0, dtpProcesar1.Value.ToShortDateString, 2)
                        Else
                            Clientes.UpdateKYC(cActivo, 0, 0, 0, 0, 0, 0, cPromo, "Credito", "", cIdGiro, 0, 0, 0, cNombre, cApePaterno, cApeMaterno, cProfGiro, drCliente("RFC"), cTipo, 1, CTOD(cFecha).ToShortDateString, Trim(drCliente("Calle")), 0, 0, Trim(drCliente("Colonia")), drCliente("Copos") _
                            , cMuni, nIdPlazam, nIDEstado, 236, 0, 0, 0, 0, 0, 0, 0, 0, (dtpProcesar1.Value).ToShortDateString, 1, Trim(drCliente("CURP")), Trim(drCliente("Telef1")), 1, Val(cCliente), (dtpProcesar1.Value).ToShortDateString, 1, CTOD(drCliente("Fecha1")).ToShortDateString, cEstado, "", Trim(drCliente("EMail1")), 236, 0, (dtpProcesar1.Value).ToShortDateString, 2, Trim(drCliente("Cliente")))
                        End If
                    Catch ex As Exception

                    End Try
                    If cTipo = "2" Then
                        cRenglon = cIdGiro & "|" & cActivo & "|1|" & cNombre & "||||" & drCliente("RFC") & "||0|" & Trim(drCliente("Telef1")) & "||||||||||||" & cCliente & "|"
                    Else
                        cRenglon = cIdGiro & "|" & cActivo & "|1||" & cNombre & "|" & cApePaterno & "|" & cApeMaterno & "|" & drCliente("RFC") & "||0|" & Trim(drCliente("Telef1")) & "||||||||||||" & cCliente & "|"
                    End If
                    'DataGridView1.DataSource = dsAgil

                End If
            End If
        Next
        cnAgil.Close()
        Cursor.Current = Cursors.Default
        MsgBox("Proceso Terminado", MsgBoxStyle.Information, "Mensaje")

    End Sub

    Private Sub BttPromo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BttPromo.Click
        Dim ta As New MINDS.ProductionDataSetTableAdapters.PromotoresTableAdapter
        Dim ta1 As New MINDS.Minds2DSTableAdapters.layoutsFuncionarioTableAdapter
        Dim PromoOrg As New ProductionDataSet.PromotoresDataTable
        Try
            ta.Fill(PromoOrg)

            For Each r As ProductionDataSet.PromotoresRow In PromoOrg.Rows
                If ta1.Existe(r.Promotor, Trim(r.APaterno)).Value = 0 Then
                    ta1.Insert(r.Promotor, Trim(r.Nombre), Trim(r.APaterno), Trim(r.AMaterno), Trim(r.Puesto), r.IDPlaza, r.Nacionalidad, CTOD(r.FechaCarga))
                Else
                    ta1.UpdateEmpleado(Trim(r.Nombre), Trim(r.APaterno), Trim(r.AMaterno), Trim(r.Puesto), r.IDPlaza, r.Nacionalidad, CTOD(r.FechaCarga), r.Promotor)
                End If

            Next
            MessageBox.Show("Terminado", "Promotores", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Sub Pagos_Tradicionales()
        Dim dsAgil As New DataSet()
        'Dim Hist As New ProductionDataSetTableAdapters.HistoriaTableAdapter
        Dim EdoCtaV As New ProductionDataSetTableAdapters.EdoctavTableAdapter
        Dim TEdoCtaV As New ProductionDataSet.EdoctavDataTable
        Dim taMAX As New ProductionDataSetTableAdapters.Vw_MAXfecVenTableAdapter
        Dim MAXfecVen As New ProductionDataSet.Vw_MAXfecVenDataTable
        Dim cnAgil As New SqlConnection(strConn)
        Dim cm1 As New SqlCommand()
        Dim drAnexo As DataRow

        Dim cDia As String
        Dim i As Integer
        Dim cAnexo As String
        Dim cCliente As String
        Dim cPromo As String
        Dim cFecha As String
        Dim cFechafin As String
        Dim cPago As String
        Dim cSucursal As String
        Dim cCheque As String
        Dim cDoc As String
        Dim nCount As Integer
        Dim nOper As Integer
        Dim nInsMon As Integer
        Dim y As Integer
        Dim nPago As Decimal
        Dim nSaldo As Decimal
        Dim x As Integer


        Dim cm2 As New SqlCommand()
        Dim dsReporte As New DataSet()
        Dim daAnexos As New SqlDataAdapter(cm1)
        Dim daEdoctav As New SqlDataAdapter(cm2)


        cDia = Mid(DTOC(Today), 7, 2) & Mid(DTOC(Today), 5, 2)
        cFecha = DTOC(dtpProcesar1.Value)
        cnAgil.Open()
        'dsAgil.Tables("Pagos").Clear()
        With cm1
            .CommandType = CommandType.Text
            .CommandText = "SELECT Serie, Numero, Fecha, Anexo,Letra, Importe, Cheque, Promo, Cliente, Sucursal, Tipar, EsEfectivo, Banco, minds " _
             & " FROM Minds_Pagos where fecha between '" & fecha.ToString("yyyyMMdd") & "' and '" & fechaLim.ToString("yyyyMMdd") _
             & "' and anexo <> 'X038790001' order by Serie, Numero"
            .Connection = cnAgil
        End With

        daAnexos.Fill(dsAgil, "Pagos")
        nCount = 1
        nPago = 0
        For Each drAnexo In dsAgil.Tables("Pagos").Rows

            cAnexo = drAnexo("Anexo")
            cCliente = drAnexo("Cliente")
            cSucursal = drAnexo("Sucursal")
            cPromo = drAnexo("Promo")
            cCheque = drAnexo("Cheque")
            cFecha = CTOD(drAnexo("Fecha")).ToShortDateString

            nInsMon = InstrumentoMonetario(cCheque, drAnexo("EsEfectivo"), IIf(IsDBNull(drAnexo("Minds")), 0, drAnexo("Minds")))

            EdoCtaV.Fill(TEdoCtaV, cAnexo, drAnexo("letra"))
            If TEdoCtaV.Rows.Count > 0 Then
                nSaldo = TEdoCtaV.Rows(0).Item("Saldo")
            ElseIf drAnexo("letra") = "888" Then
                nSaldo = EdoCtaV.ScalarSaldoII(cAnexo, drAnexo("Fecha"))
            Else
                nSaldo = 0
            End If


            taMAX.Fill(MAXfecVen, cAnexo)
            If MAXfecVen.Rows.Count > 0 Then
                cFechafin = CTOD(MAXfecVen.Rows(0).Item("Feven")).ToShortDateString
            Else
                cFechafin = cFecha
            End If
            nPago = drAnexo("Importe")

            If drAnexo("Tipar") = "F" Then
                nOper = 27
            ElseIf drAnexo("Tipar") = "P" Then
                nOper = 27
            Else
                nOper = 9
            End If

            If drAnexo("letra") = "888" Then
                nOper = 42
            ElseIf drAnexo("letra") = "999" Then
                nOper = 41
            End If
            cPago = Stuff(nPago.ToString, i, " ", 10)
            nCount = drAnexo("Numero")
            'If nCount = 999999 Or nCount = 888888 Or nCount = 777777 Then
            '    cDoc = Trim(drAnexo("Serie")) & Trim(drAnexo("Numero")) & "-" & cAnexo
            'Else
            '    cDoc = Trim(drAnexo("Serie")) & "-" & Trim(drAnexo("Numero")) & "-" & Trim(drAnexo("Letra"))
            'End If
            cDoc = Trim(drAnexo("Serie")) & Trim(drAnexo("Numero")) & "-" & cCheque ''& "-" & drAnexo("Banco")

            If drAnexo("letra") = "888" Or drAnexo("letra") = "999" Then
                cDoc = Trim(drAnexo("Serie")) & Trim(drAnexo("Numero")) & "-" & cCheque.Trim & "-" & drAnexo("Anexo")
                cDoc = cDoc.Substring(0, 28) & CInt(Math.Ceiling(Rnd() * 9)) + 1
            End If

            If nPago <> 0 Then
                Try
                    Pagos.Insert(cDoc, cAnexo, nOper, nInsMon, 1, cFecha, nPago, nPago, drAnexo("promo"), cSucursal, cFechafin, nSaldo, 0)
                    Contador += 1
                Catch ex As Exception
                    MessageBox.Show(ex.Message & " " & cDoc, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            End If
            nPago = 0
            nSaldo = 0
        Next
        'RELEVANTES CONCETRADOS TRADICIONALES++++++++++++++++++++++++++++++++++++++++
        dsAgil.Tables("Pagos").Clear()
        With cm1
            .CommandType = CommandType.Text
            .CommandText = "SELECT Serie, Numero, Fecha, Anexo,Letra, Importe, Cheque, Promo, Cliente, Sucursal, Tipar, EsEfectivo, minds " _
            & "FROM Minds_PagosReleConcetradorasTRA where fecha between'" & fecha.ToString("yyyyMMdd") & "' and '" & fechaLim.ToString("yyyyMMdd") & "' order by Serie, Numero"
            .Connection = cnAgil
        End With
        daAnexos.Fill(dsAgil, "Pagos")
        nCount = 1
        nPago = 0
        'MessageBox.Show(dsAgil.Tables("Pagos").Rows.Count.ToString & " pagos")
        For Each drAnexo In dsAgil.Tables("Pagos").Rows

            cAnexo = drAnexo("Anexo")
            cCliente = drAnexo("Cliente")
            cSucursal = drAnexo("Sucursal")
            cPromo = drAnexo("Promo")
            cCheque = Trim(drAnexo("Cheque"))
            cFecha = CTOD(drAnexo("Fecha")).ToShortDateString

            nInsMon = InstrumentoMonetario(cCheque, drAnexo("EsEfectivo"), IIf(IsDBNull(drAnexo("Minds")), 0, drAnexo("Minds")))
            'nInsMon = 1 ' para que todo se reporte en minds como efectivo Karla Sanchez


            nSaldo = 1
            cFechafin = CTOD("01/01/2030")
            nPago = drAnexo("Importe")
            nOper = 9
            cPago = Stuff(nPago.ToString, i, " ", 10)

            nCount = drAnexo("Numero")
            'cDoc = Trim(drAnexo("Serie")) & Trim(drAnexo("Numero")) & "-" & cCheque
            cDoc = cAnexo & "-" & cCheque
            If nPago > 0 Then
                x = Pagos.Existe(cDoc)
                If x = 0 Then
                    Try
                        Pagos.Insert(cDoc, cAnexo, nOper, nInsMon, 1, cFecha, nPago, nPago, drAnexo("promo"), cSucursal, cFechafin, nSaldo, 0)
                        Contador += 1
                    Catch ex As Exception

                    End Try
                End If
            End If
            nPago = 0
            nSaldo = 0
        Next
        'RELEVANTES++++++++++++++++++++++++++++++++++++++++++++++++++++
        dsAgil.Tables("Pagos").Clear()
        With cm1
            .CommandType = CommandType.Text
            .CommandText = "SELECT * " _
            & "FROM Minds_PagosRelevantesTRA where fecha between'" & fecha.ToString("yyyyMMdd") & "' and '" & fechaLim.ToString("yyyyMMdd") & "' order by Serie, Numero"
            .Connection = cnAgil
        End With
        daAnexos.Fill(dsAgil, "Pagos")
        nCount = 1
        nPago = 0
        For Each drAnexo In dsAgil.Tables("Pagos").Rows

            cAnexo = drAnexo("Anexo")
            cCliente = drAnexo("Cliente")
            cSucursal = drAnexo("Sucursal")
            cPromo = drAnexo("Promo")
            cCheque = drAnexo("Cheque")
            cFecha = CTOD(drAnexo("Fecha")).ToShortDateString

            nInsMon = InstrumentoMonetario(cCheque, drAnexo("EsEfectivo"), IIf(IsDBNull(drAnexo("Minds")), 0, drAnexo("Minds")))

            EdoCtaV.Fill(TEdoCtaV, cAnexo, drAnexo("letra"))
            If TEdoCtaV.Rows.Count > 0 Then
                nSaldo = TEdoCtaV.Rows(0).Item("Saldo")
            ElseIf drAnexo("letra") = "888" Then
                nSaldo = EdoCtaV.ScalarSaldoII(cAnexo, drAnexo("Fecha"))
            Else
                nSaldo = 0
            End If


            taMAX.Fill(MAXfecVen, cAnexo)
            If MAXfecVen.Rows.Count > 0 Then
                cFechafin = CTOD(MAXfecVen.Rows(0).Item("Feven")).ToShortDateString
            Else
                cFechafin = cFecha
            End If
            nPago = drAnexo("Importe")

            If drAnexo("Tipar") = "F" Then
                nOper = 27
            Else
                nOper = 9
            End If
            cPago = Stuff(nPago.ToString, i, " ", 10)
            nCount = drAnexo("Numero")
            'If nCount = 999999 Or nCount = 888888 Or nCount = 777777 Then
            '    cDoc = Trim(drAnexo("Serie")) & Trim(drAnexo("Numero")) & "-" & cAnexo
            'Else
            '    cDoc = Trim(drAnexo("Serie")) & "-" & Trim(drAnexo("Numero")) & "-" & Trim(drAnexo("Letra"))
            'End If
            cDoc = Trim(drAnexo("Serie")) & Trim(drAnexo("Numero")) & "-" & cCheque

            If nPago > 0 Then
                x = Pagos.Existe(cDoc)
                If x = 0 Then
                    Try
                        Pagos.Insert(cDoc, cAnexo, nOper, nInsMon, 1, cFecha, nPago, nPago, drAnexo("promo"), cSucursal, cFechafin, nSaldo, 0)
                        Contador += 1
                    Catch ex As Exception

                    End Try
                End If
            End If
            nPago = 0
            nSaldo = 0
        Next
        cnAgil.Close()
    End Sub

    Sub Pagos_Avio()
        Dim dsAgil As New DataSet()
        'Dim Hist As New ProductionDataSetTableAdapters.HistoriaTableAdapter
        Dim EdoCtaV As New ProductionDataSetTableAdapters.EdoctavTableAdapter
        Dim TEdoCtaV As New ProductionDataSet.EdoctavDataTable
        Dim taMAX As New ProductionDataSetTableAdapters.Vw_MAXfecVenTableAdapter
        Dim MAXfecVen As New ProductionDataSet.Vw_MAXfecVenDataTable
        Dim cnAgil As New SqlConnection(strConn)
        Dim cm1 As New SqlCommand()
        Dim drAnexo As DataRow

        Dim cDia As String
        Dim i As Integer
        Dim cAnexo As String
        Dim cCliente As String
        Dim cPromo As String
        Dim cFechafin As String
        Dim cFecha As String
        Dim cPago As String
        Dim cSucursal As String
        Dim cCheque As String
        Dim cDoc As String
        Dim nCount As Integer
        Dim nOper As Integer
        Dim nInsMon As Integer
        Dim y As Integer
        Dim nPago As Decimal
        Dim nSaldo As Decimal
        Dim x As Integer


        Dim cm2 As New SqlCommand()
        Dim dsReporte As New DataSet()
        Dim daAnexos As New SqlDataAdapter(cm1)
        Dim daEdoctav As New SqlDataAdapter(cm2)


        cDia = Mid(DTOC(Today), 7, 2) & Mid(DTOC(Today), 5, 2)
        cnAgil.Open()

        'RELEVANTES CONCETRADOS AVIO++++++++++++++++++++++++++++++++++++++++
        With cm1
            .CommandType = CommandType.Text
            .CommandText = "SELECT * " _
            & "FROM Minds_PagosReleConcetradorasAVI where fecha between '" & fecha.ToString("yyyyMMdd") & "' and '" & fechaLim.ToString("yyyyMMdd") & "'  order by Serie, Numero"
            .Connection = cnAgil
        End With
        daAnexos.Fill(dsAgil, "Pagos")
        nCount = 1
        nPago = 0
        'MessageBox.Show(dsAgil.Tables("Pagos").Rows.Count.ToString & " pagos")
        For Each drAnexo In dsAgil.Tables("Pagos").Rows

            'If drAnexo("aNEXO") <> "CLI05430" Then Continue For

            cAnexo = drAnexo("Anexo")
            cCliente = drAnexo("Cliente")
            cSucursal = drAnexo("Sucursal")
            cPromo = drAnexo("Promo")
            cCheque = Trim(drAnexo("Cheque"))
            cFecha = CTOD(drAnexo("Fecha")).ToShortDateString

            nInsMon = InstrumentoMonetario(cCheque, drAnexo("EsEfectivo"), IIf(IsDBNull(drAnexo("Minds")), 0, drAnexo("Minds")))

            'nInsMon = 1 ' para que todo se reporte en minds como efectivo Karla Sanchez

            nSaldo = 1
            cFechafin = CTOD("01/01/2030")
            nPago = drAnexo("Importe")
            nOper = 9
            cPago = Stuff(nPago.ToString, i, " ", 10)

            nCount = drAnexo("Numero")
            'cDoc = Trim(drAnexo("Serie")) & Trim(drAnexo("Numero")) & "-" & cCheque
            cDoc = cAnexo & "-" & cCheque
            If nPago > 0 Then
                x = Pagos.Existe(cDoc)
                If x = 0 Then
                    Try
                        Pagos.Insert(cDoc, cAnexo, nOper, nInsMon, 1, cFecha, nPago, nPago, drAnexo("promo"), cSucursal, cFechafin, nSaldo, 0)
                        Contador += 1
                    Catch ex As Exception

                    End Try
                End If
            End If
            nPago = 0
            nSaldo = 0
        Next
        'RELEVANTES++++++++++++++++++++++++++++++++++++++++++++++++++++
        dsAgil.Tables("Pagos").Clear()
        With cm1
            .CommandType = CommandType.Text
            .CommandText = "SELECT * FROM Minds_PagosRelevantesAVI"
            .Connection = cnAgil
        End With

        daAnexos.Fill(dsAgil, "Pagos")
        nCount = 1
        nPago = 0
        For Each drAnexo In dsAgil.Tables("Pagos").Rows
            'se filtra desde codigo para tuning de query
            If drAnexo("Fecha") >= fecha.ToString("yyyyMMdd") And drAnexo("Fecha") <= fechaLim.ToString("yyyyMMdd") Then

                cAnexo = drAnexo("Anexo")
                cCliente = drAnexo("Cliente")
                cSucursal = drAnexo("Sucursal")
                cPromo = drAnexo("Promo")
                cCheque = drAnexo("Cheque")
                cFecha = CTOD(drAnexo("Fecha")).ToShortDateString

                nInsMon = InstrumentoMonetario(cCheque, drAnexo("EsEfectivo"), IIf(IsDBNull(drAnexo("Minds")), 0, drAnexo("Minds")))

                taMAX.Fill(MAXfecVen, cAnexo)
                If MAXfecVen.Rows.Count > 0 Then
                    cFechafin = CTOD(MAXfecVen.Rows(0).Item("Feven")).ToShortDateString
                Else
                    cFechafin = cFecha
                End If
                nPago = drAnexo("Importe")
                nOper = 9
                cPago = Stuff(nPago.ToString, i, " ", 10)
                nCount = drAnexo("Numero")
                cDoc = Trim(drAnexo("Serie")) & Trim(drAnexo("Numero")) & "-" & cCheque
                If nPago > 0 Then
                    x = Pagos.Existe(cDoc)
                    If x = 0 Then
                        Try
                            Pagos.Insert(cDoc, cAnexo, nOper, nInsMon, 1, cFecha, nPago, nPago, drAnexo("promo"), cSucursal, cFechafin, nSaldo, 0)
                            Contador += 1
                        Catch ex As Exception

                        End Try
                    End If
                End If
            End If
            nPago = 0
            nSaldo = 0
        Next
        'AVIO++++++++++++++++++++++++++++++++++++++++++++++++++++++
        Dim Avio As New ProductionDataSetTableAdapters.Minds_Pagos_AvioTableAdapter
        Dim TAvio As New ProductionDataSet.Minds_Pagos_AvioDataTable
        Dim cCiclo As String = ""
        Avio.Fill(TAvio, fecha.ToString("yyyyMMdd"), fechaLim.ToString("yyyyMMdd"))
        For Each r As ProductionDataSet.Minds_Pagos_AvioRow In TAvio.Rows
            cAnexo = r.Anexo
            cCliente = r.Cliente
            cSucursal = r.Sucursal
            cPromo = r.Promo
            cCheque = r.Cheque
            cFecha = CTOD(r.Fecha).ToShortDateString

            nInsMon = InstrumentoMonetario(cCheque, r.EsEfectivo, r.MINDS)
            nSaldo = Avio.ScalarSaldo(r.Anexo, r.Fecha)
            cFechafin = r.FechaTerminacion
            nPago = r.Importe
            nOper = 9

            cPago = Stuff(nPago.ToString, i, " ", 10)
            nCount = r.Numero
            'If nCount = 999999 Or nCount = 888888 Or nCount = 777777 Then
            '    cDoc = Trim(r.Serie) & Trim(r.Numero) & "-" & cAnexo
            'Else
            '    cDoc = Trim(r.Serie) & "-" & Trim(r.Numero) & "-" & Trim(r.Anexo) & Trim(r.Tipar)
            'End If
            cDoc = Trim(r.Serie) & Trim(r.Numero) & "-" & cCheque
            'x = Pagos.Existe(cDoc)
            If nPago > 0 Then
                Try
                    Pagos.Insert(cDoc, cAnexo, nOper, nInsMon, 1, cFecha, nPago, nPago, r.Promo, cSucursal, cFechafin, nSaldo, 0)
                    Contador += 1
                Catch ex As Exception

                End Try
            End If
            nPago = 0
            nSaldo = 0
        Next
        cnAgil.Close()
    End Sub

    Private Sub FrmMINDS_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        fecha = Date.Now.AddMonths(-1).AddDays((Date.Now.Day - 1) * -1)
        fechaLim = Date.Now.AddDays(Date.Now.Day * -1)
        dtpProcesar1.Value = fecha
        dtpProcesar2.Value = fechaLim
    End Sub

    Function InstrumentoMonetario(Rereferncia As String, EsEfectivo As Boolean, IM_minds As Integer) As Integer

        Select Case IM_minds
            Case 0
                If Mid(Rereferncia, 1, 2) = "CH" Then
                    InstrumentoMonetario = 2
                ElseIf Mid(Rereferncia, 1, 2) = "EF" Then
                    InstrumentoMonetario = 2 '1 se reporte como tranferencia
                ElseIf Mid(Rereferncia, 1, 3) = "DEP" Or IsNumeric(Rereferncia) = True Then
                    InstrumentoMonetario = 3
                ElseIf Mid(Rereferncia, 1, 2) = "EQ" Or Mid(Rereferncia, 1, 6) = "DACION" Or Mid(Rereferncia, 1, 6) = "dacion" Then
                    InstrumentoMonetario = 9
                Else
                    InstrumentoMonetario = 8
                End If

                If EsEfectivo = True Then
                    InstrumentoMonetario = 2 '1 se reporte como tranferencia
                End If
            Case Else
                InstrumentoMonetario = IM_minds
        End Select


    End Function
End Class
