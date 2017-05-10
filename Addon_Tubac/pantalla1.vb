Imports System.Globalization.CultureInfo
Public Class pantalla1
    Dim XmlForm As String = Replace(System.Windows.Forms.Application.StartupPath & "\pantalla1.srf", "\\", "\")

    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oForm As SAPbouiCOM.Form
    Private oCompany As SAPbobsCOM.Company
    Private oFilters As SAPbouiCOM.EventFilters
    Private oFilter As SAPbouiCOM.EventFilter
    Dim lineinioriginal As Integer
    Dim linefinoriginal As Integer
    Dim oGrid As SAPbouiCOM.Grid


    Public Sub New()
        MyBase.New()
        Try
            Me.SBO_Application = Utiles.SBOApplication
            Me.oCompany = Utiles.Company

            If Utiles.ActivateFormIsOpen(SBO_Application, "FrmValor") = False Then
                LoadFromXML(XmlForm)
                oForm = SBO_Application.Forms.Item("FrmValor")
                oForm.Left = 400
                oForm.DataSources.DataTables.Add("MyDataTable")

                oForm.DataSources.UserDataSources.Add("Date", SAPbouiCOM.BoDataType.dt_PRICE)
                Dim oPrecioUpdate As SAPbouiCOM.EditText
                oPrecioUpdate = oForm.Items.Item("ItemVal").Specific
                oPrecioUpdate.DataBind.SetBound(True, "", "Date")

                oForm.DataSources.UserDataSources.Add("Date1", SAPbouiCOM.BoDataType.dt_PRICE)
                Dim oDescuento As SAPbouiCOM.EditText
                oDescuento = oForm.Items.Item("ItemPorc").Specific
                oDescuento.DataBind.SetBound(True, "", "Date1")
                         
                oGrid = oForm.Items.Item("grdDatos").Specific
                oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

                Dim oChkPor As SAPbouiCOM.CheckBox
                oChkPor = oForm.Items.Item("ChkPor").Specific
                oForm.DataSources.UserDataSources.Add("ChkPor", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                oChkPor.DataBind.SetBound(True, "", "ChkPor")
                oForm.DataSources.UserDataSources.Item("ChkPor").Value = "N"

                Dim oChkTon As SAPbouiCOM.CheckBox
                oChkTon = oForm.Items.Item("ChkTon").Specific
                oForm.DataSources.UserDataSources.Add("ChkTon", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                oChkTon.DataBind.SetBound(True, "", "ChkTon")
                oForm.DataSources.UserDataSources.Item("ChkTon").Value = "N"

            Else
                oForm = Me.SBO_Application.Forms.Item("FrmValor")
                oForm.Left = 400
                oForm.Visible = true
            End If

        Catch ex As Exception
            SBOApplication.SetStatusBarMessage(ex.Message)
        End Try
    End Sub

    Private Sub LoadFromXML(ByRef FileName As String)

        Dim oXmlDoc As Xml.XmlDocument
        oXmlDoc = New Xml.XmlDocument
        oXmlDoc.Load(FileName)
        SBO_Application.LoadBatchActions(oXmlDoc.InnerXml)
    End Sub

    Private Sub LlenaGrid(valor As String)
        Try
            Dim QryStr As String


            QryStr = (String.Format("select Itemcode,(LineNum + 1) 'Linea', Dscription 'Descripcion', Quantity 'Cantidad', Price 'Precio', DiscPrcnt 'Descuento' from QUT1 where DocEntry = '{0}'", valor))
            oForm.DataSources.DataTables.Item(0).ExecuteQuery(QryStr)
            oGrid = oForm.Items.Item("grdDatos").Specific
            oGrid.DataTable = oForm.DataSources.DataTables.Item("MyDataTable")
            oGrid.Columns.GetEnumerator()
            oGrid.Columns.Item(0).Editable = False
            oGrid.Columns.Item(1).Editable = False
            oGrid.Columns.Item(2).Editable = False
            oGrid.Columns.Item(3).Editable = False
            oGrid.Columns.Item(4).Editable = False
            oGrid.Columns.Item(5).Editable = False
            CType(oGrid.Columns.Item(0), SAPbouiCOM.EditTextColumn).LinkedObjectType = 4
            linefinoriginal = oGrid.Rows.Count.ToString()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid)
            oGrid = Nothing
            GC.Collect()
        Catch ex As Exception
            'SBO_Application.MessageBox(ex.Message)
        End Try
    End Sub

    Private Sub CambiaTon(doc As String, val As String, line1 As Integer, line2 As Integer)
        Try
            Dim orecord As SAPbobsCOM.Recordset
            Dim linea1 As Integer
            Dim linea2 As Integer
            Dim precio As Decimal
            Dim precio2 As Decimal
            linea1 = Convert.ToInt32(line1)
            linea2 = Convert.ToInt32(line2)
            Dim valores As string

            precio = GetDouble(val)

            valores = Replace(Convert.ToString(precio), ",", ".")

            Dim oQuote As SAPbobsCOM.Documents
            Dim oError As Integer = -1
            Dim message As String = ""
            oQuote = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQuotations)

            For a As Integer = linea1 To linea2

                orecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                orecord.DoQuery("select (((" + valores + ")*(isnull(IWeight1,0)*QU.Quantity/1000))/Qu.Quantity) from oitm OI join QUT1 QU on OI.ItemCode = QU.ItemCode where QU.DocEntry = '" + doc + "' and LineNum = '" + a.ToString + "'")
                precio2 = Convert.ToDecimal(orecord.Fields.Item(0).Value)
                If oQuote.GetByKey(doc) Then
                    oQuote.Lines.SetCurrentLine(a)
                    oQuote.Lines.UnitPrice = precio2
                    oQuote.Lines.DiscountPercent = 0
                    oError = oQuote.Update()
                    If oError <> 0 Then
                        SBOApplication.SetStatusBarMessage("Error al actualizar Precio", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    End If

                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(orecord)
                orecord = Nothing
                GC.Collect()
            Next
            'Dim orecord As SAPbobsCOM.Recordset
            'orecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'orecord.DoQuery("Update QUT1 set Price = " + val + " where DocEntry = " + doc + " and LineNum between " + line1 + " and " + line2 + "")
            'orecord = Nothing
            'GC.Collect()
        Catch ex As Exception
            SBO_Application.MessageBox(ex.Message)
        End Try
    End Sub
    Private Sub CambiaPor(doc As String, porc As String, line1 As Integer, line2 As Integer)
        Try
            Dim linea1 As Integer
            Dim linea2 As Integer
            Dim Desc As Double
            linea1 = Convert.ToInt32(line1)
            linea2 = Convert.ToInt32(line2)

            Desc = Convert.ToDecimal(porc)

            Dim oQuote As SAPbobsCOM.Documents
            Dim oError As Integer = -1
            Dim message As String = ""
            oQuote = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQuotations)

            For a As Integer = linea1 To linea2


                If oQuote.GetByKey(doc) Then
                    oQuote.Lines.SetCurrentLine(a)
                    oQuote.Lines.DiscountPercent = Desc
                    oError = oQuote.Update()
                    If oError <> 0 Then
                        SBOApplication.SetStatusBarMessage(oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    End If

                End If
            Next
        Catch ex As Exception
            SBO_Application.MessageBox(ex.Message)
        End Try
    End Sub
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent

        Try
            If pVal.FormUID = "FrmValor" Then
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                    Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                    oCFLEvento = pVal
                    Dim sCFL_ID As String
                    sCFL_ID = oCFLEvento.ChooseFromListUID
                    Dim oForm As SAPbouiCOM.Form
                    oForm = SBO_Application.Forms.Item(FormUID)
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                    If oCFLEvento.BeforeAction = False Then
                        Dim oDataTable As SAPbouiCOM.DataTable
                        oDataTable = oCFLEvento.SelectedObjects
                        Dim val As String


                        If (pVal.ItemUID = "Item_0") Then
                            Try
                                Dim txtFactura As SAPbouiCOM.EditText = oForm.Items.Item("Item_0").Specific
                                val = oDataTable.GetValue("DocEntry", 0)
                                LlenaGrid(val)
                                txtFactura.Value = val
                            Catch ex As Exception

                            End Try

                        End If
                    End If
                End If

            End If
#Region "muestra porcentaje o valor"
            If pVal.ItemUID = "ChkPor" And pVal.FormUID = "FrmValor" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                Dim ChkPor As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkPor").Specific
                Dim ChkTon As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkTon").Specific
                Dim TxtPorc As SAPbouiCOM.EditText = oForm.Items.Item("ItemPorc").Specific
                Dim Lblpor As SAPbouiCOM.StaticText = oForm.Items.Item("Itempor").Specific
                Dim LblVal As SAPbouiCOM.StaticText = oForm.Items.Item("lblValor").Specific
                Dim txtValor As SAPbouiCOM.EditText = oForm.Items.Item("ItemVal").Specific
                If ChkPor.Checked = true and ChkTon.Checked = True then
                    TxtPorc.Item.Visible = False
                    Lblpor.Item.Visible = False
                    txtValor.Item.Visible = False
                    LblVal.Item.Visible = False
                    Return
                End if
                If ChkPor.Checked = True and ChkTon.Checked = False then
                    TxtPorc.Item.Visible = True
                    Lblpor.Item.Visible = True
                    txtValor.Item.Visible = False
                    LblVal.Item.Visible = False
                    Return
                End If
                If ChkPor.Checked = False and ChkTon.Checked = False then
                    TxtPorc.Item.Visible = False
                    Lblpor.Item.Visible = False
                    txtValor.Item.Visible = False
                    LblVal.Item.Visible = False
                    Return
                End If
                If ChkPor.Checked = False and ChkTon.Checked = True then
                    TxtPorc.Item.Visible = False
                    Lblpor.Item.Visible = False
                    txtValor.Item.Visible = True
                    LblVal.Item.Visible = True
                    Return
                End If
            End if

            If pVal.ItemUID = "ChkTon" And pVal.FormUID = "FrmValor" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                oForm.Items.Item("Item_0").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                Dim ChkTon As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkTon").Specific
                Dim ChkPor As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkPor").Specific
                Dim TxtPorc As SAPbouiCOM.EditText = oForm.Items.Item("ItemPorc").Specific
                Dim Lblpor As SAPbouiCOM.StaticText = oForm.Items.Item("Itempor").Specific
                Dim txtValor As SAPbouiCOM.EditText = oForm.Items.Item("ItemVal").Specific
                Dim LblVal As SAPbouiCOM.StaticText = oForm.Items.Item("lblValor").Specific
                If ChkTon.Checked = True and ChkPor.Checked = False then
                    TxtPorc.Item.Visible = False
                    Lblpor.Item.Visible = False
                    txtValor.Item.Visible = True
                    LblVal.Item.Visible = True
                    Return
                End If
                If ChkPor.Checked = true and ChkTon.Checked = True then
                    TxtPorc.Item.Visible = False
                    Lblpor.Item.Visible = False
                    txtValor.Item.Visible = False
                    LblVal.Item.Visible = False
                    Return
                End If
                If ChkTon.Checked = False And ChkPor.Checked = True then
                    TxtPorc.Item.Visible = True
                    Lblpor.Item.Visible = True
                    txtValor.Item.Visible = False
                    LblVal.Item.Visible = False
                    Return
                End If

                If ChkTon.Checked = False And ChkPor.Checked = False then
                    TxtPorc.Item.Visible = False
                    Lblpor.Item.Visible = False
                    txtValor.Item.Visible = False
                    LblVal.Item.Visible = False
                    Return
                End If

            End if
#End Region

            If pVal.ItemUID = "cmdOk" And pVal.FormUID = "FrmValor" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then


                Dim txtDocum As SAPbouiCOM.EditText = oForm.Items.Item("Item_0").Specific
                Dim txtValor As SAPbouiCOM.EditText = oForm.Items.Item("ItemVal").Specific
                Dim TxtLineini As SAPbouiCOM.EditText = oForm.Items.Item("ItemIni").Specific
                Dim TxtLineFin As SAPbouiCOM.EditText = oForm.Items.Item("ItemFin").Specific
                Dim TxtPorce As SAPbouiCOM.EditText = oForm.Items.Item("ItemPorc").Specific
                Dim LblVal As SAPbouiCOM.StaticText = oForm.Items.Item("lblValor").Specific
                Dim ChkTon As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkTon").Specific
                Dim ChkPor As SAPbouiCOM.CheckBox = oForm.Items.Item("ChkPor").Specific
                Dim TxtPorc As SAPbouiCOM.EditText = oForm.Items.Item("ItemPorc").Specific
                Dim Lblpor As SAPbouiCOM.StaticText = oForm.Items.Item("Itempor").Specific
                Dim Porce As Double
                Dim Docnum As String
                Dim Valor As Double
                Dim lineini As Integer
                Dim linefin As Integer
                 
#Region "valida campos en blanco"
                If ChkPor.Checked = True And ChkTon.Checked = True Then
                    SBO_Application.SetStatusBarMessage("Seleccione Unicamente Una Casilla", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Return
                End If
                If ChkPor.Checked = False And ChkTon.Checked = False Then
                    SBO_Application.SetStatusBarMessage("Seleccione Una Casilla", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Return
                End If
                If txtDocum.Value = "" Then
                    SBO_Application.SetStatusBarMessage("Debe de seleccionar un Documento", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Return
                End If
                'If txtValor.Value = "" Then
                '    SBO_Application.SetStatusBarMessage("Debe de ingresar un Valor", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                '    Return
                'End If
                If TxtLineini.Value = "" Then
                    SBO_Application.SetStatusBarMessage("Debe de ingresar un valor de Rango Inicial", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Return
                End If
                If TxtLineFin.Value = "" Then
                    SBO_Application.SetStatusBarMessage("Debe de ingresar un valor de Rango Final", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Return
                End If
#End Region


                Docnum = txtDocum.Value.Trim
                Valor = GetDouble(txtValor.Value)
                lineini = (TxtLineini.Value) - 1
                linefin = (TxtLineFin.Value) - 1
#Region "ChkTon"
                If ChkTon.Checked = True then
                    If txtValor.Value = "" or txtValor.Value <= 0 Then
                        SBO_Application.SetStatusBarMessage("Debe de ingresar un Valor", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        Return
                    End If
                    If lineini >= 0 And TxtLineFin.Value <= linefinoriginal and (TxtLineini.Value <= TxtLineFin.Value) Then
                        Dim resp = SBO_Application.MessageBox("Guardara el documento con NO." & txtDocum.Value.Trim, 1, "SI", "NO")
                        If resp = 1 Then
                            CambiaTon(Docnum, Valor, lineini, linefin)
                            LlenaGrid(Docnum)
                            oForm.Items.Item("Item_0").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                            txtValor.Value = ""
                            ChkPor.Checked = False
                            ChkTon.Checked = False
                            TxtPorc.Item.Visible = False
                            Lblpor.Item.Visible = False
                            TxtPorc.Item.Visible = False
                            Lblpor.Item.Visible = False
                            txtValor.Item.Visible = False
                            LblVal.Item.Visible = False
                            BubbleEvent = False
                            Return
                        End If
                        BubbleEvent = False
                        Return
                    Else
                        SBO_Application.SetStatusBarMessage("El numero debe de ser mayor a: 0 y menor o igual a: " + linefinoriginal.ToString())
                        BubbleEvent = False
                        Return
                    End If
                End If
#End Region
#Region "ChkPor"
                Porce = GetDouble(TxtPorce.Value)
                If ChkPor.Checked = True then
                    If TxtPorce.Value = "" or TxtPorce.Value <= 0 Then
                        SBO_Application.SetStatusBarMessage("Debe de ingresar un Porcentaje", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        Return
                    End If
                    If (lineini >= 0 And TxtLineFin.Value <= linefinoriginal) and (TxtLineini.Value <= TxtLineFin.Value) Then
                        Dim resp = SBO_Application.MessageBox("Guardara el documento con NO." & txtDocum.Value.Trim, 1, "SI", "NO")
                        If resp = 1 Then
                            CambiaPor(Docnum, Porce, lineini, linefin)
                            LlenaGrid(Docnum)
                            TxtPorc.Value = ""
                            oForm.Items.Item("Item_0").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                            ChkPor.Checked = False
                            ChkTon.Checked = False
                            TxtPorc.Item.Visible = False
                            Lblpor.Item.Visible = False
                            txtValor.Item.Visible = False
                            LblVal.Item.Visible = False
                            
                            txtDocum.Item.Refresh
                            BubbleEvent = False
                            Return
                        End If
                        BubbleEvent = False
                        Return
                    Else
                        SBO_Application.SetStatusBarMessage("El numero debe de ser mayor a: 0 y menor o igual a: " + linefinoriginal.ToString())
                        BubbleEvent = False
                        Return
                    End If
                End If
#End Region

            End If

            'If pVal.ItemUID = "grdDatos" And pVal.FormUID = "FrmValor" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
            '   Dim txtValor As SAPbouiCOM.EditText = oForm.Items.Item("ItemVal").Specific
            '    Dim val As string

            '    oGrid = oForm.Items.Item("grdDatos").Specific
            '    val = oGrid.DataTable.GetValue(1,oGrid.GetDataTableRowIndex(pVal.Row)).ToString
            '    SBO_Application.SetStatusBarMessage(val, SAPbouiCOM.BoMessageTime.bmt_Medium, True)

            'End If

            If pVal.ItemUID = "btnCancel" And pVal.FormUID = "FrmValor" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                Try
                    oForm.Close
                    BubbleEvent = False
                    Return
                Catch ex As Exception

                End Try
                
            End If
        Catch ex As Exception
            'SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            BubbleEvent = False
            Return
        End Try
    End Sub
    Public Shared Function GetDouble(ByVal doublestring As String) As Double
        Dim retval As Decimal
        Dim sep As String = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator

        Double.TryParse(Replace(Replace(doublestring, ".", sep), ",", sep), retval)
        Return retval
    End Function
End Class
