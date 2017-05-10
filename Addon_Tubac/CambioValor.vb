Public Class CambioValor

    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oCompany As SAPbobsCOM.Company
    Private oBusinessForm As SAPbouiCOM.Form
    Public oFilters As SAPbouiCOM.EventFilters
    Public oFilter As SAPbouiCOM.EventFilter

    Private AddStarted As Boolean                ' Flag that indicates "Add" process started

    Private RedFlag As Boolean

    Private Function SetConnectionContext() As Integer

        Dim sCookie As String

        Dim sConnectionContext As String


        Try

            '// First initialize the Company object

            oCompany = New SAPbobsCOM.Company

            '// Acquire the connection context cookie from the DI API.

            sCookie = oCompany.GetContextCookie

            '// Retrieve the connection context string from the UI API using the

            '// acquired cookie.

            sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie)

            '// before setting the SBO Login Context make sure the company is not

            '// connected

            If oCompany.Connected = True Then

                oCompany.Disconnect()

            End If

            '// Set the connection context information to the DI API.

            SetConnectionContext = oCompany.SetSboLoginContext(sConnectionContext)

        Catch ex As Exception

            MsgBox(ex.Message)

        End Try

    End Function

    Private Function ConnectToCompany() As Integer

        '// Establish the connection to the company database.

        ConnectToCompany = oCompany.Connect

    End Function

    Private Sub Class_Init()
        Try
            '//*************************************************************

            '// set SBO_Application with an initialized application object

            '//*************************************************************

            SetApplication()

            '//*************************************************************

            '// Set The Connection Context

            '//*************************************************************

            If Not SetConnectionContext() = 0 Then

                SBO_Application.MessageBox("Failed setting a connection to DI API")

                End ' Terminating the Add-On Application

            End If

            '//*************************************************************

            '// Connect To The Company Data Base

            '//*************************************************************

            If Not ConnectToCompany() = 0 Then

                SBO_Application.MessageBox("Failed connecting to the company's Data Base")

                End ' Terminating the Add-On Application

            End If

            '//*************************************************************

            '// send an "hello world" message

            '//*************************************************************

            SBO_Application.SetStatusBarMessage("DI Connected To: " & oCompany.CompanyName & vbNewLine & "Add-on is loaded", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            'SetNewTax("01", "512 0% a 22 % pago al exterior", SAPbobsCOM.WithholdingTaxCodeCategoryEnum.wtcc_Invoice, SAPbobsCOM.WithholdingTaxCodeBaseTypeEnum.wtcbt_Net, 100, "512", "1-1-010-10-000")
            'SetNewTax("02", "513 0% a 22 % pago al exterior", SAPbobsCOM.WithholdingTaxCodeCategoryEnum.wtcc_Invoice, SAPbobsCOM.WithholdingTaxCodeBaseTypeEnum.wtcbt_Net, 100, "513", "1-1-010-10-000")
            'SetNewTax("03", "513A 0% a 22 % pago al exterior", SAPbobsCOM.WithholdingTaxCodeCategoryEnum.wtcc_Invoice, SAPbobsCOM.WithholdingTaxCodeBaseTypeEnum.wtcbt_Net, 100, "513A", "_SYS00000000128")
            'SetNewTax("04", "514 0% a 22 % pago al exterior", SAPbobsCOM.WithholdingTaxCodeCategoryEnum.wtcc_Invoice, SAPbobsCOM.WithholdingTaxCodeBaseTypeEnum.wtcbt_Net, 100, "514", "_SYS00000000128")

            Utiles.SBOApplication = Me.SBO_Application
            Utiles.Company = Me.oCompany
            'PROBAR()
            SBOApplication.StatusBar.SetText("Iniciando Addon Cambio de Valor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
        End Try
    End Sub

    'Public Sub SetApplication()


    '    Dim SboGuiApi As SAPbouiCOM.SboGuiApi
    '    Dim sConnectionString As String
    '    Try
    '        SboGuiApi = New SAPbouiCOM.SboGuiApi
    '        sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"

    '        '// get an initialized application object

    '        SboGuiApi.Connect(sConnectionString)
    '        SBO_Application = SboGuiApi.GetApplication()
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Ocurrio un error")
    '    End Try

    'End Sub

    Private Sub SetApplication()

        AddStarted = False

        RedFlag = False

        '*******************************************************************

        '// Use an SboGuiApi object to establish connection

        '// with the SAP Business One application and return an

        '// initialized application object

        '*******************************************************************
        Try
            Dim SboGuiApi As SAPbouiCOM.SboGuiApi

            Dim sConnectionString As String

            SboGuiApi = New SAPbouiCOM.SboGuiApi

            '// by following the steps specified above, the following

            '// statement should be sufficient for either development or run mode
            If Environment.GetCommandLineArgs.Length > 1 Then
                sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
            Else
                sConnectionString = Environment.GetCommandLineArgs.GetValue(0)
            End If

            'sConnectionString = Environment.GetCommandLineArgs.GetValue(1) '"0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"

            '// connect to a running SBO Application

            SboGuiApi.Connect(sConnectionString)

            '// get an initialized application object

            SBO_Application = SboGuiApi.GetApplication()
        Catch ex As Exception

        End Try


    End Sub

    Private Sub SetFilters()

        '// Create a new EventFilters object

        oFilters = New SAPbouiCOM.EventFilters



        '// add an event type to the container

        '// this method returns an EventFilter object

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CLICK)

        'oFilter = oFilter.Add(SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK)
        'oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)

        'oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_VALIDATE)

        'oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS)

        'oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN)

        'oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)

        ' oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK)
        ' oFilter.AddEx("60006") 'Quotation Form


        oFilter.AddEx("60004")
        'oFilter.AddEx("139") 'Orders Form
        'oFilter.AddEx("133") 'Invoice Form
        'oFilter.AddEx("169") 'Main Menu
        SBO_Application.SetFilter(oFilters)

    End Sub
    Public Sub New()

        MyBase.New()

        Class_Init()

        AddMenuItems()

        SetFilters()


    End Sub

    Private Sub AddMenuItems()

        Try

            '//******************************************************************
            '// Let's add a separator, a pop-up menu item and a string menu item
            '//******************************************************************

            Dim oMenus As SAPbouiCOM.Menus
            Dim oMenuItem As SAPbouiCOM.MenuItem

            '// Get the menus collection from the application
            oMenus = SBO_Application.Menus
            '--------------------------------------------
            'Save an XML file containing the menus...
            '--------------------------------------------
            'sXML = SBO_Application.Menus.GetAsXML
            'Dim xmlD As System.Xml.XmlDocument
            'xmlD = New System.Xml.XmlDocument
            'xmlD.LoadXml(sXML)
            'xmlD.Save("c:\\mnu.xml")
            '--------------------------------------------


            Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
            oCreationPackage = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
            oMenuItem = SBO_Application.Menus.Item("43520") 'modulo de menu'

            Dim sPath As String

            sPath = System.Windows.Forms.Application.StartupPath
            sPath = sPath.Remove(sPath.Length - 3, 3)

            '// find the place in wich you want to add your menu item
            '// in this example I chose to add my menu item under
            '// SAP Business One.
            If SBO_Application.Menus.Exists("CambioValor") Then
                SBO_Application.Menus.RemoveEx("CambioValor")
            End If
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            oCreationPackage.UniqueID = "CambioValor"
            oCreationPackage.String = "Cambio de Valor"
            oCreationPackage.Enabled = True
            oCreationPackage.Image = Replace(System.Windows.Forms.Application.StartupPath & "\valor.png", "\\", "\")
            oCreationPackage.Position = 1

            oMenus = oMenuItem.SubMenus

            Try ' If the manu already exists this code will fail
                oMenus.AddEx(oCreationPackage)

                '// Get the menu collection of the newly added pop-up item
                oMenuItem = SBO_Application.Menus.Item("CambioValor")
                oMenus = oMenuItem.SubMenus

                '// Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "pantalla1"
                oCreationPackage.String = "Cambio de Valor"
                oMenus.AddEx(oCreationPackage)


            Catch er As Exception ' Menu already exists
                'SBO_Application.MessageBox("Menu Already Exists")
            End Try

        Catch ex As Exception
            SBO_Application.MessageBox(ex.Message)
        End Try
    End Sub

    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent

        If (pVal.MenuUID = "pantalla1") And (pVal.BeforeAction = False) Then
            Dim initFrm As New pantalla1
            BubbleEvent = False
        End If

    End Sub
End Class
