Module Utiles
    Private _SBOApplication As SAPbouiCOM.Application
    Public Property SBOApplication() As SAPbouiCOM.Application
        Get
            Return _SBOApplication
        End Get
        Set(ByVal value As SAPbouiCOM.Application)
            _SBOApplication = value
        End Set
    End Property

    Private _Company As SAPbobsCOM.Company

    Public Property Company() As SAPbobsCOM.Company
        Get
            Return _Company
        End Get
        Set(ByVal value As SAPbobsCOM.Company)
            _Company = value
        End Set
    End Property

    Public Function ActivateFormIsOpen(ByVal SboApplication As SAPbouiCOM.Application, ByVal FormID As String) As Boolean
        Try
            Dim result As Boolean = False
            For x = 0 To SboApplication.Forms.Count - 1
                If SboApplication.Forms.Item(x).UniqueID = FormID Then
                    SboApplication.Forms.Item(x).Select()
                    result = True
                    Exit For
                End If
            Next
            Return result
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

End Module