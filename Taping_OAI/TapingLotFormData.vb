Public Class TapingLotFormData

    Private _PRODUCT As String
    Public Property A_PRODUCT() As String
        Get
            Return _PRODUCT
        End Get
        Set(ByVal value As String)
            _PRODUCT = value
        End Set
    End Property

    Private _SPEC As String
    Public Property B_SPEC() As String
        Get
            Return _SPEC
        End Get
        Set(ByVal value As String)
            _SPEC = value
        End Set
    End Property

    Private _FREQ As String
    Public Property C_FREQ() As String
        Get
            Return _FREQ
        End Get
        Set(ByVal value As String)
            _FREQ = value
        End Set
    End Property

    Private _INA_CODE As String
    Public Property D_INA_CODE() As String
        Get
            Return _INA_CODE
        End Get
        Set(ByVal value As String)
            _INA_CODE = value
        End Set
    End Property

    Private _OGLOTNO As String
    Public Property E_OGLOTNO() As String
        Get
            Return _OGLOTNO
        End Get
        Set(ByVal value As String)
            _OGLOTNO = value
        End Set
    End Property

    Private _LOTNO As String
    Public Property F_LOTNO() As String
        Get
            Return _LOTNO
        End Get
        Set(ByVal value As String)
            _LOTNO = value
        End Set
    End Property

    Private _WKCODE As String
    Public Property G_WKCODE() As String
        Get
            Return _WKCODE
        End Get
        Set(ByVal value As String)
            _WKCODE = value
        End Set
    End Property

    Private _USEDQTY As String
    Public Property H_USEDQTY() As Int32
        Get
            Return _USEDQTY
        End Get
        Set(ByVal value As Int32)
            _USEDQTY = value
        End Set
    End Property

End Class
