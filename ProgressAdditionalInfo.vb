Public Class ProgressAdditionalInfo
    Private _displayName As String
    Public Property DisplayNameProgress() As String
        Get
            Return _displayName
        End Get
        Set(ByVal value As String)
            _displayName = value
        End Set
    End Property

    Private _folderCount As Integer
    Public Property FolderCount() As Integer
        Get
            Return _folderCount
        End Get
        Set(ByVal value As Integer)
            _folderCount = value
        End Set
    End Property

    Public Sub New(ByVal count As Integer, ByVal name As String)
        Me.DisplayNameProgress = name
        Me.FolderCount = count
    End Sub
End Class
