Attribute VB_Name = "ModuleConnection"
Global cn As New ADODB.Connection
Sub User_Connection()
    cn.Provider = "SQLOLEDB"
    cn.Properties("Data Source").Value = "localhost"
    cn.Properties("Initial Catalog").Value = "SistemaCorretor"
    cn.Properties("User ID").Value = "sa"
    cn.Properties("Password").Value = "1q2w3e4r@#$"
    cn.Open
End Sub
