Imports System.IO
Imports System.Xml

Public Class Utils
    Public Shared Function RemFileExt(ByVal FileName As String) As String
        Dim FileNameNoExt As String = Path.GetFileNameWithoutExtension(FileName)
        Dim DirPath As String = Path.GetDirectoryName(FileName)
        If DirPath = String.Empty Then Return FileNameNoExt
        Return DirPath + "\" + FileNameNoExt
    End Function

    Public Shared Function GetBaseName(ByVal PathStr As String) As String
        Return Path.GetDirectoryName(PathStr)
    End Function

    Public Shared Function GetLeafName(ByVal PathStr As String) As String
        Return Path.GetFileName(PathStr)
    End Function

    Public Shared Function MakeValidDbTableName(ByVal TableName As String) As String
        'Return "Table_" + TableName.Replace("-", "_").Replace("&", "_").Replace(Chr(32), "_")
        'Return TableName.Replace("-", "_").Replace("&", "_").Replace(Chr(32), "_")
        Return TableName.Replace("&", "_").Replace(Chr(32), "_")
    End Function

    Public Shared Sub CreateAppExeConfig(AppCfgFile As String)
        Dim AppCfg As New XmlDocument
        Dim CfgNode As XmlNode, StartupNode As XmlNode
        Dim SubNode As XmlNode, Attr As XmlAttribute

        '<configuration>
        '  <startup useLegacyV2RuntimeActivationPolicy="true">
        '    <supportedRuntime version="v4.0" sku=".NETFramework,Version=v4.0"/>
        '    <requiredRuntime version="v4.0.20506" />
        '  </startup>
        '</configuration>

        Dim Node As XmlNode = AppCfg.CreateXmlDeclaration("1.0", Nothing, Nothing)
        AppCfg.AppendChild(Node)

        '<configuration>
        CfgNode = AppCfg.CreateElement("configuration")

        '  <startup useLegacyV2RuntimeActivationPolicy="true">
        StartupNode = AppCfg.CreateElement("startup")
        Attr = AppCfg.CreateAttribute("useLegacyV2RuntimeActivationPolicy")
        Attr.Value = "true"
        StartupNode.Attributes.Append(Attr)

        '    <supportedRuntime version="v4.0" sku=".NETFramework,Version=v4.0"/>
        SubNode = AppCfg.CreateElement("supportedRuntime")
        Attr = AppCfg.CreateAttribute("version")
        Attr.Value = "v4.0"
        SubNode.Attributes.Append(Attr)

        Attr = AppCfg.CreateAttribute("sku")
        Attr.Value = ".NETFramework,Version=v4.0"
        SubNode.Attributes.Append(Attr)
        StartupNode.AppendChild(SubNode)

        '    <requiredRuntime version="v4.0.20506" />
        SubNode = AppCfg.CreateElement("requiredRuntime")
        Attr = AppCfg.CreateAttribute("version")
        Attr.Value = "v4.0.20506"
        SubNode.Attributes.Append(Attr)
        StartupNode.AppendChild(SubNode)

        CfgNode.AppendChild(StartupNode)

        AppCfg.AppendChild(CfgNode)

        AppCfg.Save(AppCfgFile)
    End Sub

    Public Shared Function TruncatePath(InPath As String, Optional ByVal DirLevel As Short = 3) As String
        Dim i As Integer, TruncPath As String = ""
        Dim FileName As String, Count As Integer

        If String.IsNullOrEmpty(InPath) Then Return String.Empty

        FileName = "\" + Utils.GetLeafName(InPath)
        InPath = Utils.GetBaseName(InPath)

        Count = 0
        For i = Len(InPath) - 1 To 0 Step -1
            Select Case InPath(i).ToString
                Case ":"
                    TruncPath = "..." + InPath.Substring(i + 1) + FileName
                    Exit For
                Case "\"
                    TruncPath = "..." + InPath.Substring(i) + FileName
                    Count += 1
            End Select
            If Count = DirLevel Then Exit For
        Next

        Return TruncPath
    End Function
End Class
