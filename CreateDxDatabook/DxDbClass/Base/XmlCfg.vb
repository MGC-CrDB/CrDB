Imports System.IO
Imports System.Xml
Imports System.Windows.Forms

Public Class XmlCfg

#Region "Declarations"
    Public Event Notify(Msg As String, LeadingSpaces As Integer, Type As MsgType)

    ' ini file sections
    Public Enum Sections
        General
        PropertyImport
    End Enum

    ' ini file key names
    Public Enum Keys
        DynamicGUI
        LibraryLmcPath
        AllowDuplicatePartNumbers
        UseCentralLibrarySymbols
        UseOdbcAlias
        OdbcAliasName
        UseAlternateDbFolder
        AlternateDbFolder
        DbcFileInLibFolder
        UseCellPins
        PropertySideFile
        PropertyImportEnabled
        PropertyImportFieldDelimiter
        NotValid
    End Enum

    Public Enum KeyDefs
        TAB
        AccessMDB
        SQLiteDB3
        Undefined
    End Enum

    Private CfgXmlFile As String
    Private CfgXmlDoc As XmlDocument
    Private RootNode As XmlNode
#End Region

#Region "Constructor"
    Public Sub New(nCfgXmlFile As String)
        Me.CfgXmlDoc = New XmlDocument
        Me.CfgXmlFile = nCfgXmlFile
        If File.Exists(nCfgXmlFile) Then
            Me.CfgXmlDoc.Load(Me.CfgXmlFile)
            Me.RootNode = Me.CfgXmlDoc.LastChild
            Exit Sub
        End If

        Dim Assembly As String = Utils.RemFileExt(Utils.GetLeafName(Me.CfgXmlFile))
        Me.CreateDefaultConfigXml(Assembly)
    End Sub
#End Region

#Region "Public Methods"
    Public Sub PrintConfigEntries(Optional Saved As Boolean = False)
        Dim Section As XmlNode, cKey As XmlNode
        Dim SecEnum As System.Collections.IEnumerator
        Dim KeyEnum As System.Collections.IEnumerator
        Dim Val As String, KeyValue As String

        '##########################
        ' Get sections
        SecEnum = Me.RootNode.GetEnumerator()

        RaiseEvent Notify("Current Config:", 5, MsgType.Txt)
        While SecEnum.MoveNext()
            Section = CType(SecEnum.Current, XmlNode)
            RaiseEvent Notify("Section: " + Section.Name, 10, MsgType.Txt)

            KeyEnum = Section.ChildNodes.GetEnumerator()
            While KeyEnum.MoveNext()
                cKey = CType(KeyEnum.Current, XmlNode)
                KeyValue = Me.GetXmlCfgNodeAttrValue("Value", cKey)

                Select Case Me.GetConfigKeyName(cKey.Name)
                    Case Keys.AllowDuplicatePartNumbers '[GENERAL]
                        RaiseEvent Notify("... Allow duplicate Partnumbers = " + KeyValue, 15, MsgType.Txt)

                    'Case Keys.UseCentralLibrarySymbols '[GENERAL]
                    '    RaiseEvent Notify("... Use Central Library Symbols = " + KeyValue, 15, MsgType.Txt)

                    Case Keys.UseOdbcAlias '[GENERAL]
                        RaiseEvent Notify("... Use ODBC Alias = " + KeyValue, 15, MsgType.Txt)

                    Case Keys.OdbcAliasName '[GENERAL]
                        Val = KeyValue
                        If Val = String.Empty Then Val = "<not set>"
                        RaiseEvent Notify("... ODBC Alias Name = " + Val, 15, MsgType.Txt)

                    Case Keys.UseAlternateDbFolder  '[GENERAL]
                        RaiseEvent Notify("... Use alternate Database Folder = " + KeyValue, 15, MsgType.Txt)

                    Case Keys.DbcFileInLibFolder  '[GENERAL]
                        RaiseEvent Notify("... Store DBC File in Library Folder = " + KeyValue, 15, MsgType.Txt)

                    Case Keys.UseCellPins '[GENERAL]
                        RaiseEvent Notify("... Use Cell Pins = " + KeyValue, 15, MsgType.Txt)

                    Case Keys.PropertyImportEnabled '[PRP_IMPORT]
                        RaiseEvent Notify("... Import additional Properties Enabled = " + KeyValue, 15, MsgType.Txt)

                    Case Keys.PropertyImportFieldDelimiter '[PRP_IMPORT]
                        RaiseEvent Notify("... Property Import Field Delimiter = " + KeyValue, 15, MsgType.Txt)
                End Select
            End While
        End While
        RaiseEvent Notify("End Config", 5, MsgType.Txt)
        RaiseEvent Notify("", 0, MsgType.Txt)

        If Saved Then
            RaiseEvent Notify("Config File saved: " + Me.CfgXmlFile, 5, MsgType.Txt)
            RaiseEvent Notify("", 0, MsgType.Txt)
        Else
            RaiseEvent Notify("Execute 'Help > Commandline Arguments' for batch usage.", 5, MsgType.Txt)
            RaiseEvent Notify("", 0, MsgType.Txt)
        End If
    End Sub

    Public Function SaveConfig() As Boolean
        If IsNothing(Me.CfgXmlDoc) Then Return False
        If Directory.Exists(Path.GetDirectoryName(Me.CfgXmlFile)) Then
            Me.CfgXmlDoc.Save(Me.CfgXmlFile)
            Return True
        Else
            Return False
        End If
    End Function

    Public Function GetXmlConfigKeyValue(Section As Sections, Key As Keys) As String
        Dim SectionNode As XmlNode, SecKeyNode As XmlNode
        Dim KeyEnum As System.Collections.IEnumerator
        Dim KeyValue As String

        '##########################
        ' Get sections
        SectionNode = Me.GetXmlConfigSectionNode(Section)
        If IsNothing(SectionNode) Then Return KeyDefs.Undefined.ToString

        KeyEnum = SectionNode.ChildNodes.GetEnumerator()
        While KeyEnum.MoveNext()
            SecKeyNode = CType(KeyEnum.Current, XmlNode)
            If SecKeyNode.Name.ToLower = Key.ToString.ToLower Then
                KeyValue = Me.GetXmlCfgNodeAttrValue("Value", SecKeyNode)
                If String.IsNullOrEmpty(KeyValue) Then
                    Return KeyDefs.Undefined.ToString
                Else
                    Return KeyValue
                End If
            End If
        End While

        Return KeyDefs.Undefined.ToString
    End Function

    Public Function GetIniDynamicGui() As Boolean
        Dim RetVal As String = Me.GetXmlConfigSectionKeyValue(Sections.General, Keys.DynamicGUI, Boolean.TrueString)
        If RetVal = String.Empty Then Return False
        If RetVal = Boolean.TrueString Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function GetIniCentLibLmcFile() As String
        Return Me.GetXmlConfigSectionKeyValue(Sections.General, Keys.LibraryLmcPath, String.Empty)
    End Function

    Public Function GetIniAllowDuplPartnos() As Boolean
        Dim RetVal As String = Me.GetXmlConfigSectionKeyValue(Sections.General, Keys.AllowDuplicatePartNumbers, Boolean.FalseString)
        If RetVal = String.Empty Then Return False
        If RetVal = Boolean.TrueString Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function GetIniUseOdbcAlias() As Boolean
        Dim RetVal As String = Me.GetXmlConfigSectionKeyValue(Sections.General, Keys.UseOdbcAlias, Boolean.FalseString)
        If RetVal = Boolean.TrueString Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function GetIniOdbcAliasName() As String
        Dim RetVal As String = Me.GetXmlConfigSectionKeyValue(Sections.General, Keys.OdbcAliasName, String.Empty)
        If RetVal = String.Empty Then RetVal = KeyDefs.Undefined.ToString
        Return RetVal
    End Function

    Public Function GetIniUseOtherMdbDir() As Boolean
        Dim RetVal As String = Me.GetXmlConfigSectionKeyValue(Sections.General, Keys.UseAlternateDbFolder, Boolean.FalseString)
        If RetVal = Boolean.TrueString Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function GetIniUseCentLinSymbols() As Boolean
        Dim RetVal As String = Me.GetXmlConfigSectionKeyValue(Sections.General, Keys.UseCentralLibrarySymbols, Boolean.FalseString)
        If RetVal = Boolean.TrueString Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function GetIniOtherMdbDir() As String
        Return Me.GetXmlConfigSectionKeyValue(Sections.General, Keys.AlternateDbFolder, String.Empty)
    End Function

    Public Function GetIniDbcInLibDir() As Boolean
        Dim RetVal As String = Me.GetXmlConfigSectionKeyValue(Sections.General, Keys.DbcFileInLibFolder, Boolean.TrueString)
        If RetVal = Boolean.TrueString Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function GetIniUseCellPins() As Boolean
        Dim RetVal As String = Me.GetXmlConfigSectionKeyValue(Sections.General, Keys.UseCellPins, Boolean.FalseString)
        If RetVal = Boolean.TrueString Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function GetIniMdbImpPrpSideFile() As String
        Return Me.GetXmlConfigSectionKeyValue(Sections.PropertyImport, Keys.PropertySideFile, String.Empty)
    End Function

    Public Function GetIniMdbImpPrpEnable() As Boolean
        Dim RetVal As String = Me.GetXmlConfigSectionKeyValue(Sections.PropertyImport, Keys.PropertyImportEnabled, Boolean.FalseString)
        If RetVal = Boolean.TrueString Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function GetIniMdbImpFldDelim() As String
        Dim RetVal As String = Me.GetXmlConfigSectionKeyValue(Sections.PropertyImport, Keys.PropertyImportFieldDelimiter, KeyDefs.TAB.ToString)
        If RetVal = String.Empty Then RetVal = KeyDefs.TAB.ToString
        Return RetVal
    End Function
#End Region

#Region "Set .INI Config From Control Methods"
    Public Sub SetIniDbcInLibDir(ByRef DbcInLibDir As CheckBox, Optional SaveConfig As Boolean = False)
        Me.SetIniKeyFromCheckBox(DbcInLibDir, Sections.General, Keys.DbcFileInLibFolder)
        If SaveConfig Then Me.SaveConfig()
    End Sub

    Public Sub SetIniDynamicGui(ByRef ChkDynamicGui As CheckBox, Optional SaveConfig As Boolean = False)
        Me.SetIniKeyFromCheckBox(ChkDynamicGui, Sections.General, Keys.DynamicGUI)
        If SaveConfig Then Me.SaveConfig()
    End Sub

    Public Sub SetIniAllowDuplPartnumbers(ByRef ChkAllowDuplPartnos As CheckBox, Optional SaveConfig As Boolean = False)
        Me.SetIniKeyFromCheckBox(ChkAllowDuplPartnos, Sections.General, Keys.AllowDuplicatePartNumbers)
        If SaveConfig Then Me.SaveConfig()
    End Sub

    Public Sub SetIniUseCenLibSymbols(ByRef ChkAllowDuplPartnos As CheckBox, Optional SaveConfig As Boolean = False)
        Me.SetIniKeyFromCheckBox(ChkAllowDuplPartnos, Sections.General, Keys.UseCentralLibrarySymbols)
        If SaveConfig Then Me.SaveConfig()
    End Sub

    Public Sub SetIniUseOtherMdbDir(ByRef UseOtherMdbDir As CheckBox, Optional SaveConfig As Boolean = False)
        Me.SetIniKeyFromCheckBox(UseOtherMdbDir, Sections.General, Keys.UseAlternateDbFolder)
        If SaveConfig Then Me.SaveConfig()
    End Sub

    Public Sub SetIniOdbcAliasParms(UseOdbcAlias As CheckBox, ComboOdbcAlias As ComboBox, Optional SaveConfig As Boolean = False)
        Me.SetIniKeyFromCheckBox(UseOdbcAlias, Sections.General, Keys.UseOdbcAlias)
        If UseOdbcAlias.Checked Then
            Me.SetIniKeyFromComboBox(ComboOdbcAlias, Sections.General, Keys.OdbcAliasName)
        Else
            Me.SetXmlConfigSectionKeyValue(Sections.General, Keys.OdbcAliasName, String.Empty)
        End If
        If SaveConfig Then Me.SaveConfig()
    End Sub

    Public Sub SetIniPrpImportParms(PrpImpEnable As CheckBox, ComboFieldDelim As ComboBox, Optional SaveConfig As Boolean = False)
        Me.SetIniKeyFromCheckBox(PrpImpEnable, Sections.PropertyImport, Keys.PropertyImportEnabled)
        If PrpImpEnable.Checked Then
            Me.SetIniKeyFromComboBox(ComboFieldDelim, Sections.PropertyImport, Keys.PropertyImportFieldDelimiter, KeyDefs.TAB.ToString)
        Else
            Me.SetXmlConfigSectionKeyValue(Sections.PropertyImport, Keys.PropertyImportFieldDelimiter, KeyDefs.TAB.ToString)
        End If
        If SaveConfig Then Me.SaveConfig()
    End Sub

    Public Sub SetIniUseCellPinCount(ByRef UseCellPinCount As CheckBox, Optional SaveConfig As Boolean = False)
        Me.SetIniKeyFromCheckBox(UseCellPinCount, Sections.General, Keys.UseCellPins)
        If SaveConfig Then Me.SaveConfig()
    End Sub

    Public Sub SetIniCentLibLmcPath(ByRef EntryLmcPath As TextBox, Optional SaveConfig As Boolean = False)
        Me.SetIniKeyFromTextBox(EntryLmcPath, Sections.General, Keys.LibraryLmcPath)
        If SaveConfig Then Me.SaveConfig()
    End Sub

    Public Sub SetIniOtherMdbDir(ByRef EntryMdbDir As TextBox, Optional SaveConfig As Boolean = False)
        Me.SetIniKeyFromTextBox(EntryMdbDir, Sections.General, Keys.AlternateDbFolder)
        If SaveConfig Then Me.SaveConfig()
    End Sub

    Public Sub SetIniPrpSideFilePath(ByRef EntryPrpFile As TextBox, Optional SaveConfig As Boolean = False)
        Me.SetIniKeyFromTextBox(EntryPrpFile, Sections.PropertyImport, Keys.PropertySideFile)
        If SaveConfig Then Me.SaveConfig()
    End Sub
#End Region

#Region "Private Methods"
    Private Sub CreateDefaultConfigXml(Assembly As String)
        Dim SectionNode As XmlNode
        If Me.CfgXmlDoc.ChildNodes.Count > 0 Then Exit Sub

        ' create Declaration
        Dim Decl As XmlDeclaration = Me.CfgXmlDoc.CreateXmlDeclaration("1.0", "", "")
        Me.CfgXmlDoc.AppendChild(CType(Decl, XmlDeclaration))

        Me.RootNode = CType(Me.CfgXmlDoc.CreateElement(Assembly), XmlNode)
        Me.CfgXmlDoc.AppendChild(Me.RootNode)

        '------------------
        ' add SECTION names
        SectionNode = Me.AddXmlConfigSectionNode(Sections.General)

        '----------------
        ' add SOURCE keys
        Me.AddXmlConfigSectionKey(SectionNode, Keys.DynamicGUI)
        Me.AddXmlConfigSectionKey(SectionNode, Keys.LibraryLmcPath)
        Me.AddXmlConfigSectionKey(SectionNode, Keys.AllowDuplicatePartNumbers)
        Me.AddXmlConfigSectionKey(SectionNode, Keys.UseCentralLibrarySymbols)

        Me.AddXmlConfigSectionKey(SectionNode, Keys.UseOdbcAlias)
        Me.AddXmlConfigSectionKey(SectionNode, Keys.OdbcAliasName)

        Me.AddXmlConfigSectionKey(SectionNode, Keys.UseAlternateDbFolder)
        Me.AddXmlConfigSectionKey(SectionNode, Keys.AlternateDbFolder)
        Me.AddXmlConfigSectionKey(SectionNode, Keys.DbcFileInLibFolder)
        Me.AddXmlConfigSectionKey(SectionNode, Keys.UseCellPins)

        SectionNode = Me.AddXmlConfigSectionNode(Sections.PropertyImport)
        Me.AddXmlConfigSectionKey(SectionNode, Keys.PropertySideFile)
        Me.AddXmlConfigSectionKey(SectionNode, Keys.PropertyImportEnabled)
        Me.AddXmlConfigSectionKey(SectionNode, Keys.PropertyImportFieldDelimiter)

        Me.SaveConfig()
    End Sub

    Private Function GetConfigKeyName(KeyName As String) As Keys
        If System.Enum.IsDefined(GetType(Keys), KeyName) Then
            Return CType(System.[Enum].Parse(GetType(Keys), KeyName), Keys)
        Else
            Return Keys.NotValid
        End If
    End Function

    Private Function GetConfigKeyDefault(KeyName As Keys) As String
        Select Case KeyName
            Case Keys.AllowDuplicatePartNumbers
                Return Boolean.FalseString
            Case Keys.AlternateDbFolder
                Return String.Empty
            Case Keys.DbcFileInLibFolder
                Return Boolean.TrueString
            Case Keys.DynamicGUI
                Return Boolean.TrueString
            Case Keys.LibraryLmcPath
                Return String.Empty
            Case Keys.OdbcAliasName
                Return String.Empty
            Case Keys.PropertyImportEnabled
                Return Boolean.FalseString
            Case Keys.PropertyImportFieldDelimiter
                Return KeyDefs.TAB.ToString()
            Case Keys.PropertySideFile
                Return String.Empty
            Case Keys.UseAlternateDbFolder
                Return Boolean.FalseString
            Case Keys.UseCellPins
                Return Boolean.FalseString
            Case Keys.UseCentralLibrarySymbols
                Return Boolean.TrueString
            Case Keys.UseOdbcAlias
                Return Boolean.FalseString
            Case Else
                Return String.Empty
        End Select
    End Function
#End Region

#Region "Private Get/Set Control Methods"
    Private Sub SetIniKeyFromCheckBox(ByRef CheckBox As CheckBox, Section As Sections, SectionKey As Keys)
        If CheckBox.Checked Then
            Me.SetXmlConfigSectionKeyValue(Section, SectionKey, Boolean.TrueString)
        Else
            Me.SetXmlConfigSectionKeyValue(Section, SectionKey, Boolean.FalseString)
        End If
    End Sub

    Private Sub SetIniKeyFromComboBox(ByRef ComboBox As ComboBox, Section As Sections, SectionKey As Keys, Optional DefaultValue As String = "NONE")
        If ComboBox.Items.Count = 0 Then
            Me.SetXmlConfigSectionKeyValue(Section, SectionKey, DefaultValue)
        End If
        Select Case ComboBox.Text
            Case ""
                Me.SetXmlConfigSectionKeyValue(Section, SectionKey, DefaultValue)
            Case Else
                Me.SetXmlConfigSectionKeyValue(Section, SectionKey, UCase(ComboBox.Text))
        End Select
    End Sub

    Private Sub SetIniKeyFromListbox(Listbox As ListBox, Section As Sections, SectionKey As Keys)
        Dim i As Integer, KeyValue As String = ""
        For i = 0 To Listbox.Items.Count - 1
            Select Case i
                Case 0
                    If i = Listbox.Items.Count - 1 Then
                        KeyValue = Listbox.Items.Item(i).ToString
                    Else
                        KeyValue = Listbox.Items.Item(i).ToString + "|"
                    End If
                Case Listbox.Items.Count - 1
                    KeyValue += Listbox.Items.Item(i).ToString
                Case Else
                    KeyValue += Listbox.Items.Item(i).ToString + "|"
            End Select
        Next
        Me.SetXmlConfigSectionKeyValue(Section, SectionKey, KeyValue)
    End Sub

    Private Sub SetIniKeyFromRadioBut(RadioBut As RadioButton, Section As Sections, SectionKey As Keys, SetValue As String)
        If RadioBut.Checked Then
            Me.SetXmlConfigSectionKeyValue(Section, SectionKey, SetValue)
        End If
    End Sub

    Private Sub SetIniKeyFromTextBox(TextBox As TextBox, Section As Sections, SectionKey As Keys)
        Dim DefaultValue As String = Me.GetConfigKeyDefault(SectionKey)
        If TextBox.Text = String.Empty Then
            Me.SetXmlConfigSectionKeyValue(Section, SectionKey, DefaultValue)
        Else
            Me.SetXmlConfigSectionKeyValue(Section, SectionKey, TextBox.Text)
        End If
    End Sub
#End Region

#Region "Private XML Methods"
    Private Function GetXmlConfigSectionKeyValue(Section As Sections, SectionKey As Keys, DefaultValue As String) As String
        Dim KeyValue As String = Me.GetXmlConfigKeyValue(Section, SectionKey)
        If KeyValue = KeyDefs.Undefined.ToString Or KeyValue = String.Empty Then
            Return DefaultValue
        End If
        Return KeyValue
    End Function

    Private Sub SetXmlConfigSectionKeyValue(Section As Sections, SectionKey As Keys, KeyValue As String)
        Dim SectionNode As XmlNode, SecEnum As IEnumerator
        Dim SecKeyNode As XmlNode, KeyEnum As IEnumerator

        SecEnum = Me.RootNode.ChildNodes.GetEnumerator()
        While SecEnum.MoveNext()
            SectionNode = CType(SecEnum.Current, XmlNode)
            If SectionNode.Name.ToLower = Section.ToString.ToLower Then
                KeyEnum = SectionNode.ChildNodes.GetEnumerator()
                While KeyEnum.MoveNext()
                    SecKeyNode = CType(KeyEnum.Current, XmlNode)
                    If SecKeyNode.Name.ToLower = SectionKey.ToString.ToLower Then
                        Me.SetXmlCfgNodeAttrValue("Value", KeyValue, SecKeyNode)
                    End If
                End While
            End If
        End While
    End Sub

    Private Function AddXmlConfigSectionNode(Section As Sections) As XmlNode
        Dim ND As XmlNode
        If IsNothing(Me.CfgXmlDoc) Then Return Nothing

        ND = CType(Me.CfgXmlDoc.CreateElement(Section.ToString), XmlNode)
        Me.RootNode.AppendChild(ND)

        Return ND
    End Function

    Private Function AddXmlConfigSectionKey(SectionNode As XmlNode, NewKey As Keys) As Boolean
        Dim EL As XmlElement, Attr As XmlAttribute

        ' DynamicGUI,  True 
        ' LibraryLmcPath, String.Empty 
        ' AllowDuplicatePartNumbers,  False 
        ' UseCentralLibrarySymbols,  True 

        ' DatabaseType,  SQLiteDB3 
        ' UseOdbcAlias,  False 
        ' OdbcAliasName, String.Empty 

        ' UseAlternateDbFolder,  False 
        ' AlternateDbFolder, String.Empty 
        ' DbcFileInLibFolder,  True 
        ' UseCellPins,  False 

        ' PropertySideFile, String.Empty)
        ' PropertyImportEnabled,  False 
        ' PropertyImportFieldDelimiter,  tab 
        Try
            EL = Me.CfgXmlDoc.CreateElement(NewKey.ToString)
            Attr = Me.CfgXmlDoc.CreateAttribute("Value")
            Attr.Value = Me.GetConfigKeyDefault(NewKey)
            EL.Attributes.Append(Attr)

            SectionNode.AppendChild(CType(EL, XmlNode))
        Catch ex As Exception
            Return False
        End Try

        Return True
    End Function

    Private Function GetXmlConfigSectionNode(Section As Sections) As XmlNode
        Dim SectionNode As XmlNode, SecEnum As IEnumerator
        SecEnum = Me.RootNode.ChildNodes.GetEnumerator()
        While SecEnum.MoveNext()
            SectionNode = CType(SecEnum.Current, XmlNode)
            If SectionNode.Name.ToLower = Section.ToString.ToLower Then
                Return SectionNode
            End If
        End While
        Return Nothing
    End Function

    Private Function GetXmlCfgNodeAttrValue(AttrName As String, ND As XmlNode) As String
        Dim Element As XmlElement = CType(ND, XmlElement)
        For Each Attr As XmlAttribute In Element.Attributes
            If Attr.Name.ToLower = AttrName.ToLower Then
                Return Attr.Value
            End If
        Next
        Return String.Empty
    End Function

    Private Sub SetXmlCfgNodeAttrValue(AttrName As String, AttrValue As String, ND As XmlNode)
        Dim Element As XmlElement = CType(ND, XmlElement)
        For Each Attr As XmlAttribute In Element.Attributes
            If Attr.Name.ToLower = AttrName.ToLower Then
                Attr.Value = AttrValue
                Exit Sub
            End If
        Next
    End Sub
#End Region

End Class
