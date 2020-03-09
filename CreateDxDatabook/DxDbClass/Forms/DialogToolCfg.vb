Imports Microsoft.Win32
Imports System.Windows.Forms
Imports Keys = DxDbClass.XmlCfg.Keys
Imports KeyDefs = DxDbClass.XmlCfg.KeyDefs
Imports Sections = DxDbClass.XmlCfg.Sections

Public Class DialogToolCfg

#Region "Properties And Constructor"
    Private Config As XmlCfg

    Public Shared Function CreateAndShow(ByRef Config As XmlCfg) As DialogResult
        Dim Dlg As New DialogToolCfg
        Dlg.Config = Config
        Return Dlg.ShowDialog()
    End Function
#End Region

#Region "Form Events"
    Private Sub DialogToolCfg_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ' set config controls
        Me.SetCheckBoxControl(Me.ChkDynamicGui, Sections.General, Keys.DynamicGUI, True)
        Me.SetCheckBoxControl(Me.ChkDuplPartnos, Sections.General, Keys.AllowDuplicatePartNumbers)

        Me.SetCtrlsOdbcAlias(Me.ChkUseOdbcAlias, Me.CmbOdbcAlias)
        Me.SetCtrlsPdbImpEnable(Me.ChkEnaPrpImport, Me.CmbFieldDelim)
        Me.SetCheckBoxControl(ChkUseCellPinCount, Sections.General, Keys.UseCellPins)

        Me.SetCheckBoxControl(Me.ChkUseCLibSymbols, Sections.General, Keys.UseCentralLibrarySymbols)
        Me.ChkUseCLibSymbols.Checked = True
        Me.ChkUseCLibSymbols.Visible = False

        Me.SetCheckBoxControl(ChkUseOtherMdbDir, Sections.General, Keys.UseAlternateDbFolder)
        Me.SetCheckBoxControl(Me.ChkDbcInLibDir, Sections.General, Keys.DbcFileInLibFolder, True)
        If Not Me.ChkUseOtherMdbDir.Checked Then
            Me.ChkDbcInLibDir.Checked = True
            Me.Config.SetIniDbcInLibDir(Me.ChkDbcInLibDir)
        End If
    End Sub

    Private Sub BtnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOK.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK

        ' set config values
        Me.Config.SetIniDynamicGui(Me.ChkDynamicGui)
        Me.Config.SetIniAllowDuplPartnumbers(Me.ChkDuplPartnos)
        Me.Config.SetIniUseCenLibSymbols(Me.ChkUseCLibSymbols)
        Me.Config.SetIniUseOtherMdbDir(Me.ChkUseOtherMdbDir)
        Me.Config.SetIniDbcInLibDir(Me.ChkDbcInLibDir)
        Me.Config.SetIniOdbcAliasParms(Me.ChkUseOdbcAlias, Me.CmbOdbcAlias)
        Me.Config.SetIniPrpImportParms(Me.ChkEnaPrpImport, Me.CmbFieldDelim)
        Me.Config.SetIniUseCellPinCount(Me.ChkUseCellPinCount)

        Me.Config.SaveConfig()

        ' close dialog
        Me.Close()
    End Sub

    Private Sub BtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCancel.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub ChkUseOtherMdbDir_CheckedChanged(sender As Object, e As System.EventArgs) Handles ChkUseOtherMdbDir.CheckedChanged
        Dim CB As CheckBox = CType(sender, CheckBox)
        If CB.Checked Then
            Me.ChkDbcInLibDir.Enabled = True
        Else
            Me.ChkDbcInLibDir.Checked = True
            Me.ChkDbcInLibDir.Enabled = False
        End If
    End Sub

    Private Sub ChkUseOdbcAlias_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkUseOdbcAlias.CheckedChanged
        Dim CB As CheckBox = CType(sender, CheckBox)
        If CB.Checked Then
            Me.CmbOdbcAlias.Enabled = True
        Else
            Me.CmbOdbcAlias.Enabled = False
        End If
    End Sub

    Private Sub ChkEnaPrpImport_CheckedChanged(sender As Object, e As System.EventArgs) Handles ChkEnaPrpImport.CheckedChanged
        If Me.ChkEnaPrpImport.Checked Then
            Me.LblFieldDelim.Enabled = True
            Me.CmbFieldDelim.Enabled = True
        Else
            Me.LblFieldDelim.Enabled = False
            Me.CmbFieldDelim.Enabled = False
        End If
    End Sub
#End Region

#Region "Private Set Control Methods"
    Private Sub SetCtrlsOdbcAlias(ByRef ChkUseOdbcAlias As CheckBox, ByRef ComboOdbcAlias As ComboBox)
        Me.SetCheckBoxControl(ChkUseOdbcAlias, Sections.General, Keys.UseOdbcAlias)

        Dim KeyValue As String = Me.Config.GetXmlConfigKeyValue(Sections.General, Keys.OdbcAliasName)
        Dim OdbcDrivers As List(Of String) = Me.GetODBCDrivers()

        For Each OdbcDriver As String In OdbcDrivers
            ComboOdbcAlias.Items.Add(OdbcDriver)
            If KeyValue = OdbcDriver Then
                ComboOdbcAlias.Text = KeyValue
            End If
        Next

        If ComboOdbcAlias.Text = String.Empty And ComboOdbcAlias.Items.Count > 0 Then
            ComboOdbcAlias.Text = ComboOdbcAlias.Items(0).ToString
        End If

        If ChkUseOdbcAlias.Checked Then
            ComboOdbcAlias.Enabled = True
        Else
            ComboOdbcAlias.Enabled = False
        End If
    End Sub

    Private Sub SetCtrlsPdbImpEnable(ByRef ChkSapEnable As CheckBox, ByRef ComboFieldDelim As ComboBox)
        Me.SetCheckBoxControl(ChkSapEnable, Sections.PropertyImport, Keys.PropertyImportEnabled)
        ComboFieldDelim.Text = Me.Config.GetXmlConfigKeyValue(Sections.PropertyImport, Keys.PropertyImportFieldDelimiter)
    End Sub

    Private Function GetODBCDrivers() As List(Of String)
        Dim RegKey As RegistryKey, RegArray As String()
        Dim ValArray As New List(Of String)

        RegKey = My.Computer.Registry.CurrentUser.OpenSubKey("Software\ODBC\ODBC.INI\ODBC Data Sources", False)
        RegArray = RegKey.GetValueNames()
        RegKey.Close()

        If Not IsNothing(RegArray) Then
            For Each Val As String In RegArray
                ValArray.Add(Val)
            Next
        End If

        Return ValArray
    End Function

    Private Sub SetCheckBoxControl(ByRef CheckBox As CheckBox, Section As Sections, SectionKey As Keys, Optional DefaultValue As Boolean = False)
        Dim KeyValue As String = Me.Config.GetXmlConfigKeyValue(Section, SectionKey)
        If KeyValue = KeyDefs.Undefined.ToString Or KeyValue = String.Empty Then
            CheckBox.Checked = DefaultValue
        Else
            If KeyValue = Boolean.TrueString Then
                CheckBox.Checked = True
            Else
                CheckBox.Checked = False
            End If
        End If
    End Sub

    Public Sub SetRadioButtonControl(ByRef RadioBut As RadioButton, Section As Sections, SectionKey As Keys, Optional DefaultValue As Boolean = False)
        Dim KeyValue As String = Me.Config.GetXmlConfigKeyValue(Section, SectionKey)
        If KeyValue = String.Empty Then RadioBut.Checked = DefaultValue
        If KeyValue = Boolean.TrueString Then
            RadioBut.Checked = True
        Else
            RadioBut.Checked = False
        End If
    End Sub

    Public Sub SetTextBoxControl(ByRef TextBox As TextBox, Section As Sections, SectionKey As Keys, Optional DefaultValue As String = "")
        Dim KeyValue As String
        KeyValue = Me.Config.GetXmlConfigKeyValue(Section, SectionKey)
        If KeyValue = KeyDefs.Undefined.ToString Or KeyValue = String.Empty Then
            TextBox.Text = DefaultValue
        Else
            TextBox.Text = KeyValue
        End If
    End Sub

    Public Sub SetComboBoxControl(ByRef ComboBox As ComboBox, Section As Sections, SectionKey As Keys, Optional DefaultIndex As Integer = 0)
        Dim KeyValue As String
        KeyValue = Me.Config.GetXmlConfigKeyValue(Section, SectionKey)
        If ComboBox.Items.Count = 0 Then Exit Sub
        If KeyValue = KeyDefs.Undefined.ToString Or KeyValue = String.Empty Then
            ComboBox.SelectedIndex = DefaultIndex
        Else
            ComboBox.Text = KeyValue
        End If
    End Sub

    Public Sub SetListboxControl(ByRef ListBox As ListBox, Section As Sections, SectionKey As Keys, Optional DefaultValue As String = "")
        Dim i As Integer, SplitVect As String(), KeyValue As String

        KeyValue = Me.Config.GetXmlConfigKeyValue(Section, SectionKey)

        If KeyValue = KeyDefs.Undefined.ToString Or KeyValue = String.Empty Then
            KeyValue = DefaultValue
        End If
        ListBox.Items.Clear()

        If KeyValue = "" Then Exit Sub

        SplitVect = Split(KeyValue, "|")
        For i = 0 To UBound(SplitVect)
            ListBox.Items.Add(SplitVect(i))
        Next
    End Sub
#End Region

End Class
