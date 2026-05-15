Imports System.ComponentModel
Imports System.Windows.Forms
Imports Infragistics.win

<ProvideProperty("ABSHasButton", GetType(Component))> _
<ProvideProperty("ABSTableName", GetType(Component))> _
<ProvideProperty("ABSLookUpTableName", GetType(Component))> _
<ProvideProperty("ABSColumnName", GetType(Component))> _
<ProvideProperty("ABSColumnCaption", GetType(Component))> _
<ProvideProperty("ABSParentColumnName", GetType(Component))> _
<ProvideProperty("ABSPrecedentKeys", GetType(Component))> _
<ProvideProperty("ABSSecurityCodes", GetType(Component))> _
<ProvideProperty("ABSValueRequired", GetType(Component))> _
<ProvideProperty("ABSViewName", GetType(Component))> _
<ProvideProperty("ABSBindToTable", GetType(Component))> _
Public Class ABSX
    Inherits System.ComponentModel.Component
    Implements IExtenderProvider
    Private htbABSHasButton As New Hashtable
    Private htbABSTableName As New Hashtable
    Private htbABSLookUpTableName As New Hashtable
    Private htbABSColumnName As New Hashtable
    Private htbABSColumnCaption As New Hashtable
    Private htbABSParentColumnName As New Hashtable
    Private htbABSPrecedentKeys As New Hashtable
    Public htbABSSecurityCodes As New Hashtable
    Private htbABSValueRequired As New Hashtable
    Private htbABSViewName As New Hashtable
    Private htbABSBindToTable As New Hashtable
    Public dicCOLUMN_NAME As New Dictionary(Of String, Control)
    Public Event MyExtProvEvent(ByVal ABSColumnName As String)
    Public TABLE_NAME_base As String

    Public Function CanExtend(ByVal extendee As Object) As Boolean Implements System.ComponentModel.IExtenderProvider.CanExtend
        Return True
    End Function

#Region "Functions to Return Controls keyed to the ABSColumnName"

#Region "xxxFor - Functions returning an Ultra-typed control relating to COLUMN_NAME"

    Public Function txtFor(ByVal COLUMN_NAME As String, Optional ByVal ok_if_missing As Boolean = False) As UltraWinEditors.UltraTextEditor
        Dim txt As New UltraWinEditors.UltraTextEditor
        Dim x As Control = CtlFor(COLUMN_NAME, ok_if_missing)
        If x IsNot Nothing Then
            txt = DirectCast(x, UltraWinEditors.UltraTextEditor)
        End If
        Return txt
    End Function

    Public Function chkFor(ByVal COLUMN_NAME As String, Optional ByVal ok_if_missing As Boolean = False) As UltraWinEditors.UltraCheckEditor
        Dim chk As New UltraWinEditors.UltraCheckEditor
        Dim x As Control = CtlFor(COLUMN_NAME, ok_if_missing)
        If x IsNot Nothing Then
            chk = DirectCast(x, UltraWinEditors.UltraCheckEditor)
        End If
        Return chk
    End Function

    Public Function optFor(ByVal COLUMN_NAME As String, Optional ByVal ok_if_missing As Boolean = False) As UltraWinEditors.UltraOptionSet
        Dim opt As New UltraWinEditors.UltraOptionSet
        Dim x As Control = CtlFor(COLUMN_NAME, ok_if_missing)
        If x IsNot Nothing Then
            opt = DirectCast(x, UltraWinEditors.UltraOptionSet)
        End If

        Return opt
    End Function

    Public Function numFor(ByVal COLUMN_NAME As String, Optional ByVal ok_if_missing As Boolean = False) As UltraWinEditors.UltraNumericEditor
        'Dim num As New UltraWinEditors.UltraNumericEditor
        Dim num As UltraWinEditors.UltraNumericEditor = Nothing
        Dim x As Control = CtlFor(COLUMN_NAME, ok_if_missing)
        If x IsNot Nothing Then
            num = DirectCast(x, UltraWinEditors.UltraNumericEditor)
        Else
            MsgBox("Problem with Control related to " & COLUMN_NAME, MsgBoxStyle.OkOnly, "Design Issue")
            Throw New Exception("Problem with " & COLUMN_NAME, Nothing)
        End If
        Return num
    End Function

    Public Function medFor(ByVal COLUMN_NAME As String, Optional ByVal ok_if_missing As Boolean = False) As UltraWinMaskedEdit.UltraMaskedEdit
        Dim med As New UltraWinMaskedEdit.UltraMaskedEdit
        Dim x As Control = CtlFor(COLUMN_NAME, ok_if_missing)
        If x IsNot Nothing Then
            med = DirectCast(x, UltraWinMaskedEdit.UltraMaskedEdit)
        End If
        Return med
    End Function

    Public Function dteFor(ByVal COLUMN_NAME As String, Optional ByVal ok_if_missing As Boolean = False) As UltraWinEditors.UltraDateTimeEditor
        Dim dte As New UltraWinEditors.UltraDateTimeEditor
        Dim x As Control = CtlFor(COLUMN_NAME, ok_if_missing)
        If x IsNot Nothing Then
            dte = DirectCast(x, UltraWinEditors.UltraDateTimeEditor)
        End If
        Return dte
    End Function

    Public Function cbeFor(ByVal COLUMN_NAME As String, Optional ByVal ok_if_missing As Boolean = False) As UltraWinEditors.UltraComboEditor
        Dim cbe As New UltraWinEditors.UltraComboEditor
        Dim x As Control = CtlFor(COLUMN_NAME, ok_if_missing)
        If x IsNot Nothing Then
            cbe = DirectCast(x, UltraWinEditors.UltraComboEditor)
        End If
        Return cbe
    End Function

    Public Function cmbFor(ByVal COLUMN_NAME As String, Optional ByVal ok_if_missing As Boolean = False) As UltraWinGrid.UltraCombo
        Dim cmb As New UltraWinGrid.UltraCombo
        Dim x As Control = CtlFor(COLUMN_NAME, ok_if_missing)
        If x IsNot Nothing Then
            cmb = DirectCast(x, UltraWinGrid.UltraCombo)
        End If
        Return cmb
    End Function
#End Region

#Region "ctlFor - Functions returning generic controls relating to COLUMN_NAME or PARENT_COLUMN_NAME"
    Public Function CtlFor(ByVal COLUMN_NAME As String, Optional ByVal ok_if_missing As Boolean = False) As Control

        If dicCOLUMN_NAME.ContainsKey(COLUMN_NAME) Then
            Return dicCOLUMN_NAME(COLUMN_NAME)
        Else
            If InStr(COLUMN_NAME, ".") = 0 Then
                If dicCOLUMN_NAME.ContainsKey(TABLE_NAME_base & "." & COLUMN_NAME) Then
                    Return dicCOLUMN_NAME(TABLE_NAME_base & "." & COLUMN_NAME)
                Else
                    For Each TC As String In dicCOLUMN_NAME.Keys
                        If InStr(TC, ".") <> 0 Then
                            Dim C As String = Split(TC, ".")(1)
                            If COLUMN_NAME = C Then
                                Return dicCOLUMN_NAME(TC)
                            End If
                        End If
                    Next
                End If
            End If
            If Not ok_if_missing Then
                MsgBox("Problem with Control related to " & COLUMN_NAME, MsgBoxStyle.OkOnly, "Design Issue")
                Throw New Exception("Problem with " & COLUMN_NAME, Nothing)
            End If
            Return Nothing
        End If
    End Function

    Public Function CtlsFor(ByVal parent_COLUMN_NAME As String) As List(Of Control)
        Dim CTLs As New List(Of Control)
        For Each CTLx As System.Collections.DictionaryEntry In htbABSParentColumnName
            If CTLx.Value = parent_COLUMN_NAME Then
                CTLs.Add(DirectCast(CTLx.Key, Control))
            End If
        Next
        Return CTLs
    End Function
#End Region

#End Region

#Region "Functions to Set/Get Properties"

#Region "ABSHasButton"
    Public Sub SetABSHasButton(ByVal ctrl As Component, ByVal value As Boolean)
        'This method will be called each time you set a value through the properties window or when you directly call the SetDescription method. The control (ctrl) and the new description (value) are received as arguments. This method uses the hashtable (htbABSColumnName) to store the description. You can also see that we are adding a handler on specific event of the controls. This is needed to be able to raise our own event (more on this later). 

        If TypeOf ctrl Is UltraWinEditors.UltraTextEditor Then
        Else
            value = False
        End If

        If Not value Then
            If htbABSHasButton.Contains(ctrl) Then
                'remove the control from the hashtable if its value is deleted
                htbABSHasButton.Remove(ctrl)
                RemoveHandler CType(ctrl, Control).Enter, _
                            AddressOf ShowABSHasButton

            End If
        Else
            If Not htbABSHasButton.Contains(ctrl) Then
                If TypeOf ctrl Is Control Then

                    AddHandler CType(ctrl, Control).Enter, _
                            AddressOf ShowABSHasButton
                ElseIf TypeOf ctrl Is MenuItem Then
                    AddHandler CType(ctrl, MenuItem).Select, _
                            AddressOf ShowABSHasButton
                End If
            End If
            htbABSHasButton.Item(ctrl) = value
        End If
    End Sub
    <DisplayName("ABSHasButton")> _
    <Category("ABS")> _
    <Description("For use with UltraTextEditor Only! - True will cause a Code Button to appear in the Text Box.  If you want a button, you should specify a ViewName, unless the ColumnName will suffice as the ViewName.  It does not make sense to set this property to True if you leave the ColumnName and ViewName blank.")> _
    <DefaultValue(False)> _
    Public Function GetABSHasButton(ByVal ctrl As Component) As Boolean
        'The getter method is much simpler. It checks the hashtable to return the value previously set: 

        If TypeOf ctrl Is UltraWinEditors.UltraTextEditor Then
            'If htb.Contains(ctrl) Then Return DirectCast(htb(ctrl), DataRow).Item("ABSHasButton") Else Return Nothing
            If htbABSHasButton.Contains(ctrl) Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If

    End Function

    Private Sub ShowABSHasButton(ByVal sender As Object, ByVal e As System.EventArgs)
        RaiseEvent MyExtProvEvent(GetABSHasButton(CType(sender, Component)))
    End Sub

#End Region

#Region "ABSTableName"
    Public Sub SetABSTableName(ByVal ctrl As Component, ByVal value As String)
        'This method will be called each time you set a value through the properties window or when you directly call the SetDescription method. The control (ctrl) and the new description (value) are received as arguments. This method uses the hashtable (htbABSTableName) to store the description. You can also see that we are adding a handler on specific event of the controls. This is needed to be able to raise our own event (more on this later). 

        'To be sure we don't have the Nothing value
        If value Is Nothing Then value = ""

        If value.Length = 0 AndAlso htbABSTableName.Contains(ctrl) Then
            'remove the control from the hashtable if its value is deleted
            htbABSTableName.Remove(ctrl)
            RemoveHandler CType(ctrl, Control).Enter, _
                        AddressOf ShowABSTableName

        ElseIf value.Length > 0 Then
            If Not htbABSTableName.Contains(ctrl) Then
                If TypeOf ctrl Is Control Then

                    AddHandler CType(ctrl, Control).Enter, _
                            AddressOf ShowABSTableName
                ElseIf TypeOf ctrl Is MenuItem Then
                    AddHandler CType(ctrl, MenuItem).Select, _
                            AddressOf ShowABSTableName
                End If
            End If
            htbABSTableName.Item(ctrl) = value
        End If
    End Sub

    <DisplayName("ABSTableName")> _
    <Category("ABS")> _
    <Description("Contains the Name of the Table associated with this Control, for DataBinding (not for LookUp).")> _
    <DefaultValue("")> _
    Public Function GetABSTableName(ByVal ctrl As Component) As String
        'The getter method is much simpler. It checks the hashtable to return the value previously set: 

        If htbABSTableName.Contains(ctrl) Then
            Return htbABSTableName(ctrl).ToString
        Else
            Return ""
        End If
    End Function

    Private Sub ShowABSTableName(ByVal sender As Object, ByVal e As System.EventArgs)
        RaiseEvent MyExtProvEvent(GetABSTableName(CType(sender, Component)))
    End Sub

#End Region

#Region "ABSLookUpTableName"
    Public Sub SetABSLookUpTableName(ByVal ctrl As Component, ByVal value As String)
        'This method will be called each time you set a value through the properties window or when you directly call the SetDescription method. The control (ctrl) and the new description (value) are received as arguments. This method uses the hashtable (htbABSLookUpTableName) to store the description. You can also see that we are adding a handler on specific event of the controls. This is needed to be able to raise our own event (more on this later). 

        'To be sure we don't have the Nothing value
        If value Is Nothing Then value = ""

        If value.Length = 0 AndAlso htbABSLookUpTableName.Contains(ctrl) Then
            'remove the control from the hashtable if its value is deleted
            htbABSLookUpTableName.Remove(ctrl)
            RemoveHandler CType(ctrl, Control).Enter, _
                        AddressOf ShowABSLookUpTableName

        ElseIf value.Length > 0 Then
            If Not htbABSLookUpTableName.Contains(ctrl) Then
                If TypeOf ctrl Is Control Then

                    AddHandler CType(ctrl, Control).Enter, _
                            AddressOf ShowABSLookUpTableName
                ElseIf TypeOf ctrl Is MenuItem Then
                    AddHandler CType(ctrl, MenuItem).Select, _
                            AddressOf ShowABSLookUpTableName
                End If
            End If
            htbABSLookUpTableName.Item(ctrl) = value
        End If
    End Sub

    <DisplayName("ABSLookUpTableName")> _
    <Category("ABS")> _
    <Description("Contains the Name of the Table be used to Look Up or Validate the Code, or to Populate Other Columns whose Parent Column is this column.")> _
    <DefaultValue("")> _
    Public Function GetABSLookUpTableName(ByVal ctrl As Component) As String
        'The getter method is much simpler. It checks the hashtable to return the value previously set: 

        If htbABSLookUpTableName.Contains(ctrl) Then
            Return htbABSLookUpTableName(ctrl).ToString
        Else
            Return ""
        End If
    End Function

    Private Sub ShowABSLookUpTableName(ByVal sender As Object, ByVal e As System.EventArgs)
        RaiseEvent MyExtProvEvent(GetABSLookUpTableName(CType(sender, Component)))
    End Sub

#End Region

#Region "ABSColumnName"
    Public Sub SetABSColumnName(ByVal ctrl As Component, ByVal value As String)
        'This method will be called each time you set a value through the properties window or when you directly call the SetDescription method. The control (ctrl) and the new description (value) are received as arguments. This method uses the hashtable (htbABSColumnName) to store the description. You can also see that we are adding a handler on specific event of the controls. This is needed to be able to raise our own event (more on this later). 
        'If value = "BY_SEG4" Then Stop
        'To be sure we don't have the Nothing value
        If value Is Nothing Then value = ""

        If value.Length = 0 AndAlso htbABSColumnName.Contains(ctrl) Then
            'remove the control from the hashtable if its value is deleted
            htbABSColumnName.Remove(ctrl)
            RemoveHandler CType(ctrl, Control).Enter, _
                        AddressOf ShowABSColumnName

        ElseIf value.Length > 0 Then
            If Not htbABSColumnName.Contains(ctrl) Then
                If TypeOf ctrl Is Control Then

                    AddHandler CType(ctrl, Control).Enter, _
                            AddressOf ShowABSColumnName
                ElseIf TypeOf ctrl Is MenuItem Then
                    AddHandler CType(ctrl, MenuItem).Select, _
                            AddressOf ShowABSColumnName
                End If
            End If

            Try
                htbABSColumnName.Item(ctrl) = value
            Catch ex As Exception
                Stop
            End Try

            'Try
            '    htbABSColumnName.Item(DirectCast(ctrl, Control)) = value
            'Catch ex As Exception
            '    Stop
            'End Try

        End If
    End Sub

    <DisplayName("ABSColumnName")> _
    <Category("ABS")> _
    <Description("Contains the Name of the Column associated with this Control.")> _
    <DefaultValue("")> _
    Public Function GetABSColumnName(ByVal ctrl As Component) As String
        'The getter method is much simpler. It checks the hashtable to return the value previously set: 

        If htbABSColumnName.Contains(ctrl) Then
            Return htbABSColumnName(ctrl).ToString
        Else
            Return ""
        End If
    End Function

    Private Sub ShowABSColumnName(ByVal sender As Object, ByVal e As System.EventArgs)
        RaiseEvent MyExtProvEvent(GetABSColumnName(CType(sender, Component)))
    End Sub

#End Region

#Region "ABSColumnCaption"
    Public Sub SetABSColumnCaption(ByVal ctrl As Component, ByVal value As String)
        'This method will be called each time you set a value through the properties window or when you directly call the SetDescription method. The control (ctrl) and the new description (value) are received as arguments. This method uses the hashtable (htbABSColumnName) to store the description. You can also see that we are adding a handler on specific event of the controls. This is needed to be able to raise our own event (more on this later). 

        'To be sure we don't have the Nothing value
        If value Is Nothing Then value = ""

        If value.Length = 0 AndAlso htbABSColumnCaption.Contains(ctrl) Then
            'remove the control from the hashtable if its value is deleted
            htbABSColumnCaption.Remove(ctrl)
            RemoveHandler CType(ctrl, Control).Enter, _
                        AddressOf ShowABSColumnCaption

        ElseIf value.Length > 0 Then
            If Not htbABSColumnCaption.Contains(ctrl) Then
                If TypeOf ctrl Is Control Then

                    AddHandler CType(ctrl, Control).Enter, _
                            AddressOf ShowABSColumnCaption
                ElseIf TypeOf ctrl Is MenuItem Then
                    AddHandler CType(ctrl, MenuItem).Select, _
                            AddressOf ShowABSColumnCaption
                End If
            End If
            htbABSColumnCaption.Item(ctrl) = value
        End If
    End Sub
    <DisplayName("ABSColumnCaption")> _
    <Category("ABS")> _
    <Description("Contains the Caption for the Column associated with this Control.")> _
    <DefaultValue("")> _
    Public Function GetABSColumnCaption(ByVal ctrl As Component) As String
        'The getter method is much simpler. It checks the hashtable to return the value previously set: 

        If htbABSColumnCaption.Contains(ctrl) Then
            Return htbABSColumnCaption(ctrl).ToString
        Else
            Return ""
        End If
    End Function
    Private Sub ShowABSColumnCaption(ByVal sender As Object, ByVal e As System.EventArgs)
        RaiseEvent MyExtProvEvent(GetABSColumnCaption(CType(sender, Component)))
    End Sub
#End Region

#Region "ABSParentColumnName"

    Public Sub SetABSParentColumnName(ByVal ctrl As Component, ByVal value As String)
        'This method will be called each time you set a value through the properties window or when you directly call the SetDescription method. The control (ctrl) and the new description (value) are received as arguments. This method uses the hashtable (htbABSColumnName) to store the description. You can also see that we are adding a handler on specific event of the controls. This is needed to be able to raise our own event (more on this later). 
        'If value = "VEND_CODE" Then Stop
        'To be sure we don't have the Nothing value
        If value Is Nothing Then value = ""

        If value.Length = 0 AndAlso htbABSParentColumnName.Contains(ctrl) Then
            'remove the control from the hashtable if its value is deleted
            htbABSParentColumnName.Remove(ctrl)
            RemoveHandler CType(ctrl, Control).Enter, _
                        AddressOf ShowABSParentColumnName

        ElseIf value.Length > 0 Then
            If Not htbABSParentColumnName.Contains(ctrl) Then
                If TypeOf ctrl Is Control Then

                    AddHandler CType(ctrl, Control).Enter, _
                            AddressOf ShowABSParentColumnName
                ElseIf TypeOf ctrl Is MenuItem Then
                    AddHandler CType(ctrl, MenuItem).Select, _
                            AddressOf ShowABSParentColumnName
                End If
            End If
            htbABSParentColumnName.Item(ctrl) = value
        End If
    End Sub

    <DisplayName("ABSParentColumnName")> _
    <Category("ABS")> _
    <Description("Contains the Name of the Parent Column related to this Control.  When the value in the control keyed to the Parent Column Name is Changed, the value in this control will change to the corresponding value it the Column Name indicated in this Control.")> _
    <DefaultValue("")> _
    Public Function GetABSParentColumnName(ByVal ctrl As Component) As String
        'The getter method is much simpler. It checks the hashtable to return the value previously set: 

        If htbABSParentColumnName.Contains(ctrl) Then
            Return htbABSParentColumnName(ctrl).ToString
        Else
            Return ""
        End If
    End Function

    Private Sub ShowABSParentColumnName(ByVal sender As Object, ByVal e As System.EventArgs)
        RaiseEvent MyExtProvEvent(GetABSParentColumnName(CType(sender, Component)))
    End Sub

#End Region

#Region "ABSPrecedentKeys"

    Public Sub SetABSPrecedentKeys(ByVal ctrl As Component, ByVal value As String)
        'This method will be called each time you set a value through the properties window or when you directly call the SetDescription method. The control (ctrl) and the new description (value) are received as arguments. This method uses the hashtable (htbABSColumnName) to store the description. You can also see that we are adding a handler on specific event of the controls. This is needed to be able to raise our own event (more on this later). 
        'If value = "VEND_CODE" Then Stop
        'To be sure we don't have the Nothing value
        If value Is Nothing Then value = ""

        If value.Length = 0 AndAlso htbABSPrecedentKeys.Contains(ctrl) Then
            'remove the control from the hashtable if its value is deleted
            htbABSPrecedentKeys.Remove(ctrl)
            RemoveHandler CType(ctrl, Control).Enter, _
                        AddressOf ShowABSPrecedentKeys

        ElseIf value.Length > 0 Then
            If Not htbABSPrecedentKeys.Contains(ctrl) Then
                If TypeOf ctrl Is Control Then

                    AddHandler CType(ctrl, Control).Enter, _
                            AddressOf ShowABSPrecedentKeys
                ElseIf TypeOf ctrl Is MenuItem Then
                    AddHandler CType(ctrl, MenuItem).Select, _
                            AddressOf ShowABSPrecedentKeys
                End If
            End If
            htbABSPrecedentKeys.Item(ctrl) = value
        End If
    End Sub

    <DisplayName("ABSPrecedentKeys")> _
    <Category("ABS")> _
    <Description("Contains the Name of the Parent Column related to this Control.  When the value in the control keyed to the Parent Column Name is Changed, the value in this control will change to the corresponding value it the Column Name indicated in this Control.")> _
    <DefaultValue("")> _
    Public Function GetABSPrecedentKeys(ByVal ctrl As Component) As String
        'The getter method is much simpler. It checks the hashtable to return the value previously set: 

        If htbABSPrecedentKeys.Contains(ctrl) Then
            Return htbABSPrecedentKeys(ctrl).ToString
        Else
            Return ""
        End If
    End Function

    Private Sub ShowABSPrecedentKeys(ByVal sender As Object, ByVal e As System.EventArgs)
        RaiseEvent MyExtProvEvent(GetABSPrecedentKeys(CType(sender, Component)))
    End Sub

#End Region

#Region "ABSSecurityCodes"

    Public Sub SetABSSecurityCodes(ByVal ctrl As Component, ByVal value As String)
        'This method will be called each time you set a value through the properties window or when you directly call the SetDescription method. The control (ctrl) and the new description (value) are received as arguments. This method uses the hashtable (htbABSColumnName) to store the description. You can also see that we are adding a handler on specific event of the controls. This is needed to be able to raise our own event (more on this later). 
        'If value = "VEND_CODE" Then Stop
        'To be sure we don't have the Nothing value
        If value Is Nothing Then value = ""

        If value.Length = 0 AndAlso htbABSSecurityCodes.Contains(ctrl) Then
            'remove the control from the hashtable if its value is deleted
            htbABSSecurityCodes.Remove(ctrl)
            RemoveHandler CType(ctrl, Control).Enter, _
                        AddressOf ShowABSSecurityCodes

        ElseIf value.Length > 0 Then
            If Not htbABSSecurityCodes.Contains(ctrl) Then
                If TypeOf ctrl Is Control Then

                    AddHandler CType(ctrl, Control).Enter, _
                            AddressOf ShowABSSecurityCodes
                ElseIf TypeOf ctrl Is MenuItem Then
                    AddHandler CType(ctrl, MenuItem).Select, _
                            AddressOf ShowABSSecurityCodes
                End If
            End If
            htbABSSecurityCodes.Item(ctrl) = value
        End If
    End Sub

    <DisplayName("ABSSecurityCodes")> _
    <Category("ABS")> _
    <Description("Contains the Name of the Parent Column related to this Control.  When the value in the control keyed to the Parent Column Name is Changed, the value in this control will change to the corresponding value it the Column Name indicated in this Control.")> _
    <DefaultValue("")> _
    Public Function GetABSSecurityCodes(ByVal ctrl As Component) As String
        'The getter method is much simpler. It checks the hashtable to return the value previously set: 

        If htbABSSecurityCodes.Contains(ctrl) Then
            Return htbABSSecurityCodes(ctrl).ToString
        Else
            Return ""
        End If
    End Function

    Private Sub ShowABSSecurityCodes(ByVal sender As Object, ByVal e As System.EventArgs)
        RaiseEvent MyExtProvEvent(GetABSSecurityCodes(CType(sender, Component)))
    End Sub

#End Region

#Region "ABSValueRequired"
    Public Sub SetABSValueRequired(ByVal ctrl As Component, ByVal value As Boolean)
        'This method will be called each time you set a value through the properties window or when you directly call the SetDescription method. The control (ctrl) and the new description (value) are received as arguments. This method uses the hashtable (htbABSColumnName) to store the description. You can also see that we are adding a handler on specific event of the controls. This is needed to be able to raise our own event (more on this later). 

        If TypeOf ctrl Is UltraWinEditors.UltraTextEditor Then
        Else
            value = False
        End If

        If Not value Then
            If htbABSValueRequired.Contains(ctrl) Then
                'remove the control from the hashtable if its value is deleted
                htbABSValueRequired.Remove(ctrl)
                RemoveHandler CType(ctrl, Control).Enter, _
                            AddressOf ShowABSValueRequired

            End If
        Else
            If Not htbABSValueRequired.Contains(ctrl) Then
                If TypeOf ctrl Is Control Then

                    AddHandler CType(ctrl, Control).Enter, _
                            AddressOf ShowABSValueRequired
                ElseIf TypeOf ctrl Is MenuItem Then
                    AddHandler CType(ctrl, MenuItem).Select, _
                            AddressOf ShowABSValueRequired
                End If
            End If
            htbABSValueRequired.Item(ctrl) = value
        End If
    End Sub
    <DisplayName("ABSValueRequired")> _
    <Category("ABS")> _
    <Description("True will cause Validation to Fail when Clicking Update.  This property will be used only if there is a value specified in the ABSColumnName property.")> _
    <DefaultValue(False)> _
    Public Function GetABSValueRequired(ByVal ctrl As Component) As Boolean
        'The getter method is much simpler. It checks the hashtable to return the value previously set: 

        If TypeOf ctrl Is UltraWinEditors.UltraTextEditor Then
            If htbABSValueRequired.Contains(ctrl) Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If

    End Function

    Private Sub ShowABSValueRequired(ByVal sender As Object, ByVal e As System.EventArgs)
        RaiseEvent MyExtProvEvent(GetABSValueRequired(CType(sender, Component)))
    End Sub

#End Region

#Region "ABSViewName"

    Public Sub SetABSViewName(ByVal ctrl As Component, ByVal value As String)
        'This method will be called each time you set a value through the properties window or when you directly call the SetDescription method. The control (ctrl) and the new description (value) are received as arguments. This method uses the hashtable (htbABSColumnName) to store the description. You can also see that we are adding a handler on specific event of the controls. This is needed to be able to raise our own event (more on this later). 

        'To be sure we don't have the Nothing value
        If value Is Nothing Then value = ""

        If value.Length = 0 AndAlso htbABSViewName.Contains(ctrl) Then
            'remove the control from the hashtable if its value is deleted
            htbABSViewName.Remove(ctrl)
            RemoveHandler CType(ctrl, Control).Enter, _
                        AddressOf ShowABSViewName

        ElseIf value.Length > 0 Then
            If Not htbABSViewName.Contains(ctrl) Then
                If TypeOf ctrl Is Control Then

                    AddHandler CType(ctrl, Control).Enter, _
                            AddressOf ShowABSViewName
                ElseIf TypeOf ctrl Is MenuItem Then
                    AddHandler CType(ctrl, MenuItem).Select, _
                            AddressOf ShowABSViewName
                End If
            End If
            htbABSViewName.Item(ctrl) = value
        End If
    End Sub

    <DisplayName("ABSViewName")> _
    <Category("ABS")> _
    <Description("Contains the Name of the View to be used to bring up a list of Codes.  If left blank, the Column Name will be used as the View Name.")> _
    <DefaultValue("")> _
    Public Function GetABSViewName(ByVal ctrl As Component) As String
        'The getter method is much simpler. It checks the hashtable to return the value previously set: 

        If htbABSViewName.Contains(ctrl) Then
            Return htbABSViewName(ctrl).ToString
        Else
            Return ""
        End If
    End Function

    Private Sub ShowABSViewName(ByVal sender As Object, ByVal e As System.EventArgs)
        RaiseEvent MyExtProvEvent(GetABSViewName(CType(sender, Component)))
    End Sub

#End Region

#Region "ABSBindToTable"
    Public Sub SetABSBindToTable(ByVal ctrl As Component, ByVal value As Boolean)
        'This method will be called each time you set a value through the properties window or when you directly call the SetDescription method. The control (ctrl) and the new description (value) are received as arguments. This method uses the hashtable (htbABSColumnName) to store the description. You can also see that we are adding a handler on specific event of the controls. This is needed to be able to raise our own event (more on this later). 

        If value Then
            If htbABSBindToTable.Contains(ctrl) Then
                'remove the control from the hashtable if its value is deleted
                htbABSBindToTable.Remove(ctrl)
                RemoveHandler CType(ctrl, Control).Enter, _
                            AddressOf ShowABSBindToTable
            End If
        Else
            If Not htbABSBindToTable.Contains(ctrl) Then
                If TypeOf ctrl Is Control Then
                    AddHandler CType(ctrl, Control).Enter, _
                            AddressOf ShowABSBindToTable
                ElseIf TypeOf ctrl Is MenuItem Then
                    AddHandler CType(ctrl, MenuItem).Select, _
                            AddressOf ShowABSBindToTable
                End If
            End If
            htbABSBindToTable.Item(ctrl) = value
        End If
    End Sub
    <DisplayName("ABSBindToTable")> _
    <Category("ABS")> _
    <Description("True will enable the Form or Control for Binding.  False for a Control which may be referenced using the ABSColumnName without any DataBinding.")> _
    <DefaultValue(True)> _
    Public Function GetABSBindToTable(ByVal ctrl As Component) As Boolean

        ' NOTE THAT THE HASH TABLE CONTAINS THE CONTROL WHEN THE VALUE IS FALSE
        If htbABSBindToTable.Contains(ctrl) Then
            Return False
        Else
            Return True
        End If

    End Function

    Private Sub ShowABSBindToTable(ByVal sender As Object, ByVal e As System.EventArgs)
        RaiseEvent MyExtProvEvent(GetABSBindToTable(CType(sender, Component)))
    End Sub

#End Region

#End Region

    Public Sub Load_COLUMN_NAMEs()
        For Each CTLx As System.Collections.DictionaryEntry In htbABSColumnName
            Dim ctl As Control = DirectCast(CTLx.Key, Control)
            Dim ctlTABLE_NAME As String = GetABSTableName(ctl)
            Dim ABSBindToTable As Boolean = Me.GetABSBindToTable(ctl)
            If Not ABSBindToTable Then
                ctlTABLE_NAME = ""
            Else
                If ctlTABLE_NAME = "" Then
                    ctlTABLE_NAME = TABLE_NAME_base
                End If
            End If
            Dim TC As String = CTLx.Value
            If ctlTABLE_NAME <> "" Then
                TC = ctlTABLE_NAME & "." & TC
            End If
            If Not dicCOLUMN_NAME.ContainsKey(TC) Then
                Try
                    dicCOLUMN_NAME.Add(TC, ctl)
                    'Debug.Print(TC)
                Catch ex As Exception
                    'Stop
                End Try
            End If
        Next
    End Sub
    Public Sub Set_Value_For(COLUMN_NAME As String, value As Object)
        Dim ctlType As Type = Me.CtlFor(COLUMN_NAME).GetType
        If ctlType Is GetType(UltraWinEditors.UltraTextEditor) Then
            Me.txtFor(COLUMN_NAME).Value = value
        ElseIf ctlType Is GetType(UltraWinEditors.UltraOptionSet) Then
            Me.optFor(COLUMN_NAME).Value = value
        ElseIf ctlType Is GetType(UltraWinEditors.UltraDateTimeEditor) Then
            Me.dteFor(COLUMN_NAME).Value = value
        ElseIf ctlType Is GetType(UltraWinEditors.UltraNumericEditor) Then
            Me.numFor(COLUMN_NAME).Value = value
        ElseIf ctlType Is GetType(UltraWinEditors.UltraComboEditor) Then
            Me.cbeFor(COLUMN_NAME).Value = value
        ElseIf ctlType Is GetType(UltraWinGrid.UltraCombo) Then
            Me.cmbFor(COLUMN_NAME).Value = value
        ElseIf ctlType Is GetType(UltraWinEditors.UltraCheckEditor) OrElse ctlType.Name = "ABSCheckBox" Then
            Me.chkFor(COLUMN_NAME).Checked = value & "" = "1"
        ElseIf ctlType Is GetType(UltraWinMaskedEdit.UltraMaskedEdit) Then
            Me.medFor(COLUMN_NAME).Value = value
        End If
    End Sub
End Class
