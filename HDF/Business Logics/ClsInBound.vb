Imports System.Data.SqlClient
Imports System.Configuration
Imports System.IO
Imports System.Xml
Imports System.Text
Imports SAPbobsCOM

Public Class ClsInBound
    Inherits clsBase

    Private myConnection As New SqlConnection
    Private oCommand As SqlCommand
    Private oSqlAdap As SqlDataAdapter
    Dim oDBDataSource As SAPbouiCOM.DBDataSource
    Dim oDBDataSource_1 As SAPbouiCOM.DBDataSource
    Private oOrderGrid As SAPbouiCOM.Grid
    Private oSuccessGrid As SAPbouiCOM.Grid
    Private oFailureGrid As SAPbouiCOM.Grid
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oComoColumn As SAPbouiCOM.ComboBoxColumn
    Private oDtOrder As SAPbouiCOM.DataTable
    Private oDTSuccess As SAPbouiCOM.DataTable
    Private oDTFailure As SAPbouiCOM.DataTable
    Private oCombo As SAPbouiCOM.ComboBox
    Private strqry As String
    Private ds As DataSet
    Private oDetailDt As DataTable

    Dim strQuery As String
    Private oRecordSet As SAPbobsCOM.Recordset
    Private oRS As SAPbobsCOM.Recordset
    Dim oUserTable As SAPbobsCOM.UserTable

    Private oHeader As DataTable
    Private oRowDetails As DataTable
    Private oBatchDetails As DataTable
    Private intStatus As String

    Public Sub LoadForm()
        Try
            oForm = oApplication.Utilities.LoadForm(xml_Z_InBound, frm_Z_InBound)
            oForm = oApplication.SBO_Application.Forms.ActiveForm
            oForm.Freeze(True)
            oForm.PaneLevel = 1
            initialize(oForm)
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub initialize(ByRef oForm As SAPbouiCOM.Form)
        Try
            oForm.Items.Item("4").TextStyle = SAPbouiCOM.BoTextStyle.ts_BOLD
            oForm.Items.Item("8").TextStyle = 5
            oForm.Items.Item("_10").TextStyle = 5
            oForm.Items.Item("_19").TextStyle = 5
            oForm.DataSources.DataTables.Add("dtOrder")
            oForm.DataSources.DataTables.Add("dtSuccess")
            oForm.DataSources.DataTables.Add("dtFailure")
            loadCombo(oForm)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_Z_InBound
                    LoadForm()
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Dim oDataTable As SAPbouiCOM.DataTable
            If pVal.FormTypeEx = frm_Z_InBound Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If (pVal.ItemUID = "7" Or pVal.ItemUID = "3") And oForm.PaneLevel > 1 Then
                                    If validation(oForm) Then
                                        If pVal.ItemUID = "7" Then
                                            If oApplication.SBO_Application.MessageBox("Do you want to Proceed?", , "Yes", "No") = 2 Then
                                                BubbleEvent = False
                                                Exit Sub
                                            Else
                                                If oForm.PaneLevel = 3 Then
                                                    If ExportGrid(oForm) Then
                                                        oSuccessGrid = oForm.Items.Item("9").Specific
                                                        oSuccessGrid.DataTable = oForm.DataSources.DataTables.Item("dtSuccess")
                                                        oSuccessGrid = oForm.Items.Item("10").Specific
                                                        oSuccessGrid.DataTable = oForm.DataSources.DataTables.Item("dtFailure")
                                                        oForm.PaneLevel = 4
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "3" And (oForm.PaneLevel = 1) Then
                                    oForm.PaneLevel = oForm.PaneLevel + 1
                                ElseIf pVal.ItemUID = "6" And (oForm.PaneLevel > 1) Then
                                    oForm.PaneLevel = oForm.PaneLevel - 1
                                ElseIf pVal.ItemUID = "3" And (oForm.PaneLevel = 2) Then
                                    LoadGrid(oForm)
                                    oForm.PaneLevel = oForm.PaneLevel + 1
                                ElseIf pVal.ItemUID = "11" And (oForm.PaneLevel = 4) Then
                                    oForm.Close()
                                    Exit Sub

                                ElseIf pVal.ItemUID = "17" Then
                                    selectAll(oForm)
                                ElseIf pVal.ItemUID = "18" Then
                                    clearAll(oForm)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS

                            Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                                'If oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized Or oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Restore Then
                                '    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                '    Try
                                '        reDrawForm(oForm)
                                '    Catch ex As Exception

                                '    End Try
                                'End If
                        End Select
                End Select
            End If
        Catch ex As Exception
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm = oApplication.SBO_Application.Forms.Item(FormUID)
        End Try
    End Sub
#End Region

    Private Sub LoadGrid(ByVal aform As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)

            Dim intID_O As Integer = 0
            Dim strType As String = String.Empty
            Dim EXPath As String
            Dim dt As New DataTable

            Try
                strType = CType(oForm.Items.Item("13").Specific, SAPbouiCOM.ComboBox).Selected.Value
            Catch ex As Exception
            End Try


            oOrderGrid = oForm.Items.Item("5").Specific
            oOrderGrid.DataTable = oForm.DataSources.DataTables.Item("dtOrder")

            oDtOrder = oForm.DataSources.DataTables.Item("dtOrder")

            Try
                oDtOrder.Columns.Add("Select", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 1)
                oDtOrder.Columns.Add("File Name", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
            Catch ex As Exception
            End Try
            

            If strType = "Item-Add" Or strType = "Item-Update" Or strType = "Item-Delete" Or strType = "Consignment" Or strType = "JE" Then
                EXPath = GetOpenPath(strType)
                Dim DP As New IO.DirectoryInfo(EXPath)
                For Each file As IO.FileInfo In New IO.DirectoryInfo(DP.ToString).GetFiles("*.xml")

                    oDtOrder.Rows.Add(1)
                    oDtOrder.SetValue("Select", intID_O, "Y")
                    oDtOrder.SetValue("File Name", intID_O, file.Name)
                    intID_O += 1

                Next

            End If

            gridFormat(oForm, strType)
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
        End Try

    End Sub

    Private Sub gridFormat(ByVal oform As SAPbouiCOM.Form, ByVal strType As String)
        Try
            If strType = "Item-Add" Or strType = "Item-Update" Or strType = "Item-Delete" Or strType = "Consignment" Or strType = "JE" Then

                oOrderGrid = oform.Items.Item("5").Specific

                oOrderGrid.Columns.Item("Select").TitleObject.Caption = "Select"
                oOrderGrid.Columns.Item("Select").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                oOrderGrid.Columns.Item("Select").Editable = True
                oOrderGrid.Columns.Item("Select").Width = 200

                oOrderGrid.Columns.Item("File Name").TitleObject.Caption = "File Name"
                oOrderGrid.Columns.Item("File Name").Editable = False
                oOrderGrid.Columns.Item("File Name").Width = 700

            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim _retVal As Boolean = True
        Try
            If oForm.PaneLevel = 2 Then

                Dim strStype As String = oForm.Items.Item("13").Specific.value
                If strStype.Trim() = "" Then
                    oApplication.Utilities.Message("Select Type...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return _retVal = False
                End If
            End If


            If oForm.PaneLevel = 3 Then
                oOrderGrid = oForm.Items.Item("5").Specific
                For intRow As Integer = 0 To oOrderGrid.Rows.Count - 1
                    If oOrderGrid.DataTable.GetValue("Select", intRow).ToString() = "Y" Then
                        _retVal = True
                        Exit For
                    End If
                Next

                If _retVal = False Then
                    oApplication.Utilities.Message("No Records Selected", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If

            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function ExportGrid(ByVal oForm As SAPbouiCOM.Form) As Boolean
        Try

            Dim _retVal As Boolean = False
            Dim intDocEntry As Integer
            Dim strDocNum As String = String.Empty
            Dim strCardName As String = String.Empty
            Dim strDocDate As String = String.Empty
            Dim strDFNP As String

            Dim strType As String = String.Empty


            oDTSuccess = oForm.DataSources.DataTables.Item("dtSuccess")
            oDTFailure = oForm.DataSources.DataTables.Item("dtFailure")

            strType = CType(oForm.Items.Item("13").Specific, SAPbouiCOM.ComboBox).Selected.Value

            Try
                If strType <> "Consignment" And strType <> "JE" Then
                    oDTSuccess.Columns.Add("ItemName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                    oDTSuccess.Columns.Add("ItemDescription", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 200)
                    oDTFailure.Columns.Add("ItemName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                    oDTFailure.Columns.Add("ItemDescription", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 200)
                    oDTFailure.Columns.Add("FailedReason", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 500)
                Else
                    oDTSuccess.Columns.Add("File Name", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                    oDTSuccess.Columns.Add("Status", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)

                    oDTFailure.Columns.Add("File Name", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                    oDTFailure.Columns.Add("FailedReason", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 500)
                End If

            Catch ex As Exception

            End Try


            If strType = "Item-Add" Or strType = "Item-Update" Or strType = "Item-Delete" Or strType = "Consignment" Or strType = "JE" Then
                XMLTODS(oForm, strType)
            End If

            _retVal = True
            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Sub loadCombo(ByVal oForm As SAPbouiCOM.Form)
        Try
            oCombo = oForm.Items.Item("13").Specific
            With oCombo
                .ValidValues.Add("Item-Add", "Item-Add")
                .ValidValues.Add("Item-Update", "Item-Update")
                .ValidValues.Add("Item-Delete", "Item-Delete")
                .ValidValues.Add("Consignment", "Consignment Issue")
                .ValidValues.Add("JE", "Journal Entry")
            End With
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub selectAll(ByVal oForm As SAPbouiCOM.Form)
        Try
            oOrderGrid = oForm.Items.Item("5").Specific
            oForm.Freeze(True)
            For index As Integer = 0 To oOrderGrid.DataTable.Rows.Count - 1
                oOrderGrid.DataTable.SetValue("Select", index, "Y")
            Next
            oForm.Freeze(False)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub clearAll(ByVal oForm As SAPbouiCOM.Form)
        Try
            oOrderGrid = oForm.Items.Item("5").Specific
            oForm.Freeze(True)
            For index As Integer = 0 To oOrderGrid.DataTable.Rows.Count - 1
                oOrderGrid.DataTable.SetValue("Select", index, "N")
            Next
            oForm.Freeze(False)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Function GetOpenPath(ByVal Type As String) As String
        Dim _retVal As String
        Try
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select T1.U_OPath From [@Z_HDF_IBND] T1 Where T1.U_Type='" & Type & "'")
            If Not oRecordSet.EoF Then
                _retVal = oRecordSet.Fields.Item(0).Value
            Else
                Throw New Exception("")
            End If
        Catch generatedExceptionName As Exception
            Throw
        End Try
        Return _retVal
    End Function

    Public Function GetSuccessPath(ByVal Type As String) As String
        Dim _retVal As String
        Try
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select T1.U_SPath From [@Z_HDF_IBND] T1 Where T1.U_Type='" & Type & "'")
            If Not oRecordSet.EoF Then
                _retVal = oRecordSet.Fields.Item(0).Value
            Else
                Throw New Exception("")
            End If
        Catch generatedExceptionName As Exception
            Throw
        End Try
        Return _retVal
    End Function

    Private Sub XMLTODS(ByVal oForm As SAPbouiCOM.Form, ByVal strType As String)
        Dim FName As String
        Dim intID_S As Integer = 0
        Dim intID_F As Integer = 0
        Dim intStatus As Integer
        Dim OPath As String
        Dim ds As New DataSet
        Dim oItem As SAPbobsCOM.Items
        Dim Consign As String
        oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)

        Try
            If strType = "Item-Add" Or strType = "Item-Update" Or strType = "Item-Delete" Or strType = "Consignment" Or strType = "JE" Then
                OPath = GetOpenPath(strType)

                For intRow As Integer = 0 To oOrderGrid.Rows.Count - 1
                    If oOrderGrid.DataTable.GetValue("Select", intRow).ToString() = "Y" Then

                        Try
                            ds.Clear()

                            If OPath.Trim() = "" Then
                                Throw New Exception("Path Not Defined...")
                            End If

                            FName = oOrderGrid.DataTable.GetValue("File Name", intRow).ToString()
                            Dim strDPath As String = OPath & "\" + FName

                            'Dim fileReader As String = My.Computer.FileSystem.ReadAllText(strDPath).Replace("utf-8", "utf-7")
                            'My.Computer.FileSystem.WriteAllText(strDPath, fileReader, True, System.Text.Encoding.UTF7)

                            ds.ReadXml(strDPath)

                            If strType = "Item-Add" Then

                                oItem.ItemCode = ds.Tables("Article").Rows(0)("CodeArticle")
                                oItem.ItemName = ds.Tables("Article").Rows(0)("DescriptionArticle")

                                If ds.Tables("Article").Rows(0)("GererParLot") = "Y" Then
                                    oItem.ManageBatchNumbers = SAPbobsCOM.BoYesNoEnum.tYES
                                Else
                                    oItem.ManageBatchNumbers = SAPbobsCOM.BoYesNoEnum.tNO
                                End If
                                oItem.SalesUnit = ds.Tables("Article").Rows(0)("unitesDinventaireDeMesure")

                                intStatus = oItem.Add()

                            ElseIf strType = "Item-Update" Or strType = "Item-Delete" Then
                                Dim strItemCode As String = ds.Tables("Article").Rows(0)("CodeArticle")
                                If oItem.GetByKey(strItemCode) Then

                                    oItem.ItemName = ds.Tables("Article").Rows(0)("DescriptionArticle")

                                    If ds.Tables("Article").Rows(0)("GererParLot") = "Y" Then
                                        oItem.ManageBatchNumbers = SAPbobsCOM.BoYesNoEnum.tYES
                                    Else
                                        oItem.ManageBatchNumbers = SAPbobsCOM.BoYesNoEnum.tNO
                                    End If
                                    oItem.SalesUnit = ds.Tables("Article").Rows(0)("unitesDinventaireDeMesure")
                                    If ds.Tables("Article").Rows(0)("Active") = "Y" Then
                                        oItem.Valid = SAPbobsCOM.BoYesNoEnum.tYES
                                        oItem.Frozen = SAPbobsCOM.BoYesNoEnum.tNO
                                    Else
                                        oItem.Valid = SAPbobsCOM.BoYesNoEnum.tNO
                                        oItem.Frozen = SAPbobsCOM.BoYesNoEnum.tYES
                                    End If

                                    intStatus = oItem.Update()
                                End If

                            ElseIf strType = "Consignment" Then

                                'For Consignment Header
                                Dim DocRef As String = ds.Tables("Bon_livraison").Rows(0)("Id_bon_Commande")
                                Dim DocDate As String = ds.Tables("Elément_livraison").Rows(0)("Dh_livraison")
                                Dim Remarks As String = ds.Tables("Elément_livraison").Rows(0)("Commentaire")
                                Dim ItemCode1 As String = ds.Tables("Elément_livraison").Rows(0)("Id_élément_livraison")
                                strQuery = "Insert into ConsignmentHeader(DocRef, DocDate, Remarks)"
                                strQuery += "Values('" + DocRef + "','" + DocDate + "','" + Remarks + "')"
                                oRecordSet.DoQuery(strQuery)

                                strQuery = "Select (Max(DocEntry)) AS DocEntry From ConsignmentHeader"
                                oRecordSet.DoQuery(strQuery)
                                Dim DocEntry As Integer = oRecordSet.Fields.Item("DocEntry").Value

                                'For Consignment Detail
                                For iRow As Integer = 0 To ds.Tables("Composant_livré").Rows.Count - 1
                                    Dim ItemCode As String = ItemCode1 'ds.Tables("Composant_livré").Rows(iRow)("Code_Fournisseur")
                                    Dim Qty As Double = ds.Tables("Quantité_composant").Rows(iRow)("Nombre")
                                    Dim UOM As String = ds.Tables("Quantité_composant").Rows(iRow)("Unité")


                                    strQuery = "Select ItemCode From OITM Where ItemCode='" & ItemCode & "'"
                                    oRecordSet.DoQuery(strQuery)

                                    If oRecordSet.EoF Then
                                        Throw New Exception("Item in XML File Not Found...Product : " & ItemCode)
                                    End If

                                    strQuery = "Select U_IsConsign,CardCode,DfltWH From OITM Where ItemCode='" & ItemCode & "'"
                                    oRecordSet.DoQuery(strQuery)

                                    If oRecordSet.RecordCount > 0 Then
                                        Consign = oRecordSet.Fields.Item("U_IsConsign").Value
                                        Dim strVendor As String = oRecordSet.Fields.Item("CardCode").Value
                                        Dim strWareHouse As String = oRecordSet.Fields.Item("DfltWH").Value

                                        If strWareHouse.Trim().Length = 0 Then
                                            Throw New Exception("Default WareHouse Not Found...Product : " & ItemCode)
                                        End If

                                        If Consign = "Y" Then

                                            'Check for Open Purchase Order
                                            strqry = "Select ISNULL(OnHand,0) AS OpenQty From OITW Where ItemCode = '" & ItemCode & "' And WhsCode = '" & strWareHouse & "'"
                                            oRecordSet.DoQuery(strqry)
                                            Dim OpenQty As Double = oRecordSet.Fields.Item("OpenQty").Value
                                            If oRecordSet.RecordCount > 0 Then
                                                If OpenQty >= Qty Then
                                                    strQuery = "Insert into ConsignmentDetail(DocEntry, DocRef, ItemCode, ItemName, Qty, UOM, IsCon, IsPur, IsGRPO, IsGI, IsGI_S)"
                                                    strQuery += "Values(" + DocEntry.ToString() + ",'" + DocRef + "','" + ItemCode + "','" + ItemCode + "'," + Qty.ToString() + ",'" + UOM + "'," + "'Y'" + "," + "'Y'" + "," + "'Y'" + "," + "'N'" + "," + "'Y'" + ")"
                                                    oRecordSet.DoQuery(strQuery)
                                                Else
                                                    strQuery = "Insert into ConsignmentDetail(DocEntry, DocRef, ItemCode, ItemName, Qty, UOM, IsCon, IsPur, IsGRPO, IsGI)"
                                                    strQuery += "Values(" + DocEntry.ToString() + ",'" + DocRef + "','" + ItemCode + "','" + ItemCode + "'," + Qty.ToString() + ",'" + UOM + "'," + "'Y'" + "," + "'Y'" + "," + "'Y'" + "," + "'Y'" + ")"
                                                    oRecordSet.DoQuery(strQuery)
                                                End If
                                            End If

                                        Else

                                            strQuery = "Insert into ConsignmentDetail(DocEntry, DocRef, ItemCode, ItemName, Qty, UOM, IsCon, IsPur, IsGRPO, IsGI)"
                                            strQuery += "Values(" + DocEntry.ToString() + ",'" + DocRef + "','" + ItemCode + "','" + ItemCode + "'," + Qty.ToString() + ",'" + UOM + "'," + "'N'" + "," + "'N'" + "," + "'N'" + "," + "'Y'" + ")"
                                            oRecordSet.DoQuery(strQuery)

                                        End If
                                    End If
                                Next


                                'For Consignment Batch
                                strQuery = "Select (Max(DocEntry)) AS DocEntry From ConsignmentHeader"
                                oRecordSet.DoQuery(strQuery)
                                Dim DEntry As Integer = oRecordSet.Fields.Item("DocEntry").Value


                                For iRow As Integer = 0 To ds.Tables("Composant_livré").Rows.Count - 1
                                    For JRow As Integer = 0 To ds.Tables("Quantité_composant").Rows.Count - 1
                                        If ds.Tables("Composant_livré").Rows(iRow)("Composant_livré_Id").ToString() = ds.Tables("Quantité_composant").Rows(JRow)("Composant_livré_Id").ToString() Then
                                            Dim ItemCode As String = ItemCode1 'ds.Tables("Composant_livré").Rows(iRow)("Code_Fournisseur")
                                            Dim Batch As String = ds.Tables("Composant_livré").Rows(iRow)("Lot")
                                            Dim Qty As Double = ds.Tables("Quantité_composant").Rows(JRow)("Nombre")
                                            strQuery = "Insert INTO ConsignmentBatch(DocEntry, ItemCode, Batch, Qty)"
                                            strQuery += " Values(" + DEntry.ToString() + ",'" + ItemCode + "','" + Batch + "'," + Qty.ToString() + ")"
                                            oRecordSet.DoQuery(strQuery)
                                        End If
                                    Next
                                Next


                                oHeader = New DataTable()
                                oHeader.Columns.Add("ItemCode", GetType(String))
                                oHeader.Columns.Add("CardCode", GetType(String))
                                oHeader.Columns.Add("CardName", GetType(String))
                                oHeader.Columns.Add("DocDate", GetType(String))
                                oHeader.Columns.Add("DocRef", GetType(String))
                                oHeader.Columns.Add("WareHouse", GetType(String))

                                'Header Construction
                                strqry = "Select Distinct T1.CardCode,T1.DfltWH from ConsignmentDetail T0 JOIN OITM T1 ON T0.ItemCode = T1.ItemCode"
                                strqry += " And T1.CardCode Is Not Null And T0.DocEntry = " & DEntry & ""
                                Dim OH As DataSet = ExecuteDataSet(strqry)
                                If Not IsNothing(OH) Then
                                    If OH.Tables(0).Rows.Count > 0 Then
                                        Dim oDr As DataRow
                                        For Each dr As DataRow In OH.Tables(0).Rows
                                            oDr = oHeader.NewRow()
                                            oDr("ItemCode") = ItemCode1
                                            oDr("CardCode") = dr("CardCode")
                                            oDr("CardName") = dr("CardCode")
                                            oDr("DocDate") = DocDate
                                            oDr("DocRef") = DocRef
                                            oDr("WareHouse") = dr("DfltWH")
                                            oHeader.Rows.Add(oDr)
                                        Next
                                    Else
                                        Throw New Exception("Preferred Vendor Not Found...")
                                    End If
                                End If
                                oHeader.AcceptChanges()

                                oRowDetails = New DataTable()
                                oRowDetails.Columns.Add("ItemCode", GetType(String))
                                oRowDetails.Columns.Add("ItemName", GetType(String))
                                oRowDetails.Columns.Add("Quantity", GetType(Double))
                                oRowDetails.Columns.Add("UnitPrice", GetType(Double))
                                oRowDetails.Columns.Add("UOM", GetType(String))
                                oRowDetails.Columns.Add("BaseEntry", GetType(String))
                                oRowDetails.Columns.Add("BaseLine", GetType(String))

                                oBatchDetails = New DataTable()
                                oBatchDetails.Columns.Add("ItemCode", GetType(String))
                                oBatchDetails.Columns.Add("LineNo", GetType(String))
                                oBatchDetails.Columns.Add("Quantity", GetType(Double))
                                oBatchDetails.Columns.Add("BatchNo", GetType(String))

                                'If Consignment is No then It will Create Goods Issue Alone.
                                strQuery = "Select ItemCode, ItemName, ISNULL(Qty,0) As Qty, UOM, IsCon, IsPur, IsGRPO, IsGI From ConsignmentDetail Where DocEntry = " & DEntry & " and IsCon = 'N' "
                                Dim CD As DataSet = ExecuteDataSet(strQuery)
                                If Not IsNothing(CD) Then
                                    oDetailDt = CD.Tables(0)
                                    If oDetailDt.Rows.Count > 0 Then
                                        If CreatePurchaseOrder(oHeader, ds.Tables("Composant_livré"), ds.Tables("Quantité_composant"), oRowDetails, oBatchDetails, "BT") Then ' 
                                            If CreateGI(oHeader, oRowDetails, oBatchDetails) Then

                                            Else
                                                intStatus = -1
                                            End If
                                        End If
                                    End If
                                End If

                                'If Consignment Is Yes and Stock Exist for the Item 
                                strQuery = "Select ItemCode, ItemName, ISNULL(Qty,0) As Qty, UOM, IsCon, IsPur, IsGRPO, IsGI From ConsignmentDetail Where DocEntry = " & DEntry & " and IsCon ='Y' and ISNULL(IsGI_S,'N') = 'Y' "
                                Dim CD1 As DataSet = ExecuteDataSet(strQuery)
                                If Not IsNothing(CD1) Then
                                    oDetailDt = CD1.Tables(0)
                                    If oDetailDt.Rows.Count > 0 Then
                                        oApplication.Company.StartTransaction()
                                        If CreatePurchaseOrder(oHeader, ds.Tables("Composant_livré"), ds.Tables("Quantité_composant"), oRowDetails, oBatchDetails, "BT") Then ' 
                                            If CreateGI_Clear(oHeader, oRowDetails, CDbl(oDetailDt.Rows(0)("Qty").ToString)) Then
                                                If CreatePurchaseOrder(oHeader, ds.Tables("Composant_livré"), ds.Tables("Quantité_composant"), oRowDetails, oBatchDetails, "CPO") Then ' 
                                                    If CreateGRPO(oHeader, oRowDetails, oBatchDetails, "CPO") Then ' Stand alone GRPO
                                                        If 1 = 1 Then 'If CreateGI(oHeader, oRowDetails, oBatchDetails) Then
                                                            If oApplication.Company.InTransaction Then
                                                                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                                            End If
                                                            intStatus = 0
                                                        Else
                                                            If oApplication.Company.InTransaction Then
                                                                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                            End If
                                                            intStatus = -1 'oApplication.Company.GetLastErrorCode()
                                                        End If
                                                    Else
                                                        If oApplication.Company.InTransaction Then
                                                            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                        End If
                                                        intStatus = -1 'oApplication.Company.GetLastErrorCode()
                                                    End If
                                                Else
                                                    If oApplication.Company.InTransaction Then
                                                        oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                    End If
                                                    intStatus = -1 'oApplication.Company.GetLastErrorCode()
                                                End If
                                            Else
                                                If oApplication.Company.InTransaction Then
                                                    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                End If
                                                intStatus = -1 'oApplication.Company.GetLastErrorCode()
                                            End If

                                        End If
                                        
                                    End If
                                End If


                                'If Consignment Is Yes and Stock not Exist for the Item 
                                strQuery = "Select ItemCode, ItemName, ISNULL(Qty,0) As Qty, UOM, IsCon, IsPur, IsGRPO, IsGI From ConsignmentDetail Where DocEntry = " & DEntry & " and IsCon ='Y' and ISNULL(IsGI_S,'N') = 'N' "
                                Dim CD2 As DataSet = ExecuteDataSet(strQuery)
                                If Not IsNothing(CD2) Then
                                    oDetailDt = CD2.Tables(0)
                                    If oDetailDt.Rows.Count > 0 Then
                                        oApplication.Company.StartTransaction()
                                        If CreatePurchaseOrder(oHeader, ds.Tables("Composant_livré"), ds.Tables("Quantité_composant"), oRowDetails, oBatchDetails, "CPO") Then ' 
                                            If CreateGRPO(oHeader, oRowDetails, oBatchDetails, "CPO") Then ' Stand alone GRPO
                                                If CreateGI(oHeader, oRowDetails, oBatchDetails) Then
                                                    If oApplication.Company.InTransaction Then
                                                        oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                                    End If
                                                    intStatus = 0
                                                Else
                                                    If oApplication.Company.InTransaction Then
                                                        oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                    End If
                                                    intStatus = -1 'oApplication.Company.GetLastErrorCode()
                                                End If
                                            Else
                                                If oApplication.Company.InTransaction Then
                                                    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                End If
                                                intStatus = -1 'oApplication.Company.GetLastErrorCode()
                                            End If
                                        Else
                                            If oApplication.Company.InTransaction Then
                                                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                            End If
                                            intStatus = -1 'oApplication.Company.GetLastErrorCode()
                                        End If
                                    End If
                                End If

                                'strQuery = "Select ItemCode, ItemName, Qty, UOM, IsCon, IsPur, IsGRPO, IsGI From ConsignmentDetail Where DocEntry = " & DEntry & " and IsCon='Y' and IsPur='N' "
                                'Dim CD2 As DataSet = ExecuteDataSet(strQuery)
                                'If Not IsNothing(CD2) Then
                                '    oDetailDt = CD2.Tables(0)
                                '    If oDetailDt.Rows.Count > 0 Then
                                '        oApplication.Company.StartTransaction()

                                '        If CreatePurchaseOrder(oHeader, ds.Tables("Composant_livré"), ds.Tables("Quantité_composant"), oRowDetails, oBatchDetails, "OPO") Then ' Purchase order link to Row Details & BatchDetails
                                '            If CreateGRPO(oHeader, oRowDetails, oBatchDetails, "OPO") Then ' GRPO Based on Open PO
                                '                If CreateGI(oHeader, oRowDetails, oBatchDetails) Then
                                '                    If oApplication.Company.InTransaction Then
                                '                        oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                '                    End If
                                '                    intStatus = 0
                                '                Else
                                '                    If oApplication.Company.InTransaction Then
                                '                        oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                '                    End If
                                '                    intStatus = oApplication.Company.GetLastErrorCode()
                                '                End If
                                '            Else
                                '                If oApplication.Company.InTransaction Then
                                '                    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                '                End If
                                '                intStatus = oApplication.Company.GetLastErrorCode()
                                '            End If
                                '        Else
                                '            If oApplication.Company.InTransaction Then
                                '                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                '            End If
                                '            intStatus = oApplication.Company.GetLastErrorCode()
                                '        End If
                                '    End If
                                'End If
                            ElseIf strType = "JE" Then
                                If addJournal_Entry(ds.Tables(1), ds.Tables(3)) Then

                                Else
                                    intStatus = oApplication.Company.GetLastErrorCode()
                                End If
                            End If

                            If strType = "Item-Add" Or strType = "Item-Update" Or strType = "Item-Delete" Or strType = "Consignment" Or strType = "JE" Then
                                If intStatus = 0 Then
                                    'Updating Log
                                    oUserTable = oApplication.Company.UserTables.Item("Z_HDF_IBND_Log")

                                    strQuery = "Select count(*) As Code From [@Z_HDF_IBND_Log]"
                                    oRecordSet.DoQuery(strQuery)
                                    'Set default, mandatory fields
                                    If oRecordSet.RecordCount > 0 Then
                                        oUserTable.Code = (CInt(oRecordSet.Fields.Item("Code").Value) + 1).ToString()
                                        oUserTable.Name = (CInt(oRecordSet.Fields.Item("Code").Value) + 1).ToString()
                                    Else
                                        oUserTable.Code = "1"
                                        oUserTable.Name = "1"
                                    End If
                                    'Set user field
                                    oUserTable.UserFields.Fields.Item("U_Type").Value = strType
                                    oUserTable.UserFields.Fields.Item("U_Status").Value = "Y"
                                    oUserTable.UserFields.Fields.Item("U_FileName").Value = FName
                                    Dim now As DateTime = DateTime.Now
                                    oUserTable.UserFields.Fields.Item("U_ProDate").Value = now.ToString("d")
                                    oUserTable.UserFields.Fields.Item("U_ProTime").Value = now.ToString("t")
                                    oUserTable.Add()

                                    oDTSuccess.Rows.Add(1)
                                    If strType <> "Consignment" And strType <> "JE" Then
                                        oDTSuccess.SetValue("ItemName", intID_S, ds.Tables("Article").Rows(0)("CodeArticle"))
                                        oDTSuccess.SetValue("ItemDescription", intID_S, ds.Tables("Article").Rows(0)("DescriptionArticle"))
                                    Else
                                        oDTSuccess.SetValue("File Name", intID_S, FName.ToString)
                                        oDTSuccess.SetValue("Status", intID_S, "Successfully Exported")
                                    End If
                                    intID_S += 1

                                    MoveToFolder(strType, FName)

                                Else

                                    oUserTable = oApplication.Company.UserTables.Item("Z_HDF_IBND_LOG")

                                    strQuery = "Select count(*) As Code From [@Z_HDF_IBND_Log]"
                                    oRecordSet.DoQuery(strQuery)
                                    'Set default, mandatory fields
                                    If oRecordSet.RecordCount > 0 Then
                                        oUserTable.Code = (CInt(oRecordSet.Fields.Item("Code").Value) + 1).ToString()
                                        oUserTable.Name = (CInt(oRecordSet.Fields.Item("Code").Value) + 1).ToString()
                                    Else
                                        oUserTable.Code = "1"
                                        oUserTable.Name = "1"
                                    End If
                                    'Set user field
                                    oUserTable.UserFields.Fields.Item("U_Type").Value = strType
                                    oUserTable.UserFields.Fields.Item("U_Status").Value = "N"
                                    oUserTable.UserFields.Fields.Item("U_FileName").Value = FName
                                    Dim now As DateTime = DateTime.Now
                                    oUserTable.UserFields.Fields.Item("U_ProDate").Value = now.ToString("d")
                                    oUserTable.UserFields.Fields.Item("U_ProTime").Value = now.ToString("t")
                                    oUserTable.UserFields.Fields.Item("U_SAPEC").Value = oApplication.Company.GetLastErrorCode().ToString()
                                    oUserTable.UserFields.Fields.Item("U_SAPEM").Value = oApplication.Company.GetLastErrorDescription().ToString()
                                    oUserTable.Add()

                                    oDTFailure.Rows.Add(1)
                                    If strType <> "Consignment" And strType <> "JE" Then
                                        oDTFailure.SetValue("ItemName", intID_F, ds.Tables("Article").Rows(0)("CodeArticle"))
                                        oDTFailure.SetValue("ItemDescription", intID_F, ds.Tables("Article").Rows(0)("DescriptionArticle"))
                                    Else
                                        oDTFailure.SetValue("File Name", intID_F, FName.ToString)
                                    End If

                                    oDTFailure.SetValue("FailedReason", intID_F, oApplication.Company.GetLastErrorDescription().ToString())
                                    intID_F += 1

                                End If
                            End If
                        Catch ex As Exception

                            If oApplication.Company.InTransaction Then
                                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            End If

                            oUserTable = oApplication.Company.UserTables.Item("Z_HDF_IBND_LOG")

                            strQuery = "Select count(*) As Code From [@Z_HDF_IBND_Log]"
                            oRecordSet.DoQuery(strQuery)
                            'Set default, mandatory fields
                            If oRecordSet.RecordCount > 0 Then
                                oUserTable.Code = (CInt(oRecordSet.Fields.Item("Code").Value) + 1).ToString()
                                oUserTable.Name = (CInt(oRecordSet.Fields.Item("Code").Value) + 1).ToString()
                            Else
                                oUserTable.Code = "1"
                                oUserTable.Name = "1"
                            End If
                            'Set user field
                            oUserTable.UserFields.Fields.Item("U_Type").Value = strType
                            oUserTable.UserFields.Fields.Item("U_Status").Value = "N"
                            oUserTable.UserFields.Fields.Item("U_FileName").Value = FName
                            Dim now As DateTime = DateTime.Now
                            oUserTable.UserFields.Fields.Item("U_ProDate").Value = now.ToString("d")
                            oUserTable.UserFields.Fields.Item("U_ProTime").Value = now.ToString("t")
                            oUserTable.UserFields.Fields.Item("U_SAPEC").Value = oApplication.Company.GetLastErrorCode().ToString()
                            oUserTable.UserFields.Fields.Item("U_SAPEM").Value = oApplication.Company.GetLastErrorDescription().ToString()
                            oUserTable.Add()

                            oDTFailure.Rows.Add(1)
                            If strType <> "Consignment" And strType <> "JE" Then
                                oDTFailure.SetValue("ItemName", intID_F, ds.Tables("Article").Rows(0)("CodeArticle"))
                                oDTFailure.SetValue("ItemDescription", intID_F, ds.Tables("Article").Rows(0)("DescriptionArticle"))
                            Else
                                oDTFailure.SetValue("File Name", intID_F, FName.ToString)
                            End If

                            oDTFailure.SetValue("FailedReason", intID_F, ex.Message)
                            intID_F += 1

                        End Try

                    End If
                Next
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub MoveToFolder(ByVal strType As String, ByVal FName As String)
        Try
            Dim OpenFolder As String
            Dim SuccessFolder As String

            OpenFolder = GetOpenPath(strType)
            SuccessFolder = GetSuccessPath(strType)

            Dim strFileName As String = OpenFolder & "\" + FName
            If File.Exists(strFileName) = True Then
                System.IO.File.Move(strFileName, SuccessFolder + "\" + FName)
            End If
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Public Function ExecuteDataSet(ByVal strQuery As String) As DataSet

        Dim args() As Object = {oApplication.Company.Server, oApplication.Company.CompanyDB, oApplication.Company.DbUserName}
        Dim strConnection As String = String.Format(ConfigurationManager.AppSettings("ConnectionString").ToString(), args)

        myConnection = New SqlConnection(strConnection)
        Dim _retVal As New DataSet()
        Try
            myConnection.Open()
            oCommand = New SqlCommand()
            If myConnection.State = ConnectionState.Open Then
                oCommand.Connection = myConnection
                oCommand.CommandText = strQuery
                oCommand.CommandType = CommandType.Text
                oSqlAdap = New SqlDataAdapter(oCommand)
                oSqlAdap.Fill(_retVal)
            Else
                Throw New Exception("")
            End If
        Catch generatedExceptionName As Exception
            Throw
        Finally
            myConnection.Close()
            oCommand = Nothing
            oSqlAdap = Nothing
        End Try
        Return _retVal
    End Function

    Public Function CreatePurchaseOrder(ByVal oHeader As DataTable, ByVal oRow As DataTable, ByVal oBatch As DataTable, ByRef oRowDetails As DataTable, ByRef oBatchDetails As DataTable, ByVal strType As String) As Boolean
        Try
            Dim OPOR As SAPbobsCOM.Documents
            OPOR = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
            oRowDetails.Rows.Clear()
            oBatchDetails.Rows.Clear()
            If strType = "CPO" Then
                Dim odr As DataRow

                For Each oHdr As DataRow In oHeader.Rows
                    OPOR = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                    OPOR.CardCode = oHdr("CardCode")
                    ' OPOR.CardName = oHdr("CardName")
                    OPOR.DocDate = System.DateTime.Now 'oHdr("DocDate")

                    Dim intRow As Integer = 0
                    For Each oRdr As DataRow In oRow.Rows

                        If intRow > 0 Then
                            OPOR.Lines.Add()
                        End If

                        OPOR.Lines.SetCurrentLine(intRow)
                        OPOR.Lines.ItemCode = oHdr("ItemCode") 'oRdr("Code_Fournisseur")
                        ' OPOR.Lines.ItemDescription = oRdr("Code_Fournisseur")

                        odr = oRowDetails.NewRow()
                        odr("ITemCode") = oHdr("ItemCode") 'oRdr("Code_Fournisseur")
                        odr("ItemName") = oRdr("Code_Fournisseur")

                        For Each oBdr As DataRow In oBatch.Rows
                            If oRdr("Composant_livré_Id").ToString() = oBdr("Composant_livré_Id").ToString() Then 'Element Con

                                OPOR.Lines.Quantity = oBdr("Nombre")
                                OPOR.Lines.UnitPrice = GetPrice(oHdr("CardCode"), oHdr("ItemCode")) '1 ' 
                                OPOR.Lines.UoMEntry = getUOMEntry(oBdr("Unité"))


                                odr("Quantity") = oBdr("Nombre") 'To be taken from batch table
                                odr("UOM") = oBdr("Unité")
                                odr("BaseLine") = intRow
                                oRowDetails.Rows.Add(odr)

                                odr = oBatchDetails.NewRow()
                                odr("ITemCode") = oHdr("ItemCode")
                                odr("Quantity") = oBdr("Nombre")
                                odr("LineNo") = intRow
                                odr("BatchNo") = oRdr("Lot")
                                oBatchDetails.Rows.Add(odr)

                            End If
                        Next

                        intRow += 1

                    Next
                    intStatus = OPOR.Add()

                    If intStatus = 0 Then
                        For Each oRdr As DataRow In oRowDetails.Rows
                            oRdr("BaseEntry") = oApplication.Company.GetNewObjectKey() ' Purchase Order Ref
                        Next
                        oRowDetails.AcceptChanges()
                        oBatchDetails.AcceptChanges()
                        Return True
                    Else
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription(), SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        Return False
                    End If
                Next

            ElseIf strType = "BT" Then
                Dim odr As DataRow
                Dim intRow As Integer = 0
                For Each oHdr As DataRow In oHeader.Rows
                    For Each oRdr As DataRow In oRow.Rows

                        odr = oRowDetails.NewRow()

                        odr("ITemCode") = oHdr("ItemCode") 'oRdr("Code_Fournisseur")
                        odr("ItemName") = oHdr("ItemCode")

                        For Each oBdr As DataRow In oBatch.Rows
                            If oRdr("Composant_livré_Id").ToString() = oBdr("Composant_livré_Id").ToString() Then  'Element Con

                                odr("Quantity") = oBdr("Nombre")
                                odr("UnitPrice") = 1 '
                                odr("UOM") = oBdr("Unité")
                                oRowDetails.Rows.Add(odr)

                                odr = oBatchDetails.NewRow()
                                odr("ITemCode") = oHdr("ItemCode")
                                odr("Quantity") = oBdr("Nombre")
                                odr("LineNo") = intRow
                                odr("BatchNo") = oRdr("Lot")
                                oBatchDetails.Rows.Add(odr)

                            End If
                        Next

                        oRowDetails.AcceptChanges()
                        oBatchDetails.AcceptChanges()
                        intRow += 1
                    Next
                Next

                Return True

            ElseIf strType = "OPO" Then
                Dim odr As DataRow
                Dim intRow As Integer = 0

                For Each oHdr As DataRow In oHeader.Rows

                    Dim strCardCode As String = oHdr("CardCode")

                    For Each oRdr As DataRow In oRow.Rows
                        Dim strItemCode As String = oHdr("ItemCode").ToString() 'oRdr("Code_Fournisseur")


                        'Open Purchase Order Details
                        strqry = " Select T0.OpenQty,T0.DocEntry,T0.LineNum From POR1 T0 "
                        strqry += " Where T0.BaseCard = '" & strCardCode & "' And T0.LineStatus = 'O' "
                        strqry += " And T0.ItemCode = '" & strItemCode & "'"
                        strqry += " Order By DocEntry "
                        Dim OPO As DataSet = ExecuteDataSet(strqry)
                        If Not IsNothing(OPO) Then
                            If OPO.Tables(0).Rows.Count > 0 Then
                                intRow = 0


                                For Each oBdr As DataRow In oBatch.Rows

                                    If oRdr("Composant_livré_Id").ToString() = oBdr("Composant_livré_Id").ToString() Then  'Element Con

                                        Dim dblQty As Double = CDbl(oBdr("Nombre"))
                                        Dim dblToAppliedQty As Double = dblQty

                                        For Each dr As DataRow In OPO.Tables(0).Rows


                                            If dblToAppliedQty > 0 Then

                                                Dim dblSetQty As Double

                                                odr = oRowDetails.NewRow()
                                                odr("ITemCode") = strItemCode
                                                odr("ItemName") = strItemCode

                                                Dim dblOQty As Double = CDbl(dr("OpenQty"))

                                                If dblToAppliedQty <= dblOQty Then
                                                    odr("Quantity") = dblToAppliedQty
                                                    dblSetQty = dblToAppliedQty
                                                    dblToAppliedQty -= dblQty
                                                Else
                                                    odr("Quantity") = dblOQty
                                                    dblSetQty = dblOQty
                                                    dblToAppliedQty -= dblOQty
                                                End If

                                                odr("BaseEntry") = dr("DocEntry") ' Purchase Order Ref
                                                odr("BaseLine") = dr("LineNum")
                                                oRowDetails.Rows.Add(odr)


                                                odr = oBatchDetails.NewRow()
                                                odr("ITemCode") = strItemCode
                                                odr("Quantity") = dblSetQty
                                                odr("LineNo") = intRow
                                                odr("BatchNo") = oRdr("Lot")
                                                oBatchDetails.Rows.Add(odr)

                                                oRowDetails.AcceptChanges()
                                                oBatchDetails.AcceptChanges()

                                                intRow += 1
                                            End If

                                        Next

                                    End If
                                Next


                            End If
                        End If

                    Next

                Next
                Return True
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function CreateGRPO(ByVal oHeader As DataTable, ByVal oRowDetails As DataTable, ByVal oBatchDetails As DataTable, ByVal strCType As String) As Boolean
        Try

            Dim oGRPO As SAPbobsCOM.Documents
            oGRPO = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)


            For Each oHdr As DataRow In oHeader.Rows

                oGRPO.CardCode = oHdr("CardCode")
                'oGRPO.CardName = oHdr("CardName")
                oGRPO.DocDate = System.DateTime.Now 'oHdr("DocDate")


                Dim intRow As Integer = 0
                For Each oRdr As DataRow In oRowDetails.Rows

                    If intRow > 0 Then
                        oGRPO.Lines.Add()
                    End If

                    oGRPO.Lines.SetCurrentLine(intRow)
                    oGRPO.Lines.ItemCode = oRdr("ItemCode")
                    'oGRPO.Lines.ItemDescription = oRdr("ItemName")
                    oGRPO.Lines.Quantity = oRdr("Quantity")
                    'oGRPO.Lines.UnitPrice = 1 'oRdr("UnitPrice")
                    oGRPO.Lines.UoMEntry = getUOMEntry(oRdr("UOM"))
                    'oGRPO.Lines.TaxCode = "X1"
                    'oGRPO.Lines.WarehouseCode = "01"

                    If strCType <> "BT" Then
                        oGRPO.Lines.BaseType = 22
                        oGRPO.Lines.BaseEntry = oRdr("BaseEntry")
                        oGRPO.Lines.BaseLine = oRdr("BaseLine")
                    End If

                    Dim oItem As SAPbobsCOM.Items
                    oItem = oApplication.Company.GetBusinessObject(BoObjectTypes.oItems)
                    If oItem.GetByKey(oRdr("ItemCode")) Then
                        If oItem.ManageBatchNumbers = BoYesNoEnum.tYES Then
                            Dim inTbatchLine As Integer = 0
                            Dim oBatchFilter As New DataView(oBatchDetails)
                            oBatchFilter.RowFilter = " LineNo = '" & intRow.ToString() & "'"
                            For Each oBdr As DataRow In oBatchFilter.ToTable().Rows
                                If inTbatchLine > 0 Then
                                    oGRPO.Lines.BatchNumbers.Add()
                                End If
                                oGRPO.Lines.BatchNumbers.SetCurrentLine(inTbatchLine)
                                oGRPO.Lines.BatchNumbers.BatchNumber = oBdr("BatchNo")
                                oGRPO.Lines.BatchNumbers.Quantity = oBdr("Quantity")
                                inTbatchLine = inTbatchLine + 1
                            Next
                        End If
                    End If

                    intRow += 1

                Next

                intStatus = oGRPO.Add()
                If intStatus = 0 Then
                    Return True
                Else
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription(), SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Return False
                End If

            Next



        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function CreateGI(ByVal oHeader As DataTable, ByVal oRowDetails As DataTable, ByVal oBatchDetails As DataTable) As Boolean
        Try
            Dim oGI As SAPbobsCOM.Documents
            oGI = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)

            For Each oHdr As DataRow In oHeader.Rows
                Dim intRow As Integer = 0
                oGI.DocDate = System.DateTime.Now
                oGI.Reference1 = oHdr("DocRef")

                For Each oRdr As DataRow In oRowDetails.Rows
                    If intRow > 0 Then
                        oGI.Lines.Add()
                    End If

                    oGI.Lines.SetCurrentLine(intRow)
                    oGI.Lines.ItemCode = oRdr("ItemCode")
                    oGI.Lines.ItemDescription = oRdr("ItemName")
                    oGI.Lines.Quantity = oRdr("Quantity")
                    'oGI.Lines.UnitPrice = oRdr("UnitPrice")
                    oGI.Lines.UoMEntry = getUOMEntry(oRdr("UOM"))

                    Dim oItem As SAPbobsCOM.Items
                    oItem = oApplication.Company.GetBusinessObject(BoObjectTypes.oItems)
                    If oItem.GetByKey(oRdr("ItemCode")) Then
                        If oItem.ManageBatchNumbers = BoYesNoEnum.tYES Then
                            Dim inTbatchLine As Integer = 0
                            Dim oBatchFilter As New DataView(oBatchDetails)
                            oBatchFilter.RowFilter = " LineNo = '" & intRow.ToString() & "'"
                            For Each oBdr As DataRow In oBatchFilter.ToTable().Rows
                                If inTbatchLine > 0 Then
                                    oGI.Lines.BatchNumbers.Add()
                                End If
                                oGI.Lines.BatchNumbers.SetCurrentLine(inTbatchLine)
                                oGI.Lines.BatchNumbers.BatchNumber = oBdr("BatchNo")
                                oGI.Lines.BatchNumbers.Quantity = oBdr("Quantity")
                                inTbatchLine = inTbatchLine + 1
                            Next
                        End If
                    End If

                    intRow += 1

                Next
            Next
            intStatus = oGI.Add()

            If intStatus = 0 Then
                Return True
            Else
                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription(), SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Return False
            End If

        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function CreateGI_Clear(ByVal oHeader As DataTable, ByVal oRowDetails As DataTable, ByVal dblQty As Double) As Boolean
        Try
            Dim oGI As SAPbobsCOM.Documents
            oGI = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oBatchSet As SAPbobsCOM.Recordset
            oBatchSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            For Each oHdr As DataRow In oHeader.Rows
                oGI = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)

                Dim intRow As Integer = 0
                oGI.DocDate = System.DateTime.Now
                oGI.Reference1 = oHdr("DocRef")
                Dim strItemCode As String = oHdr("ItemCode").ToString()
                Dim strWare As String = oHdr("WareHouse").ToString()

                For Each oRdr As DataRow In oRowDetails.Rows

                    strQuery = "Select ISNULL(OnHand,0) AS OpenQty From OITW Where ItemCode = '" & strItemCode & "' And WhsCode = '" & strWare & "'"
                    oRecordSet.DoQuery(strQuery)

                    If Not oRecordSet.EoF Then

                        oGI.Lines.SetCurrentLine(intRow)
                        oGI.Lines.ItemCode = strItemCode
                        ' oGI.Lines.ItemDescription = oRdr("ItemName")
                        oGI.Lines.Quantity = dblQty 'CDbl(oRecordSet.Fields.Item("OpenQty").Value)
                        oGI.Lines.UoMEntry = getUOMEntry(oRdr("UOM"))

                        'Dim inTbatchLine As Integer = 0
                        'Dim oBatchFilter As New DataView(oBatchDetails)
                        'oBatchFilter.RowFilter = " LineNo = '" & intRow.ToString() & "'"

                        Dim oItem As SAPbobsCOM.Items
                        oItem = oApplication.Company.GetBusinessObject(BoObjectTypes.oItems)
                        If oItem.GetByKey(strItemCode) Then
                            If oItem.ManageBatchNumbers = BoYesNoEnum.tYES Then

                                strQuery = " Select T3.BatchNum,Convert(Decimal(18,2),T3.Quantity) AS 'BQuantity' "
                                strQuery += " From  OIBT T3 Where T3.ItemCode = '" & strItemCode & "' "
                                strQuery += " And T3.WhsCode = '" & strWare & "' "
                                strQuery += " And T3.Quantity > 0 "
                                oBatchSet.DoQuery(strQuery)
                                Dim inTbatchLine As Integer
                                Dim batchquantity As Double
                                Dim dblAssignqty As Double = 0

                                Dim dblBatchRequiredQty As Double = dblQty

                                For intBatch As Integer = 0 To oBatchSet.RecordCount - 1
                                    While (dblBatchRequiredQty > 0 And Not oBatchSet.EoF)
                                        batchquantity = oBatchSet.Fields.Item("BQuantity").Value

                                        If batchquantity >= dblBatchRequiredQty Then
                                            dblAssignqty = dblBatchRequiredQty
                                        Else
                                            dblAssignqty = batchquantity
                                        End If

                                        If inTbatchLine > 0 Then
                                            oGI.Lines.BatchNumbers.Add()
                                        End If

                                        oGI.Lines.BatchNumbers.SetCurrentLine(inTbatchLine)
                                        oGI.Lines.BatchNumbers.BatchNumber = oBatchSet.Fields.Item("BatchNum").Value
                                        oGI.Lines.BatchNumbers.Quantity = dblAssignqty 'oBatchSet.Fields.Item("BQuantity").Value

                                        inTbatchLine = inTbatchLine + 1
                                        dblBatchRequiredQty = dblBatchRequiredQty - dblAssignqty
                                        oBatchSet.MoveNext()
                                    End While
                                Next

                            End If
                        End If
                    End If

                Next

            Next
            intStatus = oGI.Add()

            If intStatus = 0 Then
                Return True
            Else
                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription(), SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Return False
            End If

        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function getUOMEntry(ByVal strCode As String) As String
        Dim _retVal As String = String.Empty
        Try
            Dim strQuery As String = String.Empty
            Dim oRecordSet As SAPbobsCOM.Recordset

            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If strCode <> "" Then
                oRecordSet.DoQuery("Select UomEntry From OUOM Where UomCode = '" & strCode & "'")
                If Not oRecordSet.EoF Then
                    _retVal = (oRecordSet.Fields.Item(0).Value)
                Else
                    _retVal = -1
                End If
            Else
                _retVal = -1
            End If

            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function GetPrice(ByVal strCardCode As String, ByVal strItemCode As String) As Double
        Try
            Dim ItemPrice As Double
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oSBOBOB As SAPbobsCOM.SBObob
            oSBOBOB = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
            oRecordSet = oSBOBOB.GetItemPrice(strCardCode, strItemCode, 1, DateTime.Today)
            If oRecordSet.RecordCount > 0 Then
                If Not oRecordSet.EoF Then
                    ItemPrice = oRecordSet.Fields.Item(0).Value
                End If
            End If
            Return ItemPrice
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function addJournal_Entry(ByVal oHeader As DataTable, ByVal oRowDetails As DataTable) As Boolean
        Try
            Dim _retVal As Boolean = False
            Dim intStatus As Int16 = 0
            Dim strQuery As String = String.Empty
            Dim intRow As Integer = 1
            Dim dblDebit As Double
            Dim dblCredit As Double
            Dim blnHasRow As Boolean = False

            Dim oRecordSet As SAPbobsCOM.Recordset
            Dim oJE As SAPbobsCOM.JournalEntries
            oJE = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            ' Dim dtCheqDate As Date = CDate(oRecordSet.Fields.Item("CheqDate").Value)
            oJE.ReferenceDate = System.DateTime.Now
            oJE.TaxDate = System.DateTime.Now
            oJE.DueDate = System.DateTime.Now
            oJE.TransactionCode = "GE"
            oJE.Memo = "General Journal Import"

            For Each oRdr As DataRow In oRowDetails.Rows

                'Dim strGL As String = oRecordSet.Fields.Item("GL").Value
                'Dim strEntity As String = oRecordSet.Fields.Item("Entity").Value
                'Dim strRemarks As String = oRecordSet.Fields.Item("AcctDesc").Value

                blnHasRow = True
                If intRow > 1 Then
                    oJE.Lines.Add()
                End If

                oJE.Lines.AccountCode = getAccount(oRdr("AccCode"))
                dblDebit = CDbl(oRdr("SARDRAmount"))
                dblCredit = CDbl(oRdr("SARCRAmount"))

                If dblDebit > 0 Then
                    oJE.Lines.Debit = dblDebit
                ElseIf dblCredit > 0 Then
                    oJE.Lines.Credit = dblCredit
                End If

                oJE.Lines.TaxDate = System.DateTime.Now
                oJE.Lines.ReferenceDate1 = System.DateTime.Now
                oJE.Lines.LineMemo = oRdr("Remarks")

                intRow = intRow + 1

            Next

            If blnHasRow Then
                intStatus = oJE.Add()
                If intStatus <> 0 Then
                    _retVal = False
                Else
                    _retVal = True
                End If
            End If
            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function getAccount(ByVal strFormat As String) As String
        Dim _retVal As String = String.Empty
        Try
            Dim strQuery As String = String.Empty
            Dim oARecordSet As SAPbobsCOM.Recordset
            oARecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = " Select AcctCode From OACT Where FormatCode = "
            strQuery &= "'" + strFormat + "'"
            oARecordSet.DoQuery(strQuery)
            If Not oARecordSet.EoF Then
                _retVal = oARecordSet.Fields.Item(0).Value
            End If
        Catch ex As Exception
            Throw ex
        End Try
        Return _retVal
    End Function

End Class

'oApplication.Company.StartTransaction()
'If CreatePurchaseOrder(oHeader, ds.Tables("Composant_livré"), ds.Tables("Quantité_composant"), oRowDetails, oBatchDetails, "CPO") Then ' Create PO & Batch Construction
'    If CreateGRPO(oHeader, oRowDetails, oBatchDetails, "CPO") Then ' GRPO Based on Created PO
'        If CreateGI(oHeader, oRowDetails, oBatchDetails) Then
'            If oApplication.Company.InTransaction Then
'                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
'            End If
'        Else
'            If oApplication.Company.InTransaction Then
'                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
'            End If
'            intStatus = oApplication.Company.GetLastErrorCode()
'        End If
'    Else
'        If oApplication.Company.InTransaction Then
'            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
'        End If
'        intStatus = oApplication.Company.GetLastErrorCode()
'    End If
'Else
'    If oApplication.Company.InTransaction Then
'        oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
'    End If
'    intStatus = oApplication.Company.GetLastErrorCode()
'End If
