﻿Imports System.IO
Imports HtmlAgilityPack
Imports System.Net.Http
Imports System.Threading
Imports Utilities.DAL
Imports Utilities.Network
Imports System.Net

Public Class frmMain

#Region "Common Delegates"
    Delegate Sub SetObjectEnableDisable_Delegate(ByVal [obj] As Object, ByVal [value] As Boolean)
    Public Sub SetObjectEnableDisable_ThreadSafe(ByVal [obj] As Object, ByVal [value] As Boolean)
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [obj].InvokeRequired Then
            Dim MyDelegate As New SetObjectEnableDisable_Delegate(AddressOf SetObjectEnableDisable_ThreadSafe)
            Me.Invoke(MyDelegate, New Object() {[obj], [value]})
        Else
            [obj].Enabled = [value]
        End If
    End Sub

    Delegate Sub SetObjectVisible_Delegate(ByVal [obj] As Object, ByVal [value] As Boolean)
    Public Sub SetObjectVisible_ThreadSafe(ByVal [obj] As Object, ByVal [value] As Boolean)
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [obj].InvokeRequired Then
            Dim MyDelegate As New SetObjectVisible_Delegate(AddressOf SetObjectVisible_ThreadSafe)
            Me.Invoke(MyDelegate, New Object() {[obj], [value]})
        Else
            [obj].Visible = [value]
        End If
    End Sub

    Delegate Sub SetLabelText_Delegate(ByVal [label] As Label, ByVal [text] As String)
    Public Sub SetLabelText_ThreadSafe(ByVal [label] As Label, ByVal [text] As String)
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [label].InvokeRequired Then
            Dim MyDelegate As New SetLabelText_Delegate(AddressOf SetLabelText_ThreadSafe)
            Me.Invoke(MyDelegate, New Object() {[label], [text]})
        Else
            [label].Text = [text]
        End If
    End Sub

    Delegate Function GetLabelText_Delegate(ByVal [label] As Label) As String
    Public Function GetLabelText_ThreadSafe(ByVal [label] As Label) As String
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [label].InvokeRequired Then
            Dim MyDelegate As New GetLabelText_Delegate(AddressOf GetLabelText_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[label]})
        Else
            Return [label].Text
        End If
    End Function

    Delegate Sub SetLabelTag_Delegate(ByVal [label] As Label, ByVal [tag] As String)
    Public Sub SetLabelTag_ThreadSafe(ByVal [label] As Label, ByVal [tag] As String)
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [label].InvokeRequired Then
            Dim MyDelegate As New SetLabelTag_Delegate(AddressOf SetLabelTag_ThreadSafe)
            Me.Invoke(MyDelegate, New Object() {[label], [tag]})
        Else
            [label].Tag = [tag]
        End If
    End Sub

    Delegate Function GetLabelTag_Delegate(ByVal [label] As Label) As String
    Public Function GetLabelTag_ThreadSafe(ByVal [label] As Label) As String
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [label].InvokeRequired Then
            Dim MyDelegate As New GetLabelTag_Delegate(AddressOf GetLabelTag_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[label]})
        Else
            Return [label].Tag
        End If
    End Function
    Delegate Sub SetToolStripLabel_Delegate(ByVal [toolStrip] As StatusStrip, ByVal [label] As ToolStripStatusLabel, ByVal [text] As String)
    Public Sub SetToolStripLabel_ThreadSafe(ByVal [toolStrip] As StatusStrip, ByVal [label] As ToolStripStatusLabel, ByVal [text] As String)
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [toolStrip].InvokeRequired Then
            Dim MyDelegate As New SetToolStripLabel_Delegate(AddressOf SetToolStripLabel_ThreadSafe)
            Me.Invoke(MyDelegate, New Object() {[toolStrip], [label], [text]})
        Else
            [label].Text = [text]
        End If
    End Sub

    Delegate Function GetToolStripLabel_Delegate(ByVal [toolStrip] As StatusStrip, ByVal [label] As ToolStripLabel) As String
    Public Function GetToolStripLabel_ThreadSafe(ByVal [toolStrip] As StatusStrip, ByVal [label] As ToolStripLabel) As String
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [toolStrip].InvokeRequired Then
            Dim MyDelegate As New GetToolStripLabel_Delegate(AddressOf GetToolStripLabel_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[toolStrip], [label]})
        Else
            Return [label].Text
        End If
    End Function

    Delegate Function GetDateTimePickerValue_Delegate(ByVal [dateTimePicker] As DateTimePicker) As Date
    Public Function GetDateTimePickerValue_ThreadSafe(ByVal [dateTimePicker] As DateTimePicker) As Date
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [dateTimePicker].InvokeRequired Then
            Dim MyDelegate As New GetDateTimePickerValue_Delegate(AddressOf GetDateTimePickerValue_ThreadSafe)
            Return Me.Invoke(MyDelegate, New DateTimePicker() {[dateTimePicker]})
        Else
            Return [dateTimePicker].Value
        End If
    End Function

    Delegate Function GetNumericUpDownValue_Delegate(ByVal [numericUpDown] As NumericUpDown) As Integer
    Public Function GetNumericUpDownValue_ThreadSafe(ByVal [numericUpDown] As NumericUpDown) As Integer
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [numericUpDown].InvokeRequired Then
            Dim MyDelegate As New GetNumericUpDownValue_Delegate(AddressOf GetNumericUpDownValue_ThreadSafe)
            Return Me.Invoke(MyDelegate, New NumericUpDown() {[numericUpDown]})
        Else
            Return [numericUpDown].Value
        End If
    End Function

    Delegate Function GetComboBoxIndex_Delegate(ByVal [combobox] As ComboBox) As Integer
    Public Function GetComboBoxIndex_ThreadSafe(ByVal [combobox] As ComboBox) As Integer
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [combobox].InvokeRequired Then
            Dim MyDelegate As New GetComboBoxIndex_Delegate(AddressOf GetComboBoxIndex_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[combobox]})
        Else
            Return [combobox].SelectedIndex
        End If
    End Function

    Delegate Function GetComboBoxItem_Delegate(ByVal [ComboBox] As ComboBox) As String
    Public Function GetComboBoxItem_ThreadSafe(ByVal [ComboBox] As ComboBox) As String
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [ComboBox].InvokeRequired Then
            Dim MyDelegate As New GetComboBoxItem_Delegate(AddressOf GetComboBoxItem_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[ComboBox]})
        Else
            Return [ComboBox].SelectedItem.ToString
        End If
    End Function

    Delegate Function GetTextBoxText_Delegate(ByVal [textBox] As TextBox) As String
    Public Function GetTextBoxText_ThreadSafe(ByVal [textBox] As TextBox) As String
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [textBox].InvokeRequired Then
            Dim MyDelegate As New GetTextBoxText_Delegate(AddressOf GetTextBoxText_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[textBox]})
        Else
            Return [textBox].Text
        End If
    End Function

    Delegate Function GetCheckBoxChecked_Delegate(ByVal [checkBox] As CheckBox) As Boolean
    Public Function GetCheckBoxChecked_ThreadSafe(ByVal [checkBox] As CheckBox) As Boolean
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [checkBox].InvokeRequired Then
            Dim MyDelegate As New GetCheckBoxChecked_Delegate(AddressOf GetCheckBoxChecked_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[checkBox]})
        Else
            Return [checkBox].Checked
        End If
    End Function

    Delegate Function GetRadioButtonChecked_Delegate(ByVal [radioButton] As RadioButton) As Boolean
    Public Function GetRadioButtonChecked_ThreadSafe(ByVal [radioButton] As RadioButton) As Boolean
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [radioButton].InvokeRequired Then
            Dim MyDelegate As New GetRadioButtonChecked_Delegate(AddressOf GetRadioButtonChecked_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[radioButton]})
        Else
            Return [radioButton].Checked
        End If
    End Function

    Delegate Sub SetDatagridBindDatatable_Delegate(ByVal [datagrid] As DataGridView, ByVal [table] As DataTable)
    Public Sub SetDatagridBindDatatable_ThreadSafe(ByVal [datagrid] As DataGridView, ByVal [table] As DataTable)
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [datagrid].InvokeRequired Then
            Dim MyDelegate As New SetDatagridBindDatatable_Delegate(AddressOf SetDatagridBindDatatable_ThreadSafe)
            Me.Invoke(MyDelegate, New Object() {[datagrid], [table]})
        Else
            [datagrid].DataSource = [table]
            [datagrid].Refresh()
        End If
    End Sub
#End Region

#Region "Event Handlers"
    Private Sub OnHeartbeatMain(message As String)
        SetLabelText_ThreadSafe(lblMainProgress, message)
    End Sub
    Private Sub OnHeartbeat(message As String)
        SetLabelText_ThreadSafe(lblProgress, message)
    End Sub
    Private Sub OnDocumentDownloadComplete()
        'OnHeartbeat("Document download compelete")
    End Sub
    Private Sub OnDocumentRetryStatus(currentTry As Integer, totalTries As Integer)
        OnHeartbeat(String.Format("Try #{0}/{1}: Connecting...", currentTry, totalTries))
    End Sub
    Public Sub OnWaitingFor(ByVal elapsedSecs As Integer, ByVal totalSecs As Integer, ByVal msg As String)
        OnHeartbeat(String.Format("{0}, waiting {1}/{2} secs", msg, elapsedSecs, totalSecs))
    End Sub
#End Region

    Private canceller As CancellationTokenSource
    Private NSEIDXOptionChainURL As String = "https://www.nseindia.com/live_market/dynaContent/live_watch/option_chain/optionKeys.jsp?symbolCode=-0&symbol={0}&symbol={0}&instrument=OPTIDX&date=-&segmentLink=17&segmentLink=17"
    Private NSESTKOptionChainURL As String = "https://www.nseindia.com/live_market/dynaContent/live_watch/option_chain/optionKeys.jsp?symbolCode=-0&symbol={0}&symbol={0}&instrument=OPTSTK&date=-&segmentLink=17&segmentLink=17"

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetObjectEnableDisable_ThreadSafe(btnStop, False)
    End Sub

    Private Sub btnStop_Click(sender As Object, e As EventArgs) Handles btnStop.Click
        canceller.Cancel()
    End Sub

    Private Async Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        SetObjectEnableDisable_ThreadSafe(btnStop, True)
        SetObjectEnableDisable_ThreadSafe(btnStart, False)
        canceller = New CancellationTokenSource
        Await Task.Run(AddressOf StartProcessingAsync).ConfigureAwait(False)
    End Sub

    Private Async Function StartProcessingAsync() As Task
        Try
            Dim stockList As List(Of String) = Await GetAllStockList().ConfigureAwait(False)
            If stockList IsNot Nothing AndAlso stockList.Count > 0 Then
                Using sqlHlpr As New MySQLDBHelper(My.Settings.ServerName, "local_stock", "3306", "rio", "speech123", canceller)
                    AddHandler sqlHlpr.Heartbeat, AddressOf OnHeartbeat
                    AddHandler sqlHlpr.DocumentDownloadComplete, AddressOf OnDocumentDownloadComplete
                    AddHandler sqlHlpr.DocumentRetryStatus, AddressOf OnDocumentRetryStatus
                    AddHandler sqlHlpr.WaitingFor, AddressOf OnWaitingFor

                    Dim stkCtr As Integer = 0
                    For Each runningStock In stockList
                        stkCtr += 1
                        OnHeartbeatMain(String.Format("Running for: {0} #{1}/{2}", runningStock, stkCtr, stockList.Count))
                        Dim openPositionDataURL As String = Nothing
                        If runningStock = "NIFTY" OrElse runningStock = "BANKNIFTY" OrElse runningStock = "NIFTYIT" Then
                            openPositionDataURL = String.Format(NSEIDXOptionChainURL, runningStock)
                        Else
                            openPositionDataURL = String.Format(NSESTKOptionChainURL, runningStock)
                        End If
                        Dim outputResponse As HtmlDocument = Nothing
                        Dim proxyToBeUsed As HttpProxy = Nothing
                        Using browser As New HttpBrowser(proxyToBeUsed, Net.DecompressionMethods.GZip, New TimeSpan(0, 1, 0), canceller)
                            AddHandler browser.DocumentDownloadComplete, AddressOf OnDocumentDownloadComplete
                            AddHandler browser.Heartbeat, AddressOf OnHeartbeat
                            AddHandler browser.WaitingFor, AddressOf OnWaitingFor
                            AddHandler browser.DocumentRetryStatus, AddressOf OnDocumentRetryStatus

                            browser.KeepAlive = True
                            Dim headersToBeSent As New Dictionary(Of String, String)
                            headersToBeSent.Add("Host", "www.nseindia.com")
                            headersToBeSent.Add("Upgrade-Insecure-Requests", "1")
                            headersToBeSent.Add("Sec-Fetch-Mode", "navigate")
                            headersToBeSent.Add("Sec-Fetch-Site", "none")

                            Dim l As Tuple(Of Uri, Object) = Await browser.NonPOSTRequestAsync(openPositionDataURL,
                                                                                            HttpMethod.Get,
                                                                                            Nothing,
                                                                                            False,
                                                                                            headersToBeSent,
                                                                                            True,
                                                                                            "text/html").ConfigureAwait(False)
                            If l IsNot Nothing AndAlso l.Item2 IsNot Nothing Then
                                outputResponse = l.Item2
                            End If
                            RemoveHandler browser.DocumentDownloadComplete, AddressOf OnDocumentDownloadComplete
                            RemoveHandler browser.Heartbeat, AddressOf OnHeartbeat
                            RemoveHandler browser.WaitingFor, AddressOf OnWaitingFor
                            RemoveHandler browser.DocumentRetryStatus, AddressOf OnDocumentRetryStatus
                        End Using

                        If outputResponse IsNot Nothing AndAlso outputResponse.DocumentNode IsNot Nothing Then
                            OnHeartbeat("Extracting Option Chain from HTML")
                            Dim calls As List(Of OptionChain) = Nothing
                            Dim puts As List(Of OptionChain) = Nothing
                            If outputResponse.DocumentNode.SelectNodes("//table[@id='octable']") IsNot Nothing Then
                                For Each table As HtmlNode In outputResponse.DocumentNode.SelectNodes("//table[@id='octable']")
                                    canceller.Token.ThrowIfCancellationRequested()
                                    If table IsNot Nothing AndAlso table.SelectNodes("tr") IsNot Nothing Then
                                        For Each row As HtmlNode In table.SelectNodes("tr")
                                            canceller.Token.ThrowIfCancellationRequested()
                                            If row IsNot Nothing AndAlso row.SelectNodes("td") IsNot Nothing Then
                                                Dim callData As OptionChain = Nothing
                                                Dim putData As OptionChain = Nothing
                                                Dim counter As Integer = 0
                                                For Each cell As HtmlNode In row.SelectNodes("td")
                                                    canceller.Token.ThrowIfCancellationRequested()
                                                    If cell IsNot Nothing AndAlso cell.InnerText IsNot Nothing AndAlso cell.InnerText <> "" Then
                                                        If cell.InnerText.Trim = "Total" Then
                                                            Exit For
                                                        End If
                                                        If callData Is Nothing Then callData = New OptionChain
                                                        If putData Is Nothing Then putData = New OptionChain
                                                        counter += 1
                                                        If counter = 1 Then
                                                            callData.OI = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                                        ElseIf counter = 2 Then
                                                            callData.ChangeInOI = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                                        ElseIf counter = 3 Then
                                                            callData.Volume = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                                        ElseIf counter = 4 Then
                                                            callData.IV = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                                        ElseIf counter = 5 Then
                                                            callData.LTP = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                                        ElseIf counter = 6 Then
                                                            callData.NetChange = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                                        ElseIf counter = 7 Then
                                                            callData.BidQuantity = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                                        ElseIf counter = 8 Then
                                                            callData.BidPrice = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                                        ElseIf counter = 9 Then
                                                            callData.AskPrice = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                                        ElseIf counter = 10 Then
                                                            callData.AskQuantity = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                                        ElseIf counter = 11 Then
                                                            callData.StrikePrice = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                                            putData.StrikePrice = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                                        ElseIf counter = 12 Then
                                                            putData.BidQuantity = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                                        ElseIf counter = 13 Then
                                                            putData.BidPrice = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                                        ElseIf counter = 14 Then
                                                            putData.AskPrice = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                                        ElseIf counter = 15 Then
                                                            putData.AskQuantity = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                                        ElseIf counter = 16 Then
                                                            putData.NetChange = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                                        ElseIf counter = 17 Then
                                                            putData.LTP = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                                        ElseIf counter = 18 Then
                                                            putData.IV = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                                        ElseIf counter = 19 Then
                                                            putData.Volume = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                                        ElseIf counter = 20 Then
                                                            putData.ChangeInOI = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                                        ElseIf counter = 21 Then
                                                            putData.OI = If(cell.InnerText.Trim = "-", Decimal.MinValue, cell.InnerText.Trim)
                                                        End If
                                                    End If
                                                Next
                                                If callData IsNot Nothing Then
                                                    If calls Is Nothing Then calls = New List(Of OptionChain)
                                                    calls.Add(callData)
                                                End If
                                                If putData IsNot Nothing Then
                                                    If puts Is Nothing Then puts = New List(Of OptionChain)
                                                    puts.Add(putData)
                                                End If
                                            End If
                                        Next
                                    End If
                                Next
                            End If
                            canceller.Token.ThrowIfCancellationRequested()

                            OnHeartbeat("Creating table of option chain data")
                            canceller.Token.ThrowIfCancellationRequested()
                            If calls IsNot Nothing AndAlso calls.Count > 0 AndAlso puts IsNot Nothing AndAlso puts.Count > 0 AndAlso calls.Count = puts.Count Then
                                canceller.Token.ThrowIfCancellationRequested()
                                Dim insertDataString As String = Nothing
                                For Each runningCall In calls
                                    canceller.Token.ThrowIfCancellationRequested()
                                    insertDataString = String.Format("{0},('{1}','{2}',{3},'{4}',{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},TIMESTAMP(CURRENT_TIME))",
                                                                        insertDataString,
                                                                        Now.ToString("yyyy-MM-dd"),
                                                                        runningStock.ToUpper,
                                                                        runningCall.StrikePrice,
                                                                        "CE",
                                                                        If(runningCall.OI <> Decimal.MinValue, runningCall.OI, Nothing),
                                                                        If(runningCall.ChangeInOI <> Decimal.MinValue, runningCall.ChangeInOI, Nothing),
                                                                        If(runningCall.Volume <> Decimal.MinValue, runningCall.Volume, Nothing),
                                                                        If(runningCall.IV <> Decimal.MinValue, runningCall.IV, Nothing),
                                                                        If(runningCall.LTP <> Decimal.MinValue, runningCall.LTP, Nothing),
                                                                        If(runningCall.NetChange <> Decimal.MinValue, runningCall.NetChange, Nothing),
                                                                        If(runningCall.BidQuantity <> Decimal.MinValue, runningCall.BidQuantity, Nothing),
                                                                        If(runningCall.BidPrice <> Decimal.MinValue, runningCall.BidPrice, Nothing),
                                                                        If(runningCall.AskPrice <> Decimal.MinValue, runningCall.AskPrice, Nothing),
                                                                        If(runningCall.AskQuantity <> Decimal.MinValue, runningCall.AskQuantity, Nothing))
                                Next
                                For Each runningPut In puts
                                    canceller.Token.ThrowIfCancellationRequested()
                                    insertDataString = String.Format("{0},('{1}','{2}',{3},'{4}',{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},TIMESTAMP(CURRENT_TIME))",
                                                                        insertDataString,
                                                                        Now.ToString("yyyy-MM-dd"),
                                                                        runningStock.ToUpper,
                                                                        runningPut.StrikePrice,
                                                                        "PE",
                                                                        If(runningPut.OI <> Decimal.MinValue, runningPut.OI, Nothing),
                                                                        If(runningPut.ChangeInOI <> Decimal.MinValue, runningPut.ChangeInOI, Nothing),
                                                                        If(runningPut.Volume <> Decimal.MinValue, runningPut.Volume, Nothing),
                                                                        If(runningPut.IV <> Decimal.MinValue, runningPut.IV, Nothing),
                                                                        If(runningPut.LTP <> Decimal.MinValue, runningPut.LTP, Nothing),
                                                                        If(runningPut.NetChange <> Decimal.MinValue, runningPut.NetChange, Nothing),
                                                                        If(runningPut.BidQuantity <> Decimal.MinValue, runningPut.BidQuantity, Nothing),
                                                                        If(runningPut.BidPrice <> Decimal.MinValue, runningPut.BidPrice, Nothing),
                                                                        If(runningPut.AskPrice <> Decimal.MinValue, runningPut.AskPrice, Nothing),
                                                                        If(runningPut.AskQuantity <> Decimal.MinValue, runningPut.AskQuantity, Nothing))
                                Next

                                If insertDataString IsNot Nothing Then
                                    Dim insertString As String = String.Format("INSERT INTO `nse_option_chain` 
                                                                                (`SnapshotDate`,
                                                                                `InstrumentName`,
                                                                                `StrikePrice`,
                                                                                `CallPut`,
                                                                                `OI`,
                                                                                `ChangeInOI`,
                                                                                `Volume`,
                                                                                `IV`,
                                                                                `LTP`,
                                                                                `NetChange`,
                                                                                `BidQuantity`,
                                                                                `BidPrice`,
                                                                                `AskPrice`,
                                                                                `AskQuantity`,
                                                                                `UpdateToDBTime`) 
                                                                                VALUES {0} 
                                                                                ON DUPLICATE KEY UPDATE 
                                                                                `SnapshotDate`=VALUES(`SnapshotDate`),
                                                                                `InstrumentName`=VALUES(`InstrumentName`),
                                                                                `StrikePrice`=VALUES(`StrikePrice`),
                                                                                `CallPut`=VALUES(`CallPut`),
                                                                                `OI`=VALUES(`OI`),
                                                                                `ChangeInOI`=VALUES(`ChangeInOI`),
                                                                                `Volume`=VALUES(`Volume`),
                                                                                `IV`=VALUES(`IV`),
                                                                                `LTP`=VALUES(`LTP`),
                                                                                `NetChange`=VALUES(`NetChange`),
                                                                                `BidQuantity`=VALUES(`BidQuantity`),
                                                                                `BidPrice`=VALUES(`BidPrice`),
                                                                                `AskPrice`=VALUES(`AskPrice`),
                                                                                `AskQuantity`=VALUES(`AskQuantity`),
                                                                                `UpdateToDBTime`=VALUES(`UpdateToDBTime`);", insertDataString.Substring(1))

                                    canceller.Token.ThrowIfCancellationRequested()
                                    Dim numberOfdata As Integer = Await sqlHlpr.RunUpdateAsync(insertString).ConfigureAwait(False)
                                End If
                                canceller.Token.ThrowIfCancellationRequested()
                            End If
                            canceller.Token.ThrowIfCancellationRequested()
                        End If
                    Next
                End Using
            End If
        Catch cex As OperationCanceledException
            MsgBox(cex.Message)
        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally
            SetObjectEnableDisable_ThreadSafe(btnStop, False)
            SetObjectEnableDisable_ThreadSafe(btnStart, True)
            SetLabelText_ThreadSafe(lblProgress, String.Format("Process Complete"))
        End Try
    End Function

    Private Async Function GetAllStockList() As Task(Of List(Of String))
        Dim ret As List(Of String) = Nothing
        Dim selectString As String = "SELECT `TRADING_SYMBOL` FROM `active_instruments_futures` WHERE `AS_ON_DATE`='{0}' AND `SEGMENT`='NFO-FUT'"
        selectString = String.Format(selectString, Now.ToString("yyyy-MM-dd"))
        Using sqlHlpr As New MySQLDBHelper(My.Settings.ServerName, "local_stock", "3306", "rio", "speech123", canceller)
            AddHandler sqlHlpr.Heartbeat, AddressOf OnHeartbeat
            AddHandler sqlHlpr.DocumentDownloadComplete, AddressOf OnDocumentDownloadComplete
            AddHandler sqlHlpr.DocumentRetryStatus, AddressOf OnDocumentRetryStatus
            AddHandler sqlHlpr.WaitingFor, AddressOf OnWaitingFor

            Dim dt As DataTable = Await sqlHlpr.RunSelectAsync(selectString).ConfigureAwait(False)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                For i = 0 To dt.Rows.Count - 1
                    Dim tradingSymbol As String = dt.Rows(i).Item(0).ToString.ToUpper
                    Dim instrumentName As String = Nothing
                    If tradingSymbol.Contains("FUT") Then
                        instrumentName = tradingSymbol.Remove(tradingSymbol.Count - 8)
                    Else
                        instrumentName = tradingSymbol
                    End If
                    If ret Is Nothing Then ret = New List(Of String)
                    If Not ret.Contains(instrumentName) Then ret.Add(instrumentName)
                Next
            End If
        End Using
        Return ret
    End Function
End Class
