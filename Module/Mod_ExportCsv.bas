Attribute VB_Name = "Mod_ExportCsv"
'****************************************************************
'Microsoft SQL Server 2000
Option Explicit
Public goPackageOld As New DTS.Package
Public goPackage As DTS.Package2

Public Sub CsvFileExport_Process(ByVal vFileName As String)
        Set goPackage = goPackageOld

        goPackage.Name = "New Package"
        goPackage.Description = "DTS package description"
        goPackage.WriteCompletionStatusToNTEventLog = False
        goPackage.FailOnError = False
        goPackage.PackagePriorityClass = 2
        goPackage.MaxConcurrentSteps = 4
        goPackage.LineageOptions = 0
        goPackage.UseTransaction = True
        goPackage.TransactionIsolationLevel = 4096
        goPackage.AutoCommitTransaction = True
        goPackage.RepositoryMetadataOptions = 0
        goPackage.UseOLEDBServiceComponents = True
        goPackage.LogToSQLServer = False
        goPackage.LogServerFlags = 0
        goPackage.FailPackageOnLogFailure = False
        goPackage.ExplicitGlobalVariables = False
        goPackage.PackageType = 0
        

Dim oConnProperty As DTS.OleDBProperty

'---------------------------------------------------------------------------
' create package connection information
'---------------------------------------------------------------------------

Dim oConnection As DTS.Connection2

'------------- a new connection defined below.
'For security purposes, the password is never scripted

Set oConnection = goPackage.Connections.New("SQLOLEDB")

        oConnection.ConnectionProperties("Persist Security Info") = True
        oConnection.ConnectionProperties("User ID") = g_DatabaseUser
        oConnection.ConnectionProperties("Initial Catalog") = g_Database
        oConnection.ConnectionProperties("Data Source") = g_Server
        oConnection.ConnectionProperties("Application Name") = "DTS  Import/Export Wizard"
        
        oConnection.Name = "Connection 1"
        oConnection.ID = 1
        oConnection.Reusable = True
        oConnection.ConnectImmediate = False
        oConnection.DataSource = g_Server
        oConnection.UserID = g_DatabaseUser
        oConnection.ConnectionTimeout = 60
        oConnection.Catalog = g_Database
        oConnection.UseTrustedConnection = False
        oConnection.UseDSL = False
        oConnection.Password = g_DatabasePwd

goPackage.Connections.Add oConnection
Set oConnection = Nothing

'------------- a new connection defined below.
'For security purposes, the password is never scripted

Set oConnection = goPackage.Connections.New("DTSFlatFile")

        oConnection.ConnectionProperties("Data Source") = vFileName
        oConnection.ConnectionProperties("Mode") = 3
        oConnection.ConnectionProperties("Row Delimiter") = vbCrLf
        oConnection.ConnectionProperties("File Format") = 1
        oConnection.ConnectionProperties("Column Delimiter") = ","
        oConnection.ConnectionProperties("File Type") = 1
        oConnection.ConnectionProperties("Skip Rows") = 0
        oConnection.ConnectionProperties("First Row Column Name") = False
        oConnection.ConnectionProperties("Column Names") = "c_csv"
        oConnection.ConnectionProperties("Number of Column") = 1
        oConnection.ConnectionProperties("Text Qualifier Col Mask: 0=no, 1=yes, e.g. 0101") = "1"
        oConnection.ConnectionProperties("Max characters per delimited column") = 8000
        oConnection.ConnectionProperties("Blob Col Mask: 0=no, 1=yes, e.g. 0101") = "0"
        
        oConnection.Name = "Connection 2"
        oConnection.ID = 2
        oConnection.Reusable = True
        oConnection.ConnectImmediate = False
        oConnection.DataSource = vFileName
        oConnection.ConnectionTimeout = 60
        oConnection.UseTrustedConnection = False
        oConnection.UseDSL = False
        
        'If you have a password for this connection, please uncomment and add your password below.
        'oConnection.Password = "<put the password here>"

goPackage.Connections.Add oConnection
Set oConnection = Nothing

'---------------------------------------------------------------------------
' create package steps information
'---------------------------------------------------------------------------

Dim oStep As DTS.Step2
Dim oPrecConstraint As DTS.PrecedenceConstraint

'------------- a new step defined below

Set oStep = goPackage.Steps.New

        oStep.Name = "Copy Data from Results to C:\Users\Nithin\Desktop\textcsv.csv Step"
        oStep.Description = "Copy Data from Results to C:\Users\Nithin\Desktop\textcsv.csv Step"
        oStep.ExecutionStatus = 1
        oStep.TaskName = "Copied data in table C:\Users\Nithin\Desktop\textcsv.csv"
        oStep.CommitSuccess = False
        oStep.RollbackFailure = False
        oStep.ScriptLanguage = "VBScript"
        oStep.AddGlobalVariables = True
        oStep.RelativePriority = 3
        oStep.CloseConnection = False
        oStep.ExecuteInMainThread = False
        oStep.IsPackageDSORowset = False
        oStep.JoinTransactionIfPresent = False
        oStep.DisableStep = False
        oStep.FailPackageOnError = False
        
goPackage.Steps.Add oStep
Set oStep = Nothing

'---------------------------------------------------------------------------
' create package tasks information
'---------------------------------------------------------------------------

'------------- call Task_Sub1 for task Copied data in table C:\Users\Nithin\Desktop\textcsv.csv (Copied data in table C:\Users\Nithin\Desktop\textcsv.csv)
Call Task_Sub1(goPackage, vFileName)

'---------------------------------------------------------------------------
' Save or execute package
'---------------------------------------------------------------------------

'goPackage.SaveToSQLServer "(local)", "sa", ""
goPackage.Execute
tracePackageError goPackage
goPackage.UnInitialize
'to save a package instead of executing it, comment out the executing package line above and uncomment the saving package line
Set goPackage = Nothing

Set goPackageOld = Nothing

End Sub


'-----------------------------------------------------------------------------
' error reporting using step.GetExecutionErrorInfo after execution
'-----------------------------------------------------------------------------
Public Sub tracePackageError(oPackage As DTS.Package)
Dim ErrorCode As Long
Dim ErrorSource As String
Dim ErrorDescription As String
Dim ErrorHelpFile As String
Dim ErrorHelpContext As Long
Dim ErrorIDofInterfaceWithError As String
Dim i As Integer

        For i = 1 To oPackage.Steps.Count
                If oPackage.Steps(i).ExecutionResult = DTSStepExecResult_Failure Then
                        oPackage.Steps(i).GetExecutionErrorInfo ErrorCode, ErrorSource, ErrorDescription, _
                                        ErrorHelpFile, ErrorHelpContext, ErrorIDofInterfaceWithError
                        MsgBox oPackage.Steps(i).Name & " failed" & vbCrLf & ErrorSource & vbCrLf & ErrorDescription
                End If
        Next i

End Sub

'------------- define Task_Sub1 for task Copied data in table C:\Users\Nithin\Desktop\textcsv.csv (Copied data in table C:\Users\Nithin\Desktop\textcsv.csv)
Public Sub Task_Sub1(ByVal goPackage As Object, ByVal vFileName As String)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask1 As DTS.DataPumpTask2
Set oTask = goPackage.Tasks.New("DTSDataPumpTask")
oTask.Name = "Copied data in table C:\Users\Nithin\Desktop\textcsv.csv"
Set oCustomTask1 = oTask.CustomTask

        oCustomTask1.Name = "Copied data in table C:\Users\Nithin\Desktop\textcsv.csv"
        oCustomTask1.Description = "Copied data in table C:\Users\Nithin\Desktop\textcsv.csv"
        oCustomTask1.SourceConnectionID = 1
        oCustomTask1.SourceSQLStatement = "Select c_csv From pr_export_csv"
        oCustomTask1.DestinationConnectionID = 2
        oCustomTask1.DestinationObjectName = vFileName
        oCustomTask1.ProgressRowCount = 1000
        oCustomTask1.MaximumErrorCount = 0
        oCustomTask1.FetchBufferSize = 1
        oCustomTask1.UseFastLoad = True
        oCustomTask1.InsertCommitSize = 0
        oCustomTask1.ExceptionFileColumnDelimiter = "|"
        oCustomTask1.ExceptionFileRowDelimiter = vbCrLf
        oCustomTask1.AllowIdentityInserts = False
        oCustomTask1.FirstRow = 0
        oCustomTask1.LastRow = 0
        oCustomTask1.FastLoadOptions = 2
        oCustomTask1.ExceptionFileOptions = 1
        oCustomTask1.DataPumpOptions = 0
        
Call oCustomTask1_Trans_Sub1(oCustomTask1)
                
                
goPackage.Tasks.Add oTask
Set oCustomTask1 = Nothing
Set oTask = Nothing

End Sub

Public Sub oCustomTask1_Trans_Sub1(ByVal oCustomTask1 As Object)

        Dim oTransformation As DTS.Transformation2
        Dim oTransProps As DTS.Properties
        Dim oColumn As DTS.Column
        Set oTransformation = oCustomTask1.Transformations.New("DTS.DataPumpTransformCopy")
                oTransformation.Name = "DirectCopyXform"
                oTransformation.TransformFlags = 63
                oTransformation.ForceSourceBlobsBuffered = 0
                oTransformation.ForceBlobsInMemory = False
                oTransformation.InMemoryBlobSize = 1048576
                oTransformation.TransformPhases = 4
                
                Set oColumn = oTransformation.SourceColumns.New("c_csv", 1)
                        oColumn.Name = "c_csv"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 104
                        oColumn.size = 7500
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("c_csv", 1)
                        oColumn.Name = "c_csv"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 104
                        oColumn.size = 7500
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

        Set oTransProps = oTransformation.TransformServerProperties

                
        Set oTransProps = Nothing

        oCustomTask1.Transformations.Add oTransformation
        Set oTransformation = Nothing

End Sub

