Imports System.Windows.Forms
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Linq
Imports System.Collections.Generic
Imports System


'   LockTools.vb
'
'   How to know the process locking a file
'   http://tinyurl.com/jgp57p6
'   Lizenziert unter der Apache 2.0-Lizenz ("Lizenz");
'   Sie dürfen diese Datei nur im Rahmen der Lizenzbedingungen verwenden.
'   Ein Exemplar der Lizenz erhalten Sie unter
'   http://www.apache.org/licenses/LICENSE-2.0
'   Downloaded from code.msdn.microsoft.com and converted from C# to VB.net
'
Public Class LockTools
		' A system restart is not required.
		Private Const RmRebootReasonNone As Integer = 0
		' maximum character count of application friendly name.
		Private Const CCH_RM_MAX_APP_NAME As Integer = 255
		' maximum character count of service short name.
		Private Const CCH_RM_MAX_SVC_NAME As Integer = 63
    '		Private Delegate Sub AddTreeNode(node As TreeNode)

    ''' <summary>
    ''' Uniquely identifies a process by its PID and the time the process began. 
    ''' An array of RM_UNIQUE_PROCESS structures can be passed
    ''' to the RmRegisterResources function.
    ''' </summary>
    <StructLayout(LayoutKind.Sequential)> _
		Structure RM_UNIQUE_PROCESS
			' The product identifier (PID).
			Public dwProcessId As Integer
			' The creation time of the process.
			Public ProcessStartTime As System.Runtime.InteropServices.ComTypes.FILETIME
		End Structure

		''' <summary>
		''' Describes an application that is to be registered with the Restart Manager.
		''' </summary>
		<StructLayout(LayoutKind.Sequential, CharSet := CharSet.Auto)> _
		Structure RM_PROCESS_INFO
			' Contains an RM_UNIQUE_PROCESS structure that uniquely identifies the
			' application by its PID and the time the process began.
			Public Process As RM_UNIQUE_PROCESS
			' If the process is a service, this parameter returns the 
			' long name for the service.
			<MarshalAs(UnmanagedType.ByValTStr, SizeConst := CCH_RM_MAX_APP_NAME + 1)> _
			Public strAppName As String
			' If the process is a service, this is the short name for the service.
			<MarshalAs(UnmanagedType.ByValTStr, SizeConst := CCH_RM_MAX_SVC_NAME + 1)> _
			Public strServiceShortName As String
			' Contains an RM_APP_TYPE enumeration value.
			Public ApplicationType As RM_APP_TYPE
			' Contains a bit mask that describes the current status of the application.
			Public AppStatus As UInteger
			' Contains the Terminal Services session ID of the process.
			Public TSSessionId As UInteger
			' TRUE if the application can be restarted by the 
			' Restart Manager; otherwise, FALSE.
			<MarshalAs(UnmanagedType.Bool)> _
			Public bRestartable As Boolean
		End Structure

		''' <summary>
		''' Specifies the type of application that is described by
		''' the RM_PROCESS_INFO structure.
		''' </summary>
		Enum RM_APP_TYPE
			' The application cannot be classified as any other type.
			RmUnknownApp = 0
			' A Windows application run as a stand-alone process that
			' displays a top-level window.
			RmMainWindow = 1
			' A Windows application that does not run as a stand-alone
			' process and does not display a top-level window.
			RmOtherWindow = 2
			' The application is a Windows service.
			RmService = 3
			' The application is Windows Explorer.
			RmExplorer = 4
			' The application is a stand-alone console application.
			RmConsole = 5
			' A system restart is required to complete the installation because
			' a process cannot be shut down.
			RmCritical = 1000
		End Enum

		''' <summary>
		''' Registers resources to a Restart Manager session. The Restart Manager uses 
		''' the list of resources registered with the session to determine which 
		''' applications and services must be shut down and restarted. Resources can be 
		''' identified by filenames, service short names, or RM_UNIQUE_PROCESS structures
		''' that describe running applications.
		''' </summary>
		''' <param name="pSessionHandle">
		''' A handle to an existing Restart Manager session.
		''' </param>
		''' <param name="nFiles">The number of files being registered</param>
		''' <param name="rgsFilenames">
		''' An array of null-terminated strings of full filename paths.
		''' </param>
		''' <param name="nApplications">The number of processes being registered</param>
		''' <param name="rgApplications">An array of RM_UNIQUE_PROCESS structures</param>
		''' <param name="nServices">The number of services to be registered</param>
		''' <param name="rgsServiceNames">
		''' An array of null-terminated strings of service short names.
		''' </param>
		''' <returns>The function can return one of the system error codes that 
		''' are defined in Winerror.h
		''' </returns>
		<DllImport("rstrtmgr.dll", CharSet := CharSet.Auto, SetLastError := True)> _
		Shared Function RmRegisterResources(pSessionHandle As UInteger, nFiles As UInt32, rgsFilenames As String(), nApplications As UInt32, <[In]> rgApplications As RM_UNIQUE_PROCESS(), nServices As UInt32, _
			rgsServiceNames As String()) As Integer
		End Function

		''' <summary>
		''' Starts a new Restart Manager session. A maximum of 64 Restart Manager 
		''' sessions per user session can be open on the system at the same time. 
		''' When this function starts a session, it returns a session handle and 
		''' session key that can be used in subsequent calls to the Restart Manager API.
		''' </summary>
		''' <param name="pSessionHandle">
		''' A pointer to the handle of a Restart Manager session.
		''' </param>
		''' <param name="dwSessionFlags">Reserved. This parameter should be 0.</param>
		''' <param name="strSessionKey">
		''' A null-terminated string that contains the session key to the new session.
		''' </param>
		''' <returns></returns>
		<DllImport("rstrtmgr.dll", CharSet := CharSet.Auto, SetLastError := True)> _
		Shared Function RmStartSession(ByRef pSessionHandle As UInteger, dwSessionFlags As Integer, strSessionKey As String) As Integer
		End Function

		''' <summary>
		''' Ends the Restart Manager session. This function should be called by the 
		''' primary installer that has previously started the session by calling the 
		''' RmStartSession function. The RmEndSession function can be called by a 
		''' secondary installer that is joined to the session once no more resources 
		''' need to be registered by the secondary installer.
		''' </summary>
		''' <param name="pSessionHandle">
		''' A handle to an existing Restart Manager session.
		''' </param>
		''' <returns>
		''' The function can return one of the system error codes
		''' that are defined in Winerror.h.
		''' </returns>
		<DllImport("rstrtmgr.dll", CharSet := CharSet.Auto, SetLastError := True)> _
		Shared Function RmEndSession(pSessionHandle As UInteger) As Integer
		End Function

		''' <summary>
		''' Gets a list of all applications and services that are currently using 
		''' resources that have been registered with the Restart Manager session.
		''' </summary>
		''' <param name="dwSessionHandle">
		''' A handle to an existing Restart Manager session.
		''' </param>
		''' <param name="pnProcInfoNeeded">A pointer to an array size necessary to 
		''' receive RM_PROCESS_INFO structures required to return information for 
		''' all affected applications and services.
		''' </param>
		''' <param name="pnProcInfo">
		''' A pointer to the total number of RM_PROCESS_INFO structures in an array
		''' and number of structures filled.
		''' </param>
		''' <param name="rgAffectedApps">
		''' An array of RM_PROCESS_INFO structures that list the applications and 
		''' services using resources that have been registered with the session.
		''' </param>
		''' <param name="lpdwRebootReasons">
		''' Pointer to location that receives a value of the RM_REBOOT_REASON
		''' enumeration that describes the reason a system restart is needed.
		''' </param>
		''' <returns></returns>
		<DllImport("rstrtmgr.dll", CharSet := CharSet.Auto, SetLastError := True)> _
		Shared Function RmGetList(dwSessionHandle As UInteger, ByRef pnProcInfoNeeded As UInteger, ByRef pnProcInfo As UInteger, <[In], Out> rgAffectedApps As RM_PROCESS_INFO(), ByRef lpdwRebootReasons As UInteger) As Integer
		End Function

		' Return a list of processes that have locks on a file.
		Public Shared Function FindLockers(filename As String) As List(Of Process)
			' Start a new Restart Manager session.
			Dim session_handle As UInteger
			Dim session_key As String = Guid.NewGuid().ToString()
			Dim result As Integer = RmStartSession(session_handle, 0, session_key)
			If result <> 0 Then
				Throw New Exception("Error " + result + " starting a Restart Manager session.")
			End If

			Dim processes As New List(Of Process)()
			Try
				Const  ERROR_MORE_DATA As Integer = 234
				Dim pnProcInfoNeeded As UInteger = 0, num_procs As UInteger = 0, lpdwRebootReasons As UInteger = RmRebootReasonNone
				Dim resources As String() = New String() {filename}
				result = RmRegisterResources(session_handle, CType(resources.Length, UInteger), resources, 0, Nothing, 0, _
					Nothing)
				If result <> 0 Then
					Throw New Exception("Could not register resource.")
				End If

				' There's a race around condition here. The first call to RmGetList()
				' returns the total number of process. However, when we call RmGetList()
				' again to get the actual processes this number may have increased.
				result = RmGetList(session_handle, pnProcInfoNeeded, num_procs, Nothing, lpdwRebootReasons)
				If result = ERROR_MORE_DATA Then
					' Create an array to store the process results.
					Dim processInfo As RM_PROCESS_INFO() = New RM_PROCESS_INFO(pnProcInfoNeeded) {}
					num_procs = pnProcInfoNeeded

					' Get the list.
					result = RmGetList(session_handle, pnProcInfoNeeded, num_procs, processInfo, lpdwRebootReasons)
					If result <> 0 Then
						Throw New Exception("Error " + result + " listing lock processes")
					End If

					' Add the results to the list.
					Dim i As Integer = 0
					While i < num_procs
						Try
							processes.Add(Process.GetProcessById(processInfo(i).Process.dwProcessId))
						' Catch the error in case the process is no longer running.
						Catch generatedExceptionName As ArgumentException
						End Try
						i += 1
					End While
				ElseIf result <> 0 Then
					Throw New Exception("Error " + result + " getting the size of the result.")
				End If
			Catch ex As Exception
            '   MessageBox.Show(ex.Message)
            MsgBox(ex.Message)
        Finally
            RmEndSession(session_handle)
			End Try

			Return processes
		End Function
	End Class
