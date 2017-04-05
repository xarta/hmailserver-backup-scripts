' Todo: move all this to a json file instead!
Function StartStopServiceError(errnum)
    
    Dim errDescription

    Select Case errNum
    Case 0
        errDescription = "The request was accepted"
    Case 1
        errDescription = "The request is not supported"
    Case 2
        errDescription = "The user did not have the necessary access"
    Case 3
        errDescription = "The service cannot be stopped because other services that are running are dependent on it"
    Case 4
        errDescription = "The requested control code is not valid, or it is unacceptable to the service"
    Case 5
        errDescription = "The requested control code cannot be sent to the service because the state of the service (Win32_BaseService.State property) is equal to 0, 1, or 2."
    Case 6
        errDescription = "The service has not been started"
    Case 7
        errDescription = "The service did not respond to the start request in a timely fashion"
    Case 8
        errDescription = "Unknown failure when starting the service"
    Case 9
        errDescription = "The directory path to the service executable file was not found"
    Case 10
        errDescription = "The service is already running"
    Case 11
        errDescription = "The database to add a new service is locked"
    Case 12
        errDescription = "A dependency this service relies on has been removed from the system"
    Case 13
        errDescription = "The service failed to find the service needed from a dependent service"
    Case 14
        errDescription = "The service has been disabled from the system"
    Case 15
        errDescription = "The service does not have the correct authentication to run on the system"
    Case 16
        errDescription = "This service is being removed from the system"
    Case 17
        errDescription = "The service has no execution thread"
    Case 18
        errDescription = "The service has circular dependencies when it starts"
    Case 19
        errDescription = "A service is running under the same name"
    Case 20
        errDescription = "The service name has invalid characters"
    Case 21
        errDescription = "Invalid parameters have been passed to the service"
    Case 22
        errDescription = "The account under which this service runs is either invalid or lacks the permissions to run the service"
    Case 23
        errDescription = "The service exists in the database of services available from the system"
    Case 24
        errDescription = "The service is currently paused in the system"
    Case Else
        errDescription = "Not recognised"
    End Select

    StartStopServiceError = errDescription

End Function


Function ChangeServiceUserError(errnum)
    
    Dim errDescription

    Select Case errNum
    Case 0
        errDescription = "None. (Success)"
    Case 1
        errDescription = "Not Supported"
    Case 2
        errDescription = "Access Denied"
    Case 3
        errDescription = "Dependent Services Running"
    Case 4
        errDescription = "Invalid Service Control"
    Case 5
        errDescription = "Service Cannot Accept Control"
    Case 6
        errDescription = "Service Not Active"
    Case 7
        errDescription = "Service Request Timeout"
    Case 8
        errDescription = "Unknown Failure"
    Case 9
        errDescription = "Path Not Found"
    Case 10
        errDescription = "Service Already Running"
    Case 11
        errDescription = "Service Database Locked"
    Case 12
        errDescription = "Service Dependency Deleted"
    Case 13
        errDescription = "Service Dependency Failure"
    Case 14
        errDescription = "Service Disabled"
    Case 15
        errDescription = "Service Logon Failed"
    Case 16
        errDescription = "Serivce Marked For Deletion"
    Case 17
        errDescription = "Serivce No Thread"
    Case 18
        errDescription = "Status Circular Dependency"
    Case 19
        errDescription = "Status Duplicate Name"
    Case 20
        errDescription = "Status Invalid Name"
    Case 21
        errDescription = "Status Invalid Parameter"
    Case 22
        errDescription = "Status Invalid Service Account"
    Case 23
        errDescription = "Status Service Exists"
    Case 24
        errDescription = "Service Already Paused"
    Case Else
        errDescription = "Not recognised"
    End Select

    ChangeServiceUserError = errDescription

End Function