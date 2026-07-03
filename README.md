# Excel-Nexus
Backend architechture for B2B custom Excel applications.  Includes a four part remote workbook updater,  error logging, and error telemetry. 


Project Overview

  This project is a deployment and maintenance framework built to manage VBA based enterprise tools. In production environments, deploying critical updates to client workbooks without interrupting workflow or corrupting active modules is a notorious issue. This system bridges that gap by establishing a live secure bridge between client side workbooks and a lightweight python staging server.
  Core capabilities include: 
  - Zero downtime hot swapping of core modules programmatically without triggering VBE file locks or runtime crashes.
  - Resillient infrastucture that uses dynamic local path resolution and late bindings when possible to ensure maximum compatibility across systems
  - Automated diagnostic telemetry that traps runtime exceptions , parses the environment state into a structured json and sends it directly to a designated administrative email for real time monitoring.




Remote updating system execution pipeline and explanation:

MasterUpdater:

  Starting off, a dialogue window is opened to choose a .bas file(s) (exported module file). The file is then loaded into a byte array and made into the body of an http post request. Its also using the http headers to carry over the filename and metadata so the server and ClientUpdator knows what its looking at. After everything is properly formatted the code opens up the connection via an http post method and sends the file to the web address where the flask server is hosted.

VBAflaskUpdater:  

   When first being spun up the python/flask server creates a  temporary virtual  directory to hold incoming files. It then waits to react depending on the method employed.  If it gets a post request it takes the data as-is and stores it in the directory. It uses the header where the file name is to save it as /"filename"
    If it recieves a get request, the server automatically triggers a download sending it directly from the directory to the get request source as-is.

ClientUpdater:

  When first triggered, the updater runs a check to see if the trust access to the object model library is checked. If not it gives a message with information on how to get to it. Ideally this would be done before hand on a call with a client. this is important because it allows for the code to programatically change itself. However this is a security risk. The "Melissa" virus and others like it were deployed this way before trust access was a thing. This works as an advantage between the seller and the client as it ensures trust and transparency. The client would be informed that as long as this box is unchecked, their software cannot be altered, mitigating potential malicious intent.
      
  After the initial check, the subroutine sends a get request to the server. It then recieves the body and headers. It uses the split method to separate both the name properties from the url and also the headers and pass one into filobj and setting the other one via the.name property. The fileobj.code is assigned the body or contens of the get response.  If succesfull it adds the fileobj to a collection called updates.

  Next using two nested for-each loops, each module in the workbook is cycled through and compared against each fileobj in the updates collection. If it finds a matching module it removes it. The for-each loops close. Then another is started, this time adding all fileobj in the updates collection. Finally, the macro hotkey for the bootloader subroutine is then set.

  ClsFileModule:
  A class object consisting of two string fields acting as a container holding two distinct data points from the http stream.

  Bootloader:
  
  Essentially identical to the ClientUpdater. The macro hotkey assignment is setting the hotkey for the ClientUpdater and vice-versa.  This setup is important as the AB setup solves the issue of trying to hotswap a module currently in use. if the ClientUpdater needs to be updated then the Bootloader is used. The inverse is also true. 

Logging and Telemtry explanation:

  ErrLogger:
  
   Taking variables passed through the error handler, this subroutine parses everything into a json string. these variables include system source, project name, subroutine, time, error type, error description, and line that it occured within the subroutine. Next, the system checks to see if a folder called logs exists as well as a file named errorlog.json.  If so, the error is logged and if not then they are created. This setup ensures the logging pipeline remains unbroken.

 Telemetry: 
 
   The subroutine starts out by determining where the log file is. Two possible paths are checked. This is done assuming the possibility a computer without the one drive directory that is now standard is used. Once the correct path is determined a cdo is configured and an email template is formatted with the error log attached. This setup requires an email that has access to password creation for applications.  


Set up

  MasterUpdater - Import to your personal workbook or whichever workbook you prefer as a terminal. Executed directly from the vba environment window.

Replace "webaddress" with the url or localhost where the flask server is hosted.(line73/340)

  VBAflaskupdater - This can deployed on any machine or server that supports Python and Flask (vscode for local server or glitch for example)

  ClientUpdater, Bootloader, Clsfilemodule, Errlogger, Telemetry - Import to end user/client workbook
 
in ClientUpdater/Bootloader line (38/190)remove the ' at fileurl = x and add replace with server address.

in Telemetry (lines 44,45,52,53). replace email and password with your email and application password



Important notes

1.The python script was not built with a way to delete files so the process must be terminated in order to perform a cleanup

2. As stated in the explanation of the ClientUpdater, it is important that the end user has "Trust access to the VBA project object model" checked for only the duration of the update.

3.Telemetry may not work for larger companies due to CDO triggering IT flags or if you can't generate an application password on your email platform.



  
