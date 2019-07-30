USE [master]
GO
/****** Object:  StoredProcedure [dbo].[sp_send_cdosysmail2]    Script Date: 07/30/2019 12:42:20 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/*******************************************************************
Purpose:   Sends an SMTP email. 
 
Overview: Uses CDOSys . 
References to the CDOSYS objects are at the following MSDN Web site:
http://msdn.microsoft.com/library/default.asp?url=/library/en-us/cdosys/html/_cdosys_messaging.asp 
 
USAGE EXAMPLE
exec sp_send_cdosysmail2  @From ='adam.bartlett@johnsondiversey.com' , 
@To ='adam.bartlett@edsgfs.co.uk' , @Subject = 'Test', 
@Body = 'Though I speak with the tounges of Angles and Men, if I have not love ...',
@MailServer = 'dle-001.dc.u1973.unilever.com'
 
Inputs:  Note that @Subject, @Body , @emailtype and @mailserver have default values.
 
Affects: Makes no changes to the database.
 
Output:  None
 
Modification History:
 Created:  Microsoft, 1998   (So this sproc is excluded from the GNU liscence for this repositry, as it wasn't created entirely by me & I'm not 100% sure what the copyright status is,  but MS do seem to have released this to public domain.)
 Modified: Adam Bartlett 10 Nov 2004, minor modifications

* PRE-CREATION SETUP  Please read!

For this to work,  you need to:

1) In  the input parameters, change the default value for @MailServer  so it matches the hostname or IP address of your SMTP relay 
2) On the SMTP relay, you'll almost certianly have to set it to allow relaying from the IP address of your SQL server.

Optionally, set a sensible default value for the @From input parameter.

*******************************************************************/
 
ALTER PROCEDURE [dbo].[sp_send_cdosysmail2] 
   @From varchar(100)= 'changeME@efault.com' ,
   @To varchar(300) ,
   @CC varchar(100) = 'no',  
   @BCC varchar(100) = 'no', 
   @Subject varchar(200)=" ",
   @Body varchar(4000) =" ",
   @emailtype char (4) = 'text' ,
   @MailServer nvarchar (200) = "SMTPrelayHostnameorIP"

 
   AS
   Declare @iMsg int
   Declare @hr int
   Declare @source varchar(255)
   Declare @description varchar(500)
   Declare @output varchar(1000)
 
--************* Create the CDO.Message Object ************************
   EXEC @hr = sp_OACreate 'CDO.Message', @iMsg OUT
 
--***************Configuring the Message Object ******************
-- This is to configure a remote SMTP server.
-- http://msdn.microsoft.com/library/default.asp?url=/library/en-us/cdosys/html/_cdosys_schema_configuration_sendusing.asp
   EXEC @hr = sp_OASetProperty @iMsg, 'Configuration.fields("http://schemas.microsoft.com/cdo/configuration/sendusing").Value','2'
-- This is to configure the Server Name or IP address. 
-- Replace MailServerName by the name or IP of your SMTP Server.
   EXEC @hr = sp_OASetProperty @iMsg, 'Configuration.fields("http://schemas.microsoft.com/cdo/configuration/smtpserver").Value', @MailServer 
 
-- Save the configurations to the message object.
   EXEC @hr = sp_OAMethod @iMsg, 'Configuration.Fields.Update', null
 
-- Set the e-mail parameters.
   EXEC @hr = sp_OASetProperty @iMsg, 'To', @To
   EXEC @hr = sp_OASetProperty @iMsg, 'From', @From
   EXEC @hr = sp_OASetProperty @iMsg, 'Subject', @Subject
   
   if @CC <> 'no'
   EXEC @hr = sp_OASetProperty @iMsg, 'CC', @CC
   
   if @BCC <> 'no'
   EXEC @hr = sp_OASetProperty @iMsg, 'BCC',@BCC


   
 
-- If you require HTML e-mail, use 'HTMLBody' instead of 'TextBody'.
 if @emailtype = 'text'
   EXEC @hr = sp_OASetProperty @iMsg, 'TextBody', @Body
 else 
   EXEC @hr = sp_OASetProperty @iMsg, 'HTMLBody', @Body
 
   EXEC @hr = sp_OAMethod @iMsg, 'Send', NULL
 
-- Sample error handling.
   IF @hr <>0 
     select @hr
     BEGIN
       EXEC @hr = sp_OAGetErrorInfo NULL, @source OUT, @description OUT
       IF @hr = 0
         BEGIN
           SELECT @output = '  Source: ' + @source
           PRINT  @output
           SELECT @output = '  Description: ' + @description
           PRINT  @output
         END
       ELSE
         BEGIN
           PRINT '  sp_OAGetErrorInfo failed.'
           RETURN
         END
     END
 
-- Do some error handling after each step if you have to.
-- Clean up the objects created.
   EXEC @hr = sp_OADestroy @iMsg
