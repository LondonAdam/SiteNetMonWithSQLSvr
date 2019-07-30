*** Remote site or VPN network monitoring With SQL server ***

Say you have remote site(s) with an unreliable network connection – whether that be due to a challenging ISP,  or perhaps a VPN that is sometimes overloaded.
A continuous ping to an IP at the remote site is a good way to verify connectivity. You find the times when the issue becomes user effecting  matches the times when you get multiple ‘time outs’ on the ping.   So you want a quick & free way to get email alerts when this happens.
If you have an SQL server and an SMTP relay, this code should do the job.  

If you don't have an SMTP relay, but do have an office 365 email,  you can instead use the code in 
https://github.com/LondonAdam/SiteNetMonWithSQLSvrO365email


*** Warning about xp_cmdshell ***
The stored procedure used here uses xp_cmdshell, and if it is currently disabled on server, it will temporarily enable it. If you have other code running on the same server doing this temporary enablement & disablement, it may cause issues if two different processes run at about the same time. Running sp_configure in core hours on a busy production server may itself cause issues, though I've never seen this happen.
This has never been the case at anywhere I've worked, but I've seen on the web that a few people think it's a major security risk to enable xp_cmdshell even temporarily.


*** Set up needed  every time you set these alerts for a new site.
1)	Copy CreateSTtoredProc.SQL into SQLstudio
2)	Carry out the 3 steps in the ‘PRE-CREATION SETUP’  sub section, in the commented out block at the top of the SP.
3)	Create the SP.
4)	If this is the first time you have set up one of these email alerts, create a new SQL agent job called pingMonitoring,  schedule it to run every 5 mins, and have it call the Stored Procedure you have just created.
5)  If you don't already have sp_send_cdosysmail2 in your master DB,  copy CreateSp_send_cdosysmail2.SQL into SQL studio, carry out the 2 steps in the ‘PRE-CREATION SETUP’  sub section near the top, and run the script to create the SP.

If you already have an SQL job for the ping monitoring,  don’t create a new one,  instead add the new Stored Procedure to the existing job, so that the sprocs for different sits run sequentially, i.e.  ‘step 1’ of your pingMonitoring job might read:

exec usp_CalgaryMon
exec usp_CapeTownMon
exec usp_10VPNMon
