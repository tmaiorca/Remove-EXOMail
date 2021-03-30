# Remove-EXOMail
Email removal script using PowerShell via Exchange Online and Security &amp; Compliance Center

In order to quicken the turn-around time of searching and purging emails, the script will first connect to Exchange Online to run a Message Trace and store the resulting recipients in a variable. The variable is then passed into a KQL-formatted query after connecting to Security &amp; Compliance Center.
