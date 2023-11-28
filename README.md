### This project has been mothballed/abandoned after Microsoft implemented a much faster way of searching for and destroying emails. It is likely that things no longer work properly. 

# eDiscovery-Search-and-Destroy
Search and Destroy emails in your 0365 Enterprise tenant. 

  This project aims to allow for someone with eDiscovery Manager permissions within the 0365 Compliance Center to search for an email and then soft or hard delete the email.
  Included in the program is a Powershell script which executes a connection using WinRM to the Compliance Center and authenticates using the input of your email address.
  The program then allows for user input of certain criteria to create the Compliance Center Case, Search, and Search Action. 



It is important to note that while you may be utilizing this tool for good, it is extremely easy for someone to do real damage with this tool. As such, only use this tool
if you are qualified to do so and know what you are doing. Make sure that only people needing to perform Compliance and Security tasks in your organization have access both to the tool, and to the Compliance Center roles. 
