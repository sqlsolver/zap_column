# zap_column
Remove an unwanted column from multiple SharePoint libraries and lists.

The RemoveColumn.ps1 script is freely available for you to use and modify under the Microsoft Public License (Ms-PL) http://www.microsoft.com/en-us/openness/licenses.aspx. Use it at your own risk and with appropriate caution as it modifies SharePoint sites by removing the column you specify and will not protect you from removing the wrong one! It is assumed you will know how to retrieve values required to specify the parameters used by the script with PowerShell including the URL and ID of your site collection and the InternalName, DisplayName and ID(s) of the column you wish to remove. The script will not function without parameters that are correct for your environment added to the .XML parameters file.

<ul>
<li>Add a site node to the .XML parameters file for each site collection you wish to perform column audit or removal operations upon. </li>
<li>The script will create at least two files, a .log file that is transcript output and a .doc file that details the result of column operations. There is also output from the discovery function to a third .csv file (so that you can review the results in Excel) but it is created only if the column actually exists in your site. Each of the files are time-stamped when the operation begins, so it's easy to see which ones resulted from the same run of the RemoveColumn.ps1 script.</li>
<li>For more information type <b>Get-Help .\RemoveColumn.ps1</b> at the PowerShell prompt.</li>
</ul>
