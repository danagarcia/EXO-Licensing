<h1>Exchange Online Licensing</h1>
<h2>Table of Contents</h2>
<table>
    <tr>
        <td><b>Description</b></td>
        <td><b>Link</b></td>
    </tr>
    <tr>
        <td>Summary</td>
        <td><a href="#summary">Link</a></td>
    </tr>
    <tr>
        <td>Pre-Requisites</td>
        <td><a href="#pre-requisites">Link</a></td>
    </tr>
    <tr>
        <td>Breakdown</td>
        <td><a href="#breakdown">Link</a></td>
    </tr>
    <tr>
        <td>Credits</td>
        <td><a href="#credits">Link</a></td>
    </tr>
</table>
<h2>Summary</h2>
<p>This script utilizes Azure Automation Runbooks and Graph API to detect mailboxes without licenses and license them. The script has built in failure detection to identify licensing issues. All processed mailboxes (successful or failed) are logged into a Power BI dataset which can be used to create a report (click <a href="#prereq-powerbi">here</a> for more info).</p>
<h2>Pre-Requisites</h2>
<h3>Azure Active Directory Application Registration</h3>
<p>Before we can configure the script to run via Azure Automation Runbook we need to register an application with Azure AD to pull and set settings via Graph API.</p>
<ul>
    <li>Sign into the <a href="https://portal.azure.com">Azure Portal</a></li>
    <li>Navigate to <b>Azure Active Directory</b> > <b>App Registration</b> using the navigation blade.</li>
    <li>Click <b>New registration</b>.<br /><img src="/Resources/Powerbi1_thumb1.jpg" /></li>
    <li>Provide a <b>Name</b> and click <b>Register</b>.<br /><img src="/Resources/Powerbi2_thumb1.jpg" /></li>
    <li>After the application is created navigate to <b>Manage</b> > <b>API permissions</b>.</li>
    <li>Click <b>Add a permission</b> > <b>Microsoft Graph</b>.<br /><img src="/Resources/Powerbi3_thumb1.jpg" /></li>
    <li>Select <b>Application permissions</b> at the next prompt.<br /><img src="/Resources/Powerbi4_thumb1.jpg" /></li>
    <li>Check <b>User</b> > <b>User.ReadWrite.All</b> and <b>MailboxSettings</b> > <b>MailboxSettings.Read</b></li>
    <li>Click <b>Add permissions</b>, the API permissions list should look like this now.<br /><img src="/Resources/Powerbi5_thumb1.jpg" /></li>
    <li>Click <b>Grant admin consent for...</b> and complete the authentication and consent dialog.<br /><img src="/Resources/Powerbi6_thumb1.jpg" /></li>
    <li>The API permissions list should look like this now.<br /><img src="/Resources/Powerbi7_thumb1.jpg" /></li>
</ul>
<h2>Breakdown</h2>
<h2>Credits</h2>