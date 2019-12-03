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
        <td><a href="#prereq">Link</a></td>
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
    <li></li>
</ul>
<h2>Breakdown</h2>
<h2>Credits</h2>