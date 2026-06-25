ProjectWise PowerShell Onboarding - Step by Step

Files kept in this folder
1. Invoke-PWUserOnboarding.ps1
2. update_groups_lists.ps1
3. echo_parameters.ps1
4. Jenkinsfile
5. projectwise-project-dictionary.json
6. readme.txt

Purpose
This folder is now PowerShell-only. Jenkins runs Invoke-PWUserOnboarding.ps1, reads the selected repository from projectwise-project-dictionary.json, and returns JSON output in both the Jenkins console and an archived artifact file called pw-result.json.

1. Where to set the repository details
Open projectwise-project-dictionary.json.

For each project entry, set these values:
1. projectCode: Short code used in Jenkins, for example WWHD.
2. repositoryId: ProjectWise repository GUID.
3. datasource: ProjectWise datasource name, for example arcadis-uk-pw.bentley.com:Arcadis-UK-19.
4. allowedUserGroups: The list of groups users are allowed to be added to.
5. allowedUserLists: The list of user lists users are allowed to be added to.

The PowerShell script uses these dictionary fields:
1. datasource
2. repositoryId
3. credentials.pwUserNameEnvVar
4. credentials.pwPasswordEnvVar
5. allowedUserGroups
6. allowedUserLists

Important
The Jenkinsfile already maps one shared Jenkins credential to all username and password environment variable names expected by the dictionary. Because of that, you do not need separate credentials per repository.

2. Where to set the username and password
Do not save the password in this repository.

In Jenkins:
1. Open Manage Jenkins.
2. Open Credentials.
3. Add a new Username with password credential.
4. Set Credentials ID to pw-adl-automation.
5. Set Username to _ADL_Automation.
6. Set Password to the password for the _ADL_Automation account.

What Jenkinsfile does with that credential
1. Reads pw-adl-automation.
2. Injects it as PW_USER and PW_PASSWORD.
3. Maps PW_USER and PW_PASSWORD to PW_WWHD_USER and PW_WWHD_PASSWORD used by the dictionary.

3. Where to set the Jenkins job to use this pipeline
In Jenkins:
1. Create or open the pipeline job.
2. In Pipeline Definition, use Pipeline script from SCM or paste the Jenkinsfile contents directly.
3. If using SCM, set Script Path to python_scripts/ADL/ADLPortal_PW_Onboarding/Jenkinsfile.
4. Save the job.

Recommended agent
The Jenkinsfile expects a Windows agent because it runs PowerShell and ProjectWise PowerShell cmdlets.

4. Parameters you can pass to the Jenkins job
The Jenkins job accepts these parameters:
1. PROJECT_CODE
2. USER_NAME
3. EMAIL
4. DESCRIPTION
5. IMS_USER
6. PASSWORD_OVERRIDE
7. USER_GROUPS
8. USER_LISTS
9. SECURITY_PROVIDER
10. INITIAL
11. TB_NAME
12. ORIGIN_CODE
13. ORIGINATOR
14. DISCIPLINE
15. GRADE
16. DISCIPLINE_CODE
17. GRADE_LEVEL
18. CALLBACK_URL
19. CALLBACK_TOKEN
20. CORRELATION_ID
21. UPDATE_GROUPS_LISTS
22. DRY_RUN

Optional pre-step: update groups and lists from ProjectWise
1. Set UPDATE_GROUPS_LISTS=true.
2. Before onboarding starts, Invoke-PWUserOnboarding.ps1 calls update_groups_lists.ps1.
3. update_groups_lists.ps1 logs into the selected datasource and queries groups and user lists.
4. It updates allowedUserGroups and allowedUserLists for that project in projectwise-project-dictionary.json.

How to enter group and list parameters
1. USER_GROUPS must be comma-separated.
Example: WWHD-Admin,WWHD-Design
2. USER_LISTS must be comma-separated.
Example: WWHD-Team,WWHD-External

5. How to run from Jenkins UI
1. Open the pipeline job.
2. Click Build with Parameters.
3. Fill PROJECT_CODE, USER_NAME, and EMAIL.
4. Set DRY_RUN to true for the first test.
5. Click Build.

Recommended first test
1. PROJECT_CODE = WWHD
2. USER_NAME = test.user
3. EMAIL = test.user@arcadis.com
4. UPDATE_GROUPS_LISTS = false
5. DRY_RUN = true

6. How to call the Jenkins job through the Jenkins API
Example URL pattern:
https://jenkins.arcadis.com/job/Projects/job/Mobility/job/UK/job/Rail/job/CurzonSt/job/CST_PW_NewUserCreation_Part1/buildWithParameters

Example parameters in a POST request:
1. PROJECT_CODE=WWHD
2. USER_NAME=test.user
3. EMAIL=test.user@arcadis.com
4. UPDATE_GROUPS_LISTS=false
5. DRY_RUN=true

Example curl command:
curl -X POST "https://jenkins.arcadis.com/job/Projects/job/Mobility/job/UK/job/Rail/job/CurzonSt/job/CST_PW_NewUserCreation_Part1/buildWithParameters" ^
    --user "your_jenkins_user:your_jenkins_api_token" ^
    --data-urlencode "PROJECT_CODE=WWHD" ^
    --data-urlencode "USER_NAME=test.user" ^
    --data-urlencode "EMAIL=test.user@arcadis.com" ^
    --data-urlencode "UPDATE_GROUPS_LISTS=false" ^
    --data-urlencode "DRY_RUN=true"

7. How the output is returned
Invoke-PWUserOnboarding.ps1 always writes a JSON result to standard output.

The Jenkinsfile captures that output and saves it to:
python_scripts/ADL/ADLPortal_PW_Onboarding/pw-result.json

This file is archived as a Jenkins build artifact.

Example success output:
{
    "timestampUtc": "2026-05-19T10:00:00Z",
    "correlationId": "...",
    "result": "success",
    "message": "Dry run successful. Validation passed.",
    "project": "WWHD",
    "repositoryId": "<repository-guid>",
    "user": {
        "userName": "test.user",
        "email": "test.user@arcadis.com",
        "created": false,
        "groups": [],
        "lists": []
    }
}

8. How to see the output and logs in Jenkins
After the build starts:
1. Open the build number.
2. Click Console Output to see all PowerShell logs and the JSON result.
3. Open Artifacts.
4. Download pw-result.json for the final structured response.

What you will see in Console Output
1. Jenkins stage logs.
2. PowerShell log lines from Invoke-PWUserOnboarding.ps1.
3. Final JSON payload.

9. What to check if the build fails
1. Confirm PROJECT_CODE exists in projectwise-project-dictionary.json.
2. Confirm datasource is correct.
3. Confirm repositoryId is correct.
4. Confirm allowedUserGroups and allowedUserLists contain the values you passed.
5. Confirm Jenkins credential pw-adl-automation exists.
6. Confirm username is _ADL_Automation.
7. Confirm the password is correct.
8. Confirm the Jenkins agent has the ProjectWise PowerShell cmdlets installed.

10. Local dry-run test from PowerShell
Run this from the folder:

$env:PW_WWHD_USER = '_ADL_Automation'
$env:PW_WWHD_PASSWORD = 'your-password'
pwsh -NoProfile -ExecutionPolicy Bypass -File .\Invoke-PWUserOnboarding.ps1 -Project WWHD -UserName test.user -Email test.user@arcadis.com -DryRun

Local run with optional group/list refresh
pwsh -NoProfile -ExecutionPolicy Bypass -File .\Invoke-PWUserOnboarding.ps1 -Project WWHD -UserName test.user -Email test.user@arcadis.com -UpdateGroupsLists -DryRun

11. Simple parameter echo test
If you only want to test parameter passing without touching ProjectWise, run:

pwsh -NoProfile -ExecutionPolicy Bypass -File .\echo_parameters.ps1 -Project WWHD -UserName test.user -Email test.user@arcadis.com -RequestId REQ-100 -Source Jenkins -AsJson

This returns the same values back in JSON so you can verify the Jenkins parameter wiring first.