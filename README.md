# Microsoft Outlook Graph API Facade for Watsonx Orchestrate

## Overview
This is API exposes the Microsoft Outlook Graph API as an agent in WXO.  NOTE: this is for demo purposes only

### Install Instructions

**Step 1:** Login to Microsoft Graph URL below using your IBM credentials. (Hint: Click user sign-in icon on top left)

https://developer.microsoft.com/en-us/graph/graph-explorer

**Step 2:** Press 'Run Query'. Go to Access Token tab (Hint: key icon) and copy it for the next step:

**Step 3:** In WXO, enter Agent Builder.  Create an agent from scratch called "Outlook Agent".  Use the following description:

```
An assistant that manages personal outlook calendar entries
```

**Step 4:** Under behavior, enter the following instructions:
```
You are an assistant that manages personal outlook calendar entries.  Date time format should be: "2025-08-26T09:29:24.665Z"
```

**Step 5:** Under Toolset, select 'Add tool', then import [/api/openapi.yaml: ](./api/openapi.yaml).

**Step 6:** select the following API operations:

![Create Connection](/images/apiOperations.png)

**Step 7:** Press Next to enter the connections screen 

**Step 8:** Select 'Add a new connection'.  Enter the following and then press 'Save and continue'

![Create Connection](/images/createOutlookConnection.png)

**Step 9:**  For the draft environment, select an autentication type of BasicAuth and a server URL of www.ibm.com

**Step 10:** Select Team Credentials and press the Connect button

**Step 11:**  Enter a username of test, and the password is the access token copied in step 2:

![Create Draft Connection](/images/createCredsDraft.png)

**Step 12:**  Repeat the step above for the Live environment.  Press Save. 

![Create Draft Connection](/images/createCredsLive.png)

 **Step 13:** You should see both Draft and Live environments with a green tick next to them:

![Connection established](/images/connectedCreds.png)

**Step 14:**  Test the agent with the following: 

```
Show me my outlook meetings scheduled for this week.
```

![Output](/images/output.png)

**Step 15:**  If all works, deploy the agent

### Running the server locally

To run the server, run:

```
npm start
```

To view the Swagger UI interface:

```
open http://localhost:8080/docs
```


### Deployment to code engine on IBM Cloud

1.	Open Git Bash shell from VSC

2.	Login to IBM Cloud.

    ibmcloud login --sso

3.	In the IBM Cloud console, go to Manage > Account > Account resources > Resource groups.  Select the resource group for Code Engine. E.g. default

    ibmcloud target -g asc_watsonx

4.	Select the code engine project (e.g. NTC):  

    ibmcloud ce project select -n asc-watsonx

5.	Start Rancher Desktop as admin

    docker login -u ncrowthe -p C****!

7.	Within this folder, edit CEbuild.sh and CErun.sh and change the REGISTRY to your Docker registry.

8.	Open the Ubuntu admin shell, type

    cd /mnt/e/WatsonOrchestrate/git/meetingApi 

10. Execute the following [if script does not run, execute each step individually, or: sed -i -e 's/\r$//' CEbuild]:

./CEbuild

9.	Within Bash shell, deploy the application to Code Engine on IBM Cloud. From the app's folder do:

./CErun

10.	Open the URL using the IBM Cloud Code Engine route for the application



## Renewing expired access token

The Orchestrate token will expire every 15 minutes or so.  To renew it, do the following:

In the connection for wxo, disconnect both DEV and PROD, then re -connect. Switching from basic auth to key/value can help

