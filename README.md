# Microsoft Outlook Graph API Facade for Watsonx Orchestrate

## Author:  Nigel T. Crowther
## Email: ncrowther@uk.ibm.com

## Overview
This is API exposes the Microsoft Outlook Graph API to build an agent in WXO.  

### Install Instructions

**Step 1:** Login to Microsoft Graph URL below using your IBM credentials. (Hint: Click user sign-in icon on top left)

https://developer.microsoft.com/en-us/graph/graph-explorer

**Step 2:** Press 'Run Query'. Go to Access Token tab (Hint: key icon) and copy it for the next step

**Step 3:** In WXO, enter Agent Builder.  Create an agent from scratch called "Outlook Agent".  Use the following description:

```
An assistant that manages personal outlook calendar entries
```

**Step 4:** Under behavior, enter the following instructions:
```
You are an assistant that manages personal outlook calendar entries.  Date time format should be: "2025-08-26T09:29:24.665Z"
```

**Step 5:** Under Toolset, select 'Add tool', then import [/api/openapi.yaml: ](./api/openapi.yaml).

**Step 6:** select all the API operations:

![Create Connection](/images/apiOperations.png)

**Step 7:** Press Next to enter the connections screen 

**Step 8:** Select 'Add a new connection'.  Enter the following and press 'Save and continue'

![Create Connection](/images/createOutlookConnection.png)

**Step 9:**  For the Draft environment, select an autentication type of BasicAuth and a server URL of www.ibm.com

**Step 10:** Select Team Credentials and press the Connect button

**Step 11:**  Enter a username of test, and the password is the token copied in step 2:

![Create Draft Connection](/images/createCredsDraft.png)

**Step 12:**  Repeat steps 9-11 for the Live environment.  Press Save. 

![Create Draft Connection](/images/createCredsLive.png)

 **Step 13:** You should see both Draft and Live environments with a green tick next to them:

![Connection established](/images/connectedCreds.png)

**Step 14:**  Test the agent: 

```
Show me my outlook meetings scheduled for this week.
```

![Output](/images/output.png)

**Step 15:**  If all works, deploy the agent

## Renewing expired access token

The Orchestrate token will expire every 15 minutes.  To renew it, do the following:

In the outlook connection, disconnect both DEV and PROD, then re -connect, following steps 8-13. Switching from basic auth to key/value can help reset.

# STEPS FOR RUNNING YOUR OWN SERVER

This section is for running the Outlook API on your own server

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

4.	Select the code engine project:  

    ibmcloud ce project select -n asc-watsonx

5.	Start Rancher Desktop as admin

    docker login -u ncrowthe -p [PASSWORD]

7.	Within this folder, edit CEbuild.sh and CErun.sh and change the REGISTRY to your Docker registry.

8.	Open the Ubuntu admin shell, type

    cd /mnt/e/WatsonOrchestrate/git/meetingApi 

10. Execute the following [if script does not run, execute each step individually, or: sed -i -e 's/\r$//' CEbuild]:

./CEbuild

9.	Within Bash shell, deploy the application to Code Engine on IBM Cloud. From the app's folder do:

./CErun

10.	Open the URL using the IBM Cloud Code Engine route for the application


