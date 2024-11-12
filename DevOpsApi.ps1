# @File: DevOpsApi.ps1
# @Author: Simone Medas
# @CreatedOn: 2024-11-12
#
# Helper class that wraps and simplifies Azure DevOps API endpoints.
#

class RestClient {
    [string] hidden $url
    [AzureDevOpsClient] hidden $azureDevOpsClient

    RestClient([string]$urlTemplate, [AzureDevOpsClient] $azureDevOpsClient) {
        $this.azureDevOpsClient = $azureDevOpsClient;
        $this.url = $this.ComposeUrl($urlTemplate);
    }

    #prepares header for Rest API calls
    [hashtable] GetInitialHeaders() {
        $base64Auth = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes(":$($this.azureDevOpsClient.PAT)"))
        return @{
            "Authorization" = "Basic $($base64Auth)"            
            "Content-Type"  = "application/json"
        }
    }

    [string] Normalize([string]$uriSegment) {
        return $uriSegment.Replace(" ", "%20")
    }

    [string] ComposeUrl([string]$src) {
        $src = $src.Replace("https://", "$($this.azureDevOpsClient.Protocol)://").
        Replace("{{instance}}", $this.azureDevOpsClient.Instance).
        Replace("{{organization}}", $this.Normalize($this.azureDevOpsClient.Organization)).
        Replace("{{project-name}}", $this.Normalize($this.azureDevOpsClient.ProjectName)).
        Replace("{{team}}", $this.Normalize($this.azureDevOpsClient.Team)).
        Replace("{{api-version}}", $this.Normalize($this.azureDevOpsClient.ApiVersion))

        return $src;
    }

    [void] SetPlaceholder([string] $plaeholder, [string] $withValue) {
        $this.url = $this.url.Replace(":$($plaeholder)", $withValue)
    }

    [PSObject] Get() {
        $patHeader = $this.GetInitialHeaders();
        return Invoke-RestMethod -Uri $this.url -Method Get -Headers $patHeader
    }

    [PSObject] Post([string] $payload) {
        $patHeader = $this.GetInitialHeaders();
        $result = Invoke-RestMethod -Uri $this.url -Method Post -Headers $patHeader -Body $payload;

        return $result
    }

    [PSObject] Patch([string] $payload) {
        $patHeader = $this.GetInitialHeaders();
        $patHeader["Content-Type"] = "application/json-patch+json";
        $result = Invoke-RestMethod -Uri $this.url -Method Patch -Headers $patHeader -Body $payload;

        return $result
    }

    # use PATCH verb but actually payload is specified as normal json
    [PSObject] PatchButPost([string] $payload) {
        $patHeader = $this.GetInitialHeaders();
        $patHeader["Content-Type"] = "application/json";
        $result = Invoke-RestMethod -Uri $this.url -Method Patch -Headers $patHeader -Body $payload;

        return $result
    }

    [PSObject] Delete() {
        $patHeader = $this.GetInitialHeaders();
        $result = Invoke-RestMethod -Uri $this.url -Method Delete -Headers $patHeader;

        return $result
    }
}

class TaskId {
    [int] $Id;
    [string] $Url;

    TaskId([string]$id, [string]$url) {
        $this.Id = $id;
        $this.Url = $url;
    }
}

class  TestSubResult {
    [int] $id;
    [string] $displayName;
    [string] $errorMessage;
    [string] $stackTrace;
    [string] $outcome;
}

enum TestResultState {
    NotStarted
    InProgress
    Completed
    Aborted
    Waiting
}

enum TestResultOutcome {
    Unspecified
    None
    Passed
    Failed
    Inconclusive
    Timeout
    Aborted
    Blocked
    NotExecuted
    Warning
    Error
    NotApplicable
    Paused
    InProgress
    NotImpacted
}

enum TestResultGroupType {
    none            #Leaf node of test result.
    dataDriven      #Hierarchy type of test result.
    generic         #Unknown hierarchy type.
    orderedTest     #Hierarchy type of test result.
    rerun           #Hierarchy type of test result.
}

class AzureDevOpsClient {
    [string] $Protocol = "https";
    [string] hidden $ApiVersion = "6.0";
    [string] $Instance = "dev.azure.com";
    [string] $Organization;
    [string] $PAT;
    [string] $ProjectName;
    [string] $Team;

    [AzureDevOpsClient] static Create($configurationFile) {
        $client = [AzureDevOpsClient]::new();
        $client.Protocol = $configurationFile.Protocol;
        $client.ApiVersion = $configurationFile.ApiVersion;
        $client.Instance = $configurationFile.Instance;
        $client.Organization = $configurationFile.Organization;
        $client.PAT = $configurationFile.PAT;
        $client.ProjectName = $configurationFile.ProjectName;
        $client.Team = $configurationFile.Team;

        return $client;
    }

    #region Test Plan/Suites

    # get TestPlan info
    [PSObject] GetTestPlan([string] $testPlanId) {        
        $urlTemplate = "https://{{instance}}/{{organization}}/{{project-name}}/_apis/testplan/plans/:planId?api-version={{api-version}}";

        $restClient = [RestClient]::new($urlTemplate, $this);
        $restClient.SetPlaceholder("planId", $testPlanId);

        return $restClient.Get();
    }

    # get the whole test suites tree from given Test Plan Id
    [PSObject] GetTestSuitesByTestPlanId([string] $testPlanId) {
        $urlTemplate = "https://{{instance}}/{{organization}}/{{project-name}}/_apis/testplan/Plans/:planId/suites?api-version={{api-version}}";

        $restClient = [RestClient]::new($urlTemplate, $this);
        $restClient.SetPlaceholder("planId", $testPlanId);

        return $restClient.Get();
    }

    # get the Test Suite from given parent test suite Id, from given test plan tree.
    [PSObject] GetTestSuiteByParentSuiteId([PSObject] $testPlanTree, [string] $parentSuiteId = $null, [string] $name) {
        
        # when parent was't given the return the first test suite
        if(-not $parentSuiteId -and $testPlanTree.value.Count -gt 0) {
            $id = $testPlanTree.value.Item(0).id;
            return $this.GetTestSuiteByParentSuiteId($testPlanTree, $id, $name);
        }

        return $testPlanTree.value | Where-Object { $_.parentSuite.id -eq $parentSuiteId -and $_.name.Replace(' ', '') -eq $name.Replace(' ', '') } | Select-Object -First 1;
    }    

    [PSObject] GetTestSuiteByName([string] $testPlanId, [string] $name) {
        $testSuites = $this.GetTestSuitesByTestPlanId($testPlanId);
        foreach ($tsuite in $testSuites.value) {
            if ($tsuite.name.Replace(' ', '') -eq $name.Replace(' ', '')) {
                return $tsuite;
                break;
            }
        }

        return $null;
    }

    [PSObject] GetTestCases([string] $testPlanId, [string] $testSuiteId) {
        $urlTemplate = "https://{{instance}}/{{organization}}/{{project-name}}/_apis/testplan/Plans/:planId/Suites/:suiteId/TestCase?api-version={{api-version}}"

        $restClient = [RestClient]::new($urlTemplate, $this);
        $restClient.SetPlaceholder("planId", $testPlanId);
        $restClient.SetPlaceholder("suiteId", $testSuiteId);

        return $restClient.Get();
    }

    [TaskId] CreateTestCaseWorkItem([string] $title, [string] $description) {        

        $urlTemplate = "https://{{instance}}/{{organization}}/{{project-name}}/_apis/wit/workitems/:type?api-version={{api-version}}";

        $restClient = [RestClient]::new($urlTemplate, $this)
        $restClient.SetPlaceholder("type", '$Test Case')

        $payload = @(
            @{
                op    = "add"
                path  = "/fields/System.Title"
                value = $title
            },
            @{
                op    = "add"
                path  = "/fields/System.Description"
                value = $description
            },

            @{
                op    = "add"
                path  = "/fields/Microsoft.VSTS.TCM.AutomatedTestName"
                value = $title
            },
            @{
                op    = "add"
                path  = "/fields/Microsoft.VSTS.TCM.AutomatedTestStorage"
                value = $title
            },
            @{
                op    = "add"
                path  = "/fields/Microsoft.VSTS.TCM.AutomatedTestType"
                value = "Postman Test Case"
            },
            @{
                op    = "add"
                path  = "/fields/Microsoft.VSTS.TCM.AutomatedTestId"
                value = "$([guid]::NewGuid())"
            }
        )

        $payload = $payload | ConvertTo-Json -Depth 10
        $httpResult = $restClient.Patch($payload)

        $taskId = $null;
        if ($httpResult) {
            $taskId = [TaskId]::new($httpResult.id, $httpResult.url)
        }

        return $taskId
    }

    # Create a new Test Case in to specific TestSuite
    # return the TestCase id.
    [psobject] CreateTestCase([string] $testPlanId, [string] $testSuiteId, [string] $title, [string] $description) {

        [TaskId] $workItem = $this.CreateTestCaseWorkItem($title, $description);

        if ($null -ne $workItem) {

            $urlTemplate = "https://{{instance}}/{{organization}}/{{project-name}}/_apis/testplan/Plans/:planId/Suites/:suiteId/TestCase?api-version={{api-version}}";

            $restClient = [RestClient]::new($urlTemplate, $this);
            $restClient.SetPlaceholder("planId", $testPlanId);
            $restClient.SetPlaceholder("suiteId", $testSuiteId);

            $payload = '[{
                "pointAssignments": [],
                "workItem": {
                    "id": ' + $workItem.Id + '
                }
            }]';

            return $restClient.Post($payload);
        }

        return $null;
    }

    # Get the TestPoint info from given TestCase.
    [PSObject] GetTestPointByTestCase([string] $testPlanId, [string] $testSuiteId, [int] $testCaseId) {
        $urlTemplate = "https://{{instance}}/{{organization}}/{{project-name}}/_apis/test/Plans/:planId/Suites/:suiteId/points?testCaseId=:testCaseId&api-version={{api-version}}";

        $restClient = [RestClient]::new($urlTemplate, $this);
        $restClient.SetPlaceholder("planId", $testPlanId);
        $restClient.SetPlaceholder("suiteId", $testSuiteId);
        $restClient.SetPlaceholder("testCaseId", $testCaseId);

        return $restClient.Get();
    }

    # this method will create a test run which acts as a container for all the test results.
    # return a Run Id.
    [PSObject] CreateTestRunInTestPlan([string] $testPlanId, [string] $testRunName, [System.Collections.ArrayList] $testPointIds, [psobject] $build, [datetime] $startDate) {
        $urlTemplate = "https://{{instance}}/{{organization}}/{{project-name}}/_apis/test/runs?api-version=7.1";
        $restClient = [RestClient]::new($urlTemplate, $this);

        $payload = @{
            automated      = $true
            name           = $testRunName
            build          = @{
                id  = $build.id
                url = $build.url
            }
            buildReference = @{
                id  = $build.id
                url = $build.url
            }
            plan           = @{
                id = $testPlanId
            }
            pointIds       = $testPointIds
            startDate      = $startDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.ffZ")            
        }

        $payload = $payload | ConvertTo-Json -Depth 10

        return $restClient.Post($payload);
    }

    [PSObject] GetPassedTestResultItem([int] $id, [psobject] $testPoint, [psobject] $testCase, [psobject] $build, [datetime] $started, [datetime] $completed, [string] $comment) {

        return @{
            id            = $id
            outcome       = [TestResultOutcome]::Passed.ToString()
            testPoint     = @{
                id  = $testPoint.id
                url = $testPoint.url
            }
            startedDate   = $started.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.ffZ")
            completedDate = $completed.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.ffZ")
            state         = [TestResultState]::Completed.ToString()
        };
    }

    [PSObject] GetFailedTestResultItemWithSubResults([int] $id, [psobject] $testPoint, [psobject] $testCase, [psobject] $build, [datetime] $started, [datetime] $completed, [string] $comment, [TestSubResult[]] $subResults) {

        $generalErrorMessage = ($subResults | ForEach-Object { "$($_.displayName): $($_.errorMessage)" }) -join "`n------`n"
        $generalStackTraceNode = ($subResults | ForEach-Object { "$($_.displayName):`n $($_.stackTrace)" }) -join "`n------`n"

        # limit to 1000 chars
        $generalStackTraceNode = $generalStackTraceNode.Substring(0, [math]::Min(1000, $generalStackTraceNode.Length))

        return @{
            id              = $id
            outcome         = [TestResultOutcome]::Failed.ToString()
            startedDate     = $started.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.ffZ")
            completedDate   = $completed.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.ffZ")
            state           = [TestResultState]::Completed.ToString()
            resultGroupType = [TestResultGroupType]::orderedTest.ToString()
            subResults      = $subResults
            errorMessage    = $generalErrorMessage
            stackTrace      = $generalStackTraceNode
            testPoint       = @{
                id  = $testPoint.id
                url = $testPoint.url
            }
        };
    }

    [PSObject] UpdateTestResults([string] $runId, [System.Collections.ArrayList] $results) {

        $urlTemplate = "https://{{instance}}/{{organization}}/{{project-name}}/_apis/test/Runs/:runId/results?api-version=7.1";

        $payload = $results | ConvertTo-Json -Depth 20

        $restClient = [RestClient]::new($urlTemplate, $this);
        $restClient.SetPlaceholder("runId", $runId);

        return $restClient.PatchButPost($payload);
    }

    [PSObject] UpdateTestRun([string] $runId, [psobject] $build) {
        $urlTemplate = "https://{{instance}}/{{organization}}/{{project-name}}/_apis/test/runs/:runId?api-version={{api-version}}";

        $payload = @{
            state = [TestResultState]::Completed.ToString()
        }

        $payload = $payload | ConvertTo-Json -Depth 10

        $restClient = [RestClient]::new($urlTemplate, $this);
        $restClient.SetPlaceholder("runId", $runId);

        return $restClient.PatchButPost($payload);
    }

    #endregion


    #region Build APIs

    # Get the latest build from Build definition
    [PSObject] GetLatestBuildInfo([string] $definitionId) {
        $urlTemplate = "https://{{instance}}/{{organization}}/{{project-name}}/_apis/build/latest/:definitionId?api-version=7.1-preview.1";

        $restClient = [RestClient]::new($urlTemplate, $this);
        $restClient.SetPlaceholder("definitionId", $definitionId);

        return $restClient.Get();
    }

    [PSObject] GetBuildInfo([string] $buildId) {
        $urlTemplate = "https://{{instance}}/{{organization}}/{{project-name}}/_apis/build/builds/:buildId?api-version={{api-version}}";

        $restClient = [RestClient]::new($urlTemplate, $this);
        $restClient.SetPlaceholder("buildId", $buildId);

        return $restClient.Get();
    }

    #endregion

}




