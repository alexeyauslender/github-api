# This functionality fetches pull requests from provided github repository. It will try to open local outlook instance and generate email as a pull requests report with open, closed and draft pull requests. If there is no local outlook installed it will output the report in your console. 

# Parameters
- To(required) - email address for the generated report
- Owner(optional) - the repo owner, the default is microsoft
- RepoName(optional) - the repo name, the default is vscode
- HistoryInDays(optional) - how long back in time you want to fetch PRs, the default is 7 days
- PAT(optional) - your personal access token, if the parameter is not provided the response might be throttled by GitHub API

# Sample run
```
.\Get-PullRequestsFromGithubApi.ps1 -To "alexey@ceridian.com" -PAT "ghp_XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXxx"

```