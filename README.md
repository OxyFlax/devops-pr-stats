# Azure DevOps Pull Request Stats

This Extension to Azure DevOps will give you a new Hub in your repositories section that is aimed at providing some statistical insights in to your Pull Request process.  This will show you things like the average time a PR is open.  It will show you what branches are getting pull requests created for.  It will show you who is approving the Pull Requests, by teams and which teams are contributing the most to reviewing them.

The original extension repository is available [here](https://github.com/jeffpriz/devops-pr-stats)

The changes of this fork are mainly to provide a more readable reporting on reviewers of a project with lots of person. It adds insights on the teams reviews.

## Debug in local
In order to debug in local and not requiring to deploy the app each time you do a change to test it, you can follow those instructions:
* Build the application for testing using `yarn build:test`
* Deploy this new extension and add it to your organization (if it's already deployed, you can publish an update with `yarn publish-test-extension --token [your token]`)
* Install "Debugger for Firefox" extension on Visual Studio Code
* Restart Visual Studio Code (you need to have the .vscode/launch.json like in the debug-in-local branch)
* Install additional dependencies (yarn install)
* Run this command in the terminal to launch the webapp: `yarn webpack serve --mode development`
* From Visual Studio Code, Start debugging (F5) and run the "Launch Firefox" task
* Inside the window which opened, navigate to this URL: https://localhost:44300
* Click on "Advanced..." and "Accept the Risk and Continue" => You should be able to see "Cannot GET /" on the page after that
* Go in your azure repository and open the extension => it will target your local extension. Any change you perform and file you save will automatically reload the extension

If this doesn't work, check [this link](https://docs.microsoft.com/en-us/azure/devops/extend/get-started/node?view=azure-devops) for basics and [this link](https://github.com/microsoft/azure-devops-extension-hot-reload-and-debug) for hot reload tutorial

## Images
![json text](images/repoHub.PNG)


## Source
[GitHub](https://github.com/jeffpriz/devops-pr-stats)

## Issues
[File an issue](https://github.com/jeffpriz/devops-pr-stats/issues)

## Credits
[Jeff Przylucki](http://www.oneluckidev.com)

Arnaud Pasquelin